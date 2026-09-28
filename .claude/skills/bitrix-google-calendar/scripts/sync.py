#!/usr/bin/env python3
"""Синхронизация задач Битрикс24 -> Google Календарь.

Каждая задача с крайним сроком (или плановыми датами) становится событием
в Google Календаре с напоминаниями. Повторный запуск обновляет события,
а не создаёт дубли: связь хранится в extendedProperties.private события.

Настройки берутся из переменных окружения (см. references/setup.md):
  BITRIX_WEBHOOK_URL           входящий вебхук, напр. https://x.bitrix24.ru/rest/1/abc123/
  GOOGLE_CALENDAR_ID           id календаря (почта или ...@group.calendar.google.com)
  GOOGLE_SERVICE_ACCOUNT_JSON  JSON ключа сервисного аккаунта ИЛИ путь к файлу
  BITRIX_USER_ID               (необяз.) брать задачи только этого участника
  REMINDER_MINUTES             (необяз.) напоминания, минуты до срока, по умолч. "1440,60"
  EVENT_DURATION_MINUTES       (необяз.) длина события, если есть только срок, по умолч. 30
  LOOKBACK_DAYS                (необяз.) брать задачи, изменённые за N дней, по умолч. 90
  COMPLETED_ACTION             (необяз.) что делать с завершёнными: mark | delete, по умолч. mark
  TIMEZONE                     (необяз.) часовой пояс событий, по умолч. Europe/Moscow

Запуск:  python3 sync.py [--dry-run]
"""

import argparse
import hashlib
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from urllib.parse import urlparse

import requests

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/calendar.events"]
CALENDAR_API = "https://www.googleapis.com/calendar/v3"
SOURCE_TAG = "bitrix24"

# Статусы задач Битрикс24
STATUS_NAMES = {
    1: "Новая",
    2: "Ждёт выполнения",
    3: "Выполняется",
    4: "Ждёт контроля",
    5: "Завершена",
    6: "Отложена",
    7: "Отклонена",
}
DONE_STATUSES = {5, 7}

# Цвета Google Календаря (colorId)
COLOR_OVERDUE = "11"  # красный
COLOR_DONE = "8"  # серый


def env(name, default=None, required=False):
    value = os.environ.get(name, default)
    if required and not value:
        sys.exit(f"Не задана переменная окружения {name} (см. references/setup.md)")
    return value


# ---------------------------------------------------------------- Битрикс24


class Bitrix:
    def __init__(self, webhook_url):
        self.base = webhook_url.rstrip("/") + "/"
        parsed = urlparse(self.base)
        self.portal = f"{parsed.scheme}://{parsed.netloc}"

    def call(self, method, params):
        resp = requests.post(self.base + method + ".json", json=params, timeout=30)
        data = resp.json()
        if "error" in data:
            raise RuntimeError(f"Битрикс24 {method}: {data.get('error_description') or data['error']}")
        return data

    def tasks(self, user_id, lookback_days):
        since = (datetime.now(timezone.utc) - timedelta(days=lookback_days)).strftime("%Y-%m-%dT%H:%M:%S+00:00")
        task_filter = {">=CHANGED_DATE": since}
        if user_id:
            task_filter["MEMBER"] = user_id
        params = {
            "filter": task_filter,
            "select": [
                "ID", "TITLE", "DESCRIPTION", "DEADLINE", "START_DATE_PLAN", "END_DATE_PLAN",
                "STATUS", "RESPONSIBLE_ID", "CREATED_BY", "GROUP_ID", "CHANGED_DATE",
            ],
            "order": {"ID": "asc"},
        }
        start = 0
        while True:
            data = self.call("tasks.task.list", {**params, "start": start})
            yield from data.get("result", {}).get("tasks", [])
            if "next" not in data:
                break
            start = data["next"]

    def task_url(self, task):
        user = task.get("responsibleId") or task.get("createdBy") or 0
        return f"{self.portal}/company/personal/user/{user}/tasks/task/view/{task['id']}/"


# ---------------------------------------------------------- Google Календарь


class Calendar:
    def __init__(self, calendar_id, credentials_json):
        from google.auth.transport.requests import AuthorizedSession
        from google.oauth2 import service_account

        if os.path.isfile(credentials_json):
            creds = service_account.Credentials.from_service_account_file(credentials_json, scopes=GOOGLE_SCOPES)
        else:
            info = json.loads(credentials_json)
            creds = service_account.Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
        self.session = AuthorizedSession(creds)
        self.url = f"{CALENDAR_API}/calendars/{requests.utils.quote(calendar_id, safe='')}/events"

    def _check(self, resp):
        if resp.status_code >= 400:
            raise RuntimeError(f"Google Calendar {resp.status_code}: {resp.text[:500]}")
        return resp.json() if resp.content else {}

    def synced_events(self):
        """Все события, созданные этим скриптом: {bitrix_task_id: event}."""
        events, page = {}, None
        while True:
            params = {"privateExtendedProperty": f"source={SOURCE_TAG}", "maxResults": 2500, "showDeleted": "false"}
            if page:
                params["pageToken"] = page
            data = self._check(self.session.get(self.url, params=params))
            for ev in data.get("items", []):
                task_id = ev.get("extendedProperties", {}).get("private", {}).get("bitrixTaskId")
                if task_id:
                    events[task_id] = ev
            page = data.get("nextPageToken")
            if not page:
                return events

    def insert(self, body):
        return self._check(self.session.post(self.url, json=body))

    def update(self, event_id, body):
        return self._check(self.session.put(f"{self.url}/{event_id}", json=body))

    def delete(self, event_id):
        resp = self.session.delete(f"{self.url}/{event_id}")
        if resp.status_code not in (200, 204, 404, 410):
            self._check(resp)


# ------------------------------------------------------------------ Логика


def parse_dt(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def build_event(task, bitrix, cfg):
    """Событие для задачи или None, если у задачи нет дат."""
    start, end = parse_dt(task.get("startDatePlan")), parse_dt(task.get("endDatePlan"))
    deadline = parse_dt(task.get("deadline"))
    if deadline:
        end = deadline
        if not start or start >= end:
            start = end - timedelta(minutes=cfg["duration"])
    elif not (start and end and start < end):
        return None

    status = int(task.get("status") or 0)
    done = status in DONE_STATUSES
    overdue = not done and deadline and deadline < datetime.now(timezone.utc)

    prefix = "✅ " if done else ("⚠️ " if overdue else "")
    description = "\n".join(filter(None, [
        f"Статус: {STATUS_NAMES.get(status, status)}",
        f"Крайний срок: {deadline.strftime('%d.%m.%Y %H:%M')}" if deadline else None,
        f"Задача в Битрикс24: {bitrix.task_url(task)}",
        "",
        (task.get("description") or "").strip()[:3000],
    ]))

    body = {
        "summary": f"{prefix}[Б24] {task.get('title', '').strip()}",
        "description": description,
        "start": {"dateTime": start.isoformat(), "timeZone": cfg["tz"]},
        "end": {"dateTime": end.isoformat(), "timeZone": cfg["tz"]},
        "source": {"title": "Битрикс24", "url": bitrix.task_url(task)},
        "reminders": {
            "useDefault": False,
            # у завершённых задач напоминания не нужны
            "overrides": [] if done else [{"method": "popup", "minutes": m} for m in cfg["reminders"]],
        },
        "extendedProperties": {"private": {"source": SOURCE_TAG, "bitrixTaskId": str(task["id"])}},
    }
    if done:
        body["colorId"] = COLOR_DONE
    elif overdue:
        body["colorId"] = COLOR_OVERDUE

    body["extendedProperties"]["private"]["hash"] = hashlib.sha1(
        json.dumps(body, sort_keys=True, ensure_ascii=False).encode()
    ).hexdigest()
    return body


def sync(dry_run=False):
    cfg = {
        "tz": env("TIMEZONE", "Europe/Moscow"),
        "duration": int(env("EVENT_DURATION_MINUTES", "30")),
        "reminders": [int(m) for m in env("REMINDER_MINUTES", "1440,60").split(",") if m.strip()],
        "completed_action": env("COMPLETED_ACTION", "mark"),
    }
    bitrix = Bitrix(env("BITRIX_WEBHOOK_URL", required=True))
    calendar = Calendar(env("GOOGLE_CALENDAR_ID", required=True), env("GOOGLE_SERVICE_ACCOUNT_JSON", required=True))

    existing = calendar.synced_events()
    stats = {"created": 0, "updated": 0, "deleted": 0, "unchanged": 0, "skipped": 0}

    for task in bitrix.tasks(env("BITRIX_USER_ID"), int(env("LOOKBACK_DAYS", "90"))):
        task_id = str(task["id"])
        event = existing.get(task_id)
        body = build_event(task, bitrix, cfg)
        done = int(task.get("status") or 0) in DONE_STATUSES

        if body is None or (done and cfg["completed_action"] == "delete"):
            if event:
                print(f"удалить   #{task_id} {task.get('title')}")
                if not dry_run:
                    calendar.delete(event["id"])
                stats["deleted"] += 1
            else:
                stats["skipped"] += 1
            continue

        if event is None:
            print(f"создать   #{task_id} {body['summary']}")
            if not dry_run:
                calendar.insert(body)
            stats["created"] += 1
        elif event.get("extendedProperties", {}).get("private", {}).get("hash") != body["extendedProperties"]["private"]["hash"]:
            print(f"обновить  #{task_id} {body['summary']}")
            if not dry_run:
                calendar.update(event["id"], body)
            stats["updated"] += 1
        else:
            stats["unchanged"] += 1

    print(("[пробный запуск] " if dry_run else "") + "Итого: " + ", ".join(f"{k}={v}" for k, v in stats.items()))
    return stats


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Синхронизация задач Битрикс24 в Google Календарь")
    parser.add_argument("--dry-run", action="store_true", help="только показать изменения, ничего не записывать")
    sync(dry_run=parser.parse_args().dry_run)
