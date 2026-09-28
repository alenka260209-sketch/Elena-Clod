# Настройка синхронизации Битрикс24 → Google Календарь

Займёт ~15 минут. Нужно получить три значения:
`BITRIX_WEBHOOK_URL`, `GOOGLE_SERVICE_ACCOUNT_JSON`, `GOOGLE_CALENDAR_ID`.

## Шаг 1. Вебхук в Битрикс24

1. Битрикс24 → **Разработчикам** (или «Приложения» → «Разработчикам») →
   **Другое** → **Входящий вебхук**.
2. В «Настройке прав» отметьте **Задачи (task)**. Больше ничего не нужно.
3. Сохраните и скопируйте «Вебхук для вызова rest api», например:
   `https://mycompany.bitrix24.ru/rest/1/abc123xyz/`
   — это `BITRIX_WEBHOOK_URL`.

Вебхук работает от имени пользователя, который его создал, и видит те же
задачи, что и он. Чтобы брать только задачи конкретного сотрудника, задайте
`BITRIX_USER_ID` (его id видно в адресе профиля: `/company/personal/user/15/`).

## Шаг 2. Сервисный аккаунт Google

1. Откройте https://console.cloud.google.com/ и создайте проект (или выберите
   существующий).
2. **APIs & Services → Library** → найдите **Google Calendar API** → **Enable**.
3. **APIs & Services → Credentials → Create credentials → Service account**.
   Имя любое, роли можно не назначать.
4. Откройте созданный аккаунт → вкладка **Keys → Add key → Create new key → JSON**.
   Скачается файл `….json` — его содержимое и есть `GOOGLE_SERVICE_ACCOUNT_JSON`.
5. Запомните e-mail сервисного аккаунта (вида
   `bitrix-sync@my-project.iam.gserviceaccount.com`).

## Шаг 3. Доступ к календарю

1. Google Календарь → слева наведите на нужный календарь → ⋮ →
   **Настройки и общий доступ**. Удобно завести отдельный календарь
   «Задачи Битрикс24», чтобы не смешивать с личными событиями.
2. **Открыть доступ пользователям → Добавить** → e-mail сервисного аккаунта →
   права **«Внесение изменений в мероприятия»**.
3. Ниже, в разделе **Интеграция календаря**, скопируйте **Идентификатор
   календаря** — это `GOOGLE_CALENDAR_ID` (для основного календаря это ваша
   почта).

Уведомления придут на все устройства, где вы вошли в Google Календарь.
Проверьте, что в настройках этого календаря уведомления не отключены.

## Шаг 4. Проверка

```bash
export BITRIX_WEBHOOK_URL='https://mycompany.bitrix24.ru/rest/1/abc123xyz/'
export GOOGLE_CALENDAR_ID='...@group.calendar.google.com'
export GOOGLE_SERVICE_ACCOUNT_JSON="$HOME/keys/bitrix-sync.json"   # путь или содержимое
pip install -r .claude/skills/bitrix-google-calendar/scripts/requirements.txt
python3 .claude/skills/bitrix-google-calendar/scripts/sync.py --dry-run
```

Если список задач выглядит правильно — запустите без `--dry-run`.

## Автозапуск (GitHub Actions)

Workflow `.github/workflows/bitrix-calendar-sync.yml` запускает синхронизацию
каждые 15 минут и вручную (кнопка **Run workflow** на вкладке Actions).

В репозитории GitHub: **Settings → Secrets and variables → Actions**.

**Secrets** (обязательно):
- `BITRIX_WEBHOOK_URL`
- `GOOGLE_CALENDAR_ID`
- `GOOGLE_SERVICE_ACCOUNT_JSON` — вставьте целиком содержимое JSON-файла

**Variables** (по желанию): `BITRIX_USER_ID`, `REMINDER_MINUTES`,
`EVENT_DURATION_MINUTES`, `LOOKBACK_DAYS`, `COMPLETED_ACTION`, `TIMEZONE`.

Важно: GitHub может задерживать запуски по расписанию на 5–15 минут и
отключает расписание в репозиториях без активности 60 дней — тогда включите
workflow заново на вкладке Actions. Для надёжности можно вместо этого
поставить cron на своём сервере:

```cron
*/15 * * * * cd /opt/elena-clod && . ./.env && python3 .claude/skills/bitrix-google-calendar/scripts/sync.py >> /var/log/bitrix-sync.log 2>&1
```
