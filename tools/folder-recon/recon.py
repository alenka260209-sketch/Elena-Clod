#!/usr/bin/env python3
"""Шаг 1 — Разведка папки (только чтение, ничего не перемещает и не удаляет).

Запуск:
    python recon.py "C:\\Users\\Имя\\Downloads\\Accounting"
    python recon.py ~/Downloads/Accounting

Результат кладётся рядом со скриптом:
    recon_report.md   — сводка: типы файлов, даты, темы, юрлица, важное
    recon_files.csv   — полный список файлов с категорией и найденным юрлицом
"""
import csv
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

TYPES = {
    "Документы": {".doc", ".docx", ".odt", ".rtf", ".txt", ".pdf"},
    "Таблицы": {".xls", ".xlsx", ".xlsm", ".ods", ".csv"},
    "Презентации": {".ppt", ".pptx", ".odp", ".key"},
    "Картинки/сканы": {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".heic", ".webp"},
    "Архивы": {".zip", ".rar", ".7z", ".tar", ".gz"},
    "1С/обмен/отчётность": {".xml", ".dt", ".cf", ".1cd", ".sig", ".p7s", ".txt1c"},
    "Программы/установщики": {".exe", ".msi", ".dmg", ".apk"},
}

# Тема -> ключевые слова (ищутся в имени файла и тексте docx/xlsx), в нижнем регистре
TOPICS = {
    "Договоры и допсоглашения": ["договор", "дог.", "допсоглаш", "доп.согл", "контракт", "соглашени", "contract"],
    "Счета": ["счет", "счёт", "invoice", "счет-фактур", "сф "],
    "Акты, УПД, накладные": ["акт", "упд", "накладн", "торг-12", "торг12", "кс-2", "кс-3", "кс2", "кс3"],
    "Акты сверки": ["сверк"],
    "Банк и платежи": ["выписк", "платеж", "платёж", "п/п", "пп ", "банк", "statement", "сбер", "тинькофф", "т-банк", "альфа", "втб", "точка"],
    "Налоги и отчётность": ["ндс", "ндфл", "усн", "декларац", "фнс", "енс", "рсв", "6-ндфл", "налог", "требован", "баланс", "отчетност", "отчётност", "квитанц", "извещен"],
    "Зарплата и кадры": ["зарплат", "з/п", "зп ", "ведомост", "табель", "приказ", "трудов", "отпуск", "сотрудник", "штатн", "больничн"],
    "Учредительные / реквизиты": ["устав", "егрюл", "егрип", "огрн", "реквизит", "карточка", "свидетельств", "выписка из", "решение", "протокол", "доверенност"],
    "Сметы, КП, прайсы": ["смет", "кп ", "коммерческ", "прайс", "расчет", "расчёт", "спецификац"],
    "Претензии и суды": ["претенз", "иск", "суд", "арбитраж", "исполнит"],
}

TOPIC_RE = {t: re.compile(r"(?<![а-яёa-z0-9])(?:" + "|".join(re.escape(k.strip()) for k in kws) + ")")
            for t, kws in TOPICS.items()}

IMPORTANT_TOPICS = {"Договоры и допсоглашения", "Налоги и отчётность", "Банк и платежи",
                    "Учредительные / реквизиты", "Претензии и суды", "Акты сверки"}

ORG_RE = re.compile(
    r"\b(ООО|АО|ПАО|ЗАО|ОАО|НКО|АНО)\s*[«\"“']?\s*([A-Za-zА-Яа-яЁё0-9][^«»\"“”'\n\r_]{1,40}?)\s*[»\"”']",
)
IP_RE = re.compile(r"\bИП\s+([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ]\.?\s*[А-ЯЁ]?\.?|\s+[А-ЯЁ][а-яё]+){0,2})")
INN_RE = re.compile(r"ИНН[\s:№]*(\d{10}|\d{12})\b")


def file_type(ext):
    for name, exts in TYPES.items():
        if ext in exts:
            return name
    return "Прочее" if ext else "Без расширения"


def office_text(path, limit=200_000):
    """Текст из docx/xlsx/pptx через zip (без сторонних библиотек)."""
    try:
        with zipfile.ZipFile(path) as z:
            parts = []
            for n in z.namelist():
                if n.startswith(("word/document", "xl/sharedStrings", "ppt/slides/slide")) and n.endswith(".xml"):
                    parts.append(z.read(n)[:limit].decode("utf-8", "ignore"))
            return re.sub(r"<[^>]+>", " ", " ".join(parts))
    except Exception:
        return ""


def pdf_text(path):
    try:
        from pypdf import PdfReader  # необязательно: pip install pypdf
    except ImportError:
        return ""
    try:
        r = PdfReader(str(path))
        return " ".join((p.extract_text() or "") for p in r.pages[:3])
    except Exception:
        return ""


def find_entities(text):
    ents = set()
    for kind, name in ORG_RE.findall(text):
        ents.add(f"{kind} «{name.strip()}»")
    for name in IP_RE.findall(text):
        ents.add(f"ИП {name.strip()}")
    inns = set(INN_RE.findall(text))
    return ents, inns


def human(n):
    for u in ("Б", "КБ", "МБ", "ГБ"):
        if n < 1024:
            return f"{n:.0f} {u}"
        n /= 1024
    return f"{n:.1f} ТБ"


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    root = Path(sys.argv[1]).expanduser()
    if not root.is_dir():
        sys.exit(f"Папка не найдена: {root}")
    out_dir = Path(__file__).resolve().parent

    rows = []
    for dirpath, _, files in os.walk(root):
        for f in files:
            p = Path(dirpath) / f
            try:
                st = p.stat()
            except OSError:
                continue
            ext = p.suffix.lower()
            rel = str(p.relative_to(root))
            text = rel
            if ext in {".docx", ".xlsx", ".xlsm", ".pptx"}:
                text += " " + office_text(p)
            elif ext == ".pdf":
                text += " " + pdf_text(p)
            low = text.lower()
            topics = [t for t, rx in TOPIC_RE.items() if rx.search(low)]
            ents, inns = find_entities(text)
            rows.append({
                "path": rel,
                "ext": ext or "-",
                "type": file_type(ext),
                "size": st.st_size,
                "modified": datetime.fromtimestamp(st.st_mtime),
                "topics": topics,
                "entities": sorted(ents),
                "inns": sorted(inns),
            })

    if not rows:
        sys.exit("Папка пуста.")

    by_type = Counter(r["type"] for r in rows)
    by_ext = Counter(r["ext"] for r in rows)
    by_topic = Counter(t for r in rows for t in r["topics"])
    by_year = Counter(r["modified"].year for r in rows)
    ent_files = defaultdict(list)
    for r in rows:
        for e in r["entities"] or ["— юрлицо не определено —"]:
            ent_files[e].append(r)
    inn_counter = Counter(i for r in rows for i in r["inns"])
    rows_sorted = sorted(rows, key=lambda r: r["modified"])
    important = [r for r in rows if IMPORTANT_TOPICS.intersection(r["topics"])]
    total = sum(r["size"] for r in rows)

    L = [f"# Разведка папки: {root}", "",
         f"Сформировано: {datetime.now():%d.%m.%Y %H:%M}. Скрипт только читал файлы, ничего не перемещал.", "",
         f"**Всего файлов:** {len(rows)}, объём {human(total)}; "
         f"подпапок: {len({str(Path(r['path']).parent) for r in rows})}", "",
         "## 1. Типы файлов", "", "| Тип | Файлов |", "|---|---|"]
    L += [f"| {t} | {n} |" for t, n in by_type.most_common()]
    L += ["", "Расширения: " + ", ".join(f"{e} — {n}" for e, n in by_ext.most_common()), ""]

    L += ["## 2. Темы и категории (по имени и содержимому)", "", "| Тема | Файлов |", "|---|---|"]
    L += [f"| {t} | {n} |" for t, n in by_topic.most_common()]
    L += [f"| Тема не определена | {sum(1 for r in rows if not r['topics'])} |", ""]

    L += ["## 3. Когда всё появилось", "", "По годам (дата изменения): " +
          ", ".join(f"{y} — {n}" for y, n in sorted(by_year.items())), "", "**Самые старые:**", ""]
    L += [f"- {r['modified']:%d.%m.%Y} — {r['path']}" for r in rows_sorted[:10]]
    L += ["", "**Самые свежие:**", ""]
    L += [f"- {r['modified']:%d.%m.%Y} — {r['path']}" for r in reversed(rows_sorted[-10:])]

    L += ["", "## 4. Юрлица (найдены в именах и тексте документов)", "",
          "| Юрлицо / ИП | Файлов |", "|---|---|"]
    L += [f"| {e} | {len(fs)} |" for e, fs in sorted(ent_files.items(), key=lambda x: -len(x[1]))]
    if inn_counter:
        L += ["", "Найденные ИНН: " + ", ".join(f"{i} ({n})" for i, n in inn_counter.most_common(30))]
    L += ["", "> Одно и то же юрлицо может встретиться в разном написании — это сведём на шаге 2.", ""]

    L += [f"## 5. Похоже на важное ({len(important)} файлов)", ""]
    for r in sorted(important, key=lambda r: r["modified"], reverse=True)[:60]:
        L.append(f"- **{', '.join(t for t in r['topics'] if t in IMPORTANT_TOPICS)}** — {r['path']} "
                 f"({r['modified']:%d.%m.%Y}){' — ' + '; '.join(r['entities']) if r['entities'] else ''}")
    if len(important) > 60:
        L.append(f"- … и ещё {len(important) - 60}, полный список в recon_files.csv")

    (out_dir / "recon_report.md").write_text("\n".join(L) + "\n", encoding="utf-8")
    with open(out_dir / "recon_files.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["Путь", "Тип", "Расширение", "Размер", "Изменён", "Темы", "Юрлица", "ИНН"])
        for r in rows_sorted:
            w.writerow([r["path"], r["type"], r["ext"], r["size"], f"{r['modified']:%d.%m.%Y}",
                        ", ".join(r["topics"]), ", ".join(r["entities"]), ", ".join(r["inns"])])
    print(f"Готово: {out_dir / 'recon_report.md'} и {out_dir / 'recon_files.csv'}")


if __name__ == "__main__":
    main()
