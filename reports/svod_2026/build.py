"""Сверка РСВ ООО «Легкие деньки» за 1 квартал и полугодие 2026 г.: первичный и уточнённый расчёты по сотрудникам."""
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = "reports/svod_2026/Сверка_РСВ_2026.xlsx"

thin = Side(style="thin", color="999999")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD = PatternFill("solid", fgColor="DDE7F3")
TOTAL = PatternFill("solid", fgColor="F2F2F2")
BAD = PatternFill("solid", fgColor="F8D7DA")
OK = PatternFill("solid", fgColor="D4EDDA")
NUM = '#,##0.00;[Red]-#,##0.00;"–"'

COLS = ["Сотрудник", "Месяц",
        "Выплаты: {a}", "Выплаты: {b}", "Разница выплат ({b} − {a})",
        "Взносы: {a}", "Взносы: {b}", "Разница взносов ({b} − {a})",
        "Проверка: взносы {b} − 30% от выплат", "Комментарий"]
WIDTHS = [30, 11, 16, 16, 16, 16, 16, 16, 16, 60]

wb = Workbook()


def sheet(title, header_lines, a, b):
    ws = wb.create_sheet(title)
    r = 1
    for i, line in enumerate(header_lines):
        ws.cell(r, 1, line).font = Font(bold=(i == 0), size=13 if i == 0 else 10)
        r += 1
    r += 1
    for c, name in enumerate(COLS, 1):
        cell = ws.cell(r, c, name.format(a=a, b=b))
        cell.font = Font(bold=True)
        cell.fill = HEAD
        cell.border = BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[r].height = 62
    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(r + 1, 3)
    return ws, r + 1


def row(ws, r, name, month, pay_a, pay_b, vz_a, vz_b, note="", fill=None, bold=False):
    vals = [name, month, pay_a, pay_b, f"=D{r}-C{r}", vz_a, vz_b, f"=G{r}-F{r}",
            f"=G{r}-ROUND(D{r}*0.3,2)", note]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(r, c, v)
        cell.border = BORDER
        if 3 <= c <= 9:
            cell.number_format = NUM
        else:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            cell.fill = fill
        if bold:
            cell.font = Font(bold=True)
    for L in "EHI":
        ws.conditional_formatting.add(f"{L}{r}", CellIsRule(operator="notBetween", formula=["-0.005", "0.005"], fill=BAD))
        ws.conditional_formatting.add(f"{L}{r}", CellIsRule(operator="between", formula=["-0.005", "0.005"], fill=OK))


def person(ws, r, name, months, notes=None):
    """months: [(месяц, выплаты A, выплаты B, взносы A, взносы B)]; добавляет строку «Итого»."""
    notes = notes or {}
    start = r
    for m in months:
        row(ws, r, name if r == start else "", *m, note=notes.get(m[0], ""))
        r += 1
    row(ws, r, f"Итого {name.split()[0]}", "",
        *[f"=SUM({L}{start}:{L}{r - 1})" for L in "CDFG"], note=notes.get("итого", ""), fill=TOTAL, bold=True)
    return r + 1, r


def text(ws, r, title, lines):
    ws.cell(r, 1, title).font = Font(bold=True)
    for t in lines:
        r += 1
        ws.cell(r, 1, t)
    return r + 2


# ----------------------------------------------------------------- 1 квартал
ws, r = sheet("1 квартал", [
    "РСВ за 1 квартал 2026 г.: первичный и уточнённый — по сотрудникам",
    "ООО «Легкие деньки», ИНН 3663129244, тариф 01 (30%)",
    "Первичный: корр. 0 от 01.06.2026, КПП 366301001, ИФНС 3663.   Уточнённый: корр. 1 от 08.09.2026, КПП 366201001, ИФНС 3662.",
], "первичный", "уточнённый")

r, tot_l = person(ws, r, "Лубкова Наталья Ивановна", [
    ("Январь", 27093.00, 27093.00, 8127.90, 8127.90),
    ("Февраль", 27093.00, 27093.00, 8127.90, 8127.90),
    ("Март", 20101.26, 27093.00, 6030.38, 8127.90),
], {"Март": "Расхождение: в первичном март 20 101,26 ₽, в уточнённом — полный МРОТ 27 093,00 ₽."})
r, tot_b = person(ws, r, "Баскова Неонила Павловна", [
    ("Январь", 0, 0, 0, 0),
], {"Январь": "В обоих расчётах — код НР, выплат нет."})

row(ws, r, "ИТОГО по разделу 3", "1 кв.", *[f"={L}{tot_l}+{L}{tot_b}" for L in "CDFG"], fill=TOTAL, bold=True)
t3 = r
r += 1
row(ws, r, "Раздел 1 (подр. 1 стр. 030 / разд. 1 стр. 030)", "1 кв.", 74287.26, 81279.00, 22286.18, 24383.70,
    note="Сводные суммы раздела 1.", fill=TOTAL)
s1 = r
r += 1
for i, (m, a, b, va, vb) in enumerate([("Январь", 27093.00, 27093.00, 8127.90, 8127.90),
                                       ("Февраль", 27093.00, 27093.00, 8127.90, 8127.90),
                                       ("Март", 20101.26, 27093.00, 6030.38, 8127.90)]):
    row(ws, r, "  в т.ч. по месяцам (стр. 031–033)" if i == 0 else "", m, a, b, va, vb)
    r += 1
row(ws, r, "Контроль: раздел 1 − раздел 3", "", *[f"={L}{s1}-{L}{t3}" for L in "CDFG"],
    note="0 — внутри каждого расчёта раздел 1 сходится с разделом 3.", bold=True)
r += 2
text(ws, r, "Вывод по 1 кварталу:", [
    "1. Расхождение только по Лубковой Н.И. и только за март: выплаты +6 991,74 ₽, взносы +2 097,52 ₽ в уточнённом расчёте.",
    "2. Январь и февраль, а также Баскова Н.П. в обоих расчётах совпадают.",
    "3. В обоих расчётах взносы = 30% от выплат (арифметика верная).",
    "4. Доплата взносов за март 2 097,52 ₽ — срок уплаты был 28.04.2026, на эту сумму начисляются пени.",
    "5. Уточнённый расчёт подан с КПП 366201001 в ИФНС 3662, первичный — с КПП 366301001 в ИФНС 3663: проверьте, что так и должно быть.",
])

# ----------------------------------------------------------------- полугодие
ws, r = sheet("Полугодие", [
    "РСВ за полугодие 2026 г.: поданный расчёт и данные с учётом уточнения за 1 квартал — по сотрудникам",
    "ООО «Легкие деньки», ИНН 3663129244, тариф 01 (30%)",
    "РСВ за полугодие: корр. 0 от 24.07.2026. Уточнённого РСВ за полугодие нет — колонка «должно быть» показывает, каким он должен быть.",
    "Разд. 3 РСВ за полугодие содержит только апрель–июнь; 1 кв. внутри него (разд. 1: 169 770,03 − 95 482,77 = 74 287,26 ₽) "
    "равен первичному РСВ за 1 кв., поэтому помесячно январь–март взяты из первичного РСВ за 1 кв.",
], "РСВ п/г (подан)", "должно быть")

r, h_l = person(ws, r, "Лубкова Наталья Ивановна", [
    ("Январь", 27093.00, 27093.00, 8127.90, 8127.90),
    ("Февраль", 27093.00, 27093.00, 8127.90, 8127.90),
    ("Март", 20101.26, 27093.00, 6030.38, 8127.90),
], {"Март": "В РСВ за полугодие остались данные первичного РСВ за 1 кв.; уточнение за 1 кв. (08.09.2026) подано позже."})
r, h_t = person(ws, r, "Троянова Екатерина Викторовна", [
    ("Апрель", 19868.20, 19868.20, 5960.46, 5960.46),
    ("Май", 27093.00, 27093.00, 8127.90, 8127.90),
    ("Июнь", 27093.00, 27093.00, 8127.90, 8127.90),
], {"Апрель": "Меньше МРОТ — проверить неполное время / неполный месяц."})
r, h_d = person(ws, r, "Денисова Александра Валерьевна", [
    ("Июнь", 21428.57, 21428.57, 6428.57, 6428.57),
], {"Июнь": "Принята в июне (30 000 × 15/21 раб. дн.)."})
r, h_b = person(ws, r, "Баскова Неонила Павловна", [
    ("Январь", 0, 0, 0, 0),
], {"Январь": "Выплат нет."})

row(ws, r, "ИТОГО по сотрудникам", "п/г", *[f"={L}{h_l}+{L}{h_t}+{L}{h_d}+{L}{h_b}" for L in "CDFG"], fill=TOTAL, bold=True)
tot = r
r += 1
row(ws, r, "Раздел 1 РСВ за полугодие (стр. 030)", "п/г", 169770.03, 176761.77, 50931.01, 53028.53,
    note="«Должно быть» = уточнённый 1 кв. (81 279,00 / 24 383,70) + 2 кв. (95 482,77 / 28 644,83).", fill=TOTAL)
s1 = r
r += 1
row(ws, r, "Контроль: раздел 1 − сумма по сотрудникам", "", *[f"={L}{s1}-{L}{tot}" for L in "CDFG"], bold=True)
r += 2

ws.cell(r, 1, "Свод по кварталам").font = Font(bold=True, size=11)
r += 1
for c, name in enumerate(COLS, 1):
    cell = ws.cell(r, c, name.format(a="РСВ п/г (подан)", b="должно быть"))
    cell.font = Font(bold=True)
    cell.fill = HEAD
    cell.border = BORDER
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
ws.row_dimensions[r].height = 62
r += 1
row(ws, r, "1 квартал", "янв–мар", 74287.26, 81279.00, 22286.18, 24383.70, note="Расхождение — Лубкова, март.")
q1r = r
r += 1
row(ws, r, "2 квартал", "апр–июн", 95482.77, 95482.77, 28644.83, 28644.83, note="Совпадает.")
r += 1
row(ws, r, "Полугодие", "", f"=C{q1r}+C{r - 1}", f"=D{q1r}+D{r - 1}", f"=F{q1r}+F{r - 1}", f"=G{q1r}+G{r - 1}",
    fill=TOTAL, bold=True)
r += 2
text(ws, r, "Вывод по полугодию:", [
    "1. РСВ за полугодие сдан 24.07.2026 на данных первичного РСВ за 1 кв.; после уточнения 1 кв. (08.09.2026) он с ним не сходится.",
    "2. Расхождение: выплаты 6 991,74 ₽, взносы 2 097,52 ₽ — Лубкова Н.И., март. Апрель–июнь (Троянова, Денисова) сходятся.",
    "3. Нужен уточнённый РСВ за полугодие: стр. 030 разд. 1 — 53 028,53 ₽ вместо 50 931,01 ₽; "
    "подр. 1 стр. 030/050 — 176 761,77 ₽ вместо 169 770,03 ₽. Строки 031–033 и разд. 3 (апрель–июнь) не меняются.",
    "4. Иначе ФНС увидит несоответствие нарастающего итога (1 кв. в РСВ за полугодие ≠ уточнённому РСВ за 1 кв.) и пришлёт требование.",
])

del wb["Sheet"]
for w in wb.worksheets:
    w.sheet_view.showGridLines = False
    w.page_setup.orientation = "landscape"
    w.page_setup.fitToWidth = 1
wb.save(OUT)
print("saved", OUT)
