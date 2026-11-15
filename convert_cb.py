# -*- coding: utf-8 -*-
"""
Конвертер выписок банка в аккуратный .xlsx с умной таблицей.

Принимает файлы .xls (HTML-выгрузка или настоящий Excel) и .xlsx.
Создаёт файл  <имя>_оформлен.xlsx  рядом с исходным.
"""
import os
import re
import sys
from datetime import datetime

import openpyxl
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------- оформление ----------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial", size=10)
BOLD_FONT = Font(name="Arial", size=10, bold=True)
TITLE_FONT = Font(name="Arial", size=12, bold=True, color="1F4E78")
INFO_FONT = Font(name="Arial", size=10, color="404040")
TOTAL_BORDER = Border(top=Side(style="thin", color="1F4E78"))

MONEY_FORMAT = "#,##0.00"
DATE_FORMAT = "dd.mm.yyyy hh:mm"
DATE_FORMATS = ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y")
MONEY_PREFIXES = ("Оборот", "Сумма", "Остаток")


# ---------- утилиты ----------
def clean(value):
    """Приводит значение к строке без лишних пробелов и неразрывных пробелов."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def is_empty(value):
    return clean(value) == ""


def to_money(value):
    if is_empty(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(clean(value).replace(" ", "").replace(",", "."))
    except ValueError:
        return clean(value)


def to_date(value):
    if isinstance(value, datetime):
        return value
    text = clean(value)
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return text


def to_text(value):
    if is_empty(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    return clean(value)


def column_kind(header):
    if header == "Дата":
        return "date"
    if header.startswith(MONEY_PREFIXES):
        return "money"
    return "text"


# ---------- чтение ----------
def read_rows(path):
    """Возвращает все строки файла как списки значений."""
    with open(path, "rb") as f:
        raw = f.read()

    # Современный .xlsx (zip) — определяем по содержимому, а не по имени файла
    if raw[:2] == b"PK":
        import io
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        ws = wb.worksheets[0]
        return [list(row) for row in ws.iter_rows(values_only=True)]

    # Настоящий двоичный Excel (.xls) — читаем через xlrd
    if raw[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        import xlrd
        book = xlrd.open_workbook(file_contents=raw)
        sh = book.sheet_by_index(0)
        rows = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK,
                                    xlrd.XL_CELL_ERROR):
                    v = None
                row.append(v)
            rows.append(row)
        return rows

    # HTML-выгрузка, сохранённая как .xls (самый частый случай)
    text = None
    for enc in ("utf-8", "cp1251"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("cp1251", errors="replace")

    soup = BeautifulSoup(text, "html.parser")
    # Удаляем скрытые элементы: их текст не виден в Excel, но попадает в get_text()
    for tag in soup.find_all(style=True):
        style = tag.get("style", "").replace(" ", "").lower()
        if "display:none" in style or "mso-hide:all" in style:
            tag.decompose()
    rows = []
    for tr in soup.find_all("tr"):
        row = []
        for cell in tr.find_all(["td", "th"]):
            row.append(cell.get_text(" ", strip=True))
            span = int(cell.get("colspan", 1) or 1)
            row.extend([""] * (span - 1))
        rows.append(row)
    return rows


# ---------- разбор структуры отчёта ----------
KNOWN_WORDS = ("Счет", "ИНН", "Наименование", "№ док-та", "Оп", "МФО",
               "Оборот Дебет", "Оборот Кредит", "Назначение платежа")


def normalize_header(name):
    """
    Если в заголовке склеено несколько названий колонок
    («Счет ИНН Наименование ...»), оставляем первое: «Счет».
    """
    hits = sum(1 for w in KNOWN_WORDS if w in name)
    if hits >= 2:
        return name.split()[0]
    return name


def unique_names(names):
    """Делает имена колонок уникальными: ИНН, ИНН (2) и т. д."""
    seen, result = {}, []
    for i, name in enumerate(names, 1):
        name = name or f"Колонка {i}"
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 1
        result.append(name)
    return result


def fix_account_values(header, rows):
    """В колонке «Счет» оставляем только номер счёта (первое число в ячейке)."""
    if "Счет" not in header:
        return rows
    i = header.index("Счет")
    fixed = []
    for row in rows:
        row = list(row)
        value = clean(row[i])
        m = re.match(r"(\d{10,})", value)
        if m:
            row[i] = m.group(1)
        fixed.append(row)
    return fixed


def split_account_column(header, rows):
    """
    В HTML-выгрузке колонка «Cчет/ИНН» содержит: «счёт/ИНН/наименование».
    Разбиваем её на три отдельные колонки: Счет, ИНН, Наименование.
    """
    for i, name in enumerate(header):
        if re.fullmatch(r"[CС]чет\s*/\s*ИНН", name):
            new_header = header[:i] + ["Счет", "ИНН", "Наименование"] + header[i + 1:]
            new_rows = []
            for row in rows:
                parts = [p.strip() for p in clean(row[i]).split("/", 2)] if not is_empty(row[i]) else []
                parts += [""] * (3 - len(parts))
                new_rows.append(row[:i] + parts[:3] + row[i + 1:])
            return new_header, new_rows
    return header, rows


def split_report(rows):
    """Делит отчёт на: строки шапки (meta), заголовки колонок и данные."""
    header_idx = None
    for i, row in enumerate(rows):
        if row and clean(row[0]) == "Дата":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("не найдена строка заголовков (ячейка «Дата»)")

    # Шапка: непустые строки над таблицей
    meta = []
    for row in rows[:header_idx]:
        line = "   ".join(clean(v) for v in row if not is_empty(v))
        if line:
            meta.append(line)

    # Заголовки: убираем пустые хвостовые колонки
    raw_header = [normalize_header(clean(h)) for h in rows[header_idx]]
    width = max(i for i, h in enumerate(raw_header, 1) if h) if any(raw_header) else 0
    header = raw_header[:width]

    # Данные: строки без содержимого пропускаем
    data, footer = [], []
    for row in rows[header_idx + 1:]:
        values = list(row[:width]) + [None] * (width - len(row))
        if all(is_empty(v) for v in values):
            continue
        if clean(values[0]).lower().startswith("итог"):
            footer.append(values)  # итоговые строки банка выносим из таблицы
            continue
        data.append(values)

    if not data:
        raise ValueError("в отчёте нет строк с данными")
    return meta, header, data, footer


# ---------- запись ----------
def build(out_path, meta, header, data, footer):
    kinds = [column_kind(h) for h in header]
    ncol = len(header)
    last_col = get_column_letter(ncol)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Выписка"
    ws.sheet_view.showGridLines = False

    # Шапка отчёта
    r = 1
    for i, line in enumerate(meta):
        if ncol > 1:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        cell = ws.cell(row=r, column=1, value=line)
        cell.font = TITLE_FONT if (i == 0 or "Сведения" in line) else INFO_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
        r += 1
    r += 1  # пустая строка перед таблицей

    # Заголовки
    header_row = r
    for c, name in enumerate(header, 1):
        cell = ws.cell(row=header_row, column=c, value=name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 30

    # Данные
    first = header_row + 1
    for r_idx, values in enumerate(data, start=first):
        for c_idx, (value, kind) in enumerate(zip(values, kinds), start=1):
            if kind == "money":
                value = to_money(value)
            elif kind == "date":
                value = to_date(value)
            else:
                value = to_text(value)

            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.font = BODY_FONT
            if kind == "money":
                cell.number_format = MONEY_FORMAT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif kind == "date":
                cell.number_format = DATE_FORMAT
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
    last = first + len(data) - 1

    # Умная таблица (фильтры, полосы, автоматическое расширение)
    table = Table(displayName="Vypiska", ref=f"A{header_row}:{last_col}{last}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(table)

    # Строка «Итого» под таблицей. SUBTOTAL считает только видимые строки,
    # поэтому итог меняется при фильтрации.
    total_row = last + 1
    ws.cell(row=total_row, column=1, value="Итого по видимым строкам").font = BOLD_FONT
    for c, kind in enumerate(kinds, 1):
        if kind == "money":
            col = get_column_letter(c)
            cell = ws.cell(row=total_row, column=c, value=f"=SUBTOTAL(9,{col}{first}:{col}{last})")
            cell.font = BOLD_FONT
            cell.number_format = MONEY_FORMAT
            cell.alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, ncol + 1):
        ws.cell(row=total_row, column=c).border = TOTAL_BORDER

    # Итоговые строки самого банка (как есть, с суммами)
    fr = total_row + 1
    for values in footer:
        ws.cell(row=fr, column=1, value=clean(values[0])).font = BOLD_FONT
        for c, kind in enumerate(kinds, 1):
            if kind == "money" and c - 1 < len(values):
                cell = ws.cell(row=fr, column=c, value=to_money(values[c - 1]))
                cell.font = BOLD_FONT
                cell.number_format = MONEY_FORMAT
                cell.alignment = Alignment(horizontal="right", vertical="center")
        fr += 1

    # Ширина колонок по содержимому (с ограничением)
    for c, kind in enumerate(kinds, 1):
        if kind == "date":
            width = 18
        elif kind == "money":
            width = 18
        else:
            longest = max(
                [len(str(ws.cell(row=rr, column=c).value or "")) for rr in range(first, last + 1)]
                + [len(header[c - 1])]
            )
            width = min(max(12, longest + 2), 60)
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[header_row].height = 30

    # Длинный текст (назначение платежа и т. п.) переносим по словам
    for c, kind in enumerate(kinds, 1):
        letter = get_column_letter(c)
        if kind == "text" and ws.column_dimensions[letter].width >= 60:
            for rr in range(first, last + 1):
                ws.cell(row=rr, column=c).alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

    # Удобства просмотра и печати
    ws.freeze_panes = f"A{first}"
    ws.print_title_rows = f"{header_row}:{header_row}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    wb.save(out_path)


# ---------- запуск ----------
def convert(path):
    rows = read_rows(path)
    meta, header, data, footer = split_report(rows)
    header, combined = split_account_column(header, data + footer)
    combined = fix_account_values(header, combined)
    data, footer = combined[:len(data)], combined[len(data):]
    header = unique_names(header)
    base, _ = os.path.splitext(path)
    out = base + "_оформлен.xlsx"
    build(out, meta, header, data, footer)
    return out, len(data)


def main():
    files = sys.argv[1:]
    if not files:
        print("Перетащите файл .xls или .xlsx на convert.bat")
        input("Нажмите Enter для выхода...")
        return
    for path in files:
        try:
            out, count = convert(path)
            print(f"Готово: {out}  (строк: {count})")
        except Exception as exc:
            print(f"Ошибка с файлом {path}: {exc}")


if __name__ == "__main__":
    main()