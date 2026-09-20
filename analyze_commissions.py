"""
Скрипт разбора банковских выписок по эквайрингу.

Поддерживает ДВА разных формата выписки (автоматически определяется):

1. SMARTVISTA (Uzcard/Visa/MC через терминал, тексты содержат "SmartVista"):
   - Покупка: Оп=4, сумма в графе "Оборот Кредит" (G), текст "Покупка товаров".
   - Комиссия: Оп=1, сумма в графе "Оборот Дебет" (F), текст "0,2% выручки от
     терминалов" или "Услуги банка". Бывает 1 строка (единая ставка) или
     2 строки на одну покупку (сплит на "обычную" 1% и "клиента" 1,5-2%,
     определяется по формуле от базовой суммы товара).
   - Возврат: текст "Возврат покупки товаров", связывается с исходной
     покупкой по RRN, обнуляет её комиссию.
   - Поступление на р/с: строки, где графа B ("Cчет/ИНН") содержит номер
     расчётного счёта САМОЙ компании (определяется автоматически - это
     единственный счёт, где вместо служебной банковской формулировки
     указано название компании в кавычках).

2. HUMO (национальная платёжная система, тексты содержат "HUMO"):
   - Покупка: Оп=4, сумма в графе "Оборот Кредит" (G), текст "Оплата на сумму"
     (но не "Отмена оплата").
   - Обычная комиссия (фикс. ставка): текст "Комиссия банка за зачисления".
   - Комиссия клиента (перем. ставка, межбанк/международные карты):
     текст "Расчет по Комиссия".
   - Возврат: текст "Отмена оплата", связывается с исходной покупкой по RRN.
   - Поступление на р/с: текст "Зачисление денежных средств" (уже сумма
     нетто, за вычетом комиссии).

Для НОВОГО, ещё не встречавшегося формата - скрипт вернёт tип 'unknown'
и потребуется добавить для него отдельную ветку разбора (см. функции
analyze_smartvista / analyze_humo как образец).
"""

import re
import sys
from datetime import datetime, date

import openpyxl


TS_RE = re.compile(r'(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})')
DATE_RE = re.compile(r'(\d{2})\.(\d{2})\.(\d{4})')
RRN_RE = re.compile(r'RRN[:\-\s]+(\d+)')
SETTLEMENT_ACCOUNT_RE = re.compile(r'^([\d]+)/[^/]*/[^"]*"[^"]+"')


def extract_timestamp(text):
    """Достаёт дату+время из текста назначения платежа, если она там есть."""
    if not isinstance(text, str):
        return None
    m = TS_RE.search(text)
    if m:
        d, mo, y, H, Mi, S = m.groups()
        try:
            return datetime(int(y), int(mo), int(d), int(H), int(Mi), int(S))
        except ValueError:
            return None
    m = DATE_RE.search(text)
    if m:
        d, mo, y = m.groups()
        try:
            return datetime(int(y), int(mo), int(d))
        except ValueError:
            return None
    return None


def h_matches_month(text, month, year):
    """Текстовый фильтр по графе H: содержит ли текст дату вида '.MM.YYYY'
    (в любом месте строки) - универсальный способ отнести операцию к месяцу,
    работает одинаково для покупок, комиссий, возвратов и поступлений."""
    if not month:
        return True
    if not isinstance(text, str):
        return False
    return f'.{month:02d}.{year}' in text


def extract_rrn(text):
    if not isinstance(text, str):
        return None
    m = RRN_RE.search(text)
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# Загрузка файла и определение типа выписки
# ---------------------------------------------------------------------------

COMPANY_INN_RE = re.compile(r'ИНН\s*:\s*(\d+)')


def load_rows(path, sheet_name=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]

    header_row = None
    company_inn = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
        if row and isinstance(row[0], str) and company_inn is None:
            m = COMPANY_INN_RE.search(row[0])
            if m:
                company_inn = m.group(1)
        if row and row[0] == 'Дата':
            header_row = i
            break
    if header_row is None:
        raise ValueError("Не найдена строка заголовков ('Дата', ..., 'Назначение платежа')")

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        date_cell, acc, doc, op, mfo, debit, credit, purpose = (list(row) + [None] * 8)[:8]
        if not isinstance(date_cell, datetime):
            continue
        rows.append({
            'batch_date': date_cell,
            'account': acc,
            'doc': doc,
            'op': op,
            'debit': debit,
            'credit': credit,
            'purpose': purpose,
        })
    return rows, company_inn


def detect_file_type(rows):
    """Определяет формат выписки по ключевым словам в тексте назначения платежа."""
    sv, humo = 0, 0
    for r in rows[:500]:
        p = r.get('purpose')
        if not isinstance(p, str):
            continue
        pl = p.lower()
        if 'smartvista' in pl:
            sv += 1
        if 'humo' in pl:
            humo += 1
    if sv == 0 and humo == 0:
        return 'unknown'
    return 'smartvista' if sv >= humo else 'humo'


# ---------------------------------------------------------------------------
# Ветка 1: SmartVista (Uzcard/Visa/MC через терминал)
# ---------------------------------------------------------------------------

def detect_settlement_account(rows, company_inn=None):
    """Автоматически определяет номер расчётного счёта (р/с) компании по
    графе B. Основной метод - средний сегмент счёта (ИНН) совпадает с ИНН
    самой компании (берётся из шапки выписки, строка "Cчет: ... ИНН: XXXX").
    Это надёжнее, чем искать кавычки в тексте - кавычки бывают и у названий
    платёжных систем ("MC-UzCARD" и т.п.), что даёт ложные срабатывания.
    Запасной метод (если ИНН не найден в шапке) - ищем кавычки в тексте."""
    if company_inn:
        counts = {}
        for r in rows:
            acc = r.get('account')
            if isinstance(acc, str) and not acc.startswith('45'):
                parts = acc.split('/')
                if len(parts) >= 2 and parts[1] == company_inn:
                    counts[parts[0]] = counts.get(parts[0], 0) + 1
        if counts:
            return max(counts, key=counts.get)

    # запасной метод - по кавычкам с названием компании в тексте счёта
    counts = {}
    for r in rows:
        acc = r.get('account')
        if isinstance(acc, str) and not acc.startswith('45'):
            m = SETTLEMENT_ACCOUNT_RE.match(acc)
            if m:
                num = m.group(1)
                counts[num] = counts.get(num, 0) + 1
    if not counts:
        return None
    return max(counts, key=counts.get)


def extract_terminal_id(text):
    """ID терминала - встречается в тексте покупки как "Ид терминал X"
    и в тексте комиссии как "терминал ид X" - используем как часть ключа
    для точной привязки комиссии к своей покупке."""
    if not isinstance(text, str):
        return None
    m = re.search(r'(?:ид\s+терминал|терминал\s+ид)\s+(\d+)', text, re.IGNORECASE)
    return m.group(1) if m else None


def collect_returns(rows):
    """Универсальный сбор строк возврата/отмены - БЕЗ фильтра по счёту (все
    счета участвуют). Сначала ищем текст, содержащий "отмена" (ловит
    "Отмена оплата", "Отмена транзакции" и т.п.); если во всём файле таких
    строк с суммой в графе F не нашлось - используем "возв" (ловит "Возврат
    покупки товаров" и т.п.). Выбор идёт на уровне всего файла, а не месяца."""
    def matches(t, kw):
        return isinstance(t, str) and kw in t.lower()

    primary = [r for r in rows if r['op'] == 1 and r['debit'] and matches(r['purpose'], 'отмена')]
    keyword_rows = primary
    if not primary:
        keyword_rows = [r for r in rows if r['op'] == 1 and r['debit'] and matches(r['purpose'], 'возв')]

    out = []
    for r in keyword_rows:
        purpose = r['purpose']
        ts = extract_timestamp(purpose)
        out.append({'doc': r['doc'], 'amount': r['debit'], 'date': ts or r['batch_date'],
                    'purpose': purpose, 'rrn': extract_rrn(purpose)})
    return out


def analyze_smartvista(rows, company_inn=None):
    settlement_account = detect_settlement_account(rows, company_inn)

    grouped = []
    last_purchase = None
    deposits = []
    commission_reversals = []

    returns = collect_returns(rows)

    def is_deposit_row(r):
        if not settlement_account:
            return False
        acc = r.get('account')
        return isinstance(acc, str) and acc.split('/')[0] == settlement_account

    def is_commission_text(t):
        if not isinstance(t, str):
            return False
        tl = t.lower()
        return ('выручки от терминалов' in tl) or ('услуги банка' in tl)

    def is_own_account(r):
        """Счёт НЕ комиссионный (не начинается с "45") - используется для
        подсчёта "Сумма выручки"."""
        acc = r.get('account')
        return isinstance(acc, str) and not acc.startswith('45')

    # ПЕРВЫЙ ПРОХОД: собрать все покупки (счёт НЕ начинается с "45") и
    # построить индекс для точной привязки комиссий - ключ (время операции +
    # ID терминала). Комиссии в файле НЕ ВСЕГДА идут сразу после "своей"
    # покупки в хронологическом порядке - поэтому привязка "к последней
    # увиденной покупке" ненадёжна; точный ключ (время+терминал) совпадает у
    # покупки и её же комиссии всегда.
    purchase_index = {}  # (timestamp_str, terminal_id) -> purchase dict

    for r in rows:
        purpose = r['purpose']
        if r['op'] == 4 and r['credit'] and is_own_account(r):
            ts = extract_timestamp(purpose)
            purchase = {
                'doc': r['doc'], 'date': ts or r['batch_date'], 'amount': r['credit'],
                'reversal': False, 'rrn': extract_rrn(purpose), 'commissions': [],
            }
            grouped.append(purchase)
            term = extract_terminal_id(purpose)
            if ts and term:
                key = (ts.isoformat(), term)
                purchase_index.setdefault(key, purchase)
        elif r['op'] == 4 and r['credit'] and not is_own_account(r):
            # кредитовая строка на комиссионном счёте (например "отменная
            # операция" - возврат ранее списанной комиссии) - не выручка,
            # уменьшает "Итого комиссия"
            ts = extract_timestamp(purpose)
            commission_reversals.append({'doc': r['doc'], 'date': ts or r['batch_date'],
                                          'amount': r['credit']})

    # ВТОРОЙ ПРОХОД: привязать комиссии/поступления
    for r in rows:
        purpose = r['purpose']
        ts = extract_timestamp(purpose)

        if r['op'] == 4 and r['credit'] and is_own_account(r):
            last_purchase = purchase_index.get(
                (ts.isoformat(), extract_terminal_id(purpose))) if ts else None
            if last_purchase is None:
                last_purchase = next((g for g in grouped if g['doc'] == r['doc']), None)
        elif r['op'] == 1 and r['debit'] and is_deposit_row(r):
            deposits.append({'doc': r['doc'], 'amount': r['debit'], 'date': ts or r['batch_date']})
        elif r['op'] == 1 and r['debit'] and is_commission_text(purpose):
            term = extract_terminal_id(purpose)
            target = purchase_index.get((ts.isoformat(), term)) if (ts and term) else None
            comm_date = ts
            inherited = False
            if target is None:
                target = last_purchase
                if comm_date is None and target is not None:
                    comm_date = target['date']
                    inherited = True
            entry = {'doc': r['doc'], 'amount': r['debit'], 'date': comm_date,
                      'date_inherited': inherited,
                      'is_mps': isinstance(r.get('account'), str) and 'по пк мпс' in r['account'].lower()}
            if target is not None:
                target['commissions'].append(entry)
        elif r['op'] == 1 and r['debit']:
            last_purchase = None

    # Классификация комиссий на "обычную" (фикс.) и "клиента" (перем.):
    #
    # ГЛАВНОЕ ПРАВИЛО - по счёту в графе B ("Cчет/ИНН"):
    #   - счёт "...Комиссия за обработку транзакций...через POS Uzcard"
    #     (БЕЗ "по ПК МПС")           -> ВСЕГДА только "обычная комиссия" (0,2%),
    #                                     комиссии клиента здесь в принципе не бывает,
    #                                     даже если случайно 2 строки или процент
    #                                     похож на переменную ставку.
    #   - счёт "...POS Uzcard по ПК МПС" (международные карты)
    #                                  -> может быть разбивка на "обычную" (1%) и
    #                                     "клиента" (1,5-2%) - определяется по формуле
    #                                     (базовая_сумма = покупка - бОльший платёж).
    STANDARD_RATE = 0.01
    CLIENT_RATES = (0.015, 0.02)
    TOL = 0.0005

    standard_comms = []
    client_comms = []

    for g in grouped:
        comms = g['commissions']
        g['standard_commission'] = 0.0
        g['client_commission'] = 0.0

        plain = [c for c in comms if not c['is_mps']]   # счёт БЕЗ "по ПК МПС" - только обычная
        mps = [c for c in comms if c['is_mps']]          # счёт "по ПК МПС" - может быть сплит

        for c in plain:
            g['standard_commission'] += c['amount']
            standard_comms.append({'doc': c['doc'], 'date': c['date'], 'amount': c['amount']})

        if len(mps) == 2:
            c1, c2 = mps
            bigger, smaller = (c1, c2) if c1['amount'] >= c2['amount'] else (c2, c1)
            base = (g.get('amount') or 0) - bigger['amount']
            matched = False
            if base > 0:
                if (any(abs(bigger['amount'] / base - r) <= TOL for r in CLIENT_RATES)
                        and abs(smaller['amount'] / base - STANDARD_RATE) <= TOL):
                    g['client_commission'] += bigger['amount']
                    g['standard_commission'] += smaller['amount']
                    client_comms.append({'doc': bigger['doc'], 'date': bigger['date'], 'amount': bigger['amount']})
                    standard_comms.append({'doc': smaller['doc'], 'date': smaller['date'], 'amount': smaller['amount']})
                    matched = True
            if not matched:
                c_sorted = sorted(mps, key=lambda c: (c['doc'] is None, c['doc']))
                g['standard_commission'] += c_sorted[0]['amount']
                g['client_commission'] += c_sorted[1]['amount']
                standard_comms.append({'doc': c_sorted[0]['doc'], 'date': c_sorted[0]['date'], 'amount': c_sorted[0]['amount']})
                client_comms.append({'doc': c_sorted[1]['doc'], 'date': c_sorted[1]['date'], 'amount': c_sorted[1]['amount']})
        elif len(mps) == 1:
            c = mps[0]
            base = g.get('amount') or 0
            rate = c['amount'] / base if base else 0
            if any(abs(rate - r) <= TOL for r in CLIENT_RATES):
                g['client_commission'] += c['amount']
                client_comms.append({'doc': c['doc'], 'date': c['date'], 'amount': c['amount']})
            else:
                g['standard_commission'] += c['amount']
                standard_comms.append({'doc': c['doc'], 'date': c['date'], 'amount': c['amount']})
        elif len(mps) > 2:
            # необычный случай (>2 строк на счёте МПС) - относим всё к обычной,
            # чтобы не додумывать разбивку без чёткого правила
            for c in mps:
                g['standard_commission'] += c['amount']
                standard_comms.append({'doc': c['doc'], 'date': c['date'], 'amount': c['amount']})

    purchases = [{'doc': g['doc'], 'date': g['date'], 'amount': g['amount'],
                  'reversal': False} for g in grouped]

    return purchases, standard_comms, client_comms, returns, deposits, commission_reversals


# ---------------------------------------------------------------------------
# Ветка 2: HUMO (национальная платёжная система)
# ---------------------------------------------------------------------------

def analyze_humo(rows, company_inn=None):
    purchases, standard_comms, client_comms, deposits = [], [], [], []
    commission_reversals = []  # "Отмена Комиссии" - отдельная строка, не входит
                                # ни в "Обычную", ни в "Комиссию клиента" напрямую

    returns = collect_returns(rows)
    settlement_account = detect_settlement_account(rows, company_inn)

    for r in rows:
        purpose = r['purpose']
        if not isinstance(purpose, str):
            continue
        pl = purpose.lower()
        ts = extract_timestamp(purpose) or r['batch_date']
        acc = r.get('account') or ''

        is_comm_account = acc.startswith('45')
        is_own_account = bool(settlement_account) and acc.split('/')[0] == settlement_account

        if r['op'] == 4 and r['credit'] and not is_comm_account:
            purchases.append({'doc': r['doc'], 'date': ts, 'amount': r['credit'], 'reversal': False})
        elif is_comm_account and 'отмена ком' in pl:
            # возврат ранее списанной комиссии - отдельная строка, не входит
            # ни в "Обычную", ни в "Комиссию клиента"
            if r['credit']:
                commission_reversals.append({'doc': r['doc'], 'date': ts, 'amount': r['credit']})
        elif is_comm_account:
            # счёт "POS Humo" (комиссионный), без "Отмена Комиссии":
            # "Межбанковский расчет" -> Комиссия клиента, остальное -> Обычная
            target = client_comms if 'межбанковский расчет' in pl else standard_comms
            if r['debit']:
                target.append({'doc': r['doc'], 'date': ts, 'amount': r['debit']})
        elif r['op'] == 1 and r['debit'] and is_own_account:
            deposits.append({'doc': r['doc'], 'date': ts, 'amount': r['debit']})
        elif r['op'] == 1 and r['debit'] and 'зачисление денежных средств' in pl:
            deposits.append({'doc': r['doc'], 'date': ts, 'amount': r['debit']})

    return purchases, standard_comms, client_comms, returns, deposits, commission_reversals


# ---------------------------------------------------------------------------
# Общая точка входа
# ---------------------------------------------------------------------------

def analyze(path, sheet_name=None):
    rows, company_inn = load_rows(path, sheet_name)
    ftype = detect_file_type(rows)

    if ftype == 'humo':
        purchases, standard_comms, client_comms, returns, deposits, commission_reversals = analyze_humo(rows, company_inn)
    elif ftype == 'smartvista':
        purchases, standard_comms, client_comms, returns, deposits, commission_reversals = analyze_smartvista(rows, company_inn)
    else:
        raise ValueError(
            "Не удалось определить формат выписки (нет ни 'SmartVista', ни 'HUMO' "
            "в тексте назначения платежа). Нужно добавить в скрипт разбор нового формата."
        )

    return {
        'file_type': ftype,
        'purchases': purchases,
        'standard_commissions': standard_comms,
        'client_commissions': client_comms,
        'returns': returns,
        'deposits': deposits,
        'commission_reversals': commission_reversals,
    }


def _filter_period(items, start_date, end_date):
    """start_date/end_date - объекты datetime.date (включительно с обеих сторон).
    Если start_date не задан - фильтрация не применяется (все данные)."""
    if not start_date:
        return items
    return [x for x in items if x.get('date') and start_date <= x['date'].date() <= end_date]


def month_range(month, year):
    """Возвращает (первый день месяца, последний день месяца) как date."""
    import calendar as _cal
    last_day = _cal.monthrange(year, month)[1]
    return date(year, month, 1), date(year, month, last_day)


def summarize(data, start_date=None, end_date=None, month=None, year=None):
    """Возвращает агрегаты за период. Либо передайте start_date/end_date
    (объекты date, включительно), либо month/year (весь месяц целиком)."""
    if start_date is None and month:
        start_date, end_date = month_range(month, year)

    purchases = [p for p in _filter_period(data['purchases'], start_date, end_date) if not p.get('reversal')]
    standard_comms = _filter_period(data['standard_commissions'], start_date, end_date)
    client_comms = _filter_period(data['client_commissions'], start_date, end_date)
    returns = _filter_period(data['returns'], start_date, end_date)
    deposits = _filter_period(data['deposits'], start_date, end_date)
    commission_reversals = _filter_period(data.get('commission_reversals', []), start_date, end_date)

    total_purchase = sum(p['amount'] or 0 for p in purchases)
    total_bank = sum(c['amount'] or 0 for c in standard_comms)
    total_client = sum(c['amount'] or 0 for c in client_comms)
    total_returns = sum(r['amount'] or 0 for r in returns)
    total_deposits = sum(d['amount'] or 0 for d in deposits)
    total_commission_reversals = sum(x['amount'] or 0 for x in commission_reversals)
    total_commission = total_bank + total_client - total_commission_reversals

    net_amount = total_purchase - total_commission - total_returns

    return {
        'file_type': data['file_type'],
        'period_start': start_date,
        'period_end': end_date,
        'count': len(purchases),
        'total_purchase': total_purchase,
        'total_bank_commission': total_bank,
        'total_client_commission': total_client,
        'total_commission_reversals': total_commission_reversals,
        'total_commission': total_commission,
        'total_returns': total_returns,
        'total_deposits': total_deposits,
        'net_amount': net_amount,
        'difference': net_amount - total_deposits,
    }


def export_xlsx(data, out_path, start_date=None, end_date=None, month=None, year=None):
    """Формирует xlsx-отчёт: сводка + детализация по покупкам/комиссиям/возвратам."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    if start_date is None and month:
        start_date, end_date = month_range(month, year)

    summary = summarize(data, start_date=start_date, end_date=end_date)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Сводка'
    bold = Font(name='Arial', bold=True, size=12)
    normal = Font(name='Arial')
    header_fill = PatternFill('solid', fgColor='D9E1F2')

    period_label = f"{start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}" if start_date else "весь период"
    ws.cell(row=1, column=1, value=f"Отчёт за {period_label}").font = bold
    ws.cell(row=2, column=1, value=f"Формат выписки: {summary['file_type']}").font = normal

    rows_summary = [
        ('Сумма выручки', summary['total_purchase']),
        ('Обычная комиссия', summary['total_bank_commission']),
        ('Комиссия клиента', summary['total_client_commission']),
        ('Отмена комиссии (возврат ранее списанной)', summary['total_commission_reversals']),
        ('Итого комиссия', summary['total_commission']),
        ('Отмен (возврат покупок)', summary['total_returns']),
        ('Чистая сумма', summary['net_amount']),
        ('Поступление на р/с', summary['total_deposits']),
        ('Разница', summary['difference']),
    ]
    r_i = 4
    for label, val in rows_summary:
        c1 = ws.cell(row=r_i, column=1, value=label)
        c1.font = bold
        c1.fill = header_fill
        c2 = ws.cell(row=r_i, column=2, value=val)
        c2.font = bold
        c2.number_format = '#,##0.00'
        r_i += 1
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20

    # детальный лист по покупкам
    ws2 = wb.create_sheet('Покупки')
    headers = ['Дата', '№ док', 'Сумма']
    for i, h in enumerate(headers, start=1):
        c = ws2.cell(row=1, column=i, value=h)
        c.font = bold
        c.fill = header_fill
    ri = 2
    for p in _filter_period(data['purchases'], start_date, end_date):
        if p.get('reversal'):
            continue
        ws2.cell(row=ri, column=1, value=p['date']).number_format = 'DD.MM.YYYY HH:MM:SS'
        ws2.cell(row=ri, column=2, value=p['doc'])
        ws2.cell(row=ri, column=3, value=p['amount']).number_format = '#,##0.00'
        ri += 1
    for i, w in enumerate([20, 16, 16], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # детальный лист по комиссиям и возвратам
    ws3 = wb.create_sheet('Комиссии и возвраты')
    headers3 = ['Тип', 'Дата', '№ док', 'Сумма']
    for i, h in enumerate(headers3, start=1):
        c = ws3.cell(row=1, column=i, value=h)
        c.font = bold
        c.fill = header_fill
    ri = 2
    for label, items in [
        ('Обычная комиссия', _filter_period(data['standard_commissions'], start_date, end_date)),
        ('Комиссия клиента', _filter_period(data['client_commissions'], start_date, end_date)),
        ('Отмена комиссии', _filter_period(data.get('commission_reversals', []), start_date, end_date)),
        ('Возврат', _filter_period(data['returns'], start_date, end_date)),
        ('Поступление на р/с', _filter_period(data['deposits'], start_date, end_date)),
    ]:
        for x in items:
            ws3.cell(row=ri, column=1, value=label)
            ws3.cell(row=ri, column=2, value=x.get('date')).number_format = 'DD.MM.YYYY HH:MM:SS'
            ws3.cell(row=ri, column=3, value=x.get('doc'))
            ws3.cell(row=ri, column=4, value=x.get('amount')).number_format = '#,##0.00'
            ri += 1
    for i, w in enumerate([20, 20, 16, 16], start=1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_path)
    return out_path


def export_combined_xlsx(results, out_path, start_date, end_date):
    """Сводный xlsx по НЕСКОЛЬКИМ файлам сразу - одна строка на файл + итог.
    results: список {'name': имя файла, 'summary': словарь из summarize()}."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = 'Свод'
    bold = Font(name='Arial', bold=True)
    normal = Font(name='Arial')
    header_fill = PatternFill('solid', fgColor='D9E1F2')

    ws.cell(row=1, column=1,
            value=f"Сводный отчёт за {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}").font = bold

    headers = [
        'Файл', 'Формат', 'Кол-во покупок', 'Сумма выручки',
        'Обычная комиссия', 'Комиссия клиента', 'Отмена комиссии',
        'Итого комиссия', 'Отмен (возврат покупок)', 'Чистая сумма',
        'Поступление на р/с', 'Разница',
    ]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = bold
        c.fill = header_fill

    r_i = 4
    numeric_cols = list(range(4, 13))
    for r in results:
        s = r['summary']
        ws.cell(row=r_i, column=1, value=r['name'])
        ws.cell(row=r_i, column=2, value=s['file_type'])
        ws.cell(row=r_i, column=3, value=s['count'])
        ws.cell(row=r_i, column=4, value=s['total_purchase'])
        ws.cell(row=r_i, column=5, value=s['total_bank_commission'])
        ws.cell(row=r_i, column=6, value=s['total_client_commission'])
        ws.cell(row=r_i, column=7, value=s['total_commission_reversals'])
        ws.cell(row=r_i, column=8, value=s['total_commission'])
        ws.cell(row=r_i, column=9, value=s['total_returns'])
        ws.cell(row=r_i, column=10, value=s['net_amount'])
        ws.cell(row=r_i, column=11, value=s['total_deposits'])
        ws.cell(row=r_i, column=12, value=s['difference'])
        for col in numeric_cols:
            ws.cell(row=r_i, column=col).number_format = '#,##0.00'
        r_i += 1
    last_row = r_i - 1

    ws.cell(row=r_i, column=1, value='ИТОГО').font = bold
    for col in numeric_cols:
        letter = get_column_letter(col)
        cell = ws.cell(row=r_i, column=col, value=f'=SUM({letter}4:{letter}{last_row})')
        cell.font = bold
        cell.number_format = '#,##0.00'

    for i, w in enumerate([22, 12, 12, 16, 16, 16, 14, 14, 16, 16, 16, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_path)
    return out_path


def fmt(value):
    """Числовой формат '1 000 000,00' (пробел - разделитель тысяч, запятая - дробная часть)."""
    if abs(value) < 0.005:
        value = 0.0
    s = f"{value:,.2f}"
    s = s.replace(',', '\u00A0PLACEHOLDER\u00A0').replace('.', ',').replace('\u00A0PLACEHOLDER\u00A0', '\u00A0')
    return s


if __name__ == '__main__':
    path = sys.argv[1]
    month = int(sys.argv[2]) if len(sys.argv) > 2 else 6
    year = int(sys.argv[3]) if len(sys.argv) > 3 else 2026

    data = analyze(path)
    summary = summarize(data, month=month, year=year)

    print(f"Файл: {path}")
    print(f"Формат выписки: {summary['file_type']}")
    print(f"Период: {summary['period_start'].strftime('%d.%m.%Y')} - {summary['period_end'].strftime('%d.%m.%Y')}")
    print(f"Кол-во покупок: {summary['count']}")
    print(f"Сумма выручки: {fmt(summary['total_purchase'])}")
    print(f"Обычная комиссия: {fmt(summary['total_bank_commission'])}")
    print(f"Комиссия клиента: {fmt(summary['total_client_commission'])}")
    print(f"Отмена комиссии (возврат ранее списанной): {fmt(summary['total_commission_reversals'])}")
    print(f"Итого комиссия: {fmt(summary['total_commission'])}")
    print(f"Отмен (возврат покупок): {fmt(summary['total_returns'])}")
    print(f"Чистая сумма: {fmt(summary['net_amount'])}")
    print(f"Поступление на р/с: {fmt(summary['total_deposits'])}")
    print(f"Разница: {fmt(summary['difference'])}")
