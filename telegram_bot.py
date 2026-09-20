"""
Телеграм-бот для разбора банковских выписок по эквайрингу.

Поддерживает ДВА формата выписки (SmartVista и HUMO), определяется
автоматически для каждого файла отдельно.

НОВОЕ:
1. Можно отправлять НЕСКОЛЬКО xlsx-файлов подряд (или альбомом) - бот
   соберёт их все, подождёт пару секунд и предложит выбрать период
   через встроенный календарь (кнопки в Telegram).
2. Выбор периода - произвольный диапазон дат (не только целый месяц):
   можно выбрать один день (начало = конец) или любую декаду/диапазон.
3. Результаты по всем отправленным файлам собираются в один сводный
   xlsx-файл и присылаются прямо в Telegram.

УСТАНОВКА:
    pip install python-telegram-bot openpyxl --upgrade

НАСТРОЙКА ТЕЛЕГРАМ-БОТА:
    1. Создать бота через @BotFather в Telegram, получить токен.
    2. Вставить токен в переменную BOT_TOKEN ниже (или в переменную окружения
       TELEGRAM_BOT_TOKEN).
    3. Положить analyze_commissions.py в ту же папку, что и этот файл.

ЗАПУСК:
    python3 telegram_bot.py
"""

import os
import re
import logging
import asyncio
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from datetime import datetime, date
import calendar as _calendar_module

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ContextTypes, filters,
)

import analyze_commissions as ac

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN')
if not BOT_TOKEN:
    raise SystemExit(
        "XATOLIK: TELEGRAM_BOT_TOKEN muhit o'zgaruvchisi topilmadi.\n"
        "Server sozlamalarida (Render -> Environment) TELEGRAM_BOT_TOKEN nomli\n"
        "o'zgaruvchi qo'shing va qiymatiga @BotFather bergan tokenni yozing."
    )

DOWNLOAD_DIR = 'bot_files'
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Сколько секунд ждать после получения файла, прежде чем считать, что
# пользователь закончил присылать файлы и пора спрашивать период.
BATCH_DEBOUNCE_SECONDS = 2.5

RU_MONTHS = ['', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
             'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
RU_WEEKDAYS = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']


# ---------------------------------------------------------------------------
# Встроенный календарь (inline-клавиатура)
# ---------------------------------------------------------------------------

def build_calendar(purpose, year, month, selected_start=None):
    """purpose: 'start' или 'end' - для какой даты строим календарь.
    selected_start: если выбираем конечную дату - показывает уже выбранную
    начальную дату в заголовке."""
    kb = []

    title = "Начальная дата периода" if purpose == 'start' else "Конечная дата периода"
    kb.append([InlineKeyboardButton(f"📅 {title}", callback_data="cal|noop")])
    if selected_start:
        kb.append([InlineKeyboardButton(
            f"Начало уже выбрано: {selected_start.strftime('%d.%m.%Y')}", callback_data="cal|noop")])

    kb.append([
        InlineKeyboardButton("«", callback_data=f"cal|{purpose}|nav|{year}|{month}|-1"),
        InlineKeyboardButton(f"{RU_MONTHS[month]} {year}", callback_data="cal|noop"),
        InlineKeyboardButton("»", callback_data=f"cal|{purpose}|nav|{year}|{month}|1"),
    ])
    kb.append([InlineKeyboardButton(d, callback_data="cal|noop") for d in RU_WEEKDAYS])

    month_days = _calendar_module.Calendar(firstweekday=0).monthdayscalendar(year, month)
    for week in month_days:
        row = []
        for day in week:
            if day == 0:
                row.append(InlineKeyboardButton(" ", callback_data="cal|noop"))
            else:
                row.append(InlineKeyboardButton(
                    str(day), callback_data=f"cal|{purpose}|pick|{year}|{month}|{day}"))
        kb.append(row)

    quick_row = []
    today = date.today()
    quick_row.append(InlineKeyboardButton(
        "Сегодня", callback_data=f"cal|{purpose}|pick|{today.year}|{today.month}|{today.day}"))
    kb.append(quick_row)

    return InlineKeyboardMarkup(kb)


# ---------------------------------------------------------------------------
# Обработчики
# ---------------------------------------------------------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Пришлите один или несколько xlsx-файлов банковской выписки "
        "(SmartVista или HUMO — формат определю автоматически для каждого).\n\n"
        "После того как файлы получены, я спрошу за какой период считать — "
        "через календарь (можно выбрать день, декаду или целый месяц).\n\n"
        "Результаты по всем файлам будут собраны в один сводный xlsx-файл "
        "и присланы сюда же."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')):
        await update.message.reply_text("Пришлите файл в формате .xlsx")
        return

    local_path = os.path.join(DOWNLOAD_DIR, f"{doc.file_unique_id}_{doc.file_name}")
    file = await doc.get_file()
    await file.download_to_drive(local_path)

    pending = context.user_data.setdefault('pending_files', [])
    pending.append({'path': local_path, 'name': doc.file_name})

    # переносим отправку запроса периода на чуть позже - если пользователь
    # шлёт ещё файлы, таймер каждый раз сбрасывается (debounce).
    # Реализовано на чистом asyncio, БЕЗ job_queue - это не требует
    # дополнительного пакета APScheduler и не может "тихо" падать.
    old_task = context.user_data.get('prompt_task')
    if old_task and not old_task.done():
        old_task.cancel()

    chat_id = update.effective_chat.id
    task = asyncio.create_task(_debounced_prompt(context, chat_id))
    context.user_data['prompt_task'] = task


async def _debounced_prompt(context: ContextTypes.DEFAULT_TYPE, chat_id):
    try:
        await asyncio.sleep(BATCH_DEBOUNCE_SECONDS)
    except asyncio.CancelledError:
        return  # пришёл ещё один файл - таймер перезапущен, этот вызов отменяем

    user_data = context.user_data
    pending = user_data.get('pending_files', [])
    if not pending:
        return

    names = "\n".join(f"• {f['name']}" for f in pending)
    await context.bot.send_message(
        chat_id, f"Получено файлов: {len(pending)}\n{names}\n\nВыберите начальную дату периода:")

    today = date.today()
    await context.bot.send_message(
        chat_id, "Календарь:",
        reply_markup=build_calendar('start', today.year, today.month))


async def handle_calendar_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data.split('|')

    if data[1] == 'noop':
        return

    purpose = data[1]
    action = data[2]

    if action == 'nav':
        year, month, delta = int(data[3]), int(data[4]), int(data[5])
        month += delta
        if month == 0:
            month, year = 12, year - 1
        elif month == 13:
            month, year = 1, year + 1
        selected_start = context.user_data.get('period_start')
        await query.edit_message_reply_markup(
            reply_markup=build_calendar(purpose, year, month, selected_start))
        return

    if action == 'pick':
        year, month, day = int(data[3]), int(data[4]), int(data[5])
        picked = date(year, month, day)

        if purpose == 'start':
            context.user_data['period_start'] = picked
            await query.edit_message_text(f"Начальная дата: {picked.strftime('%d.%m.%Y')}")
            await context.bot.send_message(
                update.effective_chat.id, "Теперь выберите конечную дату периода:",
                reply_markup=build_calendar('end', picked.year, picked.month, picked))
        else:
            start_d = context.user_data.get('period_start')
            if start_d is None:
                await query.edit_message_text("Сначала выберите начальную дату (/start).")
                return
            if picked < start_d:
                start_d, picked = picked, start_d
            await query.edit_message_text(
                f"Период: {start_d.strftime('%d.%m.%Y')} - {picked.strftime('%d.%m.%Y')}\n"
                f"Считаю..."
            )
            await process_batch(update, context, start_d, picked)


# ---------------------------------------------------------------------------
# Расчёт по всем накопленным файлам + выгрузка в Google Таблицу
# ---------------------------------------------------------------------------

async def process_batch(update: Update, context: ContextTypes.DEFAULT_TYPE,
                          start_d: date, end_d: date):
    chat_id = update.effective_chat.id
    pending = context.user_data.get('pending_files', [])
    if not pending:
        await context.bot.send_message(chat_id, "Нет файлов для расчёта. Пришлите xlsx-файл(ы).")
        return

    results = []
    for f in pending:
        try:
            data = ac.analyze(f['path'])
            summary = ac.summarize(data, start_date=start_d, end_date=end_d)
            results.append({'name': f['name'], 'summary': summary})

            out_path = os.path.join(DOWNLOAD_DIR, f"report_{os.path.basename(f['path'])}.xlsx")
            ac.export_xlsx(data, out_path, start_date=start_d, end_date=end_d)

            text = (
                f"📄 {f['name']}\n"
                f"Формат: {summary['file_type']}\n"
                f"Период: {start_d.strftime('%d.%m.%Y')} - {end_d.strftime('%d.%m.%Y')}\n"
                f"Кол-во покупок: {summary['count']}\n"
                f"Сумма выручки: {ac.fmt(summary['total_purchase'])}\n"
                f"Обычная комиссия: {ac.fmt(summary['total_bank_commission'])}\n"
                f"Комиссия клиента: {ac.fmt(summary['total_client_commission'])}\n"
                f"Отмена комиссии: {ac.fmt(summary['total_commission_reversals'])}\n"
                f"Итого комиссия: {ac.fmt(summary['total_commission'])}\n"
                f"Отмен (возврат покупок): {ac.fmt(summary['total_returns'])}\n"
                f"Чистая сумма: {ac.fmt(summary['net_amount'])}\n"
                f"Поступление на р/с: {ac.fmt(summary['total_deposits'])}\n"
                f"Разница: {ac.fmt(summary['difference'])}"
            )
            await context.bot.send_message(chat_id, text)
            await context.bot.send_document(chat_id, document=open(out_path, 'rb'),
                                             filename=f"отчёт_{f['name']}")
        except Exception as e:
            logger.exception(f"Ошибка обработки файла {f['name']}")
            await context.bot.send_message(chat_id, f"❌ Ошибка при обработке {f['name']}: {e}")
        finally:
            if os.path.exists(f['path']):
                os.remove(f['path'])

    # общий итог по всем файлам
    if len(results) > 1:
        total_purchase = sum(r['summary']['total_purchase'] for r in results)
        total_bank = sum(r['summary']['total_bank_commission'] for r in results)
        total_client = sum(r['summary']['total_client_commission'] for r in results)
        total_comm_reversals = sum(r['summary']['total_commission_reversals'] for r in results)
        total_commission = sum(r['summary']['total_commission'] for r in results)
        total_returns = sum(r['summary']['total_returns'] for r in results)
        total_net = sum(r['summary']['net_amount'] for r in results)
        total_deposits = sum(r['summary']['total_deposits'] for r in results)
        total_diff = sum(r['summary']['difference'] for r in results)
        await context.bot.send_message(
            chat_id,
            f"📊 ИТОГО по {len(results)} файлам:\n"
            f"Сумма выручки: {ac.fmt(total_purchase)}\n"
            f"Обычная комиссия: {ac.fmt(total_bank)}\n"
            f"Комиссия клиента: {ac.fmt(total_client)}\n"
            f"Отмена комиссии: {ac.fmt(total_comm_reversals)}\n"
            f"Итого комиссия: {ac.fmt(total_commission)}\n"
            f"Отмен (возврат покупок): {ac.fmt(total_returns)}\n"
            f"Чистая сумма: {ac.fmt(total_net)}\n"
            f"Поступление на р/с: {ac.fmt(total_deposits)}\n"
            f"Разница: {ac.fmt(total_diff)}"
        )

    # сводный excel-файл по всем файлам сразу - отправляется прямо в Telegram
    if results:
        combined_path = os.path.join(DOWNLOAD_DIR, f"свод_{start_d.strftime('%d%m%Y')}_{end_d.strftime('%d%m%Y')}.xlsx")
        ac.export_combined_xlsx(results, combined_path, start_d, end_d)
        await context.bot.send_document(chat_id, document=open(combined_path, 'rb'),
                                         filename=os.path.basename(combined_path))

    context.user_data['pending_files'] = []
    context.user_data['period_start'] = None


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """Глобальный обработчик ошибок - чтобы бот никогда не 'молчал' при сбое,
    а сообщал о проблеме в чат (и подробности в консоль для отладки)."""
    logger.exception("Необработанная ошибка", exc_info=context.error)
    try:
        if isinstance(update, Update) and update.effective_chat:
            await context.bot.send_message(
                update.effective_chat.id,
                f"⚠️ Произошла ошибка: {context.error}\n\n"
                f"Попробуйте ещё раз или пришлите файлы заново."
            )
    except Exception:
        pass


class _PingHandler(BaseHTTPRequestHandler):
    """Render.com kabi xostingga 'bot tirik' deb ko'rsatish uchun eng oddiy
    veb-server. Botning asosiy ishiga (fayllarni tahlil qilish va h.k.)
    hech qanday aloqasi yo'q - faqat GET so'roviga 'OK' deb javob beradi."""

    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/plain; charset=utf-8')
        self.end_headers()
        self.wfile.write("Bot ishlayapti".encode('utf-8'))

    def log_message(self, format, *args):
        pass  # konsolni ping-so'rovlar bilan to'ldirmaslik uchun


def keep_alive():
    """PORT muhit o'zgaruvchisida (Render avtomatik beradi) fon rejimida
    http-server ishga tushiradi, shunda Render bu servisni 'veb-servis' deb
    tan oladi va uni o'chirib qo'ymaydi."""
    port = int(os.environ.get('PORT', '10000'))
    server = ThreadingHTTPServer(('0.0.0.0', port), _PingHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    logger.info(f"Keep-alive server {port}-portda ishga tushdi")


def main():
    keep_alive()
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler('start', start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(handle_calendar_callback, pattern=r'^cal\|'))
    app.add_error_handler(error_handler)
    print("Бот запущен...")
    app.run_polling()


if __name__ == '__main__':
    main()
