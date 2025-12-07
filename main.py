from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
import xml.etree.ElementTree as ET

from dotenv import load_dotenv

logger = logging.getLogger(__name__)

# Change this value to bump prices by a different amount when running as a script.
PRICE_INCREMENT = Decimal("10000000")

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
X14AC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
XML_NS = "http://www.w3.org/XML/1998/namespace"
SKIP_HEADER_LABELS = {"Предоп.1п"}
STAR_PATTERN = re.compile(r"^([+-]?\d+(?:[.,]\d+)?)(\*+)$")
NUMBER_PATTERN = re.compile(r"^[+-]?\d+(?:[.,]\d+)?$")

ET.register_namespace("", NS)
ET.register_namespace("r", REL_NS)
ET.register_namespace("mc", MC_NS)
ET.register_namespace("x14ac", X14AC_NS)


def column_from_cell_ref(cell_ref: str | None) -> str | None:
    if not cell_ref:
        return None
    letters = []
    for ch in cell_ref:
        if ch.isalpha():
            letters.append(ch.upper())
        else:
            break
    return "".join(letters) if letters else None


def bump_number_string(number_str: str, delta: Decimal) -> tuple[str, bool]:
    try:
        parsed = Decimal(number_str.replace(",", "."))
    except InvalidOperation:
        return number_str, False

    if parsed == 0:
        return number_str, False

    new_value = parsed + delta
    separator = "," if "," in number_str else "."
    decimals = 0
    if separator in number_str:
        decimals = len(number_str.split(separator)[1])

    if decimals:
        quant = Decimal("1").scaleb(-decimals)
        new_value = new_value.quantize(quant)
    elif new_value == new_value.to_integral():
        new_value = new_value.quantize(Decimal("1"))

    text = format(new_value, "f")
    if separator == ",":
        text = text.replace(".", ",")
    return text, True


def adjust_text_value(text: str, delta: Decimal) -> tuple[str, bool]:
    stripped = text.strip()
    if not stripped:
        return text, False

    prefix_len = len(text) - len(text.lstrip())
    suffix_len = len(text) - len(text.rstrip())
    prefix = text[:prefix_len]
    suffix = text[len(text) - suffix_len :] if suffix_len else ""
    core = text[prefix_len : len(text) - suffix_len if suffix_len else len(text)]

    star_match = STAR_PATTERN.match(core)
    if star_match:
        number_part, stars = star_match.groups()
        new_number, changed = bump_number_string(number_part, delta)
        if changed:
            return f"{prefix}{new_number}{stars}{suffix}", True
        return text, False

    if NUMBER_PATTERN.match(core):
        new_number, changed = bump_number_string(core, delta)
        if changed:
            return f"{prefix}{new_number}{suffix}", True

    return text, False


def parse_shared_strings(xml_bytes: bytes) -> list[str]:
    root = ET.fromstring(xml_bytes)
    items: list[str] = []
    for si in root.findall(f"{{{NS}}}si"):
        texts: list[str] = []
        t_elem = si.find(f"{{{NS}}}t")
        if t_elem is not None and t_elem.text is not None:
            texts.append(t_elem.text)
        else:
            for run in si.findall(f"{{{NS}}}r"):
                rt = run.find(f"{{{NS}}}t")
                if rt is not None and rt.text is not None:
                    texts.append(rt.text)
        items.append("".join(texts))
    return items


def set_text_with_space(t_elem: ET.Element, text: str) -> None:
    t_elem.text = text
    if text.startswith(" ") or text.endswith(" "):
        t_elem.set(f"{{{XML_NS}}}space", "preserve")
    elif f"{{{XML_NS}}}space" in t_elem.attrib:
        del t_elem.attrib[f"{{{XML_NS}}}space"]


def process_shared_strings(xml_bytes: bytes, delta: Decimal) -> bytes:
    # Shared strings stay untouched; adjustments happen per-cell to avoid
    # changing skipped columns that share string indices.
    return xml_bytes


def resolve_cell_text(
    cell: ET.Element, shared_strings: list[str] | None
) -> str | None:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        is_elem = cell.find(f"{{{NS}}}is")
        if is_elem is None:
            return None
        t_elem = is_elem.find(f"{{{NS}}}t")
        return t_elem.text if t_elem is not None else None
    if cell_type == "str":
        value_elem = cell.find(f"{{{NS}}}v")
        return value_elem.text if value_elem is not None else None
    if cell_type == "s":
        if shared_strings is None:
            return None
        value_elem = cell.find(f"{{{NS}}}v")
        if value_elem is None or value_elem.text is None:
            return None
        try:
            idx = int(value_elem.text)
        except ValueError:
            return None
        if 0 <= idx < len(shared_strings):
            return shared_strings[idx]
    return None


def process_sheet(
    xml_bytes: bytes, delta: Decimal, shared_strings: list[str] | None
) -> bytes:
    root = ET.fromstring(xml_bytes)
    changed = False

    skip_columns: set[str] = set()
    for cell in root.iter(f"{{{NS}}}c"):
        text_value = resolve_cell_text(cell, shared_strings)
        if text_value and text_value.strip() in SKIP_HEADER_LABELS:
            col = column_from_cell_ref(cell.get("r"))
            if col:
                skip_columns.add(col)

    for cell in root.iter(f"{{{NS}}}c"):
        if cell.find(f"{{{NS}}}f") is not None:
            continue

        cell_column = column_from_cell_ref(cell.get("r"))
        if cell_column and cell_column in skip_columns:
            continue

        cell_type = cell.get("t")
        if cell_type == "inlineStr":
            is_elem = cell.find(f"{{{NS}}}is")
            if is_elem is None:
                continue
            text_elem = is_elem.find(f"{{{NS}}}t")
            if text_elem is None or text_elem.text is None:
                continue
            new_text, updated = adjust_text_value(text_elem.text, delta)
            if updated:
                set_text_with_space(text_elem, new_text)
                changed = True

        elif cell_type == "str":
            value_elem = cell.find(f"{{{NS}}}v")
            if value_elem is None or value_elem.text is None:
                continue
            new_text, updated = adjust_text_value(value_elem.text, delta)
            if updated:
                value_elem.text = new_text
                changed = True

        elif cell_type == "s":
            if shared_strings is None:
                continue
            value_elem = cell.find(f"{{{NS}}}v")
            if value_elem is None or value_elem.text is None:
                continue
            try:
                idx = int(value_elem.text)
            except ValueError:
                continue
            if not (0 <= idx < len(shared_strings)):
                continue
            new_text, updated = adjust_text_value(shared_strings[idx], delta)
            if updated:
                cell.set("t", "inlineStr")
                if value_elem in list(cell):
                    cell.remove(value_elem)
                is_elem = ET.Element(f"{{{NS}}}is")
                t_elem = ET.SubElement(is_elem, f"{{{NS}}}t")
                set_text_with_space(t_elem, new_text)
                cell.append(is_elem)
                changed = True

        elif cell_type in (None, "n"):
            value_elem = cell.find(f"{{{NS}}}v")
            if value_elem is None or value_elem.text is None:
                continue
            new_value, updated = bump_number_string(value_elem.text, delta)
            if updated:
                value_elem.text = new_value
                changed = True

    if not changed:
        return xml_bytes
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def format_delta_suffix(delta: Decimal) -> str:
    """Convert delta into a human-friendly suffix like '+10' or '-2.5'."""
    sign = "+" if delta >= 0 else ""
    text = format(delta, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return f"{sign}{text}"


def make_target_path(src: Path, delta: Decimal, suffix: str | None = None) -> Path:
    suffix = suffix if suffix is not None else format_delta_suffix(delta)
    return src.with_name(f"{src.stem}{suffix}{src.suffix}")


def update_workbook(src_path: Path, delta: Decimal, suffix: str | None = None) -> Path:
    if not src_path.exists():
        raise FileNotFoundError(f"File not found: {src_path}")
    if src_path.suffix.lower() != ".xlsx":
        raise ValueError("Only .xlsx files are supported.")

    target_path = make_target_path(src_path, delta, suffix)
    if target_path.exists():
        raise FileExistsError(f"Target file already exists: {target_path}")

    shared_strings_bytes: bytes | None = None
    shared_strings: list[str] | None = None

    with zipfile.ZipFile(src_path, "r") as zin:
        try:
            shared_strings_bytes = zin.read("xl/sharedStrings.xml")
            shared_strings = parse_shared_strings(shared_strings_bytes)
        except KeyError:
            shared_strings_bytes = None
            shared_strings = None

        with zipfile.ZipFile(target_path, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/sharedStrings.xml" and shared_strings_bytes is not None:
                    data = process_shared_strings(shared_strings_bytes, delta)
                elif item.filename.startswith("xl/worksheets/") and re.match(
                    r"sheet\d+\.xml$", Path(item.filename).name
                ):
                    data = process_sheet(data, delta, shared_strings)

                zout.writestr(item, data)

    return target_path


def run_bot() -> int:
    try:
        from telegram import Update
        from telegram.ext import (
            Application,
            CommandHandler,
            ConversationHandler,
            ContextTypes,
            MessageHandler,
            filters,
        )
    except ModuleNotFoundError:
        print("Не хватает зависимости python-telegram-bot. Установи: pip install -r requirements.txt")
        return 1

    load_dotenv()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

    WAITING_FILE, WAITING_DELTA = range(2)
    TOKEN_ENV_NAME = "BOT_TOKEN"

    def read_token() -> str:
        value = os.getenv(TOKEN_ENV_NAME)
        if value:
            return value
        raise RuntimeError(f"Укажи токен бота в .env (строка BOT_TOKEN=...) или переменной окружения {TOKEN_ENV_NAME}.")

    def cleanup_user_state(context: ContextTypes.DEFAULT_TYPE) -> None:
        workdir = context.user_data.pop("workdir", None)
        context.user_data.pop("source_path", None)
        if workdir:
            shutil.rmtree(workdir, ignore_errors=True)

    async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        cleanup_user_state(context)
        await update.message.reply_text(
            "Привет! Пришли .xlsx файл с прайсом, я спрошу на сколько увеличить цены и верну новый файл."
        )
        return WAITING_FILE

    async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        document = update.message.document
        if document is None or not document.file_name.lower().endswith(".xlsx"):
            await update.message.reply_text("Нужен файл в формате .xlsx.")
            return WAITING_FILE

        cleanup_user_state(context)
        tmp_dir = Path(tempfile.mkdtemp(prefix="price_bot_"))
        target_file = tmp_dir / Path(document.file_name).name

        telegram_file = await document.get_file()
        await telegram_file.download_to_drive(custom_path=str(target_file))

        context.user_data["workdir"] = str(tmp_dir)
        context.user_data["source_path"] = str(target_file)

        await update.message.reply_text(
            "Сколько прибавить к каждому числу? Пример: 500 или 12.5"
        )
        return WAITING_DELTA

    async def handle_delta(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        text = (update.message.text or "").strip()
        try:
            delta = Decimal(text.replace(",", "."))
        except InvalidOperation:
            await update.message.reply_text("Не понял число. Напиши, например, 500 или 12.5")
            return WAITING_DELTA

        source_path_str = context.user_data.get("source_path")
        if not source_path_str:
            await update.message.reply_text("Сначала пришли файл .xlsx.")
            return WAITING_FILE

        source_path = Path(source_path_str)
        try:
            updated_path = update_workbook(source_path, delta)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Failed to update workbook")
            await update.message.reply_text(f"Не получилось обработать файл: {exc}")
            cleanup_user_state(context)
            return ConversationHandler.END

        try:
            with updated_path.open("rb") as fh:
                await update.message.reply_document(document=fh, filename=updated_path.name)
        finally:
            cleanup_user_state(context)

        await update.message.reply_text("Готово. Можешь прислать следующий файл.")
        return WAITING_FILE

    async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        cleanup_user_state(context)
        await update.message.reply_text("Отменено. Отправь /start, чтобы начать заново.")
        return ConversationHandler.END

    async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
        logger.exception("Unhandled error while processing update %s", update, exc_info=context.error)

    token = read_token()
    application = Application.builder().token(token).build()

    conversation = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.Document.FileExtension("xlsx"), handle_file),
        ],
        states={
            WAITING_FILE: [
                MessageHandler(filters.Document.FileExtension("xlsx"), handle_file),
                MessageHandler(filters.ALL & ~filters.COMMAND, start),
            ],
            WAITING_DELTA: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_delta),
                MessageHandler(filters.Document.ALL, handle_file),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    application.add_handler(conversation)
    application.add_error_handler(error_handler)
    application.run_polling()
    return 0


def main(argv: list[str] | None = None) -> int:
    return run_bot()


if __name__ == "__main__":
    raise SystemExit(main())
