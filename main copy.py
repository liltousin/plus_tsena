from __future__ import annotations

import re
import sys
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
import xml.etree.ElementTree as ET

# Change this value to bump prices by a different amount.
PRICE_INCREMENT = Decimal("10000000")

SUFFIX_FOR_NEW_FILE = f"+{PRICE_INCREMENT}"
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


def make_target_path(src: Path) -> Path:
    return src.with_name(f"{src.stem}{SUFFIX_FOR_NEW_FILE}{src.suffix}")


def update_workbook(src_path: Path, delta: Decimal) -> Path:
    if not src_path.exists():
        raise FileNotFoundError(f"File not found: {src_path}")
    if src_path.suffix.lower() != ".xlsx":
        raise ValueError("Only .xlsx files are supported.")

    target_path = make_target_path(src_path)
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


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python3 main.py <path-to-xlsx>")
        return 1

    source = Path(sys.argv[1]).expanduser().resolve()
    try:
        result = update_workbook(source, PRICE_INCREMENT)
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}")
        return 1

    print(f"Новый файл: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
