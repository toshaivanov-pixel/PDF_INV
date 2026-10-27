#!/usr/bin/env python3
"""Interactive invoice-to-PDF generator.

Workflow:
  1. scan ./data for CSV and JSON files and ./templates for HTML templates;
  2. show both lists and let the user pick one data file and one template;
  3. list every invoice (by invoice id) found in the selected data file;
  4. substitute the chosen invoice into the template and write a PDF to ./output;
  5. open the produced PDF with the OS default application.

CSV files are read with pandas when available and with the stdlib ``csv``
module otherwise; JSON files are read with the stdlib ``json`` module.
PDF rendering is done by WeasyPrint. Cyrillic support comes from the Roboto
font files bundled in ./assets/fonts (loaded through CSS ``@font-face``), so
nothing has to be installed system-wide. The script runs on Windows and macOS.
"""

from __future__ import annotations

import csv
import json
import os
import platform
import re
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Iterable

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

DATA_SUFFIXES = (".csv", ".json")
TEMPLATE_SUFFIXES = (".html", ".htm")

LINE_WIDTH = 66
LINE = "=" * LINE_WIDTH
THIN_LINE = "-" * LINE_WIDTH

# ---------------------------------------------------------------------------
# field aliases: canonical name -> accepted source column / key names
# ---------------------------------------------------------------------------

INVOICE_ID_ALIASES = (
    "invoice_id", "invoiceid", "invoice_no", "invoice_number", "check_id",
    "receipt_id", "document_number", "doc_number", "number", "id",
)

FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "date": ("date", "invoice_date", "issue_date", "check_date", "created_at", "receipt_date"),
    "due_date": ("due_date", "payment_due_date", "due", "pay_before"),
    "customer_name": ("customer_name", "client_name", "buyer_name", "purchaser", "payer", "customer", "client"),
    "customer_address": ("customer_address", "client_address", "address", "billing_address", "shipping_address"),
    "customer_email": ("customer_email", "client_email", "email", "mail"),
    "customer_phone": ("customer_phone", "client_phone", "phone", "tel"),
    "seller_name": ("seller_name", "vendor_name", "supplier_name", "company_name", "merchant", "seller", "vendor", "supplier", "company"),
    "seller_inn": ("seller_inn", "vendor_inn", "supplier_inn", "seller_tin", "inn", "tin"),
    "payment_method": ("payment_method", "payment_type", "pay_method", "method", "payment"),
    "currency": ("currency", "currency_code"),
    "note": ("note", "notes", "comment", "comments", "memo", "description_note"),
}

# invoice level fields that are never treated as line-item data
TOTAL_ALIASES = ("total", "grand_total", "amount_total", "total_amount", "sum_total", "summary")

ITEM_NAME_ALIASES = ("item_description", "item_name", "description", "product", "service", "goods", "title", "item", "name")
ITEM_QTY_ALIASES = ("quantity", "qty", "count", "units")
ITEM_PRICE_ALIASES = ("unit_price", "price", "rate", "cost", "tariff")
ITEM_TOTAL_ALIASES = ("line_total", "item_total", "position_total", "sum", "amount", "total")

ITEMS_LIST_ALIASES = ("items", "positions", "lines", "line_items", "products", "services", "goods")

CURRENCY_SYMBOLS = {
    "RUB": "\u20bd", "USD": "$", "EUR": "\u20ac", "KZT": "\u20b8",
    "BYN": "Br", "UAH": "\u20b4", "GBP": "\u00a3", "CNY": "\u00a5",
}

ALL_INVOICE_ALIASES = (
    set(FIELD_ALIASES["date"]) | set(FIELD_ALIASES["due_date"])
    | set(FIELD_ALIASES["customer_name"]) | set(FIELD_ALIASES["customer_address"])
    | set(FIELD_ALIASES["customer_email"]) | set(FIELD_ALIASES["customer_phone"])
    | set(FIELD_ALIASES["seller_name"]) | set(FIELD_ALIASES["seller_inn"])
    | set(FIELD_ALIASES["payment_method"]) | set(FIELD_ALIASES["currency"])
    | set(FIELD_ALIASES["note"]) | set(TOTAL_ALIASES) | set(INVOICE_ID_ALIASES)
)


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------

def normalize_key(key: Any) -> str:
    """Lower-case a column/key name, collapse separators into underscores."""
    text = str(key).replace("\ufeff", "").strip().lower()
    text = re.sub(r"[\s\-.]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_")


def to_float(value: Any) -> float | None:
    """Best-effort conversion of spreadsheet-ish numbers to float."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    text = re.sub(r"[^\d.,\-+]", "", text)
    if not text:
        return None
    if "," in text and "." in text:
        # the right-most separator is the decimal one
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def fmt_money(value: Any, currency: Any = "RUB") -> str:
    """Render a number as an amount with a currency symbol (1 234,50 ₽)."""
    number = to_float(value)
    if number is None:
        number = 0.0
    code = str(currency or "RUB").strip().upper()
    formatted = f"{number:,.2f}".replace(",", "\u00a0").replace(".", ",")
    symbol = CURRENCY_SYMBOLS.get(code, code if code else "")
    return f"{formatted} {symbol}".strip()


def fmt_quantity(value: Any) -> str:
    """Render a quantity without trailing zeroes (2 instead of 2.00)."""
    number = to_float(value)
    if number is None:
        return ""
    if float(number).is_integer():
        return str(int(number))
    return f"{number:.3f}".rstrip("0").rstrip(".").replace(".", ",")


def pick(source: dict[str, Any], aliases: Iterable[str]) -> str:
    """Return the first non-empty value among ``aliases`` in ``source``."""
    for alias in aliases:
        value = source.get(alias)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


# ---------------------------------------------------------------------------
# discovery
# ---------------------------------------------------------------------------

def discover_files(directory: Path, suffixes: tuple[str, ...]) -> list[Path]:
    """Return data/template files of the given extensions, sorted by name."""
    if not directory.is_dir():
        return []
    found = [
        p for p in directory.iterdir()
        if p.is_file() and p.suffix.lower() in suffixes
    ]
    return sorted(found, key=lambda p: p.name.lower())


# ---------------------------------------------------------------------------
# invoice assembly
# ---------------------------------------------------------------------------

def pick_number(source: dict[str, Any], aliases: Iterable[str]) -> float | None:
    """Return the first alias of ``source`` that parses as a number."""
    for alias in aliases:
        if alias in source:
            number = to_float(source[alias])
            if number is not None:
                return number
    return None


def make_item(raw: dict[str, Any]) -> dict[str, Any]:
    """Normalize one line item, computing the line total when missing."""
    name = pick(raw, ITEM_NAME_ALIASES)
    quantity = pick_number(raw, ITEM_QTY_ALIASES)
    unit_price = pick_number(raw, ITEM_PRICE_ALIASES)
    line_total = pick_number(raw, ITEM_TOTAL_ALIASES)
    if quantity is None:
        quantity = 1.0
    if unit_price is None:
        unit_price = 0.0
    if line_total is None:
        line_total = quantity * unit_price
    return {
        "name": name or "\u2014",
        "quantity": quantity,
        "unit_price": unit_price,
        "line_total": line_total,
    }


def build_invoice(invoice_id: str, fields: dict[str, Any], items: list[dict[str, Any]]) -> dict[str, Any]:
    """Merge invoice level fields and line items into the template context."""
    invoice: dict[str, Any] = {"invoice_id": invoice_id}
    for key, aliases in FIELD_ALIASES.items():
        invoice[key] = pick(fields, aliases)
    if not invoice["currency"]:
        invoice["currency"] = "RUB"

    consumed = ALL_INVOICE_ALIASES
    invoice["extras"] = {
        key: value for key, value in fields.items()
        if key not in consumed and str(value).strip()
    }

    invoice["items"] = items
    computed = sum(item["line_total"] for item in items)
    if computed:
        invoice["total"] = round(computed, 2)
    else:
        invoice["total"] = round(to_float(pick(fields, TOTAL_ALIASES)) or 0.0, 2)
    invoice["total_formatted"] = fmt_money(invoice["total"], invoice["currency"])
    invoice["items_count"] = len(items)
    return invoice


def invoices_from_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Group flat CSV rows (one row per line item) into invoices."""
    if not rows:
        return []
    columns = list(rows[0].keys())
    id_col = next((c for c in INVOICE_ID_ALIASES if c in columns), None)
    name_col = next((c for c in ITEM_NAME_ALIASES if c in columns), None)
    qty_col = next((c for c in ITEM_QTY_ALIASES if c in columns), None)
    price_col = next((c for c in ITEM_PRICE_ALIASES if c in columns), None)

    item_cols = {c for c in (name_col, qty_col, price_col) if c}
    # a total column belongs to the line item only when item columns exist
    total_col = None
    if item_cols:
        total_col = next((c for c in ITEM_TOTAL_ALIASES if c in columns and c not in item_cols), None)
        if total_col:
            item_cols.add(total_col)

    meta_cols = [c for c in columns if c != id_col and c not in item_cols]

    order: list[str] = []
    grouped: dict[str, dict[str, Any]] = {}
    for row in rows:
        invoice_id = (row.get(id_col) or "").strip() if id_col else ""
        if not invoice_id:
            continue
        if invoice_id not in grouped:
            grouped[invoice_id] = {
                "fields": {c: (row.get(c) or "") for c in meta_cols},
                "items": [],
            }
            order.append(invoice_id)

        raw_item = {}
        if name_col:
            raw_item[name_col] = row.get(name_col)
        if qty_col:
            raw_item[qty_col] = row.get(qty_col)
        if price_col:
            raw_item[price_col] = row.get(price_col)
        if total_col:
            raw_item[total_col] = row.get(total_col)
        if any(str(v or "").strip() for v in raw_item.values()):
            grouped[invoice_id]["items"].append(make_item(raw_item))

    return [
        build_invoice(invoice_id, grouped[invoice_id]["fields"], grouped[invoice_id]["items"])
        for invoice_id in order
    ]


def invoices_from_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Normalize JSON invoices (nested ``items`` supported) into the context."""
    invoices: list[dict[str, Any]] = []
    for record in records:
        if not isinstance(record, dict):
            continue
        fields = {normalize_key(k): ("" if v is None else v) for k, v in record.items()}
        items_payload: list[Any] = []
        items_key = None
        for alias in ITEMS_LIST_ALIASES:
            value = fields.get(alias)
            if isinstance(value, list):
                items_payload = value
                items_key = alias
                break
        if items_key:
            fields.pop(items_key, None)

        invoice_id = pick(fields, INVOICE_ID_ALIASES)
        if not invoice_id:
            invoice_id = f"DOC-{len(invoices) + 1:03d}"

        # scalar invoice level fields become strings, nested dicts are flattened
        flat_fields: dict[str, Any] = {}
        for key, value in fields.items():
            if isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    flat_fields[f"{key}_{sub_key}"] = sub_value
            elif isinstance(value, list):
                flat_fields[key] = ", ".join(str(v) for v in value)
            else:
                flat_fields[key] = value

        items = []
        for raw_item in items_payload:
            if isinstance(raw_item, dict):
                items.append(make_item({normalize_key(k): v for k, v in raw_item.items()}))
            else:
                items.append({"name": str(raw_item), "quantity": 1.0,
                              "unit_price": 0.0, "line_total": 0.0})
        if not items:
            # a flat record without an items list is a single line document
            name = pick(fields, ITEM_NAME_ALIASES)
            if name:
                items = [make_item(fields)]
        invoices.append(build_invoice(invoice_id, flat_fields, items))
    return invoices


# ---------------------------------------------------------------------------
# data loading
# ---------------------------------------------------------------------------

def read_csv_rows(path: Path) -> list[dict[str, Any]]:
    """Read a CSV file into a list of dicts (pandas first, stdlib csv after)."""
    try:
        import pandas as pd
    except ImportError:
        pd = None  # type: ignore[assignment]

    if pd is not None:
        frame = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        rows = [
            {normalize_key(k): str(v).strip() for k, v in row.items()}
            for row in frame.to_dict("records")
        ]
        return rows

    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return [
            {normalize_key(k): (v or "").strip() for k, v in row.items() if k is not None}
            for row in reader
        ]


def read_json_records(path: Path) -> list[dict[str, Any]]:
    """Read a JSON file and return the list of invoice objects it contains."""
    data = json.loads(path.read_text(encoding="utf-8-sig"))

    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]

    if isinstance(data, dict):
        # look for the list of documents under a well-known key
        for key, value in data.items():
            if isinstance(value, list) and value and isinstance(value[0], dict):
                normalized = normalize_key(key)
                if normalized in set(ITEMS_LIST_ALIASES) | {"invoices", "checks", "receipts", "documents", "data", "records", "orders"}:
                    return value
        # single invoice stored as one object
        return [data]

    raise ValueError("unsupported JSON structure: expected an object or an array")


def load_invoices(path: Path) -> list[dict[str, Any]]:
    """Load invoices from a CSV or JSON data file."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return invoices_from_rows(read_csv_rows(path))
    if suffix == ".json":
        return invoices_from_records(read_json_records(path))
    raise ValueError(f"неизвестный формат файла: {path.name}")


# ---------------------------------------------------------------------------
# console menu
# ---------------------------------------------------------------------------

def print_header(title: str) -> None:
    print(f"\n{LINE}")
    print(title.center(LINE_WIDTH))
    print(LINE)


def print_subheader(title: str) -> None:
    print(f"\n{THIN_LINE}")
    print(title)
    print(THIN_LINE)


def print_menu(caption: str, entries: list[str]) -> None:
    """Print a numbered menu; numbering starts at 1, 0 always means quit."""
    print(f"\n{caption}:\n")
    for index, entry in enumerate(entries, start=1):
        print(f"  {index:>3}. {entry}")
    print("\n    0. Выход")


def ask_choice(prompt: str, count: int) -> int | None:
    """Ask for a menu number. Returns a 0-based index or None when quitting."""
    while True:
        answer = input(prompt).strip()
        if answer.lower() in {"", "0", "q", "quit", "exit", "в", "выход"}:
            return None
        if re.fullmatch(r"\d+", answer) and 1 <= int(answer) <= count:
            return int(answer) - 1
        print(f"     ! Нужно целое число от 1 до {count} (или 0 для выхода).")


def human_size(path: Path) -> str:
    size = path.stat().st_size
    if size < 1024:
        return f"{size} Б"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} КБ"
    return f"{size / 1024 / 1024:.1f} МБ"


# ---------------------------------------------------------------------------
# PDF generation
# ---------------------------------------------------------------------------

def to_template_context(invoice: dict[str, Any]) -> SimpleNamespace:
    """Expose the invoice as an object so ``invoice.items`` is the line list.

    A plain dict would make ``invoice.items`` resolve to the dict method in
    Jinja, so attribute access is provided explicitly. ``extras`` stays a
    regular dict and can be iterated with ``invoice.extras.items()``.
    """
    context = dict(invoice)
    context["items"] = [SimpleNamespace(**item) for item in invoice["items"]]
    return SimpleNamespace(**context)


def render_html(template_path: Path, invoice: dict[str, Any], source_file: str) -> str:
    """Fill the Jinja HTML template with the invoice data."""
    try:
        from jinja2 import Environment, FileSystemLoader, StrictUndefined
    except ImportError as exc:  # pragma: no cover - dependency guard
        raise RuntimeError(
            "Не найдена библиотека Jinja2. Установите зависимости: pip install -r requirements.txt"
        ) from exc

    environment = Environment(
        loader=FileSystemLoader(str(TEMPLATE_DIR), encoding="utf-8"),
        autoescape=True,
        undefined=StrictUndefined,
        trim_blocks=True,
        lstrip_blocks=True,
    )
    environment.filters["money"] = fmt_money
    environment.filters["num"] = fmt_quantity
    template = environment.get_template(template_path.name)
    return template.render(
        invoice=to_template_context(invoice),
        source_file=source_file,
        template_name=template_path.name,
    )


def unique_output_path(invoice_id: str, template_stem: str) -> Path:
    """Build a non-clashing path inside ./output for the new PDF."""
    safe_id = re.sub(r"[^\w\-]+", "_", invoice_id).strip("_") or "document"
    safe_stem = re.sub(r"[^\w\-]+", "_", template_stem).strip("_") or "template"
    stem = f"{safe_id}__{safe_stem}"
    candidate = OUTPUT_DIR / f"{stem}.pdf"
    counter = 1
    while candidate.exists():
        candidate = OUTPUT_DIR / f"{stem}_{counter}.pdf"
        counter += 1
    return candidate


def write_pdf(html: str, template_path: Path, target: Path) -> None:
    """Render HTML to PDF with WeasyPrint (fonts resolve via ./assets/fonts)."""
    try:
        from weasyprint import HTML
    except (ImportError, OSError) as exc:
        raise RuntimeError(
            "WeasyPrint не может загрузить нативные библиотеки Pango/GTK.\n"
            "  Windows: установите MSYS2 и выполните `pacman -S mingw-w64-x86_64-weasyprint`\n"
            "             либо добавьте каталог с DLL в PATH.\n"
            "  macOS:   brew install pango\n"
            f"  Исходная ошибка: {exc}"
        ) from exc

    # base_url points at the folder holding the template so that relative
    # font paths such as '../assets/fonts/Roboto-Regular.ttf' resolve
    base_url = template_path.parent.as_uri() + "/"
    HTML(string=html, base_url=base_url).write_pdf(str(target))


def open_pdf(path: Path) -> bool:
    """Open the PDF in the system default application (Windows/macOS/Linux)."""
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif system == "Darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
        return True
    except Exception as exc:  # pragma: no cover - depends on the desktop env
        print(f"  ! Не удалось открыть PDF автоматически: {exc}")
        print(f"    Файл сохранён здесь: {path}")
        return False


# ---------------------------------------------------------------------------
# main flow
# ---------------------------------------------------------------------------

def describe_invoice(invoice: dict[str, Any]) -> str:
    """One line summary of an invoice for the selection list."""
    parts = [invoice["invoice_id"]]
    if invoice["date"]:
        parts.append(invoice["date"])
    client = invoice["customer_name"] or invoice["seller_name"]
    if client:
        parts.append(client)
    parts.append(f"{invoice['items_count']} поз.")
    parts.append(invoice["total_formatted"])
    return " · ".join(parts)


def scan_requirements() -> list[str]:
    """Return warnings about missing optional dependencies."""
    warnings: list[str] = []
    try:
        import pandas  # noqa: F401
    except ImportError:
        warnings.append("pandas не установлен \u2014 CSV будет читаться модулем csv")
    try:
        import jinja2  # noqa: F401
    except ImportError:
        warnings.append("Jinja2 не установлена \u2014 шаблоны не отрендерятся")
    try:
        import weasyprint  # noqa: F401
    except (ImportError, OSError):
        warnings.append("WeasyPrint недоступен \u2014 PDF не создать (нужны библиотеки Pango/GTK)")
    return warnings


def configure_console() -> None:
    """Make console output safe for Cyrillic on Windows terminals.

    A redirected Windows stream uses the locale codepage (cp1251), which has
    no ruble sign and would raise UnicodeEncodeError.  Keeping the native
    encoding but replacing unencodable characters avoids a crash.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="replace")
            except (ValueError, OSError):  # pragma: no cover - odd streams
                pass


def main() -> int:
    configure_console()

    print_header("ГЕНЕРАТОР PDF ИЗ CSV / JSON")
    print(f"  Каталог данных:   {DATA_DIR}")
    print(f"  Каталог шаблонов: {TEMPLATE_DIR}")
    print(f"  Каталог вывода:   {OUTPUT_DIR}")

    for warning in scan_requirements():
        print(f"  ! {warning}")

    data_files = discover_files(DATA_DIR, DATA_SUFFIXES)
    template_files = discover_files(TEMPLATE_DIR, TEMPLATE_SUFFIXES)

    print_subheader("ДОСТУПНЫЕ ФАЙЛЫ С ДАННЫМИ")
    if not data_files:
        print(f"  В каталоге {DATA_DIR} нет файлов .csv или .json.")
        return 1
    print_menu("Файлы данных", [
        f"{p.name}  ({p.suffix[1:].upper()}, {human_size(p)})" for p in data_files
    ])

    print_subheader("ДОСТУПНЫЕ HTML-ШАБЛОНЫ")
    if not template_files:
        print(f"  В каталоге {TEMPLATE_DIR} нет файлов .html.")
        return 1
    print_menu("HTML-шаблоны", [p.name for p in template_files])

    print_header("ВЫБОР ИСТОЧНИКА ДАННЫХ")
    data_index = ask_choice("  Выберите файл с данными [0-{}]: ".format(len(data_files)), len(data_files))
    if data_index is None:
        print("\n  Выход.")
        return 0
    data_path = data_files[data_index]

    print_header("ВЫБОР ШАБЛОНА")
    template_index = ask_choice("  Выберите HTML-шаблон [0-{}]: ".format(len(template_files)), len(template_files))
    if template_index is None:
        print("\n  Выход.")
        return 0
    template_path = template_files[template_index]

    print(f"\n  Читаем «{data_path.name}» ...")
    try:
        invoices = load_invoices(data_path)
    except Exception as exc:
        print(f"  ! Ошибка чтения файла: {exc}")
        return 1

    print_header("ДОСТУПНЫЕ ЧЕКИ (invoice id)")
    if not invoices:
        print(f"  В файле «{data_path.name}» не найдено ни одного чека.")
        print("  Проверьте, что есть колонка/ключ с идентификатором (например, invoice_id).")
        return 1
    print_menu("Чеки", [describe_invoice(inv) for inv in invoices])

    invoice_index = ask_choice("  Выберите чек по invoice id [0-{}]: ".format(len(invoices)), len(invoices))
    if invoice_index is None:
        print("\n  Выход.")
        return 0
    invoice = invoices[invoice_index]

    print_header("ГЕНЕРАЦИЯ PDF")
    print(f"  Чек:      {invoice['invoice_id']}")
    print(f"  Шаблон:   {template_path.name}")

    try:
        html = render_html(template_path, invoice, data_path.name)
    except Exception as exc:
        print(f"  ! Ошибка подстановки данных в шаблон: {exc}")
        return 1

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    target = unique_output_path(invoice["invoice_id"], template_path.stem)
    try:
        write_pdf(html, template_path, target)
    except Exception as exc:
        print(f"  ! Ошибка генерации PDF: {exc}")
        return 1

    print(f"  Готово:   {target}")
    print(f"  Размер:   {human_size(target)}")

    print("\n  Открываем PDF в системной программе ...")
    if open_pdf(target):
        print("  PDF открыт.")
    print(f"\n{LINE}")
    print("  Завершено.")
    print(LINE)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n  Прервано пользователем.")
        sys.exit(130)
    except EOFError:
        print("\n  Нет данных для ввода (нужен интерактивный запуск скрипта).")
        sys.exit(1)
