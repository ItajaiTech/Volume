import os
import re
import json
import threading
from collections import defaultdict
from functools import wraps
from uuid import uuid4

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_from_directory, session, url_for
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

from database import (
    add_box,
    add_product,
    add_shipment_history,
    create_order,
    delete_order_item,
    delete_order_with_dependencies,
    get_box_by_id,
    get_dashboard_stats,
    get_order,
    get_order_items,
    get_packing_rules,
    get_product_by_id,
    get_shipment_history_by_order,
    import_boxes_from_excel,
    import_products_from_excel,
    init_db,
    list_boxes,
    list_orders,
    list_products,
    list_shipment_history,
    normalize_name,
    replace_shipment_history,
    set_order_item_quantity,
    set_box_active,
    update_packing_rules,
    update_box,
    update_product,
)
from learning import suggest_box_from_history
from packing import (
    build_packing_3d_previews,
    calculate_order_totals,
    choose_best_box,
    describe_item_pack_profile,
    estimate_packages_for_box,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "database.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")

os.makedirs(UPLOAD_DIR, exist_ok=True)
init_db(DB_PATH)

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.secret_key = os.getenv("VOLUME_APP_SECRET", "change-this-secret")
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024

ADMIN_USER = os.getenv("VOLUME_ADMIN_USER", "admin")
ADMIN_PASS = os.getenv("VOLUME_ADMIN_PASS", "admin123")
USER_UPLOAD_TEMPLATE = "user_upload.html"
IMPORT_PRODUCTS_TEMPLATE = "import_products.html"
ORDER_NOT_FOUND_MSG = "Pedido nao encontrado."
TOKEN_EDGE_CLEANUP_PATTERN = r"(^[^A-Za-z0-9]+)|([^A-Za-z0-9-]+$)"


def safe_int_env(name, default):
    try:
        return int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


def safe_bool_env(name, default=False):
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def form_checkbox_to_int(form_data, key):
    value = str(form_data.get(key, "")).strip().lower()
    return 1 if value in {"1", "true", "on", "yes"} else 0


APP_PORT = safe_int_env("VOLUME_PORT", 6100)
APP_HTTP_PORT = safe_int_env("VOLUME_HTTP_PORT", 6080)
APP_HOST = os.getenv("VOLUME_BIND_HOST", os.getenv("VOLUME_HOST", "0.0.0.0"))
APP_PUBLIC_HOST = os.getenv("VOLUME_PUBLIC_HOST", "volume.local")
APP_SERVER_NAME = os.getenv(
    "VOLUME_SERVER_NAME",
    APP_PUBLIC_HOST if APP_PORT == 443 else f"{APP_PUBLIC_HOST}:{APP_PORT}",
)
APP_DEBUG = safe_bool_env("VOLUME_DEBUG", True)
APP_URL_SCHEME = os.getenv("VOLUME_URL_SCHEME", "http").strip().lower()
SSL_CERT_FILE = os.getenv("VOLUME_SSL_CERT_FILE", os.path.join(BASE_DIR, "certs", "volume.local.crt"))
SSL_KEY_FILE = os.getenv("VOLUME_SSL_KEY_FILE", os.path.join(BASE_DIR, "certs", "volume.local.key"))
DEFAULT_BOXES_XLSX = os.getenv(
    "VOLUME_BOXES_TEMPLATE",
    os.path.normpath(os.path.join(BASE_DIR, "..", "Caixas.xlsx")),
)

trusted_hosts = []


def add_trusted_host(raw_host):
    host = str(raw_host or "").strip()
    if not host:
        return

    host = re.sub(r"^https?://", "", host, flags=re.IGNORECASE)
    host = host.split("/", 1)[0].strip()
    if not host:
        return

    candidates = {host}
    if ":" in host:
        candidates.add(host.split(":", 1)[0].strip())
    else:
        candidates.add(f"{host}:{APP_PORT}")
        candidates.add(f"{host}:443")
        candidates.add(f"{host}:80")

    for candidate in candidates:
        if candidate and candidate not in trusted_hosts:
            trusted_hosts.append(candidate)


for raw_host in [
    APP_PUBLIC_HOST,
    APP_SERVER_NAME,
    "volume.local",
    ".local",
    "127.0.0.1",
    f"127.0.0.1:{APP_PORT}",
    "localhost",
    f"localhost:{APP_PORT}",
]:
    add_trusted_host(raw_host)

extra_trusted_hosts = os.getenv("VOLUME_TRUSTED_HOSTS", "")
for raw_host in extra_trusted_hosts.split(","):
    add_trusted_host(raw_host)

app.config["SERVER_NAME"] = APP_SERVER_NAME
app.config["TRUSTED_HOSTS"] = trusted_hosts
app.config["PREFERRED_URL_SCHEME"] = APP_URL_SCHEME


@app.route("/favicon.ico", methods=["GET"])
def favicon():
    return send_from_directory(
        os.path.join(app.root_path, "static"),
        "favicon.svg",
        mimetype="image/svg+xml",
    )


def admin_required(view_func):
    @wraps(view_func)
    def _wrapped(*args, **kwargs):
        if not session.get("is_admin"):
            next_url = request.path
            return redirect(url_for("admin_login", next=next_url))
        return view_func(*args, **kwargs)

    return _wrapped


def parse_lines_to_items(lines):
    """
    Accepts both patterns:
    - 10 SSD
    - SSD 10
    """
    parsed = {}

    for raw in lines:
        line = " ".join(str(raw).replace("\t", " ").split())
        if not line:
            continue

        qty = None
        name = None

        qty_first = re.match(r"^(\d+)\s+(.+)$", line)
        qty_last = re.match(r"^(.+?)\s+(\d+)$", line)

        if qty_first:
            qty = int(qty_first.group(1))
            name = qty_first.group(2)
        elif qty_last:
            name = qty_last.group(1)
            qty = int(qty_last.group(2))

        if not name or qty is None or qty <= 0:
            continue

        sku = extract_sku_from_identifier(name)
        key = normalize_name(sku or name)
        if key not in parsed:
            parsed[key] = {
                "sku": sku,
                "product_name": name.strip(),
                "quantity": qty,
            }
        else:
            parsed[key]["quantity"] += qty

    return list(parsed.values())


def looks_like_sku(token):
    cleaned = re.sub(
        TOKEN_EDGE_CLEANUP_PATTERN,
        "",
        str(token or "").strip().upper(),
    )
    digit_count = sum(ch.isdigit() for ch in cleaned)
    return bool(
        cleaned
        and re.match(r"^[A-Z]{2,}[A-Z0-9-]{2,}$", cleaned)
        and digit_count >= 4
        and (cleaned.startswith("PRD") or len(cleaned) == 8)
    )


def _line_has_explicit_sku_reference(text, sku):
    normalized_text = " ".join(str(text or "").replace("\t", " ").split()).strip()
    normalized_sku = str(sku or "").strip().upper()
    if not normalized_text or not normalized_sku:
        return False

    upper_text = normalized_text.upper()
    if upper_text == normalized_sku:
        return True
    if upper_text.startswith(f"{normalized_sku} - "):
        return True

    trailing_match = re.search(r"([A-Za-z0-9-]+)\s*$", normalized_text)
    if not trailing_match:
        return False

    trailing_token = re.sub(
        TOKEN_EDGE_CLEANUP_PATTERN,
        "",
        trailing_match.group(1).strip().upper(),
    )
    return trailing_token == normalized_sku


def extract_sku_from_identifier(value):
    """
    Extracts SKU from text like:
    - SKU123 - Produto X
    - SKU123
    - CPU I3 2120
    """
    text = " ".join(str(value).replace("\t", " ").split()).strip()
    if not text:
        return ""
    if " - " in text:
        prefix = text.split(" - ", 1)[0].strip()
        if looks_like_sku(prefix):
            return prefix

    for token in reversed(text.split()):
        cleaned = re.sub(TOKEN_EDGE_CLEANUP_PATTERN, "", token).strip()
        if looks_like_sku(cleaned) and _line_has_explicit_sku_reference(text, cleaned):
            return cleaned.upper()

    return ""


def extract_sku_from_catalog_name(value):
    """
    In catalog, SKU is stored as prefix in names like:
    SKU123 - Produto X
    """
    text = " ".join(str(value).replace("\t", " ").split()).strip()
    if not text:
        return ""
    if " - " in text:
        prefix = text.split(" - ", 1)[0].strip()
        if looks_like_sku(prefix):
            return prefix
    return extract_sku_from_identifier(text)


def format_product_label(product_row):
    sku = str(product_row["sku"] or "").strip() if "sku" in product_row.keys() else ""
    name = str(product_row["name"] or "").strip()
    return f"{sku} - {name}" if sku else name


def parse_quantity_token(value):
    raw = str(value).strip()
    if not raw:
        return 0
    normalized = raw.replace(".", "").replace(",", ".")
    try:
        qty = int(round(float(normalized)))
    except Exception:
        return 0
    return qty if qty > 0 else 0


def _normalize_line(raw):
    return " ".join(str(raw).replace("\t", " ").split()).strip()


def _append_parsed_invoice_item(parsed, sku, product_name, qty):
    key = normalize_name(sku)
    if key not in parsed:
        parsed[key] = {
            "sku": sku,
            "product_name": product_name or sku,
            "quantity": qty,
        }
        return

    parsed[key]["quantity"] += qty


def _consume_pending_invoice_line(line, pending_items, parsed):
    qty_match = re.match(r"^(\d+(?:[\.,]\d+)?)\s+[A-Za-z]{1,4}\b", line)
    if not (pending_items and qty_match):
        return False, pending_items

    qty = parse_quantity_token(qty_match.group(1))
    pending_item = pending_items.pop(0)
    if qty > 0:
        _append_parsed_invoice_item(
            parsed,
            pending_item["sku"],
            pending_item["product_name"],
            qty,
        )

    return True, pending_items


def _extract_pending_invoice_item(line):
    sku = extract_sku_from_identifier(line)
    if not sku:
        return "", ""

    product_name = line
    sku_pos = product_name.upper().rfind(sku.upper())
    if sku_pos >= 0:
        product_name = product_name[:sku_pos].strip(" -")

    return sku, (product_name or sku)


def parse_invoice_items_from_lines(lines):
    """
    Parses invoice-like PDF layout where product line contains SKU and the next
    line contains quantity, e.g.:
    - SSD ... PRD00040
    - 20,00 UN 211,46 4.229,20
    """
    parsed = {}
    pending_items = []

    for raw in lines:
        line = _normalize_line(raw)
        if not line:
            continue

        consumed, pending_items = _consume_pending_invoice_line(
            line,
            pending_items,
            parsed,
        )
        if consumed:
            continue

        next_sku, next_name = _extract_pending_invoice_item(line)
        if next_sku:
            pending_items.append(
                {
                    "sku": next_sku,
                    "product_name": next_name,
                }
            )

    return list(parsed.values())


def parse_items_from_pdf(file_path):
    try:
        from pypdf import PdfReader
    except Exception as exc:
        raise RuntimeError(
            "Leitura de PDF requer pypdf. Execute: pip install pypdf"
        ) from exc

    reader = PdfReader(file_path)
    lines = []

    for page in reader.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            line = line.strip()
            if line:
                lines.append(line)

    items = parse_invoice_items_from_lines(lines)
    if not items:
        items = parse_lines_to_items(lines)

    if not items:
        raise ValueError(
            "Nenhuma linha valida encontrada no PDF. Use linhas como 'SKU123 10', '10 SKU123' ou o layout de pedido com SKU e Qtd."
        )
    return items


def _parse_uploaded_excel_quantity(value):
    try:
        return int(float(value))
    except Exception:
        return 0


def _extract_uploaded_excel_item(row, has_sku):
    sku = str(row.get("sku", "")).strip() if has_sku else ""
    name = str(row.get("product_name", "")).strip()
    qty = _parse_uploaded_excel_quantity(row.get("quantity", 0))

    if qty <= 0 or (not name and not sku):
        return None

    return {
        "sku": sku,
        "product_name": name or sku,
        "quantity": qty,
    }


def _accumulate_uploaded_item(parsed, item):
    key = normalize_name(item["sku"] or item["product_name"])
    if key not in parsed:
        parsed[key] = dict(item)
        return

    parsed[key]["quantity"] += item["quantity"]


def parse_items_from_excel(file_path):
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip().lower() for c in df.columns]

    if "product_name" not in df.columns or "quantity" not in df.columns:
        raise ValueError("A planilha deve conter as colunas: product_name e quantity")

    parsed = {}
    has_sku = "sku" in df.columns

    for _, row in df.iterrows():
        item = _extract_uploaded_excel_item(row, has_sku)
        if not item:
            continue

        _accumulate_uploaded_item(parsed, item)

    if not parsed:
        raise ValueError("Nenhuma linha valida encontrada na planilha")

    return list(parsed.values())


def _catalog_sku_key(product):
    product_sku = str(product["sku"] or "").strip() if "sku" in product.keys() else ""
    return normalize_name(product_sku or extract_sku_from_catalog_name(product["name"]))


def _build_product_catalog_maps(products):
    catalog_by_name = {}
    catalog_by_sku = {}

    for product in products:
        catalog_by_name[normalize_name(product["name"])] = product
        sku_key = _catalog_sku_key(product)
        if sku_key and sku_key not in catalog_by_sku:
            catalog_by_sku[sku_key] = product

    return catalog_by_name, catalog_by_sku


def _match_product_from_catalog(raw_sku, raw_name, catalog_by_name, catalog_by_sku):
    sku_key = normalize_name(raw_sku)
    name_key = normalize_name(raw_name)

    if not sku_key and raw_name:
        extracted_sku = extract_sku_from_identifier(raw_name)
        sku_key = normalize_name(extracted_sku)

    if sku_key and sku_key in catalog_by_sku:
        return catalog_by_sku[sku_key], "sku_exato"
    if name_key in catalog_by_name:
        return catalog_by_name[name_key], "nome_exato"
    if not sku_key:
        return None, "nenhum"

    prefix = f"{sku_key} - "
    for catalog_name, candidate in catalog_by_name.items():
        if catalog_name.startswith(prefix):
            return candidate, "sku_prefixo"

    return None, "nenhum"


def _build_unknown_item(raw_sku, raw_name, qty):
    return {
        "sku": raw_sku,
        "product_name": raw_name,
        "quantity": qty,
    }


def _build_import_debug_row(raw_sku, raw_name, qty, product, match_source):
    return {
        "sku": raw_sku,
        "product_name": raw_name,
        "quantity": qty,
        "mapped_name": format_product_label(product) if product else "",
        "status": "mapeado" if product else "nao_encontrado",
        "match_source": match_source,
    }


def map_uploaded_items_to_catalog(uploaded_items, include_debug=False):
    catalog_by_name, catalog_by_sku = _build_product_catalog_maps(list_products(DB_PATH))

    order_map = defaultdict(int)
    unknown = []
    import_debug = []

    for item in uploaded_items:
        raw_sku = str(item.get("sku", "")).strip()
        raw_name = str(item.get("product_name", "")).strip()
        qty = int(item["quantity"])

        product, match_source = _match_product_from_catalog(
            raw_sku,
            raw_name,
            catalog_by_name,
            catalog_by_sku,
        )

        if product:
            order_map[int(product["id"])] += qty
        else:
            unknown.append(_build_unknown_item(raw_sku, raw_name, qty))

        if include_debug:
            import_debug.append(
                _build_import_debug_row(raw_sku, raw_name, qty, product, match_source)
            )

    order_items = [
        {"product_id": pid, "quantity": qty}
        for pid, qty in sorted(order_map.items(), key=lambda x: x[0])
    ]
    if include_debug:
        return order_items, unknown, import_debug
    return order_items, unknown


def _safe_int_value(value, default):
    try:
        return int(float(value))
    except Exception:
        return int(default)


def _safe_float_value(value, default):
    try:
        return float(value)
    except Exception:
        return float(default)


def _row_value_safe(row, key, default=""):
    if hasattr(row, "keys") and key in row.keys():
        return row[key]
    if isinstance(row, dict) and key in row:
        return row[key]
    return default


def _normalized_item_text(item_row):
    name = str(_row_value_safe(item_row, "name", "")).strip().lower()
    sku = str(_row_value_safe(item_row, "sku", "")).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", f"{name} {sku}")
    return f" {text.strip()} "


def _is_mb_order_item(item_row):
    text = _normalized_item_text(item_row)
    return (
        " mb " in text
        or " motherboard " in text
        or " mainboard " in text
        or " placa mae " in text
    )


def _is_ssdm2_order_item(item_row):
    text = _normalized_item_text(item_row).replace("m 2", "m2")
    has_m2 = " m2 " in text or " nvme " in text
    return " ssd " in text and has_m2


def _is_mb_only_order(order_items):
    has_positive_qty = False
    for item in order_items:
        qty = _safe_int_value(_row_value_safe(item, "quantity", 0), 0)
        if qty <= 0:
            continue
        has_positive_qty = True
        if not _is_mb_order_item(item):
            return False
    return has_positive_qty


def _find_box_exact(boxes_rows, hint):
    for box in boxes_rows:
        box_name = str(box["name"]).strip().lower()
        if box_name == hint:
            return box
    return None


def _find_box_contains(boxes_rows, hint):
    for box in boxes_rows:
        box_name = str(box["name"]).strip().lower()
        if hint in box_name:
            return box
    return None


def _resolve_box_by_hint(boxes_rows, name_hint, contains_hint=None):
    hint = str(name_hint or "").strip().lower()
    if hint:
        exact_match = _find_box_exact(boxes_rows, hint)
        if exact_match:
            return exact_match

        contains_match = _find_box_contains(boxes_rows, hint)
        if contains_match:
            return contains_match

    contains_value = str(contains_hint or "").strip().lower()
    if contains_value:
        return _find_box_contains(boxes_rows, contains_value)

    return None


def _resolve_mb_box(boxes_rows, name_hint):
    box = _resolve_box_by_hint(boxes_rows, name_hint, contains_hint="caixa mb")
    if box:
        return box

    for box in boxes_rows:
        if "caixa mb" in str(box["name"]).strip().lower():
            return box

    return None


def _resolve_ssdm2_master_box(boxes_rows, name_hint):
    box = _resolve_box_by_hint(boxes_rows, name_hint, contains_hint="caixa nvme")
    if box:
        return box

    for box in boxes_rows:
        box_name = str(box["name"]).strip().lower()
        if "nvme" in box_name:
            return box

    return None


def _mark_default_in_tested_boxes(tested_boxes, box_name, packages_required, note, status):
    rows = list(tested_boxes or [])
    target = str(box_name or "").strip().lower()
    for row in rows:
        if str(row.get("name", "")).strip().lower() != target:
            continue
        row["packages_required"] = packages_required
        row["status"] = status
        row["note"] = note
        return rows

    rows.append(
        {
            "name": box_name,
            "packages_required": packages_required,
            "status": status,
            "note": note,
        }
    )
    return rows


def _mark_mb_default_in_tested_boxes(tested_boxes, mb_box_name, packages_required, note):
    return _mark_default_in_tested_boxes(
        tested_boxes,
        mb_box_name,
        packages_required,
        note,
        "padrao_mb",
    )


def _mark_ssdm2_default_in_tested_boxes(tested_boxes, box_name, packages_required, note):
    return _mark_default_in_tested_boxes(
        tested_boxes,
        box_name,
        packages_required,
        note,
        "padrao_ssdm2_master",
    )


def _resolve_mb_context(order_items, boxes, packing_rules):
    mb_min_qty = max(1, _safe_int_value(packing_rules.get("mb_bundle_min_qty", 21), 21))
    mb_total_qty = sum(
        int(item["quantity"])
        for item in order_items
        if int(item["quantity"]) > 0 and _is_mb_order_item(item)
    )
    if mb_total_qty < mb_min_qty or not boxes:
        return None

    mb_box = _resolve_mb_box(
        boxes,
        packing_rules.get("mb_default_box_name", "Caixa MB"),
    )
    if not mb_box:
        return None

    return {
        "box": mb_box,
        "min_qty": mb_min_qty,
        "total_qty": mb_total_qty,
    }


def _resolve_ssdm2_context(order_items, boxes, packing_rules):
    ssdm2_min_qty = max(
        1,
        _safe_int_value(packing_rules.get("ssdm2_master_box_qty", 250), 250),
    )
    ssdm2_total_qty = sum(
        int(item["quantity"])
        for item in order_items
        if int(item["quantity"]) > 0 and _is_ssdm2_order_item(item)
    )
    if ssdm2_total_qty < ssdm2_min_qty or not boxes:
        return None

    ssdm2_box = _resolve_ssdm2_master_box(
        boxes,
        packing_rules.get("ssdm2_default_box_name", "Caixa NVME"),
    )
    if not ssdm2_box:
        return None

    return {
        "box": ssdm2_box,
        "min_qty": ssdm2_min_qty,
        "total_qty": ssdm2_total_qty,
    }


def _build_effective_rules(packing_rules, mb_context):
    rules = dict(packing_rules)
    mb_box_dims = None
    if not mb_context:
        return rules, mb_box_dims

    mb_box = mb_context["box"]
    mb_box_dims = {
        "length_cm": float(mb_box["length_cm"]),
        "width_cm": float(mb_box["width_cm"]),
        "height_cm": float(mb_box["height_cm"]),
    }
    rules["mb_box_length_cm"] = mb_box_dims["length_cm"]
    rules["mb_box_width_cm"] = mb_box_dims["width_cm"]
    rules["mb_box_height_cm"] = mb_box_dims["height_cm"]
    return rules, mb_box_dims


def _resolve_packaging_unit_weight(box_row, packing_rules):
    if not box_row:
        return 0.0

    raw_value = _safe_float_value(_row_value_safe(box_row, "max_weight", 0), 0.0)
    if raw_value <= 0:
        return 0.0

    # max_weight can represent carrying capacity in some templates.
    min_effective_weight = _safe_float_value(
        packing_rules.get("min_effective_max_weight_kg", 1.0),
        1.0,
    )
    if raw_value > max(0.0, min_effective_weight):
        return 0.0

    return raw_value


def _with_weight_breakdown(totals, box_row, packages_required, packing_rules):
    normalized = dict(totals or {})
    items_weight = _safe_float_value(normalized.get("total_weight", 0.0), 0.0)
    packages = max(0, _safe_int_value(packages_required, 0))
    package_unit_weight = _resolve_packaging_unit_weight(box_row, packing_rules)
    packages_weight = package_unit_weight * packages

    normalized["total_weight_items"] = items_weight
    normalized["total_weight_packages"] = packages_weight
    normalized["total_weight_with_packages"] = items_weight + packages_weight
    normalized["package_weight_per_box"] = package_unit_weight
    return normalized


def _empty_recommendation(source, totals, mb_box_dims, packing_rules):
    return {
        "box": None,
        "packages_required": None,
        "confidence": 0,
        "source": source,
        "evidence_count": 0,
        "totals": _with_weight_breakdown(totals, None, 0, packing_rules),
        "tested_boxes": [],
        "unpack_applied": False,
        "unpack_plan": {},
        "mb_box_dims": mb_box_dims,
    }


def _apply_mb_default_rule(algo, totals, order_items, mb_context, packing_rules):
    if not mb_context:
        return algo, False

    mb_box = mb_context["box"]
    mb_fill_factor = 1.0 if _is_mb_only_order(order_items) else 0.9
    mb_estimate = estimate_packages_for_box(
        totals["total_volume_cm3"],
        totals["total_weight"],
        mb_box,
        fill_factor=mb_fill_factor,
        order_items=order_items,
        packing_rules=packing_rules,
    )
    if not mb_estimate:
        return algo, False

    selected = dict(algo)
    selected["box"] = mb_box
    selected["packages_required"] = mb_estimate["packages_required"]
    selected["unpack_applied"] = bool(mb_estimate.get("unpack_applied"))
    selected["unpack_plan"] = dict(mb_estimate.get("unpack_plan") or {})
    selected["reason"] = "algorithm_mb_default_pack20"
    selected["confidence"] = 72

    rule_note = (
        f"Regra MB: mais de {mb_context['min_qty'] - 1} und usa {mb_box['name']} como padrao."
    )
    selected["tested_boxes"] = _mark_mb_default_in_tested_boxes(
        algo.get("tested_boxes", []),
        mb_box["name"],
        selected["packages_required"],
        rule_note,
    )
    return selected, True


def _apply_ssdm2_default_rule(algo, totals, order_items, ssdm2_context, packing_rules):
    if not ssdm2_context:
        return algo, False

    ssdm2_box = ssdm2_context["box"]
    ssdm2_estimate = estimate_packages_for_box(
        totals["total_volume_cm3"],
        totals["total_weight"],
        ssdm2_box,
        order_items=order_items,
        packing_rules=packing_rules,
    )
    if not ssdm2_estimate:
        return algo, False

    selected = dict(algo)
    selected["box"] = ssdm2_box
    selected["packages_required"] = ssdm2_estimate["packages_required"]
    selected["unpack_applied"] = bool(ssdm2_estimate.get("unpack_applied"))
    selected["unpack_plan"] = dict(ssdm2_estimate.get("unpack_plan") or {})
    selected["reason"] = "algorithm_ssdm2_master_box"
    selected["confidence"] = 74

    rule_note = (
        f"Regra SSD M.2 master: mais de {ssdm2_context['min_qty'] - 1} und usa "
        f"{ssdm2_box['name']} como padrao."
    )
    selected["tested_boxes"] = _mark_ssdm2_default_in_tested_boxes(
        algo.get("tested_boxes", []),
        ssdm2_box["name"],
        selected["packages_required"],
        rule_note,
    )
    return selected, True


def build_recommendation(order_items, packing_rules=None):
    if packing_rules is None:
        packing_rules = get_packing_rules(DB_PATH)

    boxes = list_boxes(DB_PATH)
    mb_context = _resolve_mb_context(order_items, boxes, packing_rules)
    ssdm2_context = _resolve_ssdm2_context(order_items, boxes, packing_rules)
    effective_rules, mb_box_dims = _build_effective_rules(packing_rules, mb_context)
    totals = calculate_order_totals(order_items, packing_rules=effective_rules)

    if not boxes:
        return _empty_recommendation("no_boxes", totals, mb_box_dims, effective_rules)

    algo = choose_best_box(
        totals["total_volume_cm3"],
        totals["total_weight"],
        boxes,
        order_items=order_items,
        packing_rules=effective_rules,
    )
    if not algo:
        return _empty_recommendation(
            "no_compatible_box",
            totals,
            mb_box_dims,
            effective_rules,
        )

    algo, mb_default_applied = _apply_mb_default_rule(
        algo,
        totals,
        order_items,
        mb_context,
        effective_rules,
    )
    ssdm2_default_applied = False
    if not mb_default_applied:
        algo, ssdm2_default_applied = _apply_ssdm2_default_rule(
            algo,
            totals,
            order_items,
            ssdm2_context,
            effective_rules,
        )

    source = algo["reason"]
    confidence = algo["confidence"]
    evidence_count = 0

    if not mb_default_applied and not ssdm2_default_applied:
        history = suggest_box_from_history(DB_PATH, order_items)
        if history and int(history["box_id"]) == int(algo["box"]["id"]):
            source = history["source"]
            confidence = history["confidence"]
            evidence_count = history["evidence_count"]

    totals_with_breakdown = _with_weight_breakdown(
        totals,
        algo.get("box"),
        algo.get("packages_required"),
        effective_rules,
    )

    return {
        "box": algo["box"],
        "packages_required": algo["packages_required"],
        "confidence": confidence,
        "source": source,
        "evidence_count": evidence_count,
        "totals": totals_with_breakdown,
        "tested_boxes": list(algo.get("tested_boxes", [])),
        "unpack_applied": bool(algo.get("unpack_applied")),
        "unpack_plan": dict(algo.get("unpack_plan") or {}),
        "mb_box_dims": mb_box_dims,
    }


def _attach_pack_profile_to_items(items, packing_rules):
    enriched_items = []
    for item in items:
        row = dict(item)
        profile = describe_item_pack_profile(item, packing_rules=packing_rules)
        row["pack_profile"] = profile["profile"]
        row["pack_profile_label"] = profile["label"]
        row["pack_profile_detail"] = profile["detail"]
        row["pack_profile_uses_pack"] = profile["uses_pack"]
        enriched_items.append(row)
    return enriched_items


def _packing_preview_signature_placements(preview):
    placements = []
    for placement in preview.get("placements") or []:
        placements.append(
            (
                str(placement.get("label") or ""),
                round(float(placement.get("x") or 0.0), 4),
                round(float(placement.get("y") or 0.0), 4),
                round(float(placement.get("z") or 0.0), 4),
                round(float(placement.get("length_cm") or 0.0), 4),
                round(float(placement.get("width_cm") or 0.0), 4),
                round(float(placement.get("height_cm") or 0.0), 4),
            )
        )
    return tuple(placements)


def _packing_preview_signature_legend(preview):
    legend = []
    for item in preview.get("legend") or []:
        legend.append(
            (
                str(item.get("label") or ""),
                str(item.get("color") or ""),
                int(item.get("blocks") or 0),
                int(item.get("original_quantity") or 0),
            )
        )
    return tuple(legend)


def _packing_preview_signature(preview):
    if not preview:
        return None

    box = preview.get("box") or {}

    return (
        str(box.get("name") or ""),
        round(float(box.get("length_cm") or 0.0), 4),
        round(float(box.get("width_cm") or 0.0), 4),
        round(float(box.get("height_cm") or 0.0), 4),
        _packing_preview_signature_placements(preview),
        _packing_preview_signature_legend(preview),
        int(preview.get("total_blocks") or 0),
        int(preview.get("placed_blocks") or 0),
        int(preview.get("hidden_blocks") or 0),
        int(preview.get("overflow_blocks") or 0),
        round(float(preview.get("fill_percent") or 0.0), 2),
        tuple(preview.get("notes") or []),
        bool(preview.get("unpack_applied")),
    )


def _group_packing_previews(packing_previews):
    grouped = []
    grouped_map = {}

    for preview in packing_previews or []:
        signature = _packing_preview_signature(preview)
        if signature in grouped_map:
            grouped_map[signature]["count"] += 1
            grouped_map[signature]["volume_indexes"].append(
                int(preview.get("volume_index") or (len(grouped_map[signature]["volume_indexes"]) + 1))
            )
            continue

        group = {
            "preview": preview,
            "count": 1,
            "volume_indexes": [int(preview.get("volume_index") or 1)],
        }
        grouped_map[signature] = group
        grouped.append(group)

    return grouped


def _manual_preview_empty_entry(box_row, note=None):
    box_length = float(box_row["length_cm"])
    box_width = float(box_row["width_cm"])
    box_height = float(box_row["height_cm"])
    notes = []
    if note:
        notes.append(str(note))

    return {
        "box": {
            "name": str(box_row["name"]),
            "length_cm": box_length,
            "width_cm": box_width,
            "height_cm": box_height,
        },
        "legend": [],
        "placements": [],
        "total_blocks": 0,
        "placed_blocks": 0,
        "hidden_blocks": 0,
        "overflow_blocks": 0,
        "used_volume_cm3": 0.0,
        "fill_percent": 0.0,
        "notes": notes,
        "unpack_applied": False,
        "volume_index": 1,
        "volume_count": 1,
    }


def _shipment_assignments_to_items(shipment_row, items_by_product_id):
    assignments_json = shipment_row["assignments_json"] if "assignments_json" in shipment_row.keys() else None
    if not assignments_json:
        return []

    try:
        assignments = json.loads(assignments_json)
    except Exception:
        return []

    if not isinstance(assignments, dict):
        return []

    assigned_items = []
    for product_id_raw, qty_raw in assignments.items():
        try:
            product_id = int(product_id_raw)
            qty = int(qty_raw)
        except (TypeError, ValueError):
            continue

        if qty <= 0:
            continue

        source_item = items_by_product_id.get(product_id)
        if not source_item:
            continue

        item_copy = dict(source_item)
        item_copy["quantity"] = qty
        assigned_items.append(item_copy)

    return assigned_items


def _build_manual_shipment_previews(shipment_rows, order_items, packing_rules):
    if not shipment_rows:
        return []

    items_by_product_id = {}
    for item in order_items or []:
        try:
            items_by_product_id[int(item["product_id"])] = item
        except Exception:
            continue

    expanded_shipments = []
    for shipment in shipment_rows:
        try:
            box_id = int(shipment["box_id"])
        except Exception:
            continue

        quantity = 1
        try:
            quantity = max(1, int(shipment["quantity"]))
        except Exception:
            quantity = 1

        box_row = get_box_by_id(DB_PATH, box_id)
        if not box_row:
            continue

        assigned_items = _shipment_assignments_to_items(shipment, items_by_product_id)
        for _ in range(quantity):
            expanded_shipments.append(
                {
                    "box": box_row,
                    "assigned_items": assigned_items,
                }
            )

    total_volumes = len(expanded_shipments)
    if total_volumes <= 0:
        return []

    previews = []
    for index, entry in enumerate(expanded_shipments, start=1):
        box_row = entry["box"]
        assigned_items = entry["assigned_items"]

        preview_entry = None
        if assigned_items:
            generated = build_packing_3d_previews(
                assigned_items,
                box_row,
                packages_required=1,
                packing_rules=packing_rules,
            )
            if generated:
                preview_entry = dict(generated[0])

        if not preview_entry:
            fallback_note = (
                "Plano manual sem distribuicao detalhada de itens para esta embalagem."
                if not assigned_items
                else "Nao foi possivel simular os itens atribuidos nesta embalagem."
            )
            preview_entry = _manual_preview_empty_entry(box_row, note=fallback_note)

        preview_entry["volume_index"] = index
        preview_entry["volume_count"] = total_volumes
        previews.append(preview_entry)

    return previews


@app.route("/", methods=["GET"])
def index():
    return redirect(url_for("user_upload"))


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        if username == ADMIN_USER and password == ADMIN_PASS:
            session["is_admin"] = True
            session["admin_user"] = username
            flash("Login de administrador realizado com sucesso.", "success")
            return redirect(request.args.get("next") or url_for("admin_dashboard"))

        flash("Credenciais de administrador invalidas.", "error")

    return render_template("admin_login.html")


@app.route("/admin/logout", methods=["GET"])
def admin_logout():
    session.clear()
    flash("Voce saiu do sistema.", "info")
    return redirect(url_for("admin_login"))


@app.route("/admin", methods=["GET"])
@admin_required
def admin_dashboard():
    stats = get_dashboard_stats(DB_PATH)
    packing_rules = get_packing_rules(DB_PATH)
    return render_template("dashboard.html", stats=stats, packing_rules=packing_rules)


@app.route("/admin/settings/packing", methods=["POST"])
@admin_required
def admin_packing_settings():
    rules_input = {
        "ssd25_bundle_qty": request.form.get("ssd25_bundle_qty", "10"),
        "ssd25_bundle_length_cm": request.form.get("ssd25_bundle_length_cm", "22"),
        "ssd25_bundle_width_cm": request.form.get("ssd25_bundle_width_cm", "8"),
        "ssd25_bundle_height_cm": request.form.get("ssd25_bundle_height_cm", "6.7"),
        "ssdm2_bundle_qty": request.form.get("ssdm2_bundle_qty", "10"),
        "ssdm2_bundle_length_cm": request.form.get("ssdm2_bundle_length_cm", "21.3"),
        "ssdm2_bundle_width_cm": request.form.get("ssdm2_bundle_width_cm", "5"),
        "ssdm2_bundle_height_cm": request.form.get("ssdm2_bundle_height_cm", "9"),
        "ssdm2_master_box_qty": request.form.get("ssdm2_master_box_qty", "250"),
        "ssdm2_default_box_name": request.form.get("ssdm2_default_box_name", "Caixa NVME"),
        "ram_nb_bundle_qty": request.form.get("ram_nb_bundle_qty", "10"),
        "ram_nb_bundle_min_qty": request.form.get("ram_nb_bundle_min_qty", "10"),
        "ram_desk_bundle_qty": request.form.get("ram_desk_bundle_qty", "10"),
        "ram_desk_bundle_min_qty": request.form.get("ram_desk_bundle_min_qty", "10"),
        "min_effective_max_weight_kg": request.form.get("min_effective_max_weight_kg", "1"),
    }

    update_packing_rules(DB_PATH, rules_input)
    flash("Regras de empacotamento atualizadas com sucesso.", "success")
    return redirect(url_for("admin_dashboard"))


@app.route("/products", methods=["GET", "POST"])
@admin_required
def products():
    if request.method == "POST":
        try:
            add_product(
                DB_PATH,
                request.form.get("sku", ""),
                request.form.get("name", ""),
                request.form.get("length_cm", 0),
                request.form.get("width_cm", 0),
                request.form.get("height_cm", 0),
                request.form.get("weight", 0),
            )
            flash("Produto cadastrado com sucesso.", "success")
        except Exception as exc:
            flash(f"Nao foi possivel cadastrar produto: {exc}", "error")

    # Obter parâmetros de ordenação da query string
    sort_by = request.args.get("sort_by", "name")
    order = request.args.get("order", "ASC")
    
    rows = list_products(DB_PATH, sort_by=sort_by, order=order)
    
    # Passar informações de ordenação para o template
    return render_template(
        "products.html",
        products=rows,
        current_sort_by=sort_by,
        current_order=order
    )


@app.route("/boxes", methods=["GET", "POST"])
@admin_required
def boxes():
    if request.method == "POST":
        try:
            add_box(
                DB_PATH,
                request.form.get("name", ""),
                request.form.get("length_cm", 0),
                request.form.get("width_cm", 0),
                request.form.get("height_cm", 0),
                request.form.get("max_weight", 0),
                is_active=form_checkbox_to_int(request.form, "is_active"),
            )
            flash("Embalagem cadastrada com sucesso.", "success")
        except Exception as exc:
            flash(f"Nao foi possivel cadastrar embalagem: {exc}", "error")

    rows = list_boxes(DB_PATH, active_only=False)
    return render_template("boxes.html", boxes=rows)


@app.route("/boxes/reimport-default", methods=["POST"])
@admin_required
def boxes_reimport_default():
    if not os.path.exists(DEFAULT_BOXES_XLSX):
        flash(
            f"Arquivo de caixas nao encontrado: {DEFAULT_BOXES_XLSX}",
            "error",
        )
        return redirect(url_for("boxes"))

    try:
        imported, skipped, errors = import_boxes_from_excel(DB_PATH, DEFAULT_BOXES_XLSX)
        flash(
            f"Embalagens sincronizadas: {imported}. Linhas ignoradas: {skipped}.",
            "success",
        )
        if errors:
            preview = "; ".join(errors[:3])
            if len(errors) > 3:
                preview += f"; ... e mais {len(errors) - 3} erro(s)"
            flash(f"Avisos na importacao: {preview}", "warning")
    except Exception as exc:
        flash(f"Falha ao reimportar embalagens: {exc}", "error")

    return redirect(url_for("boxes"))


@app.route("/product/<int:product_id>/edit", methods=["GET", "POST"])
@admin_required
def product_edit(product_id):
    product = get_product_by_id(DB_PATH, product_id)
    if not product:
        flash("Produto nao encontrado.", "error")
        return redirect(url_for("products"))

    if request.method == "POST":
        try:
            update_product(
                DB_PATH,
                product_id,
                request.form.get("sku", ""),
                request.form.get("name", ""),
                request.form.get("length_cm", 0),
                request.form.get("width_cm", 0),
                request.form.get("height_cm", 0),
                request.form.get("weight", 0),
            )
            flash("Produto atualizado com sucesso.", "success")
            return redirect(url_for("products"))
        except Exception as exc:
            flash(f"Nao foi possivel atualizar produto: {exc}", "error")

    return render_template("product_edit.html", product=product)


@app.route("/box/<int:box_id>/edit", methods=["GET", "POST"])
@admin_required
def box_edit(box_id):
    box = get_box_by_id(DB_PATH, box_id)
    if not box:
        flash("Embalagem nao encontrada.", "error")
        return redirect(url_for("boxes"))

    if request.method == "POST":
        try:
            update_box(
                DB_PATH,
                box_id,
                request.form.get("name", ""),
                request.form.get("length_cm", 0),
                request.form.get("width_cm", 0),
                request.form.get("height_cm", 0),
                request.form.get("max_weight", 0),
                is_active=form_checkbox_to_int(request.form, "is_active"),
            )
            flash("Embalagem atualizada com sucesso.", "success")
            return redirect(url_for("boxes"))
        except Exception as exc:
            flash(f"Nao foi possivel atualizar embalagem: {exc}", "error")

    return render_template("box_edit.html", box=box)


@app.route("/box/<int:box_id>/toggle-active", methods=["POST"])
@admin_required
def box_toggle_active(box_id):
    box = get_box_by_id(DB_PATH, box_id)
    if not box:
        flash("Embalagem nao encontrada.", "error")
        return redirect(url_for("boxes"))

    next_is_active = 0 if int(box["is_active"] or 0) else 1
    set_box_active(DB_PATH, box_id, next_is_active)

    if next_is_active:
        flash("Embalagem ativada com sucesso.", "success")
    else:
        flash("Embalagem desativada com sucesso.", "info")

    return redirect(url_for("boxes"))


def _render_import_products_template():
    return render_template(IMPORT_PRODUCTS_TEMPLATE)


def _extract_products_import_file():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return None, "Nenhum arquivo foi selecionado."
    if not uploaded.filename.lower().endswith((".xlsx", ".xls")):
        return None, "Por favor, selecione um arquivo Excel (.xlsx ou .xls)."
    return uploaded, ""


def _build_products_column_mapping(form_data):
    return {
        "sku": form_data.get("col_sku", "").strip() or None,
        "name": form_data.get("col_name", "Descrição do produto"),
        "length": form_data.get("col_length", "Comprimento"),
        "width": form_data.get("col_width", "Largura"),
        "height": form_data.get("col_height", "altura"),
        "weight": form_data.get("col_weight", "Peso"),
    }


def _flash_products_import_result(imported, errors, error_messages):
    if imported > 0:
        flash(f"Sucesso! {imported} produto(s) importado(s).", "success")

    if errors <= 0:
        return

    error_text = "; ".join(error_messages[:5])
    if len(error_messages) > 5:
        error_text += f"; ... e mais {len(error_messages) - 5} erro(s)"
    flash(f"Aviso: {errors} linhas com erro. {error_text}", "warning")


@app.route("/products/import", methods=["GET", "POST"])
@admin_required
def import_products_route():
    if request.method != "POST":
        return _render_import_products_template()

    uploaded_file, validation_error = _extract_products_import_file()
    if validation_error:
        flash(validation_error, "error")
        return _render_import_products_template()

    file_path = os.path.join(UPLOAD_DIR, secure_filename(uploaded_file.filename))
    try:
        uploaded_file.save(file_path)
        column_mapping = _build_products_column_mapping(request.form)
        imported, errors, error_messages = import_products_from_excel(
            DB_PATH,
            file_path,
            column_mapping,
        )
        _flash_products_import_result(imported, errors, error_messages)
        return redirect(url_for("products"))

    except Exception as exc:
        flash(f"Erro ao importar arquivo: {exc}", "error")
        return _render_import_products_template()

    finally:
        if os.path.exists(file_path):
            os.remove(file_path)


@app.route("/orders", methods=["GET"])
@admin_required
def orders():
    rows = list_orders(DB_PATH)
    return render_template("orders.html", orders=rows)


@app.route("/order/new", methods=["GET", "POST"])
@admin_required
def order_new():
    products_rows = list_products(DB_PATH)

    if request.method == "POST":
        items = []
        for p in products_rows:
            qty_raw = request.form.get(f"qty_{p['id']}", "0").strip() or "0"
            try:
                qty = int(qty_raw)
            except ValueError:
                qty = 0

            if qty > 0:
                items.append({"product_id": int(p["id"]), "quantity": qty})

        if not items:
            flash("Informe ao menos uma quantidade de produto.", "error")
            return render_template("order_new.html", products=products_rows)

        order_id = create_order(DB_PATH, items)
        flash(f"Pedido #{order_id} criado com sucesso.", "success")
        return redirect(url_for("order_detail", id=order_id))

    return render_template("order_new.html", products=products_rows)


@app.route("/order/<int:id>", methods=["GET"])
@admin_required
def order_detail(id):
    order = get_order(DB_PATH, id)
    if not order:
        flash(ORDER_NOT_FOUND_MSG, "error")
        return redirect(url_for("orders"))

    items = get_order_items(DB_PATH, id)
    products_rows = list_products(DB_PATH)
    shipment_rows = get_shipment_history_by_order(DB_PATH, id)
    packing_rules = get_packing_rules(DB_PATH)
    recommendation = build_recommendation(items, packing_rules=packing_rules)
    effective_rules = dict(packing_rules)
    if recommendation.get("mb_box_dims"):
        dims = recommendation["mb_box_dims"]
        effective_rules["mb_box_length_cm"] = dims["length_cm"]
        effective_rules["mb_box_width_cm"] = dims["width_cm"]
        effective_rules["mb_box_height_cm"] = dims["height_cm"]

    items_for_view = _attach_pack_profile_to_items(items, effective_rules)
    boxes_rows = list_boxes(DB_PATH)
    packing_previews = []
    if recommendation.get("box"):
        packing_previews = build_packing_3d_previews(
            items,
            recommendation["box"],
            packages_required=recommendation.get("packages_required", 1),
            packing_rules=effective_rules,
            unpack_plan=recommendation.get("unpack_plan"),
        )
    if not packing_previews and shipment_rows:
        packing_previews = _build_manual_shipment_previews(
            shipment_rows,
            items,
            effective_rules,
        )
    packing_preview = packing_previews[0] if packing_previews else None
    grouped_packing_previews = _group_packing_previews(packing_previews)

    return render_template(
        "order_detail.html",
        order=order,
        items=items_for_view,
        recommendation=recommendation,
        packing_preview=packing_preview,
        packing_previews=packing_previews,
        grouped_packing_previews=grouped_packing_previews,
        boxes=boxes_rows,
        products=products_rows,
        shipments=shipment_rows,
        user_mode=False,
        unknown_items=[],
        import_debug=[],
    )


@app.route("/order/<int:order_id>/item", methods=["POST"])
@admin_required
def order_item_upsert(order_id):
    order = get_order(DB_PATH, order_id)
    if not order:
        flash(ORDER_NOT_FOUND_MSG, "error")
        return redirect(url_for("orders"))

    product_id = request.form.get("product_id", "").strip()
    quantity = request.form.get("quantity", "").strip()

    if not product_id.isdigit():
        flash("Selecione um produto valido.", "error")
        return redirect(url_for("order_detail", id=order_id))

    try:
        quantity_value = int(quantity)
    except Exception:
        flash("Informe uma quantidade inteira valida.", "error")
        return redirect(url_for("order_detail", id=order_id))

    product = get_product_by_id(DB_PATH, int(product_id))
    if not product:
        flash("Produto nao encontrado.", "error")
        return redirect(url_for("order_detail", id=order_id))

    result = set_order_item_quantity(DB_PATH, order_id, int(product_id), quantity_value)
    product_label = format_product_label(product)

    if result == "created":
        flash(f"Item adicionado ao pedido: {product_label}.", "success")
    elif result == "updated":
        flash(f"Quantidade atualizada: {product_label}.", "success")
    elif result == "deleted":
        flash(f"Item removido do pedido: {product_label}.", "info")
    else:
        flash("Nenhuma alteracao foi aplicada ao pedido.", "warning")

    return redirect(url_for("order_detail", id=order_id))


@app.route("/order/<int:order_id>/item/<int:product_id>/delete", methods=["POST"])
@admin_required
def order_item_delete(order_id, product_id):
    order = get_order(DB_PATH, order_id)
    if not order:
        flash(ORDER_NOT_FOUND_MSG, "error")
        return redirect(url_for("orders"))

    product = get_product_by_id(DB_PATH, product_id)
    deleted = delete_order_item(DB_PATH, order_id, product_id)

    if deleted and product:
        flash(f"Item removido do pedido: {format_product_label(product)}.", "info")
    elif deleted:
        flash("Item removido do pedido.", "info")
    else:
        flash("Item nao encontrado no pedido.", "warning")

    return redirect(url_for("order_detail", id=order_id))


@app.route("/suggest_box/<int:order_id>", methods=["GET"])
@admin_required
def suggest_box(order_id):
    order = get_order(DB_PATH, order_id)
    if not order:
        return {"error": "Pedido nao encontrado"}, 404

    items = get_order_items(DB_PATH, order_id)
    recommendation = build_recommendation(items)

    if not recommendation["box"]:
        return {
            "order_id": order_id,
            "suggestion": None,
            "packages_required": recommendation["packages_required"],
            "confidence": recommendation["confidence"],
            "source": recommendation["source"],
            "totals": recommendation["totals"],
        }

    return {
        "order_id": order_id,
        "suggestion": recommendation["box"]["name"],
        "box_id": recommendation["box"]["id"],
        "packages_required": recommendation["packages_required"],
        "confidence": recommendation["confidence"],
        "source": recommendation["source"],
        "totals": recommendation["totals"],
    }


@app.route("/order/<int:order_id>/record_shipment", methods=["POST"])
@admin_required
def record_shipment(order_id):
    existing_shipments = get_shipment_history_by_order(DB_PATH, order_id)
    if existing_shipments:
        flash(
            "A expedicao deste pedido segue o plano de expedicao salvo. Atualize o plano acima ou use a volumetria interativa.",
            "warning",
        )
        return redirect(url_for("order_detail", id=order_id))

    box_id = request.form.get("box_id", "").strip()
    quantity_raw = request.form.get("quantity", "1").strip() or "1"
    if not box_id.isdigit():
        flash("Selecao de embalagem invalida.", "error")
        return redirect(url_for("order_detail", id=order_id))

    try:
        quantity = int(quantity_raw)
    except Exception:
        flash("Quantidade de expedicao invalida.", "error")
        return redirect(url_for("order_detail", id=order_id))

    if quantity <= 0:
        flash("Quantidade de expedicao deve ser maior que zero.", "error")
        return redirect(url_for("order_detail", id=order_id))

    selected_box = get_box_by_id(DB_PATH, int(box_id))
    if not selected_box or int(selected_box["is_active"] or 0) != 1:
        flash("A embalagem selecionada esta inativa ou nao existe.", "error")
        return redirect(url_for("order_detail", id=order_id))

    replace_shipment_history(
        DB_PATH,
        order_id,
        [{"box_id": int(box_id), "quantity": quantity}],
    )
    flash("Plano de expedicao definido com sucesso.", "success")
    return redirect(url_for("order_detail", id=order_id))


def _parse_shipment_plan_rows(form_data, max_rows=3):
    shipments = []
    for index in range(1, max_rows + 1):
        box_id = form_data.get(f"box_id_{index}", "").strip()
        qty_raw = form_data.get(f"quantity_{index}", "").strip()
        if not box_id and not qty_raw:
            continue
        if not box_id.isdigit():
            raise ValueError("Selecione embalagens validas para o plano de expedicao.")

        try:
            quantity = int(qty_raw or "0")
        except Exception as exc:
            raise ValueError("Informe quantidades inteiras no plano de expedicao.") from exc

        if quantity <= 0:
            continue

        shipments.append({"box_id": int(box_id), "quantity": quantity})

    return shipments


def _validate_shipment_plan_boxes(shipments):
    for shipment in shipments:
        selected_box = get_box_by_id(DB_PATH, shipment["box_id"])
        if not selected_box or int(selected_box["is_active"] or 0) != 1:
            raise ValueError("Uma das embalagens do plano esta inativa ou nao existe.")


@app.route("/order/<int:order_id>/shipment-plan", methods=["POST"])
@admin_required
def order_shipment_plan(order_id):
    order = get_order(DB_PATH, order_id)
    if not order:
        flash(ORDER_NOT_FOUND_MSG, "error")
        return redirect(url_for("orders"))

    try:
        shipments = _parse_shipment_plan_rows(request.form)
        _validate_shipment_plan_boxes(shipments)
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("order_detail", id=order_id))

    replace_shipment_history(DB_PATH, order_id, shipments)
    if shipments:
        flash("Plano de expedicao atualizado com sucesso.", "success")
    else:
        flash("Plano de expedicao removido do pedido.", "info")
    return redirect(url_for("order_detail", id=order_id))


@app.route("/order/<int:order_id>/delete", methods=["POST"])
@admin_required
def order_delete(order_id):
    if not get_order(DB_PATH, order_id):
        flash(ORDER_NOT_FOUND_MSG, "error")
        return redirect(url_for("orders"))

    deleted = delete_order_with_dependencies(DB_PATH, order_id)
    if deleted:
        flash(f"Pedido #{order_id} excluido com sucesso.", "success")
    else:
        flash("Nao foi possivel excluir o pedido.", "error")

    return redirect(url_for("orders"))


@app.route("/history", methods=["GET"])
@admin_required
def history():
    rows = list_shipment_history(DB_PATH)
    return render_template("history.html", history=rows)


@app.route("/user/upload", methods=["GET", "POST"])
def user_upload():
    if request.method == "POST":
        uploaded_file = request.files.get("order_file")
        if not uploaded_file or not uploaded_file.filename:
            flash("Selecione um arquivo PDF ou Excel.", "error")
            return render_template(USER_UPLOAD_TEMPLATE)

        original_name = secure_filename(uploaded_file.filename)
        ext = os.path.splitext(original_name)[1].lower()

        if ext not in {".pdf", ".xlsx", ".xls"}:
            flash("Tipo de arquivo nao suportado. Use PDF ou Excel.", "error")
            return render_template(USER_UPLOAD_TEMPLATE)

        temp_name = f"{uuid4().hex}{ext}"
        save_path = os.path.join(UPLOAD_DIR, temp_name)
        uploaded_file.save(save_path)

        try:
            if ext == ".pdf":
                parsed_items = parse_items_from_pdf(save_path)
            else:
                parsed_items = parse_items_from_excel(save_path)

            order_items, unknown_items, import_debug = map_uploaded_items_to_catalog(
                parsed_items,
                include_debug=True,
            )
            if not order_items:
                return render_template(
                    USER_UPLOAD_TEMPLATE,
                    import_error_debug=import_debug,
                    parsed_items=parsed_items,
                )

            order_id = create_order(DB_PATH, order_items)
            session["last_unknown_items"] = unknown_items
            session["last_import_debug"] = import_debug
            return redirect(url_for("user_order_result", order_id=order_id))

        except Exception as exc:
            flash(f"Nao foi possivel processar o arquivo: {exc}", "error")
            return render_template(USER_UPLOAD_TEMPLATE)
        finally:
            if os.path.exists(save_path):
                os.remove(save_path)

    return render_template(USER_UPLOAD_TEMPLATE)


@app.route("/user/order/<int:order_id>", methods=["GET"])
def user_order_result(order_id):
    order = get_order(DB_PATH, order_id)
    if not order:
        flash(ORDER_NOT_FOUND_MSG, "error")
        return redirect(url_for("user_upload"))

    items = get_order_items(DB_PATH, order_id)
    packing_rules = get_packing_rules(DB_PATH)
    recommendation = build_recommendation(items, packing_rules=packing_rules)
    effective_rules = dict(packing_rules)
    if recommendation.get("mb_box_dims"):
        dims = recommendation["mb_box_dims"]
        effective_rules["mb_box_length_cm"] = dims["length_cm"]
        effective_rules["mb_box_width_cm"] = dims["width_cm"]
        effective_rules["mb_box_height_cm"] = dims["height_cm"]

    items_for_view = _attach_pack_profile_to_items(items, effective_rules)
    packing_previews = []
    if recommendation.get("box"):
        packing_previews = build_packing_3d_previews(
            items,
            recommendation["box"],
            packages_required=recommendation.get("packages_required", 1),
            packing_rules=effective_rules,
            unpack_plan=recommendation.get("unpack_plan"),
        )
    packing_preview = packing_previews[0] if packing_previews else None
    grouped_packing_previews = _group_packing_previews(packing_previews)
    unknown_items = session.pop("last_unknown_items", [])
    import_debug = session.pop("last_import_debug", [])

    return render_template(
        "order_detail.html",
        order=order,
        items=items_for_view,
        recommendation=recommendation,
        packing_preview=packing_preview,
        packing_previews=packing_previews,
        grouped_packing_previews=grouped_packing_previews,
        boxes=[],
        products=[],
        shipments=[],
        user_mode=True,
        unknown_items=unknown_items,
        import_debug=import_debug,
    )



# ─── Volumetria Interativa ────────────────────────────────────────────────────

def _build_items_json_for_volumetry(items):
    """Serializa os itens do pedido para o template de volumetria."""
    import json
    result = []
    for item in items:
        l = float(item["length_cm"])
        w = float(item["width_cm"])
        h = float(item["height_cm"])
        result.append({
            "product_id": int(item["product_id"]),
            "name": str(item["name"] or ""),
            "sku": str(item["sku"] or "") if "sku" in item.keys() else "",
            "quantity": int(item["quantity"]),
            "length_cm": l,
            "width_cm": w,
            "height_cm": h,
            "weight": float(item["weight"]),
            "vol_unit": round(l * w * h, 6),
        })
    return json.dumps(result, ensure_ascii=False)


def _build_boxes_json_for_volumetry(boxes):
    """Serializa as embalagens ativas para o template de volumetria."""
    import json
    result = []
    for b in boxes:
        if not int(b["is_active"] or 0):
            continue
        l = float(b["length_cm"])
        w = float(b["width_cm"])
        h = float(b["height_cm"])
        
        # Usar max_capacity_weight se existir, senão calcular a partir de max_weight
        max_capacity = None
        if "max_capacity_weight" in b.keys():
            max_capacity = float(b["max_capacity_weight"]) if b["max_capacity_weight"] else None
        
        if max_capacity is None:
            # Fallback: considerar max_weight como peso bruto, estimar capacidade como 15x
            max_capacity = float(b["max_weight"]) * 15 if b["max_weight"] else 30.0
        
        result.append({
            "id": int(b["id"]),
            "name": str(b["name"]),
            "length_cm": l,
            "width_cm": w,
            "height_cm": h,
            "max_weight": max_capacity,  # Capacidade máxima
            "empty_weight": float(b["max_weight"]),  # Peso bruto (para exibição opcional)
            "vol": round(l * w * h, 6),
        })
    return json.dumps(result, ensure_ascii=False)


def _build_preload_packages(shipments):
    """Constrói JSON de pré-carga a partir do plano de expedição salvo.

    Se cada shipment tem assignments_json, usa-o. Caso contrário, retorna None.
    """
    import json

    if not shipments:
        return None

    packages = []
    for sh in shipments:
        box_id = int(sh["box_id"])
        assignments = {}

        # Tentar desserializar assignments_json se existir
        assignments_json = sh["assignments_json"] if "assignments_json" in sh.keys() else None
        if assignments_json:
            try:
                assignments = json.loads(assignments_json)
                # Garantir que as chaves sejam strings
                assignments = {str(k): int(v) for k, v in assignments.items()}
            except (json.JSONDecodeError, TypeError, ValueError):
                # Se falhar, deixar vazio (para que o usuário redistribua)
                assignments = {}
        
        packages.append({"box_id": box_id, "assignments": assignments})

    return json.dumps(packages, ensure_ascii=False)


@app.route("/order/<int:order_id>/volumetry", methods=["GET"])
@admin_required
def order_volumetry(order_id):
    order = get_order(DB_PATH, order_id)
    if not order:
        flash(ORDER_NOT_FOUND_MSG, "error")
        return redirect(url_for("orders"))

    items = get_order_items(DB_PATH, order_id)
    if not items:
        flash("Este pedido nao possui itens para volumetria.", "info")
        return redirect(url_for("order_detail", id=order_id))

    boxes = list_boxes(DB_PATH)
    shipments = get_shipment_history_by_order(DB_PATH, order_id)

    items_json = _build_items_json_for_volumetry(items)
    boxes_json = _build_boxes_json_for_volumetry(boxes)
    preload_packages = _build_preload_packages(shipments)

    return render_template(
        "volumetry.html",
        order=order,
        items=items,
        items_json=items_json,
        boxes_json=boxes_json,
        preload_packages=preload_packages,
    )


@app.route("/order/<int:order_id>/volumetry/save", methods=["POST"])
@admin_required
def order_volumetry_save(order_id):
    import json

    order = get_order(DB_PATH, order_id)
    if not order:
        flash(ORDER_NOT_FOUND_MSG, "error")
        return redirect(url_for("orders"))

    raw = request.form.get("volumetry_json", "").strip()
    if not raw:
        flash("Nenhum dado de volumetria recebido.", "error")
        return redirect(url_for("order_volumetry", order_id=order_id))

    try:
        payload = json.loads(raw)
        if not isinstance(payload, list):
            raise ValueError("Formato invalido")
    except (json.JSONDecodeError, ValueError) as exc:
        flash(f"Dados de volumetria invalidos: {exc}", "error")
        return redirect(url_for("order_volumetry", order_id=order_id))

    # Cada elemento do payload é um pacote físico com box_id e assignments de itens
    shipments = []
    for entry in payload:
        box_id_raw = entry.get("box_id")
        assignments_raw = entry.get("assignments", {})
        
        if not box_id_raw:
            flash("Uma das embalagens nao foi selecionada.", "error")
            return redirect(url_for("order_volumetry", order_id=order_id))

        try:
            box_id = int(box_id_raw)
        except (TypeError, ValueError):
            flash("ID de embalagem invalido.", "error")
            return redirect(url_for("order_volumetry", order_id=order_id))

        selected_box = get_box_by_id(DB_PATH, box_id)
        if not selected_box or not int(selected_box["is_active"] or 0):
            flash(f"Embalagem ID {box_id} esta inativa ou nao existe.", "error")
            return redirect(url_for("order_volumetry", order_id=order_id))

        shipments.append({
            "box_id": box_id,
            "quantity": 1,  # Cada entrada é 1 embalagem física
            "assignments": assignments_raw
        })

    if not shipments:
        flash("Nenhuma embalagem valida no plano.", "error")
        return redirect(url_for("order_volumetry", order_id=order_id))

    replace_shipment_history(DB_PATH, order_id, shipments)

    # Construir sumário agrupando por box_id
    box_counts = {}
    for sh in shipments:
        bid = sh["box_id"]
        box_counts[bid] = box_counts.get(bid, 0) + 1
    
    box_summary = ", ".join(
        f"{qty}x {get_box_by_id(DB_PATH, bid)['name']}" for bid, qty in box_counts.items()
    )
    flash(f"Plano de embalagem salvo: {box_summary}.", "success")
    return redirect(url_for("order_detail", id=order_id))




if __name__ == "__main__":
    wants_https = APP_URL_SCHEME == "https"
    has_ssl_files = os.path.exists(SSL_CERT_FILE) and os.path.exists(SSL_KEY_FILE)
    use_waitress = safe_bool_env("VOLUME_USE_WAITRESS", not APP_DEBUG)

    if use_waitress and wants_https and has_ssl_files:
        print("HTTPS detectado com certificado local. Iniciando com Flask SSL em vez de Waitress.")
        use_waitress = False

    if use_waitress:
        from waitress import serve

        serve(
            app,
            host=APP_HOST,
            port=APP_PORT,
            threads=safe_int_env("VOLUME_THREADS", 8),
        )
    else:
        ssl_context = None
        if wants_https and has_ssl_files:
            ssl_context = (SSL_CERT_FILE, SSL_KEY_FILE)

            redirect_app = Flask("volume_http_redirect")

            @redirect_app.route("/", defaults={"path": ""}, methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
            @redirect_app.route("/<path:path>", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS", "HEAD"])
            def redirect_to_https(path):
                host = request.host.split(":", 1)[0]
                target = f"https://{host}" if APP_PORT == 443 else f"https://{host}:{APP_PORT}"
                if path:
                    target = f"{target}/{path}"
                if request.query_string:
                    target = f"{target}?{request.query_string.decode()}"
                return redirect(target, code=301)

            redirect_thread = threading.Thread(
                target=redirect_app.run,
                kwargs={
                    "host": APP_HOST,
                    "port": APP_HTTP_PORT,
                    "debug": False,
                    "use_reloader": False,
                },
                daemon=True,
            )
            redirect_thread.start()
        elif wants_https:
            print("HTTPS solicitado, mas certificado/chave nao encontrados. Iniciando sem TLS.")

        app.run(host=APP_HOST, port=APP_PORT, debug=APP_DEBUG, ssl_context=ssl_context)
