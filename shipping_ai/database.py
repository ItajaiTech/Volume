import sqlite3
import re
from contextlib import closing


DEFAULT_SSDM2_MASTER_BOX_NAME = "Caixa NVME"


DEFAULT_PACKING_RULES = {
    "ssd25_bundle_qty": 10,
    "ssd25_bundle_length_cm": 22.0,
    "ssd25_bundle_width_cm": 8.0,
    "ssd25_bundle_height_cm": 6.7,
    "ssdm2_bundle_qty": 10,
    "ssdm2_bundle_length_cm": 21.3,
    "ssdm2_bundle_width_cm": 5.0,
    "ssdm2_bundle_height_cm": 9.0,
    "ssdm2_master_box_qty": 250,
    "ssdm2_default_box_name": DEFAULT_SSDM2_MASTER_BOX_NAME,
    "ram_nb_bundle_qty": 10,
    "ram_nb_bundle_min_qty": 10,
    "ram_desk_bundle_qty": 10,
    "ram_desk_bundle_min_qty": 10,
    "min_effective_max_weight_kg": 1.0,
}


def get_connection(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def _normalize_sku(value):
    sku = " ".join(str(value or "").strip().upper().split())
    return sku or None


def _looks_like_sku(value):
    text = " ".join(str(value or "").strip().upper().split())
    if not text:
        return False
    if not any(ch.isdigit() for ch in text):
        return False
    return bool(re.match(r"^[A-Z0-9][A-Z0-9\- ]{2,}$", text))


def _split_legacy_product_name(raw_name):
    name = " ".join(str(raw_name or "").strip().split())
    if " - " not in name:
        return None, name

    prefix, suffix = name.split(" - ", 1)
    sku = _normalize_sku(prefix)
    clean_name = suffix.strip() or name

    if not _looks_like_sku(sku):
        return None, name
    return sku, clean_name


def _parse_decimal_value(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        text = str(value or "").strip().replace(" ", "")
        if not text:
            return 0.0

        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        else:
            text = text.replace(",", ".")

        return float(text)


def init_db(db_path):
    with closing(get_connection(db_path)) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sku TEXT,
                name TEXT NOT NULL UNIQUE,
                length_cm REAL NOT NULL,
                width_cm REAL NOT NULL,
                height_cm REAL NOT NULL,
                weight REAL NOT NULL
            );

            CREATE TABLE IF NOT EXISTS boxes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                length_cm REAL NOT NULL,
                width_cm REAL NOT NULL,
                height_cm REAL NOT NULL,
                max_weight REAL NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS order_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                quantity INTEGER NOT NULL,
                FOREIGN KEY(order_id) REFERENCES orders(id),
                FOREIGN KEY(product_id) REFERENCES products(id)
            );

            CREATE TABLE IF NOT EXISTS shipment_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                box_id INTEGER NOT NULL,
                quantity INTEGER NOT NULL DEFAULT 1,
                FOREIGN KEY(order_id) REFERENCES orders(id),
                FOREIGN KEY(box_id) REFERENCES boxes(id)
            );

            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_order_items_order ON order_items(order_id);
            CREATE INDEX IF NOT EXISTS idx_order_items_product ON order_items(product_id);
            CREATE INDEX IF NOT EXISTS idx_shipments_order ON shipment_history(order_id);
            CREATE INDEX IF NOT EXISTS idx_shipments_box ON shipment_history(box_id);
            """
        )

        columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(products)").fetchall()
        }
        if "sku" not in columns:
            conn.execute("ALTER TABLE products ADD COLUMN sku TEXT")

        box_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(boxes)").fetchall()
        }
        if "is_active" not in box_columns:
            conn.execute("ALTER TABLE boxes ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1")

        conn.execute("UPDATE boxes SET is_active = 1 WHERE is_active IS NULL")

        conn.execute("CREATE INDEX IF NOT EXISTS idx_products_sku ON products(sku)")

        # Adicionar coluna max_capacity_weight (capacidade máxima diferente do peso bruto)
        box_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(boxes)").fetchall()
        }
        if "max_capacity_weight" not in box_columns:
            conn.execute("ALTER TABLE boxes ADD COLUMN max_capacity_weight REAL")
            # Migração: definir capacidade máxima padrão como 30kg para caixas existentes
            conn.execute("""
                UPDATE boxes 
                SET max_capacity_weight = COALESCE(max_capacity_weight, 30.0)
                WHERE max_capacity_weight IS NULL
            """)

        shipment_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(shipment_history)").fetchall()
        }
        if "quantity" not in shipment_columns:
            conn.execute(
                "ALTER TABLE shipment_history ADD COLUMN quantity INTEGER NOT NULL DEFAULT 1"
            )
        
        # Adicionar coluna para armazenar distribuição de itens por embalagem (JSON)
        if "assignments_json" not in shipment_columns:
            conn.execute(
                "ALTER TABLE shipment_history ADD COLUMN assignments_json TEXT"
            )

        legacy_rows = conn.execute(
            """
            SELECT id, name
            FROM products
            WHERE (sku IS NULL OR TRIM(sku) = '')
              AND INSTR(name, ' - ') > 0
            """
        ).fetchall()

        for row in legacy_rows:
            parsed_sku, parsed_name = _split_legacy_product_name(row["name"])
            if not parsed_sku:
                continue

            try:
                conn.execute(
                    "UPDATE products SET sku = ?, name = ? WHERE id = ?",
                    (parsed_sku, parsed_name, int(row["id"])),
                )
            except sqlite3.IntegrityError:
                # Keep current name if unique constraint would be violated.
                conn.execute(
                    "UPDATE products SET sku = ? WHERE id = ?",
                    (parsed_sku, int(row["id"])),
                )

        for key, value in DEFAULT_PACKING_RULES.items():
            conn.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES (?, ?)
                ON CONFLICT(key) DO NOTHING
                """,
                (key, str(value)),
            )

        conn.commit()


def add_product(db_path, sku, name, length_cm, width_cm, height_cm, weight):
    with closing(get_connection(db_path)) as conn:
        conn.execute(
            """
            INSERT INTO products (sku, name, length_cm, width_cm, height_cm, weight)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                _normalize_sku(sku),
                str(name).strip(),
                _parse_decimal_value(length_cm),
                _parse_decimal_value(width_cm),
                _parse_decimal_value(height_cm),
                _parse_decimal_value(weight),
            ),
        )
        conn.commit()


def list_products(db_path, sort_by="name", order="ASC"):
    """Lista produtos com suporte a ordenação.
    
    Args:
        db_path: caminho do banco de dados
        sort_by: coluna para ordenação (name, sku, length_cm, width_cm, height_cm, weight, id)
        order: ASC ou DESC
    """
    # Validar parâmetros para evitar SQL injection
    valid_columns = {"id", "sku", "name", "length_cm", "width_cm", "height_cm", "weight"}
    valid_orders = {"ASC", "DESC"}
    
    sort_by = sort_by if sort_by in valid_columns else "name"
    order = order.upper() if order.upper() in valid_orders else "ASC"
    
    with closing(get_connection(db_path)) as conn:
        query = f"SELECT * FROM products ORDER BY {sort_by} {order}"
        rows = conn.execute(query).fetchall()
    return rows


def _to_active_flag(value, default=1):
    text = str(value).strip().lower()
    if text in {"1", "true", "on", "yes", "sim"}:
        return 1
    if text in {"0", "false", "off", "no", "nao"}:
        return 0
    if value is None:
        return 1 if int(default) else 0

    try:
        return 1 if int(value) != 0 else 0
    except Exception:
        return 1 if int(default) else 0


def add_box(db_path, name, length_cm, width_cm, height_cm, max_weight, is_active=1):
    with closing(get_connection(db_path)) as conn:
        conn.execute(
            """
            INSERT INTO boxes (name, length_cm, width_cm, height_cm, max_weight, is_active)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                name.strip(),
                float(length_cm),
                float(width_cm),
                float(height_cm),
                float(max_weight),
                _to_active_flag(is_active, default=1),
            ),
        )
        conn.commit()


def list_boxes(db_path, active_only=True):
    where_clause = "WHERE is_active = 1" if active_only else ""
    status_order = "is_active DESC, " if not active_only else ""

    with closing(get_connection(db_path)) as conn:
        rows = conn.execute(
            f"SELECT * FROM boxes {where_clause} "
            f"ORDER BY {status_order}(length_cm * width_cm * height_cm) ASC, name ASC"
        ).fetchall()
    return rows


def get_box_by_id(db_path, box_id):
    with closing(get_connection(db_path)) as conn:
        row = conn.execute("SELECT * FROM boxes WHERE id = ?", (box_id,)).fetchone()
    return row


def create_order(db_path, items):
    """
    items: list of dicts [{"product_id": int, "quantity": int}, ...]
    """
    with closing(get_connection(db_path)) as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO orders DEFAULT VALUES")
        order_id = cur.lastrowid

        for item in items:
            cur.execute(
                """
                INSERT INTO order_items (order_id, product_id, quantity)
                VALUES (?, ?, ?)
                """,
                (order_id, int(item["product_id"]), int(item["quantity"])),
            )

        conn.commit()
    return order_id


def list_orders(db_path):
    with closing(get_connection(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT
                o.id,
                o.created_at,
                COUNT(oi.id) AS line_count,
                COALESCE(SUM(oi.quantity), 0) AS total_units,
                (
                    SELECT GROUP_CONCAT(
                        CASE
                            WHEN COALESCE(sh.quantity, 1) > 1 THEN b.name || ' x' || sh.quantity
                            ELSE b.name
                        END,
                        ', '
                    )
                    FROM shipment_history sh
                    JOIN boxes b ON b.id = sh.box_id
                    WHERE sh.order_id = o.id
                ) AS used_box_name
            FROM orders o
            LEFT JOIN order_items oi ON oi.order_id = o.id
            GROUP BY o.id, o.created_at
            ORDER BY o.id DESC
            """
        ).fetchall()
    return rows


def get_order(db_path, order_id):
    with closing(get_connection(db_path)) as conn:
        row = conn.execute("SELECT * FROM orders WHERE id = ?", (order_id,)).fetchone()
    return row


def get_order_items(db_path, order_id):
    with closing(get_connection(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT
                oi.id,
                oi.order_id,
                oi.product_id,
                oi.quantity,
                p.sku,
                p.name,
                p.length_cm,
                p.width_cm,
                p.height_cm,
                p.weight
            FROM order_items oi
            JOIN products p ON p.id = oi.product_id
            WHERE oi.order_id = ?
            ORDER BY p.name ASC
            """,
            (order_id,),
        ).fetchall()
    return rows


def set_order_item_quantity(db_path, order_id, product_id, quantity):
    order_id = int(order_id)
    product_id = int(product_id)
    quantity = int(quantity)

    with closing(get_connection(db_path)) as conn:
        existing = conn.execute(
            """
            SELECT id, quantity
            FROM order_items
            WHERE order_id = ? AND product_id = ?
            """,
            (order_id, product_id),
        ).fetchone()

        if quantity <= 0:
            conn.execute(
                "DELETE FROM order_items WHERE order_id = ? AND product_id = ?",
                (order_id, product_id),
            )
            conn.commit()
            return "deleted" if existing else "noop"

        if existing:
            conn.execute(
                "UPDATE order_items SET quantity = ? WHERE id = ?",
                (quantity, int(existing["id"])),
            )
            conn.commit()
            return "updated"

        conn.execute(
            """
            INSERT INTO order_items (order_id, product_id, quantity)
            VALUES (?, ?, ?)
            """,
            (order_id, product_id, quantity),
        )
        conn.commit()
        return "created"


def delete_order_item(db_path, order_id, product_id):
    with closing(get_connection(db_path)) as conn:
        cur = conn.execute(
            "DELETE FROM order_items WHERE order_id = ? AND product_id = ?",
            (int(order_id), int(product_id)),
        )
        conn.commit()
    return cur.rowcount > 0


def add_shipment_history(db_path, order_id, box_id, quantity=1):
    with closing(get_connection(db_path)) as conn:
        conn.execute(
            "INSERT INTO shipment_history (order_id, box_id, quantity) VALUES (?, ?, ?)",
            (int(order_id), int(box_id), max(1, int(quantity))),
        )
        conn.commit()


def replace_shipment_history(db_path, order_id, shipments):
    import json
    with closing(get_connection(db_path)) as conn:
        conn.execute("DELETE FROM shipment_history WHERE order_id = ?", (int(order_id),))
        for shipment in shipments or []:
            qty = max(1, int(shipment["quantity"]))
            assignments_json = None
            if "assignments" in shipment and shipment["assignments"]:
                try:
                    assignments_json = json.dumps(shipment["assignments"], ensure_ascii=False)
                except Exception:
                    assignments_json = None
            
            conn.execute(
                "INSERT INTO shipment_history (order_id, box_id, quantity, assignments_json) VALUES (?, ?, ?, ?)",
                (int(order_id), int(shipment["box_id"]), qty, assignments_json),
            )
        conn.commit()


def get_shipment_history_by_order(db_path, order_id):
    with closing(get_connection(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT
                sh.id,
                sh.order_id,
                sh.box_id,
                sh.quantity,
                sh.assignments_json,
                b.name AS box_name
            FROM shipment_history sh
            JOIN boxes b ON b.id = sh.box_id
            WHERE sh.order_id = ?
            ORDER BY sh.id ASC
            """,
            (int(order_id),),
        ).fetchall()
    return rows


def delete_order_with_dependencies(db_path, order_id):
    with closing(get_connection(db_path)) as conn:
        conn.execute("DELETE FROM shipment_history WHERE order_id = ?", (int(order_id),))
        conn.execute("DELETE FROM order_items WHERE order_id = ?", (int(order_id),))
        cur = conn.execute("DELETE FROM orders WHERE id = ?", (int(order_id),))
        conn.commit()
    return cur.rowcount > 0


def list_shipment_history(db_path):
    with closing(get_connection(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT
                sh.id,
                sh.order_id,
                sh.box_id,
                sh.quantity,
                o.created_at,
                b.name AS box_name
            FROM shipment_history sh
            JOIN orders o ON o.id = sh.order_id
            JOIN boxes b ON b.id = sh.box_id
            ORDER BY sh.id DESC
            """
        ).fetchall()
    return rows


def map_products_by_normalized_name(db_path):
    """
    Returns a dict: normalized_name -> product row
    """
    products = list_products(db_path)
    name_map = {}
    for p in products:
        key = normalize_name(p["name"])
        name_map[key] = p
    return name_map


def normalize_name(value):
    return " ".join(str(value).strip().lower().split())


def get_dashboard_stats(db_path):
    with closing(get_connection(db_path)) as conn:
        most_used_box = conn.execute(
            """
            SELECT b.name, SUM(COALESCE(sh.quantity, 1)) AS uses
            FROM shipment_history sh
            JOIN boxes b ON b.id = sh.box_id
            GROUP BY sh.box_id, b.name
            ORDER BY uses DESC
            LIMIT 1
            """
        ).fetchone()

        avg_volume = conn.execute(
            """
            SELECT AVG(order_volume) AS avg_volume
            FROM (
                SELECT
                    o.id,
                    SUM((p.length_cm * p.width_cm * p.height_cm) * oi.quantity) AS order_volume
                FROM orders o
                JOIN order_items oi ON oi.order_id = o.id
                JOIN products p ON p.id = oi.product_id
                GROUP BY o.id
            ) q
            """
        ).fetchone()

        top_products = conn.execute(
            """
            SELECT p.name, SUM(oi.quantity) AS total_qty
            FROM order_items oi
            JOIN products p ON p.id = oi.product_id
            GROUP BY p.id, p.name
            ORDER BY total_qty DESC
            LIMIT 5
            """
        ).fetchall()

    return {
        "most_used_box": most_used_box,
        "avg_volume": (avg_volume["avg_volume"] if avg_volume and avg_volume["avg_volume"] else 0),
        "top_products": top_products,
    }


def get_product_by_id(db_path, product_id):
    with closing(get_connection(db_path)) as conn:
        row = conn.execute("SELECT * FROM products WHERE id = ?", (product_id,)).fetchone()
    return row


def update_product(db_path, product_id, sku, name, length_cm, width_cm, height_cm, weight):
    with closing(get_connection(db_path)) as conn:
        conn.execute(
            """
            UPDATE products
            SET sku = ?, name = ?, length_cm = ?, width_cm = ?, height_cm = ?, weight = ?
            WHERE id = ?
            """,
            (
                _normalize_sku(sku),
                str(name).strip(),
                _parse_decimal_value(length_cm),
                _parse_decimal_value(width_cm),
                _parse_decimal_value(height_cm),
                _parse_decimal_value(weight),
                int(product_id),
            ),
        )
        conn.commit()


def update_box(db_path, box_id, name, length_cm, width_cm, height_cm, max_weight, is_active=1):
    with closing(get_connection(db_path)) as conn:
        conn.execute(
            """
            UPDATE boxes
            SET name = ?, length_cm = ?, width_cm = ?, height_cm = ?, max_weight = ?, is_active = ?
            WHERE id = ?
            """,
            (
                name.strip(),
                float(length_cm),
                float(width_cm),
                float(height_cm),
                float(max_weight),
                _to_active_flag(is_active, default=1),
                int(box_id),
            ),
        )
        conn.commit()


def set_box_active(db_path, box_id, is_active):
    with closing(get_connection(db_path)) as conn:
        conn.execute(
            "UPDATE boxes SET is_active = ? WHERE id = ?",
            (_to_active_flag(is_active, default=1), int(box_id)),
        )
        conn.commit()


def _mapped_cell(row, column_mapping, key, default=""):
    column_name = column_mapping.get(key)
    if not column_name:
        return default
    return row.get(column_name, default)


def _parse_product_import_row(row, column_mapping):
    raw_sku = _mapped_cell(row, column_mapping, "sku", "")
    sku = _normalize_sku(raw_sku)
    name = str(_mapped_cell(row, column_mapping, "name", "")).strip()

    if not name and not sku:
        return None
    if not name:
        name = sku

    length_cm = _parse_decimal_value(_mapped_cell(row, column_mapping, "length", 0))
    width_cm = _parse_decimal_value(_mapped_cell(row, column_mapping, "width", 0))
    height_cm = _parse_decimal_value(_mapped_cell(row, column_mapping, "height", 0))
    weight = _parse_decimal_value(_mapped_cell(row, column_mapping, "weight", 0))

    if not name or min(length_cm, width_cm, height_cm, weight) <= 0:
        return None

    return {
        "sku": sku,
        "name": name,
        "length_cm": length_cm,
        "width_cm": width_cm,
        "height_cm": height_cm,
        "weight": weight,
    }


def _find_existing_product_id(conn, sku, name):
    if sku:
        existing = conn.execute(
            "SELECT id FROM products WHERE sku = ?",
            (sku,),
        ).fetchone()
        if existing:
            return int(existing["id"])

    existing = conn.execute(
        "SELECT id FROM products WHERE name = ?",
        (name,),
    ).fetchone()
    if existing:
        return int(existing["id"])
    return None


def _upsert_product(conn, existing_id, product_data):
    if existing_id is not None:
        conn.execute(
            """
            UPDATE products
            SET sku = ?, name = ?, length_cm = ?, width_cm = ?, height_cm = ?, weight = ?
            WHERE id = ?
            """,
            (
                product_data["sku"],
                product_data["name"],
                product_data["length_cm"],
                product_data["width_cm"],
                product_data["height_cm"],
                product_data["weight"],
                existing_id,
            ),
        )
        return

    conn.execute(
        """
        INSERT INTO products (sku, name, length_cm, width_cm, height_cm, weight)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            product_data["sku"],
            product_data["name"],
            product_data["length_cm"],
            product_data["width_cm"],
            product_data["height_cm"],
            product_data["weight"],
        ),
    )


def import_products_from_excel(db_path, file_path, column_mapping):
    """
    Import products from Excel file.
    column_mapping: dict with keys 'name', 'length', 'width', 'height', 'weight'
                   and optional 'sku'
    Returns: (imported_count, error_count, error_messages)
    """
    import pandas as pd
    
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        raise ValueError(f"Erro ao ler arquivo Excel: {e}")
    
    # Strip whitespace from column names.
    df.columns = df.columns.str.strip()

    imported = 0
    errors = 0
    error_messages = []

    with closing(get_connection(db_path)) as conn:
        for idx, row in df.iterrows():
            try:
                product_data = _parse_product_import_row(row, column_mapping)
                if not product_data:
                    continue

                existing_id = _find_existing_product_id(
                    conn,
                    product_data["sku"],
                    product_data["name"],
                )
                _upsert_product(conn, existing_id, product_data)
                imported += 1
            except Exception as e:
                errors += 1
                error_messages.append(f"Linha {idx + 2}: {str(e)}")

        conn.commit()

    return imported, errors, error_messages


def _to_float_ptbr(value):
    text = str(value or "").strip().replace(".", "").replace(",", ".")
    try:
        return float(text)
    except Exception:
        return 0.0


def _safe_int(value, default):
    try:
        return int(float(value))
    except Exception:
        return int(default)


def _safe_float(value, default):
    try:
        return float(value)
    except Exception:
        return float(default)


def _normalize_packing_rules(rules):
    legacy_bundle_qty = rules.get("ram_bundle_qty")
    legacy_min_qty = rules.get("ram_bundle_min_qty")

    ram_nb_bundle_qty = rules.get("ram_nb_bundle_qty")
    if ram_nb_bundle_qty is None:
        ram_nb_bundle_qty = legacy_bundle_qty if legacy_bundle_qty is not None else 10

    ram_nb_bundle_min_qty = rules.get("ram_nb_bundle_min_qty")
    if ram_nb_bundle_min_qty is None:
        ram_nb_bundle_min_qty = legacy_min_qty if legacy_min_qty is not None else 10

    ram_desk_bundle_qty = rules.get("ram_desk_bundle_qty")
    if ram_desk_bundle_qty is None:
        ram_desk_bundle_qty = legacy_bundle_qty if legacy_bundle_qty is not None else 10

    ram_desk_bundle_min_qty = rules.get("ram_desk_bundle_min_qty")
    if ram_desk_bundle_min_qty is None:
        ram_desk_bundle_min_qty = legacy_min_qty if legacy_min_qty is not None else 10

    normalized = {
        "ssd25_bundle_qty": max(1, _safe_int(rules.get("ssd25_bundle_qty"), 10)),
        "ssd25_bundle_length_cm": max(0.1, _safe_float(rules.get("ssd25_bundle_length_cm"), 22.0)),
        "ssd25_bundle_width_cm": max(0.1, _safe_float(rules.get("ssd25_bundle_width_cm"), 8.0)),
        "ssd25_bundle_height_cm": max(0.1, _safe_float(rules.get("ssd25_bundle_height_cm"), 6.7)),
        "ssdm2_bundle_qty": max(1, _safe_int(rules.get("ssdm2_bundle_qty"), 10)),
        "ssdm2_bundle_length_cm": max(0.1, _safe_float(rules.get("ssdm2_bundle_length_cm"), 21.3)),
        "ssdm2_bundle_width_cm": max(0.1, _safe_float(rules.get("ssdm2_bundle_width_cm"), 5.0)),
        "ssdm2_bundle_height_cm": max(0.1, _safe_float(rules.get("ssdm2_bundle_height_cm"), 9.0)),
        "ssdm2_master_box_qty": max(1, _safe_int(rules.get("ssdm2_master_box_qty"), 250)),
        "ram_nb_bundle_qty": max(1, _safe_int(ram_nb_bundle_qty, 10)),
        "ram_nb_bundle_min_qty": max(1, _safe_int(ram_nb_bundle_min_qty, 10)),
        "ram_desk_bundle_qty": max(1, _safe_int(ram_desk_bundle_qty, 10)),
        "ram_desk_bundle_min_qty": max(1, _safe_int(ram_desk_bundle_min_qty, 10)),
        "min_effective_max_weight_kg": max(0.0, _safe_float(rules.get("min_effective_max_weight_kg"), 1.0)),
    }
    normalized["ssdm2_default_box_name"] = " ".join(
        str(rules.get("ssdm2_default_box_name") or DEFAULT_SSDM2_MASTER_BOX_NAME).split()
    ) or DEFAULT_SSDM2_MASTER_BOX_NAME
    return normalized


def get_packing_rules(db_path):
    rules = dict(DEFAULT_PACKING_RULES)
    with closing(get_connection(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT key, value
            FROM app_settings
            WHERE key IN (
                'ssd25_bundle_qty',
                'ssd25_bundle_length_cm',
                'ssd25_bundle_width_cm',
                'ssd25_bundle_height_cm',
                'ssdm2_bundle_qty',
                'ssdm2_bundle_length_cm',
                'ssdm2_bundle_width_cm',
                'ssdm2_bundle_height_cm',
                'ssdm2_master_box_qty',
                'ssdm2_default_box_name',
                'ram_nb_bundle_qty',
                'ram_nb_bundle_min_qty',
                'ram_desk_bundle_qty',
                'ram_desk_bundle_min_qty',
                'ram_bundle_qty',
                'ram_bundle_min_qty',
                'min_effective_max_weight_kg'
            )
            """
        ).fetchall()

    for row in rows:
        rules[row["key"]] = row["value"]

    return _normalize_packing_rules(rules)


def update_packing_rules(db_path, rules):
    normalized = _normalize_packing_rules(rules)
    with closing(get_connection(db_path)) as conn:
        for key, value in normalized.items():
            conn.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (key, str(value)),
            )
        conn.commit()

    return normalized


def _parse_box_dimensions(raw_text):
    numbers = re.findall(r"(\d+(?:[\.,]\d+)?)", str(raw_text or ""))
    if len(numbers) < 3:
        return 0.0, 0.0, 0.0

    values = [_to_float_ptbr(x) for x in numbers[:3]]
    # Planilha geralmente vem como A x L x C; banco usa C, L, A.
    height_cm, width_cm, length_cm = values[0], values[1], values[2]
    return length_cm, width_cm, height_cm


def import_boxes_from_excel(db_path, file_path):
    """
    Importa embalagens a partir do modelo em Excel (ex.: C:\\Volume\\Caixas.xlsx).
    Faz upsert por nome para permitir reimportacao sem duplicar.
    Retorna: (imported_or_updated, skipped, error_messages)
    """
    import pandas as pd

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        raise ValueError(f"Erro ao ler arquivo Excel de embalagens: {e}")

    imported_or_updated = 0
    skipped = 0
    error_messages = []

    with closing(get_connection(db_path)) as conn:
        for idx, row in df.iterrows():
            try:
                name = str(row.get("Unnamed: 0", "")).strip()
                measures = str(row.get("Unnamed: 1", "")).strip()
                weight_text = str(row.get("Unnamed: 2", "")).strip()

                if not name or name.lower() in {"nome", "nan"}:
                    skipped += 1
                    continue

                length_cm, width_cm, height_cm = _parse_box_dimensions(measures)
                max_weight = _to_float_ptbr((weight_text.split() or [""])[0])

                if min(length_cm, width_cm, height_cm, max_weight) <= 0:
                    skipped += 1
                    continue

                conn.execute(
                    """
                    INSERT INTO boxes (name, length_cm, width_cm, height_cm, max_weight)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(name) DO UPDATE SET
                        length_cm = excluded.length_cm,
                        width_cm = excluded.width_cm,
                        height_cm = excluded.height_cm,
                        max_weight = excluded.max_weight
                    """,
                    (name, length_cm, width_cm, height_cm, max_weight),
                )
                imported_or_updated += 1

            except Exception as e:
                error_messages.append(f"Linha {idx + 2}: {e}")

        conn.commit()

    return imported_or_updated, skipped, error_messages
