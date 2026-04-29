import re
import sqlite3

import pandas as pd

DB_PATH = r"c:\Volume\shipping_ai\database.db"
XLSX_PATH = r"c:\Volume\Caixas.xlsx"


def to_float(value):
    text = str(value or "").strip().replace(".", "").replace(",", ".")
    try:
        return float(text)
    except Exception:
        return 0.0


def parse_measurements(raw):
    numbers = re.findall(r"(\d+(?:[\.,]\d+)?)", str(raw or ""))
    values = [to_float(n) for n in numbers]
    if len(values) < 3:
        return 0.0, 0.0, 0.0

    # Planilha informa A x L x C; banco salva C, L, A.
    height_cm = values[0]
    width_cm = values[1]
    length_cm = values[2]
    return length_cm, width_cm, height_cm


def main():
    df = pd.read_excel(XLSX_PATH)
    conn = sqlite3.connect(DB_PATH)

    imported = 0
    skipped = 0

    for _, row in df.iterrows():
        name = str(row.get("Unnamed: 0", "")).strip()
        measures = str(row.get("Unnamed: 1", "")).strip()
        weight_text = str(row.get("Unnamed: 2", "")).strip()

        if not name or name.lower() in {"nome", "nan"}:
            skipped += 1
            continue

        length_cm, width_cm, height_cm = parse_measurements(measures)
        max_weight = to_float(weight_text.split()[0] if weight_text else "0")

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
        imported += 1

    conn.commit()
    conn.close()

    print(f"Embalagens importadas/atualizadas: {imported}")
    print(f"Linhas ignoradas: {skipped}")


if __name__ == "__main__":
    main()
