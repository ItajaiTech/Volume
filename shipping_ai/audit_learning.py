"""Read-only replay of shipped orders, excluding each order's own evidence."""
import argparse
import gc
import json
import sqlite3
import tempfile
from collections import Counter
from contextlib import closing
from pathlib import Path
from database import get_order_items
from learning import suggest_plan_from_history


def audit_database(database, limit=30):
    with tempfile.TemporaryDirectory() as directory:
        copied = str(Path(directory) / 'audit.db')
        uri = Path(database).resolve().as_uri() + '?mode=ro'
        with closing(sqlite3.connect(uri, uri=True)) as source, closing(sqlite3.connect(copied)) as target:
            source.backup(target)
        with closing(sqlite3.connect(copied)) as conn:
            ids = [r[0] for r in conn.execute(
                'SELECT DISTINCT order_id FROM shipment_history ORDER BY order_id DESC LIMIT ?', (limit,))]
        sources, reasons = Counter(), Counter()
        errors = []
        for oid in ids:
            audit = []
            try:
                result = suggest_plan_from_history(copied, get_order_items(copied, oid), audit=audit)
                sources[result['source'] if result else 'algorithm_fallback'] += 1
                reasons.update(r['reason'] for r in audit)
            except Exception as exc:
                errors.append(dict(order_id=oid, error=str(exc)))
        gc.collect()
    return dict(sample_orders=len(ids), recommendations=dict(sources), reasons=dict(reasons), errors=errors)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('database')
    parser.add_argument('--limit', type=int, default=30)
    args = parser.parse_args()
    print(json.dumps(audit_database(args.database, args.limit), indent=2))
