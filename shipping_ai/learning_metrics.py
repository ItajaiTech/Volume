"""Prospective metrics: observations and operator feedback, deduplicated per order."""
import json
from contextlib import closing
from collections import Counter
from database import get_connection


def _initialize(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS learning_observations (
        order_id INTEGER PRIMARY KEY, observed_at TEXT DEFAULT CURRENT_TIMESTAMP,
        signature TEXT NOT NULL, plan TEXT NOT NULL, source TEXT NOT NULL,
        evidence_count INTEGER NOT NULL, conflicts INTEGER NOT NULL,
        history_found INTEGER NOT NULL, feedback TEXT, feedback_at TEXT
    )''')


def _plan_key(plan):
    counts = Counter()
    for row in plan:
        counts[int(row['box_id'])] += int(row['quantity'])
    return json.dumps(sorted(counts.items()))


def observe(db_path, items, recommendation):
    from learning import build_signature
    ids = {int(dict(i)['order_id']) for i in items if 'order_id' in dict(i)}
    if len(ids) != 1:
        return
    signature = json.dumps(build_signature(items))
    plan = recommendation.get('history_plan') or [dict(
        box_id=int(recommendation['box']['id']), quantity=int(recommendation['packages_required']))]
    audit = recommendation.get('learning_audit', [])
    with closing(get_connection(db_path)) as conn:
        _initialize(conn)
        conn.execute('''INSERT INTO learning_observations
            (order_id,signature,plan,source,evidence_count,conflicts,history_found)
            VALUES(?,?,?,?,?,?,?) ON CONFLICT(order_id) DO UPDATE SET
            signature=excluded.signature,plan=excluded.plan,source=excluded.source,
            evidence_count=excluded.evidence_count,conflicts=excluded.conflicts,
            history_found=excluded.history_found,observed_at=CURRENT_TIMESTAMP
            WHERE learning_observations.feedback IS NULL''',
            (next(iter(ids)), signature, _plan_key(plan), recommendation['source'],
             recommendation['evidence_count'], recommendation.get('history_conflicts', 0),
             int(any(r['reason'] == 'candidate_validated' for r in audit))))
        conn.commit()


def feedback(db_path, order_id, items, shipments):
    from learning import build_signature
    with closing(get_connection(db_path)) as conn:
        _initialize(conn)
        row = conn.execute('SELECT * FROM learning_observations WHERE order_id=?', (order_id,)).fetchone()
        if row and row['signature'] == json.dumps(build_signature(items)):
            status = 'accepted' if row['plan'] == _plan_key(shipments) else 'corrected'
            conn.execute('UPDATE learning_observations SET feedback=?,feedback_at=CURRENT_TIMESTAMP WHERE order_id=?',
                         (status, order_id))
            conn.commit()


def report(db_path):
    from learning import physical_profile
    from database import get_packing_rules
    from hashlib import sha256
    with closing(get_connection(db_path)) as conn:
        exists = conn.execute("SELECT 1 FROM sqlite_master WHERE name='learning_observations'").fetchone()
        rows = conn.execute('SELECT * FROM learning_observations').fetchall() if exists else []
        products = {int(r['id']): dict(r) for r in conn.execute('SELECT * FROM products')}
        assignments = conn.execute('SELECT order_id, assignments_json FROM shipment_history WHERE assignments_json IS NOT NULL').fetchall()
    rules = get_packing_rules(db_path)
    profiles = {}
    for row in assignments:
        try:
            parsed = json.loads(row['assignments_json'])
            for pid, quantity in parsed.items():
                if type(quantity) is not int or quantity <= 0:
                    continue
                profile = physical_profile(products[int(pid)], rules)
                key = sha256(json.dumps(profile).encode()).hexdigest()[:16]
                profiles.setdefault(key, set()).add(int(row['order_id']))
        except (ValueError, KeyError, TypeError, AttributeError):
            continue
    total = len(rows)
    confirmed = [r for r in rows if r['feedback']]
    historical = [r for r in confirmed if r['source'].startswith('history_')]
    return dict(observed_orders=total, confirmed_orders=len(confirmed),
                history_found_percent=round(100*sum(r['history_found'] for r in rows)/total, 1) if total else None,
                historical_acceptance_percent=round(100*sum(r['feedback']=='accepted' for r in historical)/len(historical), 1) if historical else None,
                correction_percent=round(100*sum(r['feedback']=='corrected' for r in confirmed)/len(confirmed), 1) if confirmed else None,
                algorithm_only_orders=sum(not r['source'].startswith('history_') for r in rows),
                conflicting_orders=sum(r['conflicts']>0 for r in rows),
                profile_order_counts={key: len(orders) for key, orders in profiles.items()},
                corrected_plan_counts=dict(Counter(r['plan'] for r in rows if r['feedback']=='corrected')))


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('database')
    print(json.dumps(report(parser.parse_args().database), indent=2, ensure_ascii=False))
