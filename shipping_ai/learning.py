from collections import defaultdict

from database import get_connection


def build_signature(order_items):
    """
    order_items: iterable with keys product_id and quantity
    returns tuple sorted by product_id: ((pid, qty), ...)
    """
    parts = []
    for item in order_items:
        parts.append((int(item["product_id"]), int(item["quantity"])))
    parts.sort(key=lambda x: x[0])
    return tuple(parts)


def _signature_similarity(sig_a, sig_b):
    ids_a = {pid for pid, _ in sig_a}
    ids_b = {pid for pid, _ in sig_b}

    if ids_a != ids_b:
        return 0.0

    if not ids_a:
        return 0.0

    ratios = []
    b_map = dict(sig_b)
    for pid, qty_a in sig_a:
        qty_b = b_map[pid]
        max_v = max(qty_a, qty_b)
        min_v = min(qty_a, qty_b)
        ratios.append(min_v / max_v if max_v else 1.0)

    return sum(ratios) / len(ratios)


def _load_order_history_rows(db_path):
    with get_connection(db_path) as conn:
        return conn.execute(
            """
            SELECT
                sh.order_id,
                sh.box_id,
                oi.product_id,
                oi.quantity
            FROM shipment_history sh
            JOIN order_items oi ON oi.order_id = sh.order_id
            ORDER BY sh.order_id ASC
            """
        ).fetchall()


def _group_single_box_history(rows):
    order_data = {}
    for row in rows:
        oid = int(row["order_id"])
        if oid not in order_data:
            order_data[oid] = {
                "box_ids": set(),
                "items": [],
            }
        order_data[oid]["box_ids"].add(int(row["box_id"]))
        order_data[oid]["items"].append(
            {"product_id": int(row["product_id"]), "quantity": int(row["quantity"])}
        )

    return [rec for rec in order_data.values() if len(rec["box_ids"]) == 1]


def _find_exact_history_match(history_records, target_signature):
    exact_votes = defaultdict(int)
    exact_count = 0

    for rec in history_records:
        box_id = next(iter(rec["box_ids"]))
        if build_signature(rec["items"]) != target_signature:
            continue
        exact_votes[box_id] += 1
        exact_count += 1

    if exact_count <= 0:
        return None

    best_box_id, best_count = sorted(exact_votes.items(), key=lambda x: x[1], reverse=True)[0]
    confidence = int(round((best_count / exact_count) * 100))
    return {
        "box_id": best_box_id,
        "confidence": confidence,
        "source": "history_exact",
        "evidence_count": exact_count,
    }


def _find_similar_history_match(history_records, target_signature):
    weighted_votes = defaultdict(float)
    similar_count = 0

    for rec in history_records:
        box_id = next(iter(rec["box_ids"]))
        signature = build_signature(rec["items"])
        similarity = _signature_similarity(target_signature, signature)
        if similarity < 0.6:
            continue
        weighted_votes[box_id] += similarity
        similar_count += 1

    if not weighted_votes:
        return None

    sorted_votes = sorted(weighted_votes.items(), key=lambda x: x[1], reverse=True)
    best_box_id, best_score = sorted_votes[0]
    total_score = sum(weighted_votes.values())
    confidence = int(round((best_score / total_score) * 100)) if total_score else 0
    confidence = max(50, min(confidence, 95))
    return {
        "box_id": best_box_id,
        "confidence": confidence,
        "source": "history_similar",
        "evidence_count": similar_count,
    }


def suggest_box_from_history(db_path, target_items):
    """
    Returns dict or None:
    {
      box_id,
      confidence,
      source,
      evidence_count
    }
    """
    target_signature = build_signature(target_items)
    if not target_signature:
        return None

    rows = _load_order_history_rows(db_path)

    if not rows:
        return None

    history_records = _group_single_box_history(rows)
    exact_match = _find_exact_history_match(history_records, target_signature)
    if exact_match:
        return exact_match

    similar_match = _find_similar_history_match(history_records, target_signature)
    if similar_match:
        return similar_match

    return None
