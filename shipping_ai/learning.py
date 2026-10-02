from collections import defaultdict
import json
import math
from contextlib import closing

from database import get_connection


def build_signature(order_items):
    """
    order_items: iterable with keys product_id and quantity
    returns tuple sorted by product_id: ((pid, qty), ...)
    """
    parts = defaultdict(int)
    for item in order_items:
        parts[int(item["product_id"])] += int(item["quantity"])
    return tuple(sorted(parts.items()))


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
    with closing(get_connection(db_path)) as conn:
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


def _load_manual_plan_records(db_path):
    """Carrega pedidos que possuem um plano de expedicao salvo.

    Itens e embalagens sao lidos separadamente para que um pedido com varias
    caixas nao duplique os itens durante a comparacao.
    """
    with closing(get_connection(db_path)) as conn:
        item_rows = conn.execute(
            """
            SELECT order_id, product_id, quantity
            FROM order_items
            ORDER BY order_id ASC, product_id ASC
            """
        ).fetchall()
        plan_rows = conn.execute(
            """
            SELECT order_id, box_id, SUM(COALESCE(quantity, 1)) AS quantity
            FROM shipment_history
            GROUP BY order_id, box_id
            ORDER BY order_id ASC, box_id ASC
            """
        ).fetchall()

    records = {}
    for row in item_rows:
        order_id = int(row["order_id"])
        records.setdefault(order_id, {"order_id": order_id, "items": [], "plan": []})
        records[order_id]["items"].append(
            {"product_id": int(row["product_id"]), "quantity": int(row["quantity"])}
        )
    for row in plan_rows:
        order_id = int(row["order_id"])
        if order_id not in records:
            continue
        records[order_id]["plan"].append(
            {"box_id": int(row["box_id"]), "quantity": int(row["quantity"])}
        )

    return [record for record in records.values() if record["items"] and record["plan"]]


def _plan_signature(plan):
    counts = defaultdict(int)
    for row in plan:
        counts[int(row["box_id"])] += int(row["quantity"])
    return tuple(sorted(counts.items()))


def _target_fits_reference_plan(target_signature, reference_signature):
    """Evita recomendar sem ajuste um plano que atendia menos itens."""
    reference_quantities = dict(reference_signature)
    return all(
        product_id in reference_quantities and quantity <= reference_quantities[product_id]
        for product_id, quantity in target_signature
    )



def physical_profile(item, packing_rules=None):
    """Exact physical equivalence including the operational pack family."""
    from packing import _item_bundle_spec, _is_mb_item, _ram_profile
    row = dict(item)
    dims = tuple(sorted(float(row[k]) for k in ('length_cm', 'width_cm', 'height_cm')))
    weight = float(row['weight'])
    if any(not math.isfinite(v) or v <= 0 for v in (*dims, weight)):
        raise ValueError('invalid_physical_profile')
    row['quantity'] = 1000000  # Identify quantity-dependent families consistently.
    spec = _item_bundle_spec(row, packing_rules) or {}
    family = spec.get('profile') or ('mb' if _is_mb_item(row) else _ram_profile(row)) or 'none'
    return dims, weight, family, spec.get('bundle_qty', 1), tuple(spec.get('dims', ()))


def physical_signature(items, rules=None):
    quantities = defaultdict(int)
    for item in items:
        quantities[physical_profile(item, rules)] += int(item['quantity'])
    return tuple(sorted(quantities.items()))


def _remap_packages(packages, products, targets, rules):
    remaining = {int(i['product_id']): int(i['quantity']) for i in targets}
    result = []
    for package in packages:
        assigned = defaultdict(int)
        for pid, qty in package['assignments'].items():
            profile = physical_profile(products[int(pid)], rules)
            for item in targets:
                target_id = int(item['product_id'])
                if physical_profile(item, rules) != profile:
                    continue
                take = min(qty, remaining[target_id])
                assigned[str(target_id)] += take
                remaining[target_id] -= take
                qty -= take
                if not qty:
                    break
        assigned = {pid: qty for pid, qty in assigned.items() if qty}
        if assigned:
            result.append(dict(box_id=package['box_id'], quantity=1, assignments=assigned))
    # Preserve the learned distribution and consider additional products only
    # after assigning their full quantities; the normal capacity validator decides.
    if result and any(remaining.values()):
        for pid, qty in remaining.items():
            if qty:
                result[-1]['assignments'][str(pid)] = result[-1]['assignments'].get(str(pid), 0) + qty
    return result if result else None


def suggest_plan_from_history(db_path, target_items, packing_rules=None, audit=None, validator=None):
    """Rank historical evidence and validate every candidate before voting.

    Legacy plans have no proof of operator confirmation. Assignments are the
    stronger evidence, and are checked for complete coverage before reuse.
    """
    from database import get_packing_rules
    rules = packing_rules if packing_rules is not None else get_packing_rules(db_path)
    audit = audit if audit is not None else []
    targets = [dict(i) for i in target_items]
    if validator is None:
        from database import list_boxes
        boxes = list_boxes(db_path)
        validator = lambda plan: validate_history_plan(plan, targets, boxes, rules)
    if not targets:
        return None
    with closing(get_connection(db_path)) as conn:
        products = {int(r['id']): dict(r) for r in conn.execute('SELECT * FROM products')}
        shipments = conn.execute('SELECT * FROM shipment_history ORDER BY order_id, id').fetchall()
        items = conn.execute('SELECT * FROM order_items').fetchall()
    orders, histories = defaultdict(list), defaultdict(list)
    for row in items:
        pid = int(row['product_id'])
        if pid in products:
            orders[int(row['order_id'])].append(dict(products[pid], product_id=pid, quantity=int(row['quantity'])))
    for row in shipments:
        histories[int(row['order_id'])].append(dict(row))
    try:
        target_physical = physical_signature(targets, rules)
    except (ValueError, KeyError, TypeError) as exc:
        audit.append(dict(reason=str(exc)))
        return None
    candidates = []
    current_orders = {int(i['order_id']) for i in targets if 'order_id' in i}
    for oid, rows in histories.items():
        if oid in current_orders:
            audit.append(dict(order_id=oid, reason='current_order_excluded'))
            continue
        try:
            reference = physical_signature(orders[oid], rules)
            manual = all(r.get('assignments_json') for r in rows)
            packages = []
            if manual:
                totals = defaultdict(int)
                for row in rows:
                    assigned = json.loads(row['assignments_json'])
                    if not isinstance(assigned, dict) or int(row['quantity']) != 1:
                        raise ValueError('invalid_assignments')
                    for pid, qty in assigned.items():
                        if str(int(pid)) != str(pid) or type(qty) is not int or qty < 0:
                            raise ValueError('invalid_assignments')
                        if qty:
                            totals[int(pid)] += qty
                    packages.append(dict(box_id=int(row['box_id']), quantity=1, assignments=assigned))
                if tuple(sorted(totals.items())) != build_signature(orders[oid]):
                    raise ValueError('incomplete_assignments')
            target_map, reference_map = dict(target_physical), dict(reference)
            overlap = tuple((key, qty) for key, qty in target_physical if key in reference_map)
            additional = bool(set(target_map) - set(reference_map))
            if not overlap or not _target_fits_reference_plan(overlap, reference):
                reason = 'quantity_above_reference' if set(dict(target_physical)) <= set(dict(reference)) else 'products_incompatible'
                raise ValueError(reason)
            similarity = _signature_similarity(overlap, reference)
            if similarity < .6:
                raise ValueError('similarity_below_limit')
            exact_sku = build_signature(targets) == build_signature(orders[oid])
            exact_physical = target_physical == reference
            rank = 0 if exact_sku and manual else 1 if exact_physical else 2 if manual and additional else 3
            source = 'history_plan_exact' if exact_sku and manual else 'history_physical_exact' if exact_physical else 'history_manual_distribution' if manual else 'history_physical_similar'
            plan = _remap_packages(packages, products, targets, rules) if manual else [dict(box_id=int(r['box_id']), quantity=int(r['quantity'])) for r in rows]
            if not plan:
                raise ValueError('assignments_incompatible')
            if validator:
                reason = validator(plan)
                if reason:
                    raise ValueError(reason)
            candidates.append(dict(rank=rank, source=source, plan=plan, order_id=oid, manual=manual, score=similarity*(2 if manual else 1)))
            audit.append(dict(order_id=oid, reason='candidate_validated', manual=manual, similarity=round(similarity, 4)))
        except (ValueError, KeyError, TypeError) as exc:
            audit.append(dict(order_id=oid, reason=str(exc)))
    if not candidates:
        return None
    rank = min(c['rank'] for c in candidates)
    candidates = [c for c in candidates if c['rank'] == rank]
    votes = defaultdict(float)
    for c in candidates:
        votes[_plan_signature(c['plan'])] += c['score']
    winner = max(votes, key=lambda key: (votes[key], key))
    selected = next(c for c in candidates if _plan_signature(c['plan']) == winner)
    return dict(plan=selected['plan'], source=selected['source'],
                confidence=min(round(100*votes[winner]/sum(votes.values())), 100 if rank == 0 else 90 if rank == 2 else 95),
                evidence_count=len(candidates), conflict_count=len(votes)-1,
                reference_order_ids=[c['order_id'] for c in candidates],
                manual_evidence_count=sum(c['manual'] for c in candidates), audit=audit)

def validate_history_plan(plan, target_items, boxes, rules=None):
    """Validate capacity with packing.py; mixed boxes require explicit assignments."""
    from packing import calculate_order_totals, estimate_packages_for_box, is_box_dimension_compatible
    if not isinstance(plan, list) or not plan:
        return 'invalid_plan'
    for row in plan:
        if not isinstance(row, dict) or type(row.get('box_id')) is not int or type(row.get('quantity')) is not int:
            return 'invalid_plan'
        if 'assignments' in row and (not isinstance(row['assignments'], dict) or any(
            not str(pid).isdigit() or type(qty) is not int or qty < 0 for pid, qty in row['assignments'].items()
        )):
            return 'invalid_assignments'
    available = {int(b['id']): b for b in boxes if int(dict(b).get('is_active', 1))}
    targets = {int(i['product_id']): dict(i) for i in target_items}
    assigned_totals = defaultdict(int)
    has_assignments = all('assignments' in row for row in plan)
    if not has_assignments and len({row['box_id'] for row in plan}) > 1:
        return 'distribution_unknown'
    for row in plan:
        box = available.get(int(row['box_id']))
        if not box:
            return 'box_missing_or_inactive'
        if int(row['quantity']) <= 0:
            return 'invalid_box_quantity'
        if has_assignments:
            contents = []
            for pid, qty in row['assignments'].items():
                if int(pid) not in targets or type(qty) is not int or qty < 0:
                    return 'invalid_assignments'
                if not qty:
                    continue
                assigned_totals[int(pid)] += qty
                contents.append(dict(targets[int(pid)], quantity=qty))
        else:
            contents = list(targets.values())
        if not contents:
            return 'empty_package'
        if not is_box_dimension_compatible(contents, box, packing_rules=rules):
            return 'physical_or_pack_incompatible'
        totals = calculate_order_totals(contents, packing_rules=rules)
        estimate = estimate_packages_for_box(totals['total_volume_cm3'], totals['total_weight'], box,
                                             order_items=contents, packing_rules=rules)
        capacity = int(row['quantity']) if has_assignments else sum(int(r['quantity']) for r in plan)
        if not estimate or estimate['packages_required'] > capacity:
            return 'physical_or_pack_capacity_exceeded'
    if has_assignments and tuple(sorted(assigned_totals.items())) != build_signature(target_items):
        return 'incomplete_assignments'
    return None
