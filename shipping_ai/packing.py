import math
import re
from collections import defaultdict
from itertools import permutations


SSD_25_BUNDLE_QTY = 10
SSD_25_BUNDLE_LENGTH_CM = 22.0
SSD_25_BUNDLE_WIDTH_CM = 8.0
SSD_25_BUNDLE_HEIGHT_CM = 6.7
SSD_M2_BUNDLE_QTY = 10
SSD_M2_BUNDLE_LENGTH_CM = 21.3
SSD_M2_BUNDLE_WIDTH_CM = 5.0
SSD_M2_BUNDLE_HEIGHT_CM = 9.0
RAM_NB_BUNDLE_QTY = 10
RAM_NB_BUNDLE_MIN_QTY = 10
RAM_DESK_BUNDLE_QTY = 10
RAM_DESK_BUNDLE_MIN_QTY = 10
MB_BUNDLE_QTY = 20
MB_BUNDLE_MIN_QTY = 21
UNPACK_PACKS_SINGLE_VOLUME = 1
MIN_EFFECTIVE_MAX_WEIGHT_KG = 1.0

ACCENT_TRANSLATION = str.maketrans(
    "áàãâäéèêëíìîïóòõôöúùûüç",
    "aaaaaeeeeiiiiooooouuuuc",
)


def _rule_value(packing_rules, key, default):
    if not packing_rules:
        return default
    try:
        return packing_rules.get(key, default)
    except Exception:
        return default


def product_volume_cm3(product_row):
    return (
        float(product_row["length_cm"])
        * float(product_row["width_cm"])
        * float(product_row["height_cm"])
    )


def box_volume_cm3(box_row):
    return (
        float(box_row["length_cm"]) * float(box_row["width_cm"]) * float(box_row["height_cm"])
    )


def _box_sort_key(box_row):
    return (box_volume_cm3(box_row), str(box_row["name"]).lower())


def _row_value(row, key, default=""):
    if hasattr(row, "keys") and key in row.keys():
        return row[key]
    if isinstance(row, dict) and key in row:
        return row[key]
    return default


def _is_ssd_m2_item(item_row):
    name = str(_row_value(item_row, "name", "")).strip().lower()
    sku = str(_row_value(item_row, "sku", "")).strip().lower()
    text = f"{name} {sku}".replace(",", ".")
    has_m2 = "m.2" in text or bool(re.search(r"\bm2\b", text))
    return "ssd" in text and has_m2


def _is_ssd_25_item(item_row):
    name = str(_row_value(item_row, "name", "")).strip().lower()
    sku = str(_row_value(item_row, "sku", "")).strip().lower()
    text = f"{name} {sku}".replace(",", ".")
    return "ssd" in text and "2.5" in text and not _is_ssd_m2_item(item_row)


def _raw_item_dimensions_cm(item_row):
    return [
        float(item_row["length_cm"]),
        float(item_row["width_cm"]),
        float(item_row["height_cm"]),
    ]


def _item_text(item_row):
    name = str(_row_value(item_row, "name", "")).strip().lower()
    sku = str(_row_value(item_row, "sku", "")).strip().lower()
    raw = f" {name} {sku} ".replace(",", ".")
    raw = raw.translate(ACCENT_TRANSLATION)
    cleaned = re.sub(r"[^a-z0-9]+", " ", raw)
    return f" {cleaned.strip()} "


def _item_tokens(item_row):
    return set(re.findall(r"[a-z0-9]+", _item_text(item_row)))


def _is_mb_item(item_row):
    tokens = _item_tokens(item_row)
    if not tokens:
        return False

    if "mb" in tokens:
        return True

    if "motherboard" in tokens or "mainboard" in tokens:
        return True

    if "placa" in tokens and "mae" in tokens:
        return True

    return False


def _ram_profile(item_row):
    tokens = _item_tokens(item_row)
    if not tokens:
        return None

    has_ram_marker = (
        "ram" in tokens
        or "memoria" in tokens
        or "memory" in tokens
        or any(token.startswith("ddr") for token in tokens)
    )
    if not has_ram_marker:
        return None

    is_nb = (
        "nb" in tokens
        or "notebook" in tokens
        or "laptop" in tokens
        or "sodimm" in tokens
        or ("so" in tokens and "dimm" in tokens)
    )
    if is_nb:
        return "ram_nb"

    is_desk = (
        "desk" in tokens
        or "desktop" in tokens
        or "udimm" in tokens
        or "dimm" in tokens
    )
    if is_desk:
        return "ram_desk"

    return None


def _safe_item_qty(item_row):
    try:
        return int(_row_value(item_row, "quantity", 0))
    except Exception:
        return 0


def _item_label(item_row):
    sku = str(_row_value(item_row, "sku", "")).strip()
    name = str(_row_value(item_row, "name", "Item")).strip() or "Item"
    return f"{sku} - {name}" if sku else name


def _bundle_pack_count(item_row, bundle_spec):
    qty = _safe_item_qty(item_row)
    if qty <= 0:
        return 0
    return _bundle_breakdown(
        item_row,
        bundle_spec,
    )["full_packs"]


def _bundle_int_rule(packing_rules, key, default, legacy_key=None):
    value = _rule_value(packing_rules, key, None)
    if value is None and legacy_key:
        value = _rule_value(packing_rules, legacy_key, default)
    if value is None:
        value = default
    try:
        return max(1, int(float(value)))
    except Exception:
        return max(1, int(default))


def _ram_bundle_dimensions_cm(item_row, bundle_qty):
    dims = sorted(_raw_item_dimensions_cm(item_row))
    smallest, middle, largest = dims[0], dims[1], dims[2]
    stacked_height = smallest * float(bundle_qty)
    return [largest, middle, stacked_height]


def _mb_bundle_dimensions_cm(item_row, packing_rules=None):
    mb_length = _rule_value(packing_rules, "mb_box_length_cm", None)
    mb_width = _rule_value(packing_rules, "mb_box_width_cm", None)
    mb_height = _rule_value(packing_rules, "mb_box_height_cm", None)

    try:
        if mb_length is not None and mb_width is not None and mb_height is not None:
            length = float(mb_length)
            width = float(mb_width)
            height = float(mb_height)
            if length > 0 and width > 0 and height > 0:
                return [length, width, height]
    except Exception:
        pass

    return _raw_item_dimensions_cm(item_row)


def _item_bundle_spec(item_row, packing_rules=None):
    if _is_ssd_m2_item(item_row):
        bundle_qty = int(_rule_value(packing_rules, "ssdm2_bundle_qty", SSD_M2_BUNDLE_QTY))
        bundle_qty = max(1, bundle_qty)
        return {
            "label": "SSD M.2",
            "profile": "ssdm2",
            "bundle_qty": bundle_qty,
            "dims": [
                float(_rule_value(packing_rules, "ssdm2_bundle_length_cm", SSD_M2_BUNDLE_LENGTH_CM)),
                float(_rule_value(packing_rules, "ssdm2_bundle_width_cm", SSD_M2_BUNDLE_WIDTH_CM)),
                float(_rule_value(packing_rules, "ssdm2_bundle_height_cm", SSD_M2_BUNDLE_HEIGHT_CM)),
            ],
        }

    if _is_ssd_25_item(item_row):
        bundle_qty = int(_rule_value(packing_rules, "ssd25_bundle_qty", SSD_25_BUNDLE_QTY))
        bundle_qty = max(1, bundle_qty)
        return {
            "label": "SSD 2.5",
            "profile": "ssd25",
            "bundle_qty": bundle_qty,
            "dims": [
                float(
                    _rule_value(
                        packing_rules,
                        "ssd25_bundle_length_cm",
                        SSD_25_BUNDLE_LENGTH_CM,
                    )
                ),
                float(
                    _rule_value(
                        packing_rules,
                        "ssd25_bundle_width_cm",
                        SSD_25_BUNDLE_WIDTH_CM,
                    )
                ),
                float(
                    _rule_value(
                        packing_rules,
                        "ssd25_bundle_height_cm",
                        SSD_25_BUNDLE_HEIGHT_CM,
                    )
                ),
            ],
        }

    if _is_mb_item(item_row):
        qty = _safe_item_qty(item_row)
        bundle_qty = _bundle_int_rule(
            packing_rules,
            "mb_bundle_qty",
            MB_BUNDLE_QTY,
        )
        min_qty = _bundle_int_rule(
            packing_rules,
            "mb_bundle_min_qty",
            MB_BUNDLE_MIN_QTY,
        )

        if qty >= min_qty:
            return {
                "label": "MB",
                "profile": "mb",
                "bundle_qty": bundle_qty,
                "dims": _mb_bundle_dimensions_cm(item_row, packing_rules=packing_rules),
            }

    profile = _ram_profile(item_row)
    if not profile:
        return None

    qty = _safe_item_qty(item_row)
    if profile == "ram_nb":
        bundle_qty = _bundle_int_rule(
            packing_rules,
            "ram_nb_bundle_qty",
            RAM_NB_BUNDLE_QTY,
            legacy_key="ram_bundle_qty",
        )
        min_qty = _bundle_int_rule(
            packing_rules,
            "ram_nb_bundle_min_qty",
            RAM_NB_BUNDLE_MIN_QTY,
            legacy_key="ram_bundle_min_qty",
        )
        label = "RAM NB"
    else:
        bundle_qty = _bundle_int_rule(
            packing_rules,
            "ram_desk_bundle_qty",
            RAM_DESK_BUNDLE_QTY,
            legacy_key="ram_bundle_qty",
        )
        min_qty = _bundle_int_rule(
            packing_rules,
            "ram_desk_bundle_min_qty",
            RAM_DESK_BUNDLE_MIN_QTY,
            legacy_key="ram_bundle_min_qty",
        )
        label = "RAM DESK"

    if qty < min_qty:
        return None

    return {
        "label": label,
        "bundle_qty": bundle_qty,
        "dims": _ram_bundle_dimensions_cm(item_row, bundle_qty),
        "profile": profile,
    }


def _ram_profile_threshold(profile, packing_rules=None):
    if profile == "ram_nb":
        return _bundle_int_rule(
            packing_rules,
            "ram_nb_bundle_min_qty",
            RAM_NB_BUNDLE_MIN_QTY,
            legacy_key="ram_bundle_min_qty",
        )

    if profile == "ram_desk":
        return _bundle_int_rule(
            packing_rules,
            "ram_desk_bundle_min_qty",
            RAM_DESK_BUNDLE_MIN_QTY,
            legacy_key="ram_bundle_min_qty",
        )

    return 0


def _bundle_min_qty(bundle_spec, packing_rules=None):
    profile = str(bundle_spec.get("profile") or "")
    bundle_qty = max(1, int(bundle_spec.get("bundle_qty") or 1))

    if profile in {"ssd25", "ssdm2"}:
        return bundle_qty

    if profile == "mb":
        return _bundle_int_rule(
            packing_rules,
            "mb_bundle_min_qty",
            MB_BUNDLE_MIN_QTY,
        )

    if profile in {"ram_nb", "ram_desk"}:
        return max(1, _ram_profile_threshold(profile, packing_rules=packing_rules))

    return bundle_qty


def _bundle_breakdown(item_row, bundle_spec, packing_rules=None):
    qty = _safe_item_qty(item_row)
    bundle_qty = max(1, int(bundle_spec.get("bundle_qty") or 1))
    min_qty = max(1, _bundle_min_qty(bundle_spec, packing_rules=packing_rules))

    if qty <= 0 or qty < min_qty or qty < bundle_qty:
        return {
            "qty": qty,
            "bundle_qty": bundle_qty,
            "min_qty": min_qty,
            "full_packs": 0,
            "loose_units": max(0, qty),
            "uses_pack": False,
        }

    full_packs = qty // bundle_qty
    loose_units = qty % bundle_qty
    return {
        "qty": qty,
        "bundle_qty": bundle_qty,
        "min_qty": min_qty,
        "full_packs": int(full_packs),
        "loose_units": int(loose_units),
        "uses_pack": full_packs > 0,
    }


def describe_item_pack_profile(item_row, packing_rules=None):
    qty = _safe_item_qty(item_row)
    bundle_spec = _item_bundle_spec(item_row, packing_rules=packing_rules)

    if bundle_spec:
        breakdown = _bundle_breakdown(
            item_row,
            bundle_spec,
            packing_rules=packing_rules,
        )
        profile = bundle_spec.get("profile") or "bundle"
        pack_labels = {
            "ssdm2": "SSD M.2 (pack)",
            "ssd25": "SSD 2.5 (pack)",
            "mb": "MB (pack)",
            "ram_nb": "RAM NB (pack)",
            "ram_desk": "RAM DESK (pack)",
        }
        no_pack_labels = {
            "ssdm2": "SSD M.2 (sem pack)",
            "ssd25": "SSD 2.5 (sem pack)",
            "mb": "MB (sem pack)",
            "ram_nb": "RAM NB (sem pack)",
            "ram_desk": "RAM DESK (sem pack)",
        }

        if breakdown["uses_pack"]:
            detail_parts = []
            if breakdown["full_packs"] > 0:
                detail_parts.append(
                    f"{breakdown['full_packs']} pack(s) de {breakdown['bundle_qty']} und"
                )
            if breakdown["loose_units"] > 0:
                detail_parts.append(f"{breakdown['loose_units']} unidade(s)")

            return {
                "profile": profile,
                "label": pack_labels.get(profile, "Pack"),
                "detail": " + ".join(detail_parts) if detail_parts else "0 und",
                "uses_pack": True,
            }

        return {
            "profile": profile,
            "label": no_pack_labels.get(profile, "Sem pack"),
            "detail": f"pack completo: {breakdown['bundle_qty']} und | qtd atual: {qty}",
            "uses_pack": False,
        }

    if _is_mb_item(item_row):
        min_qty = _bundle_int_rule(
            packing_rules,
            "mb_bundle_min_qty",
            MB_BUNDLE_MIN_QTY,
        )
        return {
            "profile": "mb",
            "label": "MB (sem pack)",
            "detail": f"min para pack: {min_qty} und | qtd atual: {qty}",
            "uses_pack": False,
        }

    ram_profile = _ram_profile(item_row)
    if ram_profile in {"ram_nb", "ram_desk"}:
        min_qty = _ram_profile_threshold(ram_profile, packing_rules=packing_rules)
        labels = {
            "ram_nb": "RAM NB (sem pack)",
            "ram_desk": "RAM DESK (sem pack)",
        }
        return {
            "profile": ram_profile,
            "label": labels[ram_profile],
            "detail": f"min para pack: {min_qty} und | qtd atual: {qty}",
            "uses_pack": False,
        }

    return {
        "profile": "none",
        "label": "Sem pack",
        "detail": "",
        "uses_pack": False,
    }


def _item_dimensions_cm(item_row, packing_rules=None):
    bundle_spec = _item_bundle_spec(item_row, packing_rules=packing_rules)
    if bundle_spec and _bundle_breakdown(
        item_row,
        bundle_spec,
        packing_rules=packing_rules,
    )["uses_pack"]:
        return bundle_spec["dims"]
    return _raw_item_dimensions_cm(item_row)


def _effective_volume_for_item(item_row, packing_rules=None):
    qty = int(item_row["quantity"])
    if qty <= 0:
        return 0.0

    bundle_spec = _item_bundle_spec(item_row, packing_rules=packing_rules)
    if bundle_spec:
        breakdown = _bundle_breakdown(
            item_row,
            bundle_spec,
            packing_rules=packing_rules,
        )
        bundle_dims = bundle_spec["dims"]
        bundle_volume = bundle_dims[0] * bundle_dims[1] * bundle_dims[2]
        raw_volume = product_volume_cm3(item_row)
        return (breakdown["full_packs"] * bundle_volume) + (breakdown["loose_units"] * raw_volume)

    return product_volume_cm3(item_row) * qty


def _max_units_per_box(unit_dims, box_dims):
    max_units = 0
    for oriented_unit in set(permutations(unit_dims, 3)):
        units = 1
        for unit_side, box_side in zip(oriented_unit, box_dims):
            units *= int(math.floor(float(box_side) / float(unit_side)))
        if units > max_units:
            max_units = units
    return max_units


def _dimension_units_needed(item_row, packing_rules=None):
    qty = int(item_row["quantity"])
    if qty <= 0:
        return 0

    bundle_spec = _item_bundle_spec(item_row, packing_rules=packing_rules)
    if not bundle_spec:
        return qty

    breakdown = _bundle_breakdown(
        item_row,
        bundle_spec,
        packing_rules=packing_rules,
    )
    return int(breakdown["full_packs"] + breakdown["loose_units"])


def _dimension_packages_required(order_items, box_row, packing_rules=None):
    relevant_items = [item for item in order_items if int(item["quantity"]) > 0]
    if not relevant_items:
        return 1

    box_dims = [
        float(box_row["length_cm"]),
        float(box_row["width_cm"]),
        float(box_row["height_cm"]),
    ]

    packages_required = 1
    for item in relevant_items:
        item_groups, _ = _build_preview_groups_for_item(
            item,
            packing_rules=packing_rules,
            unpack_packs=0,
        )
        for group in item_groups:
            units_per_box = _max_units_per_box(group["dims"], box_dims)
            if units_per_box <= 0:
                return None

            group_packages = int(math.ceil(group["blocks"] / units_per_box))
            packages_required = max(packages_required, group_packages)

    return packages_required


def _sorted_dimensions(row, packing_rules=None):
    dims = _item_dimensions_cm(row, packing_rules=packing_rules)
    return sorted(dims)


def item_fits_in_box(product_row, box_row, packing_rules=None):
    """
    Checks physical fit allowing rotation of the product.
    """
    product_dims = _sorted_dimensions(product_row, packing_rules=packing_rules)
    box_dims = _sorted_dimensions(box_row)
    return all(p <= b for p, b in zip(product_dims, box_dims))


def is_box_dimension_compatible(order_items, box_row, packing_rules=None):
    """
    Returns True only if every product in the order can physically fit
    (considering rotation) inside the box.
    """
    for item in order_items:
        qty = int(item["quantity"])
        if qty <= 0:
            continue
        if not item_fits_in_box(item, box_row, packing_rules=packing_rules):
            return False
    return True


def calculate_order_totals(order_items, packing_rules=None):
    total_volume = 0.0
    total_weight = 0.0
    total_units = 0

    for item in order_items:
        qty = int(item["quantity"])
        total_units += qty
        total_volume += _effective_volume_for_item(item, packing_rules=packing_rules)
        total_weight += float(item["weight"]) * qty

    return {
        "total_volume_cm3": total_volume,
        "total_weight": total_weight,
        "total_units": total_units,
    }


def estimate_packages_for_box(
    total_volume_cm3,
    total_weight,
    box_row,
    fill_factor=0.9,
    order_items=None,
    packing_rules=None,
):
    raw_volume = box_volume_cm3(box_row)
    usable_volume = raw_volume * float(fill_factor)
    max_weight = float(box_row["max_weight"])

    if usable_volume <= 0 or max_weight <= 0:
        return None

    by_volume = int(math.ceil(total_volume_cm3 / usable_volume)) if total_volume_cm3 > 0 else 1

    # Some imported templates store the box own weight instead of carrying capacity.
    min_effective_weight = float(
        _rule_value(
            packing_rules,
            "min_effective_max_weight_kg",
            MIN_EFFECTIVE_MAX_WEIGHT_KG,
        )
    )
    if max_weight > min_effective_weight:
        by_weight = int(math.ceil(total_weight / max_weight)) if total_weight > 0 else 1
    else:
        by_weight = 1

    by_dimension = 1
    if order_items:
        by_dimension = _dimension_packages_required(
            order_items,
            box_row,
            packing_rules=packing_rules,
        )
        if by_dimension is None:
            return None

    packages_required = max(1, by_volume, by_weight, by_dimension)

    unpack_applied = False
    unpack_plan = {}
    unpack_notes = []
    if order_items and packages_required > 1 and by_weight <= 1:
        unpack_scenario = _try_single_volume_with_unpack(
            order_items,
            box_row,
            packing_rules=packing_rules,
            max_blocks=1,
        )
        if unpack_scenario:
            packages_required = 1
            unpack_applied = True
            unpack_plan = unpack_scenario.get("unpack_plan", {})
            unpack_notes = list(unpack_scenario.get("notes") or [])

    return {
        "packages_required": packages_required,
        "usable_volume_cm3": usable_volume,
        "raw_volume_cm3": raw_volume,
        "unpack_applied": unpack_applied,
        "unpack_plan": unpack_plan,
        "unpack_notes": unpack_notes,
    }


def _tested_box_entry(box, estimate):
    if estimate["packages_required"] == 1:
        status = "coube_desmontando_pack" if estimate.get("unpack_applied") else "coube"
    else:
        status = "nao_coube_quantidade"

    return {
        "name": box["name"],
        "packages_required": estimate["packages_required"],
        "status": status,
        "note": (estimate.get("unpack_notes") or [""])[0],
    }


def choose_best_box(
    total_volume_cm3,
    total_weight,
    boxes,
    order_items=None,
    packing_rules=None,
):
    """
    Rules:
    1) If a single box can fit, choose the smallest compatible box.
    2) Otherwise choose the box that needs fewer packages.
    3) Tie-break by smallest box volume.
    """
    options = []
    tested_boxes = []

    sorted_boxes = sorted(boxes, key=_box_sort_key)

    for box in sorted_boxes:
        if order_items and not is_box_dimension_compatible(
            order_items,
            box,
            packing_rules=packing_rules,
        ):
            tested_boxes.append(
                {
                    "name": box["name"],
                    "packages_required": None,
                    "status": "nao_coube_dimensao",
                }
            )
            continue

        estimate = estimate_packages_for_box(
            total_volume_cm3,
            total_weight,
            box,
            order_items=order_items,
            packing_rules=packing_rules,
        )
        if not estimate:
            tested_boxes.append(
                {
                    "name": box["name"],
                    "packages_required": None,
                    "status": "nao_coube",
                }
            )
            continue

        tested_boxes.append(_tested_box_entry(box, estimate))

        options.append(
            {
                "box": box,
                "packages_required": estimate["packages_required"],
                "raw_volume_cm3": estimate["raw_volume_cm3"],
                "usable_volume_cm3": estimate["usable_volume_cm3"],
                "unpack_applied": bool(estimate.get("unpack_applied")),
                "unpack_plan": estimate.get("unpack_plan") or {},
                "unpack_notes": list(estimate.get("unpack_notes") or []),
            }
        )

    if not options:
        return None

    # First-fit strategy from smallest to largest.
    for opt in options:
        if opt["packages_required"] == 1:
            picked = dict(opt)
            picked["reason"] = "algorithm_small_to_large_first_fit"
            picked["confidence"] = 65
            picked["tested_boxes"] = tested_boxes
            return picked

    options.sort(key=lambda x: (x["packages_required"], x["raw_volume_cm3"]))
    picked = options[0]
    picked["reason"] = "algorithm_min_packages"
    picked["confidence"] = 45
    picked["tested_boxes"] = tested_boxes
    return picked


def _stable_color_hex(text):
    text_value = str(text or "")
    seed = 0
    for idx, ch in enumerate(text_value):
        seed += (idx + 1) * ord(ch)

    r = 70 + ((seed * 37) % 150)
    g = 70 + ((seed * 57) % 150)
    b = 70 + ((seed * 79) % 150)
    return f"#{r:02x}{g:02x}{b:02x}"


def _orientation_for_space(item_dims, free_length, free_width, free_height):
    candidates = []
    for dims in set(permutations(item_dims, 3)):
        length, width, height = dims
        if length <= free_length and width <= free_width and height <= free_height:
            waste = (
                (free_length - length)
                + (free_width - width)
                + ((free_height - height) * 0.35)
            )
            candidates.append((waste, dims))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0])
    return [float(v) for v in candidates[0][1]]


def _preview_orientation_candidates(item_dims, box_dims, point, mode="balanced"):
    pos_x, pos_y, pos_z = point
    free_length = float(box_dims[0]) - float(pos_x)
    free_width = float(box_dims[1]) - float(pos_y)
    free_height = float(box_dims[2]) - float(pos_z)
    if free_length <= 0 or free_width <= 0 or free_height <= 0:
        return []

    candidates = []
    for dims in set(permutations(item_dims, 3)):
        length, width, height = [float(v) for v in dims]
        if length > free_length or width > free_width or height > free_height:
            continue

        footprint = length * width
        waste = (
            (free_length - length)
            + (free_width - width)
            + ((free_height - height) * 0.35)
        )

        if mode == "narrow_row":
            key = (width, footprint, length, -height, waste)
        elif mode == "footprint":
            key = (footprint, width, length, -height, waste)
        else:
            key = (waste, width, length, -height)

        candidates.append((key, (length, width, height)))

    candidates.sort(key=lambda x: x[0])
    return [dims for _, dims in candidates]


def _preview_point_sort_key(point, mode="z_y_x"):
    x, y, z = point
    if mode == "y_x_z":
        return (y, x, z)
    if mode == "x_y_z":
        return (x, y, z)
    return (z, y, x)


def _preview_boxes_overlap(a, b):
    return not (
        (a["x"] + a["length_cm"] <= b["x"])
        or (b["x"] + b["length_cm"] <= a["x"])
        or (a["y"] + a["width_cm"] <= b["y"])
        or (b["y"] + b["width_cm"] <= a["y"])
        or (a["z"] + a["height_cm"] <= b["z"])
        or (b["z"] + b["height_cm"] <= a["z"])
    )


def _preview_point_inside_block(point, block):
    x, y, z = point
    return (
        (block["x"] < x < (block["x"] + block["length_cm"]))
        and (block["y"] < y < (block["y"] + block["width_cm"]))
        and (block["z"] < z < (block["z"] + block["height_cm"]))
    )


def _preview_prune_points(points):
    unique_points = sorted(set(points))
    pruned = []
    for idx, point in enumerate(unique_points):
        dominated = False
        for jdx, other in enumerate(unique_points):
            if idx == jdx:
                continue

            if (
                other[0] <= point[0]
                and other[1] <= point[1]
                and other[2] <= point[2]
                and (
                    other[0] < point[0]
                    or other[1] < point[1]
                    or other[2] < point[2]
                )
            ):
                dominated = True
                break

        if not dominated:
            pruned.append(point)

    return pruned


def _preview_cleanup_points(points, occupied, box_dims):
    cleaned = []
    box_length, box_width, box_height = box_dims
    for raw_point in points:
        x, y, z = [float(v) for v in raw_point]

        if x < 0 or y < 0 or z < 0:
            continue
        if x > box_length or y > box_width or z > box_height:
            continue

        point = (round(x, 6), round(y, 6), round(z, 6))
        if any(_preview_point_inside_block(point, block) for block in occupied):
            continue

        cleaned.append(point)

    return _preview_prune_points(cleaned)


def _try_place_at(item_dims, box_dims, position):
    box_length, box_width, box_height = box_dims
    pos_x, pos_y, pos_z = position

    free_length = box_length - pos_x
    free_width = box_width - pos_y
    free_height = box_height - pos_z
    if free_length <= 0 or free_width <= 0 or free_height <= 0:
        return None

    return _orientation_for_space(item_dims, free_length, free_width, free_height)


def _attempt_next_row(item_dims, state, box_dims):
    _, y, z, row_depth, layer_height = state
    if row_depth <= 0:
        return None

    next_row_y = y + row_depth
    if next_row_y >= box_dims[1]:
        return None

    orientation = _try_place_at(item_dims, box_dims, (0.0, next_row_y, z))
    if orientation is None:
        return None

    return orientation, (0.0, next_row_y, z, 0.0, layer_height)


def _attempt_next_layer(item_dims, state, box_dims):
    _, _, z, _, layer_height = state
    if layer_height <= 0:
        return None

    next_layer_z = z + layer_height
    if next_layer_z >= box_dims[2]:
        return None

    orientation = _try_place_at(item_dims, box_dims, (0.0, 0.0, next_layer_z))
    if orientation is None:
        return None

    return orientation, (0.0, 0.0, next_layer_z, 0.0, 0.0)


def _place_block_shelf(item_dims, state, box_dims):
    x, y, z, row_depth, layer_height = state
    orientation = _try_place_at(item_dims, box_dims, (x, y, z))
    placement_state = state

    if orientation is None:
        row_attempt = _attempt_next_row(item_dims, state, box_dims)
        if row_attempt is not None:
            orientation, placement_state = row_attempt

    if orientation is None:
        layer_attempt = _attempt_next_layer(item_dims, state, box_dims)
        if layer_attempt is not None:
            orientation, placement_state = layer_attempt

    if orientation is None:
        return None, state

    x, y, z, row_depth, layer_height = placement_state
    length, width, height = orientation

    placed = {
        "x": float(x),
        "y": float(y),
        "z": float(z),
        "length_cm": float(length),
        "width_cm": float(width),
        "height_cm": float(height),
    }

    next_state = (
        float(x + length),
        float(y),
        float(z),
        float(max(row_depth, width)),
        float(max(layer_height, height)),
    )
    return placed, next_state


def _build_preview_groups_for_item(item, packing_rules=None, unpack_packs=0):
    raw_qty = _safe_item_qty(item)
    if raw_qty <= 0:
        return [], []

    label = _item_label(item)
    color = _stable_color_hex(label)
    bundle_spec = _item_bundle_spec(item, packing_rules=packing_rules)

    if not bundle_spec:
        dims = _raw_item_dimensions_cm(item)
        return (
            [
                {
                    "label": label,
                    "dims": [float(dims[0]), float(dims[1]), float(dims[2])],
                    "blocks": int(raw_qty),
                    "raw_quantity": raw_qty,
                    "color": color,
                }
            ],
            [],
        )

    breakdown = _bundle_breakdown(
        item,
        bundle_spec,
        packing_rules=packing_rules,
    )
    if not breakdown["uses_pack"]:
        raw_dims = _raw_item_dimensions_cm(item)
        return (
            [
                {
                    "label": label,
                    "dims": [float(raw_dims[0]), float(raw_dims[1]), float(raw_dims[2])],
                    "blocks": int(raw_qty),
                    "raw_quantity": raw_qty,
                    "color": color,
                }
            ],
            [f"{bundle_spec['label']} sem pack: {raw_qty} unidade(s) soltas."],
        )

    total_packs = breakdown["full_packs"]
    unpack_packs = max(0, min(int(unpack_packs), total_packs))
    packed_packs = max(0, total_packs - unpack_packs)
    loose_units = max(0, raw_qty - (packed_packs * bundle_spec["bundle_qty"]))

    groups = []
    bundle_dims = bundle_spec["dims"]
    raw_dims = _raw_item_dimensions_cm(item)

    if packed_packs > 0:
        groups.append(
            {
                "label": label,
                "dims": [float(bundle_dims[0]), float(bundle_dims[1]), float(bundle_dims[2])],
                "blocks": int(packed_packs),
                "raw_quantity": raw_qty,
                "color": color,
            }
        )

    if loose_units > 0:
        groups.append(
            {
                "label": label,
                "dims": [float(raw_dims[0]), float(raw_dims[1]), float(raw_dims[2])],
                "blocks": int(loose_units),
                "raw_quantity": raw_qty,
                "color": color,
            }
        )

    if unpack_packs > 0:
        note = (
            f"{bundle_spec['label']} com desmontagem: {raw_qty} unidade(s) "
            f"=> {packed_packs} pack(s) + {loose_units} unidade(s) soltas."
        )
    elif loose_units > 0:
        note = (
            f"{bundle_spec['label']} tratado em packs: {raw_qty} unidade(s) "
            f"=> {packed_packs} pack(s) + {loose_units} unidade(s) soltas."
        )
    else:
        note = (
            f"{bundle_spec['label']} tratado em packs: {raw_qty} unidade(s) "
            f"=> {packed_packs} pack(s) de {bundle_spec['bundle_qty']}."
        )

    return groups, [note]


def _collect_preview_groups(order_items, packing_rules=None, unpack_plan=None):
    groups = []
    notes = []
    unpack_plan = unpack_plan or {}

    for idx, item in enumerate(order_items):
        unpack_packs = int(unpack_plan.get(idx, 0))
        item_groups, item_notes = _build_preview_groups_for_item(
            item,
            packing_rules=packing_rules,
            unpack_packs=unpack_packs,
        )
        groups.extend(item_groups)
        notes.extend(item_notes)

    return groups, notes


def _expand_preview_blocks(item_groups):
    blocks_to_place = []
    for group in item_groups:
        volume = group["dims"][0] * group["dims"][1] * group["dims"][2]
        for _ in range(group["blocks"]):
            blocks_to_place.append(
                {
                    "label": group["label"],
                    "dims": group["dims"],
                    "color": group["color"],
                    "volume": volume,
                }
            )

    blocks_to_place.sort(key=lambda b: b["volume"], reverse=True)
    return blocks_to_place


def _simulate_preview(blocks_to_place, box_dims, max_blocks):
    placements = []
    overflow_blocks = 0
    total_placed_blocks = 0
    used_volume_cm3 = 0.0

    state = (0.0, 0.0, 0.0, 0.0, 0.0)
    for block in blocks_to_place:
        placed, state = _place_block_shelf(block["dims"], state, box_dims)
        if placed is None:
            overflow_blocks += 1
            continue

        total_placed_blocks += 1
        used_volume_cm3 += (
            placed["length_cm"] * placed["width_cm"] * placed["height_cm"]
        )

        if len(placements) >= int(max_blocks):
            continue

        placed["label"] = block["label"]
        placed["color"] = block["color"]
        placements.append(placed)

    return {
        "placements": placements,
        "overflow_blocks": overflow_blocks,
        "placed_blocks": total_placed_blocks,
        "used_volume_cm3": used_volume_cm3,
    }


def _simulate_preview_candidates(
    blocks_to_place,
    box_dims,
    max_blocks,
    point_mode="z_y_x",
    orientation_mode="balanced",
):
    stats = {
        "placements": [],
        "occupied": [],
        "overflow_blocks": 0,
        "placed_blocks": 0,
        "used_volume_cm3": 0.0,
    }

    points = [(0.0, 0.0, 0.0)]
    for block in blocks_to_place:
        placed = _preview_try_place_block(
            block["dims"],
            points,
            box_dims,
            stats["occupied"],
            point_mode=point_mode,
            orientation_mode=orientation_mode,
        )
        if placed is None:
            stats["overflow_blocks"] += 1
            continue

        stats["occupied"].append(placed)
        stats["placed_blocks"] += 1
        stats["used_volume_cm3"] += (
            placed["length_cm"] * placed["width_cm"] * placed["height_cm"]
        )

        if len(stats["placements"]) < int(max_blocks):
            visible_entry = dict(placed)
            visible_entry["label"] = block["label"]
            visible_entry["color"] = block["color"]
            stats["placements"].append(visible_entry)

        points = _preview_update_points(points, placed, stats["occupied"], box_dims)

    return {
        "placements": stats["placements"],
        "overflow_blocks": stats["overflow_blocks"],
        "placed_blocks": stats["placed_blocks"],
        "used_volume_cm3": stats["used_volume_cm3"],
    }


def _preview_try_place_block(
    item_dims,
    points,
    box_dims,
    occupied_placements,
    point_mode="z_y_x",
    orientation_mode="balanced",
):
    ordered_points = sorted(points, key=lambda p: _preview_point_sort_key(p, point_mode))
    for point in ordered_points:
        orientations = _preview_orientation_candidates(
            item_dims,
            box_dims,
            point,
            mode=orientation_mode,
        )
        if not orientations:
            continue

        for length, width, height in orientations:
            candidate = {
                "x": float(point[0]),
                "y": float(point[1]),
                "z": float(point[2]),
                "length_cm": float(length),
                "width_cm": float(width),
                "height_cm": float(height),
            }
            if any(_preview_boxes_overlap(candidate, other) for other in occupied_placements):
                continue

            return candidate

    return None


def _preview_update_points(points, placed, occupied_placements, box_dims):
    candidate_points = set(points)
    candidate_points.discard((placed["x"], placed["y"], placed["z"]))
    candidate_points.add(
        (
            placed["x"] + placed["length_cm"],
            placed["y"],
            placed["z"],
        )
    )
    candidate_points.add(
        (
            placed["x"],
            placed["y"] + placed["width_cm"],
            placed["z"],
        )
    )
    candidate_points.add(
        (
            placed["x"],
            placed["y"],
            placed["z"] + placed["height_cm"],
        )
    )

    return _preview_cleanup_points(candidate_points, occupied_placements, box_dims)


def _preview_better_candidate(candidate, best):
    if best is None:
        return True

    if candidate["placed_blocks"] > best["placed_blocks"]:
        return True

    if candidate["placed_blocks"] < best["placed_blocks"]:
        return False

    return candidate["used_volume_cm3"] > best["used_volume_cm3"]


def _preview_shelf_orders(blocks_to_place):
    return [
        sorted(blocks_to_place, key=lambda b: b["volume"], reverse=True),
        sorted(blocks_to_place, key=lambda b: b["volume"]),
        sorted(blocks_to_place, key=lambda b: (max(b["dims"]), b["volume"]), reverse=True),
        sorted(blocks_to_place, key=lambda b: (min(b["dims"]), b["volume"]), reverse=True),
    ]


def _preview_candidate_orders(blocks_to_place):
    return [
        sorted(blocks_to_place, key=lambda b: b["volume"], reverse=True),
        sorted(blocks_to_place, key=lambda b: b["volume"]),
        sorted(blocks_to_place, key=lambda b: (max(b["dims"]), b["volume"]), reverse=True),
    ]


def _evaluate_preview_shelf_orders(orders, box_dims, max_blocks):
    best = None
    for scenario_blocks in orders:
        candidate = _simulate_preview(scenario_blocks, box_dims, max_blocks)
        if _preview_better_candidate(candidate, best):
            best = candidate
    return best


def _evaluate_preview_candidate_orders(orders, box_dims, max_blocks, baseline_best, target_count):
    best = baseline_best
    orientation_modes = ["balanced", "narrow_row", "footprint"]
    point_modes = ["z_y_x", "y_x_z"]

    for scenario_blocks in orders:
        for orientation_mode in orientation_modes:
            for point_mode in point_modes:
                candidate = _simulate_preview_candidates(
                    scenario_blocks,
                    box_dims,
                    max_blocks,
                    point_mode=point_mode,
                    orientation_mode=orientation_mode,
                )
                if _preview_better_candidate(candidate, best):
                    best = candidate

                if best and best["placed_blocks"] >= target_count:
                    return best

    return best


def _best_preview_simulation(blocks_to_place, box_dims, max_blocks):
    if not blocks_to_place:
        return {
            "placements": [],
            "overflow_blocks": 0,
            "placed_blocks": 0,
            "used_volume_cm3": 0.0,
        }

    target_count = len(blocks_to_place)
    best = _evaluate_preview_shelf_orders(
        _preview_shelf_orders(blocks_to_place),
        box_dims,
        max_blocks,
    )

    if best and best["placed_blocks"] >= target_count:
        return best

    return _evaluate_preview_candidate_orders(
        _preview_candidate_orders(blocks_to_place),
        box_dims,
        max_blocks,
        best,
        target_count,
    )


def _candidate_unpack_plans(order_items, packing_rules=None, max_unpack_packs=1):
    plans = [{}]
    if max_unpack_packs <= 0:
        return plans

    for idx, item in enumerate(order_items):
        bundle_spec = _item_bundle_spec(item, packing_rules=packing_rules)
        if not bundle_spec:
            continue

        packs = _bundle_pack_count(item, bundle_spec)
        if packs <= 0:
            continue

        plans.append({idx: min(int(max_unpack_packs), packs)})

    return plans


def _simulate_layout_for_box(
    order_items,
    box_row,
    packing_rules=None,
    max_blocks=180,
    unpack_plan=None,
):
    if not box_row:
        return None

    item_groups, notes = _collect_preview_groups(
        order_items,
        packing_rules=packing_rules,
        unpack_plan=unpack_plan,
    )

    blocks_to_place = _expand_preview_blocks(item_groups)
    box_dims = (
        float(box_row["length_cm"]),
        float(box_row["width_cm"]),
        float(box_row["height_cm"]),
    )
    simulation = _best_preview_simulation(
        blocks_to_place,
        box_dims,
        max(1, int(max_blocks)),
    )

    return {
        "item_groups": item_groups,
        "notes": notes,
        "blocks_to_place": blocks_to_place,
        "total_blocks": len(blocks_to_place),
        "simulation": simulation,
        "unpack_plan": unpack_plan or {},
        "unpack_applied": bool(unpack_plan),
    }


def _try_single_volume_with_unpack(order_items, box_row, packing_rules=None, max_blocks=180):
    max_unpack_packs = int(
        _rule_value(
            packing_rules,
            "unpack_packs_single_volume",
            UNPACK_PACKS_SINGLE_VOLUME,
        )
    )
    max_unpack_packs = max(0, max_unpack_packs)
    if max_unpack_packs <= 0:
        return None

    plans = _candidate_unpack_plans(
        order_items,
        packing_rules=packing_rules,
        max_unpack_packs=max_unpack_packs,
    )

    for unpack_plan in plans[1:]:
        scenario = _simulate_layout_for_box(
            order_items,
            box_row,
            packing_rules=packing_rules,
            max_blocks=max_blocks,
            unpack_plan=unpack_plan,
        )
        if not scenario:
            continue

        if scenario["simulation"]["placed_blocks"] >= scenario["total_blocks"]:
            return scenario

    return None


def _preview_legend_from_groups(item_groups):
    legend_map = {}
    legend_order = []

    for group in item_groups:
        label = group["label"]
        if label not in legend_map:
            legend_map[label] = {
                "label": label,
                "color": group["color"],
                "blocks": 0,
                "original_quantity": int(group["raw_quantity"]),
            }
            legend_order.append(label)

        legend_map[label]["blocks"] += int(group["blocks"])

    return [legend_map[label] for label in legend_order]


def _select_preferred_preview_scenario(base_scenario, unpack_scenario):
    if not unpack_scenario:
        return base_scenario

    base_fit = base_scenario["simulation"]["placed_blocks"] >= base_scenario["total_blocks"]
    unpack_fit = unpack_scenario["simulation"]["placed_blocks"] >= unpack_scenario["total_blocks"]
    if unpack_fit and not base_fit:
        return unpack_scenario

    if (
        unpack_scenario["simulation"]["placed_blocks"]
        > base_scenario["simulation"]["placed_blocks"]
    ):
        return unpack_scenario

    return base_scenario


def build_packing_3d_preview(
    order_items,
    box_row,
    packing_rules=None,
    max_blocks=180,
    unpack_plan=None,
):
    if not box_row:
        return None

    box_length = float(box_row["length_cm"])
    box_width = float(box_row["width_cm"])
    box_height = float(box_row["height_cm"])
    if box_length <= 0 or box_width <= 0 or box_height <= 0:
        return None

    base_unpack_plan = unpack_plan if unpack_plan is not None else {}
    scenario = _simulate_layout_for_box(
        order_items,
        box_row,
        packing_rules=packing_rules,
        max_blocks=max_blocks,
        unpack_plan=base_unpack_plan,
    )
    if not scenario or not scenario["item_groups"]:
        return None

    if unpack_plan is None:
        unpack_scenario = _try_single_volume_with_unpack(
            order_items,
            box_row,
            packing_rules=packing_rules,
            max_blocks=max_blocks,
        )
        scenario = _select_preferred_preview_scenario(scenario, unpack_scenario)

    simulation = scenario["simulation"]
    blocks_to_place = scenario["blocks_to_place"]
    notes = list(scenario.get("notes") or [])

    hidden_blocks = max(0, simulation["placed_blocks"] - len(simulation["placements"]))
    box_volume = box_length * box_width * box_height
    fill_percent = (simulation["used_volume_cm3"] / box_volume) * 100 if box_volume > 0 else 0.0

    return {
        "box": {
            "name": str(_row_value(box_row, "name", "Caixa")),
            "length_cm": box_length,
            "width_cm": box_width,
            "height_cm": box_height,
        },
        "legend": _preview_legend_from_groups(scenario["item_groups"]),
        "placements": simulation["placements"],
        "total_blocks": len(blocks_to_place),
        "placed_blocks": simulation["placed_blocks"],
        "hidden_blocks": hidden_blocks,
        "overflow_blocks": simulation["overflow_blocks"],
        "used_volume_cm3": simulation["used_volume_cm3"],
        "fill_percent": round(fill_percent, 2),
        "notes": notes,
        "unpack_applied": bool(scenario.get("unpack_applied")),
    }


def _preview_block_signature(label, color, length, width, height):
    dims_key = tuple(
        sorted(
            (
                round(float(length), 4),
                round(float(width), 4),
                round(float(height), 4),
            )
        )
    )
    return (str(label or ""), str(color or ""), dims_key)


def _consume_preview_blocks(blocks_to_place, placed_placements):
    if not placed_placements:
        return list(blocks_to_place), 0

    target_counts = defaultdict(int)
    for placement in placed_placements:
        key = _preview_block_signature(
            placement.get("label"),
            placement.get("color"),
            placement.get("length_cm", 0),
            placement.get("width_cm", 0),
            placement.get("height_cm", 0),
        )
        target_counts[key] += 1

    remaining_blocks = []
    consumed = 0
    for block in blocks_to_place:
        dims = block.get("dims") or [0, 0, 0]
        key = _preview_block_signature(
            block.get("label"),
            block.get("color"),
            dims[0],
            dims[1],
            dims[2],
        )
        if target_counts[key] > 0:
            target_counts[key] -= 1
            consumed += 1
            continue

        remaining_blocks.append(block)

    if consumed < len(placed_placements):
        missing = min(len(remaining_blocks), len(placed_placements) - consumed)
        remaining_blocks = remaining_blocks[missing:]
        consumed += missing

    return remaining_blocks, consumed


def _preview_original_qty_map(item_groups):
    qty_map = {}
    for group in item_groups:
        qty_map[str(group.get("label") or "")] = int(group.get("raw_quantity") or 0)
    return qty_map


def _preview_legend_from_placements(placements, original_qty_map):
    legend_map = {}
    legend_order = []
    for placement in placements:
        label = str(placement.get("label") or "Item")
        if label not in legend_map:
            legend_map[label] = {
                "label": label,
                "color": str(placement.get("color") or "#64748b"),
                "blocks": 0,
                "original_quantity": int(original_qty_map.get(label, 0)),
            }
            legend_order.append(label)

        legend_map[label]["blocks"] += 1

    return [legend_map[label] for label in legend_order]


def _preview_legend_from_volume_blocks(placements, overflow_blocks, original_qty_map):
    legend_map = {}
    legend_order = []

    def ensure_entry(label, color):
        if label not in legend_map:
            legend_map[label] = {
                "label": label,
                "color": color,
                "blocks": 0,
                "overflow_blocks": 0,
                "original_quantity": int(original_qty_map.get(label, 0)),
            }
            legend_order.append(label)
        return legend_map[label]

    for placement in placements:
        label = str(placement.get("label") or "Item")
        color = str(placement.get("color") or "#64748b")
        ensure_entry(label, color)["blocks"] += 1

    for block in overflow_blocks or []:
        label = str(block.get("label") or "Item")
        color = str(block.get("color") or "#64748b")
        ensure_entry(label, color)["overflow_blocks"] += 1

    return [legend_map[label] for label in legend_order]


def _build_volume_preview_entry(
    box_row,
    placements,
    overflow_items,
    placed_blocks,
    total_blocks,
    overflow_blocks,
    used_volume_cm3,
    notes,
    unpack_applied,
    original_qty_map,
    volume_index,
    volume_count,
    max_blocks,
):
    box_length = float(box_row["length_cm"])
    box_width = float(box_row["width_cm"])
    box_height = float(box_row["height_cm"])
    box_volume = box_length * box_width * box_height
    fill_percent = (used_volume_cm3 / box_volume) * 100 if box_volume > 0 else 0.0

    visible_placements = list(placements[: max(1, int(max_blocks))])
    hidden_blocks = max(0, int(placed_blocks) - len(visible_placements))

    return {
        "box": {
            "name": str(_row_value(box_row, "name", "Caixa")),
            "length_cm": box_length,
            "width_cm": box_width,
            "height_cm": box_height,
        },
        "legend": _preview_legend_from_volume_blocks(
            visible_placements,
            overflow_items,
            original_qty_map,
        ),
        "placements": visible_placements,
        "total_blocks": int(total_blocks),
        "placed_blocks": int(placed_blocks),
        "hidden_blocks": hidden_blocks,
        "overflow_blocks": int(max(0, overflow_blocks)),
        "used_volume_cm3": float(used_volume_cm3),
        "fill_percent": round(fill_percent, 2),
        "notes": list(notes or []),
        "unpack_applied": bool(unpack_applied),
        "volume_index": int(volume_index),
        "volume_count": int(volume_count),
    }


def _preview_box_dims(box_row):
    if not box_row:
        return None

    box_length = float(box_row["length_cm"])
    box_width = float(box_row["width_cm"])
    box_height = float(box_row["height_cm"])
    if box_length <= 0 or box_width <= 0 or box_height <= 0:
        return None

    return (box_length, box_width, box_height)


def _resolve_preview_base_scenario(
    order_items,
    box_row,
    packing_rules,
    max_blocks,
    unpack_plan,
):
    base_unpack_plan = unpack_plan if unpack_plan is not None else {}
    scenario = _simulate_layout_for_box(
        order_items,
        box_row,
        packing_rules=packing_rules,
        max_blocks=max_blocks,
        unpack_plan=base_unpack_plan,
    )
    if not scenario or not scenario["item_groups"]:
        return None

    if unpack_plan is None:
        unpack_scenario = _try_single_volume_with_unpack(
            order_items,
            box_row,
            packing_rules=packing_rules,
            max_blocks=max_blocks,
        )
        scenario = _select_preferred_preview_scenario(scenario, unpack_scenario)

    return scenario


def _simulate_preview_volume_step(remaining_blocks, box_dims):
    simulation = _best_preview_simulation(
        remaining_blocks,
        box_dims,
        max(1, len(remaining_blocks)),
    )
    placed_blocks = int(simulation.get("placed_blocks") or 0)
    placements = list(simulation.get("placements") or [])
    if placed_blocks <= 0 or not placements:
        return None

    placed_blocks = min(placed_blocks, len(placements))
    placed_full = _limit_preview_placements_by_box_capacity(
        placements[:placed_blocks],
        remaining_blocks,
        box_dims,
    )
    placed_blocks = len(placed_full)
    if placed_blocks <= 0:
        return None

    next_remaining, consumed = _consume_preview_blocks(remaining_blocks, placed_full)
    if consumed <= 0:
        return None

    simulation = dict(simulation)
    simulation["placed_blocks"] = placed_blocks
    simulation["placements"] = list(placed_full)
    simulation["used_volume_cm3"] = _preview_used_volume_cm3(placed_full)

    return {
        "simulation": simulation,
        "placed_blocks": placed_blocks,
        "placements": placed_full,
        "remaining_blocks": next_remaining,
    }


def _preview_block_dims_from_signature_source(block_or_placement):
    dims = block_or_placement.get("dims")
    if dims:
        return [float(dims[0]), float(dims[1]), float(dims[2])]

    return [
        float(block_or_placement.get("length_cm", 0)),
        float(block_or_placement.get("width_cm", 0)),
        float(block_or_placement.get("height_cm", 0)),
    ]


def _preview_signature_capacity_map(blocks_to_place, box_dims):
    capacity_map = {}
    for block in blocks_to_place or []:
        dims = _preview_block_dims_from_signature_source(block)
        key = _preview_block_signature(
            block.get("label"),
            block.get("color"),
            dims[0],
            dims[1],
            dims[2],
        )
        if key in capacity_map:
            continue

        capacity_map[key] = max(0, _max_units_per_box(dims, box_dims))

    return capacity_map


def _preview_used_volume_cm3(placements):
    total = 0.0
    for placement in placements or []:
        total += (
            float(placement.get("length_cm", 0))
            * float(placement.get("width_cm", 0))
            * float(placement.get("height_cm", 0))
        )
    return total


def _limit_preview_placements_by_box_capacity(placements, blocks_to_place, box_dims):
    capacity_map = _preview_signature_capacity_map(blocks_to_place, box_dims)
    if not capacity_map:
        return list(placements)

    placed_counts = defaultdict(int)
    limited = []

    for placement in placements or []:
        dims = _preview_block_dims_from_signature_source(placement)
        key = _preview_block_signature(
            placement.get("label"),
            placement.get("color"),
            dims[0],
            dims[1],
            dims[2],
        )
        capacity = capacity_map.get(key, 0)
        if capacity <= 0:
            continue
        if placed_counts[key] >= capacity:
            continue

        placed_counts[key] += 1
        limited.append(placement)

    return limited


def _preview_signature_counts_from_blocks(blocks):
    counts = defaultdict(int)
    for block in blocks or []:
        dims = block.get("dims") or [0, 0, 0]
        key = _preview_block_signature(
            block.get("label"),
            block.get("color"),
            dims[0],
            dims[1],
            dims[2],
        )
        counts[key] += 1
    return counts


def _preview_signature_counts_from_placements(placements):
    counts = defaultdict(int)
    for placement in placements or []:
        key = _preview_block_signature(
            placement.get("label"),
            placement.get("color"),
            placement.get("length_cm", 0),
            placement.get("width_cm", 0),
            placement.get("height_cm", 0),
        )
        counts[key] += 1
    return counts


def _remaining_repeated_volume_count(placed_placements, remaining_blocks):
    placed_counts = _preview_signature_counts_from_placements(placed_placements)
    remaining_counts = _preview_signature_counts_from_blocks(remaining_blocks)

    if not placed_counts or not remaining_counts:
        return 0

    if set(remaining_counts.keys()) != set(placed_counts.keys()):
        return 0

    repeated_volume_count = None
    for key, placed_count in placed_counts.items():
        if placed_count <= 0:
            return 0

        remaining_count = remaining_counts.get(key, 0)
        if remaining_count % placed_count != 0:
            return 0

        key_repeat_count = remaining_count // placed_count
        if repeated_volume_count is None:
            repeated_volume_count = key_repeat_count
            continue

        if repeated_volume_count != key_repeat_count:
            return 0

    return max(0, int(repeated_volume_count or 0))


def _clone_preview_for_volume(preview, volume_index, volume_count, notes=None):
    cloned = {
        "box": dict(preview.get("box") or {}),
        "legend": [dict(item) for item in (preview.get("legend") or [])],
        "placements": [dict(item) for item in (preview.get("placements") or [])],
        "total_blocks": int(preview.get("total_blocks") or 0),
        "placed_blocks": int(preview.get("placed_blocks") or 0),
        "hidden_blocks": int(preview.get("hidden_blocks") or 0),
        "overflow_blocks": int(preview.get("overflow_blocks") or 0),
        "used_volume_cm3": float(preview.get("used_volume_cm3") or 0.0),
        "fill_percent": float(preview.get("fill_percent") or 0.0),
        "notes": list(preview.get("notes") or []) if notes is None else list(notes),
        "unpack_applied": bool(preview.get("unpack_applied")),
        "volume_index": int(volume_index),
        "volume_count": int(volume_count),
    }
    return cloned


def _expand_identical_previews(previews, preview_entry, start_index, repeat_count, volume_count):
    if repeat_count <= 0:
        return

    for offset in range(repeat_count):
        previews.append(
            _clone_preview_for_volume(
                preview_entry,
                volume_index=start_index + offset,
                volume_count=volume_count,
                notes=[],
            )
        )


def _try_expand_identical_remaining_previews(
    previews,
    placed_full,
    remaining_blocks,
    current_preview,
    current_index,
    volume_count,
):
    remaining_repeat_count = _remaining_repeated_volume_count(
        placed_full,
        remaining_blocks,
    )
    expected_remaining = volume_count - (current_index + 1)
    if remaining_repeat_count <= 0 or remaining_repeat_count != expected_remaining:
        return False

    _expand_identical_previews(
        previews,
        current_preview,
        start_index=current_index + 2,
        repeat_count=remaining_repeat_count,
        volume_count=volume_count,
    )
    return True


def _append_preview_volume_entry(
    previews,
    box_row,
    placed_full,
    placed_blocks,
    remaining_blocks,
    simulation,
    notes,
    unpack_applied,
    original_qty_map,
    volume_index,
    volume_count,
    max_blocks,
):
    is_last_volume = (volume_index - 1) == (volume_count - 1)
    overflow_items = list(remaining_blocks) if is_last_volume else []
    overflow_blocks = len(overflow_items)
    total_blocks = placed_blocks + overflow_blocks
    preview_notes = notes if volume_index == 1 else []

    previews.append(
        _build_volume_preview_entry(
            box_row,
            placed_full,
            overflow_items,
            placed_blocks,
            total_blocks,
            overflow_blocks,
            simulation.get("used_volume_cm3", 0.0),
            preview_notes,
            unpack_applied,
            original_qty_map,
            volume_index=volume_index,
            volume_count=volume_count,
            max_blocks=max_blocks,
        )
    )


def _append_preview_leftover(previews, remaining_blocks, volume_count):
    if not remaining_blocks or not previews:
        return

    last = previews[-1]
    if int(last.get("volume_index") or 0) >= int(volume_count or 0):
        return

    leftover = len(remaining_blocks)
    last["overflow_blocks"] = int(last.get("overflow_blocks", 0)) + leftover
    last["total_blocks"] = int(last.get("total_blocks", 0)) + leftover


def _single_preview_with_volume_meta(
    order_items,
    box_row,
    packing_rules,
    max_blocks,
    unpack_plan,
    volume_count,
):
    single = build_packing_3d_preview(
        order_items,
        box_row,
        packing_rules=packing_rules,
        max_blocks=max_blocks,
        unpack_plan=unpack_plan,
    )
    if not single:
        return []

    single["volume_index"] = 1
    single["volume_count"] = volume_count
    return [single]


def build_packing_3d_previews(
    order_items,
    box_row,
    packages_required=1,
    packing_rules=None,
    max_blocks=180,
    unpack_plan=None,
):
    box_dims = _preview_box_dims(box_row)
    if not box_dims:
        return []

    scenario = _resolve_preview_base_scenario(
        order_items,
        box_row,
        packing_rules,
        max_blocks,
        unpack_plan,
    )
    if scenario is None:
        return []

    volume_count = max(1, int(packages_required or 1))
    notes = list(scenario.get("notes") or [])
    unpack_applied = bool(scenario.get("unpack_applied"))
    original_qty_map = _preview_original_qty_map(scenario["item_groups"])
    remaining_blocks = list(scenario["blocks_to_place"])

    previews = []
    for index in range(volume_count):
        if not remaining_blocks:
            break

        step = _simulate_preview_volume_step(remaining_blocks, box_dims)
        if step is None:
            break

        remaining_blocks = step["remaining_blocks"]
        placed_blocks = step["placed_blocks"]
        placed_full = step["placements"]
        simulation = step["simulation"]

        _append_preview_volume_entry(
            previews,
            box_row,
            placed_full,
            placed_blocks,
            remaining_blocks,
            simulation,
            notes,
            unpack_applied,
            original_qty_map,
            volume_index=(index + 1),
            volume_count=volume_count,
            max_blocks=max_blocks,
        )

        current_preview = previews[-1]
        if _try_expand_identical_remaining_previews(
            previews,
            placed_full,
            remaining_blocks,
            current_preview,
            index,
            volume_count,
        ):
            remaining_blocks = []
            break

    _append_preview_leftover(previews, remaining_blocks, volume_count)

    if not previews:
        return _single_preview_with_volume_meta(
            order_items,
            box_row,
            packing_rules,
            max_blocks,
            unpack_plan,
            volume_count,
        )

    return previews
