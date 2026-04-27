from app.db import get_connection

DEFAULT_YEAR_SHORT = 26
FIRST_WEEK = 1
LAST_WEEK = 52
CURRENT_PERIOD_KEY = 2617


def clean_text(value):
    return str(value or "").strip()


def normalize_role(value):
    return clean_text(value).upper()


def period_key_from_week(week, year_short=DEFAULT_YEAR_SHORT):
    week_int = int(week or 0)
    if week_int >= 1000:
        return week_int
    return int(year_short) * 100 + week_int


def week_from_period_key(period_key):
    value = int(period_key or 0)
    if value >= 1000:
        return value % 100
    return value


def is_external_text(name="", role=""):
    name_text = normalize_role(name)
    role_text = normalize_role(role)
    return (
        "-EXT" in name_text
        or name_text.endswith(" EXT")
        or "-EXT" in role_text
        or role_text.endswith(" EXT")
    )


def is_overall_project(project):
    text = normalize_role(project.get("name") or project.get("project_name") or "")
    note = normalize_role(project.get("note") or "")
    return "OVERALL OFFICINA" in text or "OVERALL OFFICINA" in note


def is_workshop_child(project):
    text = normalize_role(project.get("name") or project.get("project_name") or "")
    note = normalize_role(project.get("note") or "")
    return "OFFICINA" in text and "OVERALL OFFICINA" not in text or "WORKSHOP_ROLLUP" in note


def resource_unavailable_reason(resource):
    if not resource:
        return "RESOURCE_NOT_FOUND"
    if is_external_text(resource.get("name", ""), resource.get("role", "")):
        return ""
    if int(resource.get("is_active") or 0) != 1:
        return "FUORI_CONTRATTO"
    note = normalize_role(resource.get("availability_note") or "")
    if "INDISP" in note or "NON DISPONIBILE" in note:
        return "INDISP"
    if "FUORI CONTRATTO" in note or "FUORI_CONTRATTO" in note or "CESS" in note:
        return "FUORI_CONTRATTO"
    return ""


def allocation_weight(allocation):
    load = float(allocation.get("load_percent") or 100)
    if load <= 0:
        return 0
    return round(load / 100, 4)


def cell_state(required, allocated):
    required = float(required or 0)
    allocated = float(allocated or 0)
    if required == 0 and allocated == 0:
        return "empty"
    if allocated < required:
        return "missing"
    if allocated == required:
        return "covered"
    return "surplus"


def cell_classes(required, allocated, external_allocated=0, period_key=None):
    state = cell_state(required, allocated)
    classes = ["planner-cell-clickable"]
    if state == "missing":
        classes.append("cell-demand")
    elif state == "covered":
        classes.append("cell-ok")
    elif state == "surplus":
        classes.extend(["cell-ok", "cell-surplus"])
    else:
        classes.append("cell-empty")
    if float(external_allocated or 0) > 0:
        classes.append("cell-ext")
    if int(period_key or 0) == CURRENT_PERIOD_KEY:
        classes.append("current-col")
    return " ".join(classes)


def fetch_rows(conn, sql, params=()):
    return [dict(row) for row in conn.execute(sql, params).fetchall()]


def load_base_data(conn):
    projects = fetch_rows(conn, "SELECT id, name, client, start_date, end_date, status, note FROM projects ORDER BY name")
    resources = fetch_rows(conn, "SELECT id, name, role, availability_note, is_active FROM resources ORDER BY name")
    demands = fetch_rows(conn, """
        SELECT d.id, d.project_id, p.name AS project_name, d.week,
               COALESCE(NULLIF(d.period_key, 0), 2600 + d.week) AS period_key,
               d.role, d.quantity, d.note
        FROM demands d JOIN projects p ON p.id = d.project_id
    """)
    allocations = fetch_rows(conn, """
        SELECT a.id, a.resource_id, r.name AS resource_name, r.role AS resource_role,
               r.availability_note AS resource_availability_note, r.is_active AS resource_is_active,
               a.project_id, p.name AS project_name, a.week,
               COALESCE(NULLIF(a.period_key, 0), 2600 + a.week) AS period_key,
               a.role, a.hours, a.load_percent, a.note
        FROM allocations a
        JOIN resources r ON r.id = a.resource_id
        JOIN projects p ON p.id = a.project_id
    """)
    history = fetch_rows(conn, """
        SELECT h.id AS history_id, h.resource_id, h.resource_name, h.resource_role,
               h.project_id, h.project_name, h.week,
               COALESCE(NULLIF(h.period_key, 0), 2600 + h.week) AS period_key,
               h.role, h.hours, h.load_percent, h.reason, h.note, h.created_at
        FROM allocation_history h
    """)
    return projects, resources, demands, allocations, history


def build_planner_matrix(show_zero_demand_projects=False, project_filter=None, role_filter=None):
    conn = get_connection()
    try:
        projects, resources, demands, allocations, history = load_base_data(conn)
    finally:
        conn.close()

    project_by_id = {int(p["id"]): p for p in projects}
    resource_by_id = {int(r["id"]): r for r in resources}
    overall_projects = [p for p in projects if is_overall_project(p)]
    overall_project = overall_projects[0] if overall_projects else None
    overall_id = int(overall_project["id"]) if overall_project else None
    workshop_child_ids = {int(p["id"]) for p in projects if is_workshop_child(p)}

    demand_map = {}
    source_map = {}
    row_keys = set()
    for d in demands:
        project_id = int(d["project_id"])
        role = normalize_role(d["role"])
        period_key = int(d["period_key"] or period_key_from_week(d["week"]))
        qty = float(d["quantity"] or 0)
        target_project_id = overall_id if overall_id and project_id in workshop_child_ids else project_id
        key = (target_project_id, role, period_key)
        demand_map[key] = demand_map.get(key, 0) + qty
        source_map.setdefault(key, []).append({
            "project_id": project_id,
            "project_name": d.get("project_name"),
            "role": role,
            "period_key": period_key,
            "week": week_from_period_key(period_key),
            "required": qty,
            "demand_id": d.get("id"),
        })
        row_keys.add((target_project_id, role))

    allocation_map = {}
    active_resource_map = {}
    historical_resource_map = {}
    for a in allocations:
        project_id = int(a["project_id"])
        role = normalize_role(a["role"] or a.get("resource_role") or "")
        period_key = int(a["period_key"] or period_key_from_week(a["week"]))
        resource = resource_by_id.get(int(a["resource_id"]))
        external = is_external_text(a.get("resource_name", ""), a.get("resource_role", ""))
        reason = resource_unavailable_reason(resource)
        target_project_id = project_id
        key = (target_project_id, role, period_key)
        row_keys.add((target_project_id, role))
        item = dict(a)
        item["status"] = "EXT" if external else (reason or "ACTIVE")
        item["counts"] = bool(external or not reason)
        item["external"] = external
        item["display_weight"] = "0:1" if reason and not external else "1:1"
        if item["counts"]:
            weight = allocation_weight(a)
            allocation_map[key] = allocation_map.get(key, {"allocated": 0, "internal": 0, "external": 0})
            allocation_map[key]["allocated"] += weight
            if external:
                allocation_map[key]["external"] += weight
            else:
                allocation_map[key]["internal"] += weight
            active_resource_map.setdefault(key, []).append(item)
        else:
            historical_resource_map.setdefault(key, []).append(item)

    for h in history:
        project_id = int(h.get("project_id") or 0)
        role = normalize_role(h.get("role") or h.get("resource_role") or "")
        period_key = int(h["period_key"] or period_key_from_week(h["week"]))
        key = (project_id, role, period_key)
        item = dict(h)
        item["status"] = normalize_role(h.get("reason") or "STORICO")
        item["counts"] = False
        item["external"] = is_external_text(h.get("resource_name", ""), h.get("resource_role", ""))
        item["display_weight"] = "0:1"
        historical_resource_map.setdefault(key, []).append(item)
        row_keys.add((project_id, role))

    if not show_zero_demand_projects:
        filtered = set()
        for project_id, role in row_keys:
            project = project_by_id.get(project_id, {})
            if project_id in workshop_child_ids:
                continue
            if any(demand_map.get((project_id, role, p), 0) > 0 or allocation_map.get((project_id, role, p), {}).get("allocated", 0) > 0 for p in range(CURRENT_PERIOD_KEY, period_key_from_week(LAST_WEEK) + 1)):
                filtered.add((project_id, role))
            elif any(demand_map.get((project_id, role, p), 0) > 0 or allocation_map.get((project_id, role, p), {}).get("allocated", 0) > 0 for p in range(CURRENT_PERIOD_KEY - 4, CURRENT_PERIOD_KEY)):
                filtered.add((project_id, role))
            elif is_overall_project(project):
                filtered.add((project_id, role))
        row_keys = filtered

    if project_filter:
        txt = normalize_role(project_filter)
        row_keys = {rk for rk in row_keys if txt in normalize_role(project_by_id.get(rk[0], {}).get("name", ""))}
    if role_filter:
        rtxt = normalize_role(role_filter)
        row_keys = {rk for rk in row_keys if rtxt in rk[1]}

    weeks = [{"week": week, "period_key": period_key_from_week(week), "label": f"W{week:02d}"} for week in range(FIRST_WEEK, LAST_WEEK + 1)]

    rows = []
    hidden_count = 0
    for project_id, role in sorted(row_keys, key=lambda x: (
        min([p for (pid, r, p), q in demand_map.items() if pid == x[0] and r == x[1] and q > 0] or [9999]),
        normalize_role(project_by_id.get(x[0], {}).get("name", "")),
        x[1],
    )):
        project = project_by_id.get(project_id, {"name": f"Project {project_id}"})
        cells = {}
        first_useful = None
        last_useful = None
        for period in weeks:
            pk = period["period_key"]
            required = round(demand_map.get((project_id, role, pk), 0), 4)
            alloc = allocation_map.get((project_id, role, pk), {"allocated": 0, "internal": 0, "external": 0})
            allocated = round(alloc["allocated"], 4)
            diff = round(required - allocated, 4)
            if required > 0 or allocated > 0:
                first_useful = pk if first_useful is None else min(first_useful, pk)
                last_useful = pk if last_useful is None else max(last_useful, pk)
            badges = []
            if alloc["external"] > 0:
                badges.append("EXT")
            if required == 0 and allocated > 0:
                badges.append("NO FAB")
            if historical_resource_map.get((project_id, role, pk)):
                badges.append("STORICO")
            cells[str(pk)] = {
                "project_id": project_id,
                "project_name": project.get("name"),
                "role": role,
                "week": period["week"],
                "period_key": pk,
                "required": required,
                "allocated": allocated,
                "internal_allocated": round(alloc["internal"], 4),
                "external_allocated": round(alloc["external"], 4),
                "diff": diff,
                "state": cell_state(required, allocated),
                "class_name": cell_classes(required, allocated, alloc["external"], pk),
                "badges": badges,
                "active_resources": active_resource_map.get((project_id, role, pk), []),
                "historical_resources": historical_resource_map.get((project_id, role, pk), []),
                "workshop_sources": source_map.get((project_id, role, pk), []),
            }
        rows.append({
            "project_id": project_id,
            "project_name": project.get("name"),
            "role": role,
            "is_overall": is_overall_project(project),
            "first_useful_period": first_useful,
            "last_useful_period": last_useful,
            "cells": cells,
        })

    return {
        "weeks": weeks,
        "rows": rows,
        "hidden_notice": "Vista operativa: commesse officina figlie e righe senza fabbisogno utile sono nascoste." if not show_zero_demand_projects else "Vista diagnostica completa attiva.",
        "resources": resources,
        "overall_details_available": overall_id is not None,
        "overall_project_id": overall_id,
        "hidden_count": hidden_count,
    }


def get_workshop_breakdown(project_id=None, role=None, week=None, period_key=None):
    matrix = build_planner_matrix(show_zero_demand_projects=True)
    pk = int(period_key or period_key_from_week(week or 0))
    wanted_role = normalize_role(role)
    for row in matrix["rows"]:
        if project_id and int(row["project_id"]) != int(project_id):
            continue
        if wanted_role and row["role"] != wanted_role:
            continue
        if not row.get("is_overall"):
            continue
        cell = row["cells"].get(str(pk))
        if not cell:
            continue
        sources = cell.get("workshop_sources") or []
        return {
            "project_id": row["project_id"],
            "project_name": row["project_name"],
            "role": row["role"],
            "week": week_from_period_key(pk),
            "period_key": pk,
            "sources": sources,
            "total_required": sum(float(s.get("required") or 0) for s in sources),
        }
    return {"project_id": project_id, "role": wanted_role, "week": week_from_period_key(pk), "period_key": pk, "sources": [], "total_required": 0}
