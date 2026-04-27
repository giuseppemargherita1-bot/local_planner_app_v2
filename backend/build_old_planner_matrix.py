import csv
import sqlite3
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict


BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "planner_v2.sqlite3"

OUT_CSV = BASE_DIR / "data" / "old_planner_matrix_preview.csv"
OUT_TXT = BASE_DIR / "data" / "old_planner_matrix_preview.txt"

YEAR = 2026
YEAR_SHORT = 26
WEEK_FROM = 1
WEEK_TO = 52
CURRENT_WEEK = 17
OVERALL_OFFICINA_NAME = "OVERALL OFFICINA"
IMPORT_ZERO_DATE = "07/04/2026"


def connect(path):
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


def clean(value):
    return str(value or "").strip()


def up(value):
    return clean(value).upper()


def norm_role(value):
    return up(value)


def norm_code(value):
    return up(value)


def period_key_from_week(week):
    return YEAR_SHORT * 100 + int(week)


def monday_of_iso_week(year, week):
    jan4 = datetime(year, 1, 4)
    monday_week_1 = jan4 - timedelta(days=jan4.isoweekday() - 1)
    return monday_week_1 + timedelta(weeks=int(week) - 1)


def parse_date(value):
    raw = clean(value)
    if not raw:
        return None

    if raw.upper() in (
        "N.D.",
        "ND",
        "N/D",
        "INDET.",
        "INDET",
        "INDETERMINATO",
        "31/12/2099",
        "2099-12-31",
    ):
        return None

    if raw.upper().startswith("W"):
        return None

    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt)
        except Exception:
            pass

    return None


def table_exists(conn, name):
    row = conn.execute(
        """
        SELECT name
        FROM sqlite_master
        WHERE type='table'
          AND name=?
        """,
        (name,),
    ).fetchone()
    return row is not None


def require_raw(conn):
    required = [
        "raw_old_projects",
        "raw_old_resources",
        "raw_old_demands",
        "raw_old_allocations",
        "raw_old_unavailability",
    ]
    missing = [x for x in required if not table_exists(conn, x)]
    if missing:
        raise RuntimeError(
            "Mancano tabelle raw: "
            + ", ".join(missing)
            + ". Prima esegui backend/copy_old_raw.py"
        )


def row_get(row, key, default=""):
    try:
        if key in row.keys():
            return row[key]
    except Exception:
        pass
    return default


def is_ext_resource(resource):
    name = up(row_get(resource, "name"))
    employer = up(row_get(resource, "employer"))
    role1 = up(row_get(resource, "role1"))
    role2 = up(row_get(resource, "role2"))

    return (
        "-EXT" in name
        or name.endswith(" EXT")
        or employer == "EXT"
        or "-EXT" in role1
        or role1.endswith(" EXT")
        or "-EXT" in role2
        or role2.endswith(" EXT")
    )


def is_overall_officina(project):
    return norm_code(row_get(project, "code")) == OVERALL_OFFICINA_NAME


def is_ws_overall_old(project):
    return norm_code(row_get(project, "code")) == "WS OVERALL"


def is_officina_child(project):
    code = norm_code(row_get(project, "code"))
    activity = norm_code(row_get(project, "activity"))
    ptype = norm_code(row_get(project, "type"))
    workshop_rollup = int(row_get(project, "workshop_rollup", 0) or 0)

    if is_overall_officina(project):
        return False

    if is_ws_overall_old(project):
        return False

    return (
        "OFFICINA" in code
        or "OFFICINA" in activity
        or ptype == "WS"
        or workshop_rollup == 1
    )


def week_status_for_resource(resource, week):
    if is_ext_resource(resource):
        return "EXT"

    hire_date = parse_date(row_get(resource, "hire_date"))
    end_date = parse_date(row_get(resource, "end_date"))

    week_start = monday_of_iso_week(YEAR, week)
    week_end = week_start + timedelta(days=6)

    if hire_date and week_end < hire_date:
        return "NON_ASSUNTO"

    if end_date and week_start > end_date:
        return "FUORI_CONTRATTO"

    return "VALIDO"


def resource_primary_role(resource):
    return norm_role(row_get(resource, "role1") or row_get(resource, "role2"))


def resource_display_name(resource):
    return clean(row_get(resource, "name"))


def short_resource_name(name):
    txt = clean(name).replace("  ", " ")
    parts = [p for p in txt.split(" ") if p]
    if not parts:
        return ""

    if len(parts) == 1:
        return parts[0].upper()

    surname = " ".join(parts[:-1]).upper()
    first = parts[-1].upper()[:3]
    return f"{surname} {first}"


def load_data(conn):
    projects = conn.execute("SELECT * FROM raw_old_projects ORDER BY id").fetchall()
    resources = conn.execute("SELECT * FROM raw_old_resources ORDER BY id").fetchall()
    demands = conn.execute("SELECT * FROM raw_old_demands ORDER BY project_id, role, week").fetchall()
    allocations = conn.execute("SELECT * FROM raw_old_allocations ORDER BY id").fetchall()
    unavailability = conn.execute("SELECT * FROM raw_old_unavailability ORDER BY id").fetchall()

    return projects, resources, demands, allocations, unavailability


def build_project_maps(projects):
    projects_by_id = {int(p["id"]): p for p in projects}

    overall = next((p for p in projects if is_overall_officina(p)), None)
    overall_old_id = int(overall["id"]) if overall else None

    officina_child_ids = {
        int(p["id"])
        for p in projects
        if is_officina_child(p)
    }

    def visual_project_id(old_project_id):
        if old_project_id in officina_child_ids and overall_old_id:
            return overall_old_id
        return old_project_id

    return projects_by_id, overall_old_id, officina_child_ids, visual_project_id


def load_unavailability_by_resource_week(unavailability):
    by_resource_week = defaultdict(list)

    for row in unavailability:
      resource_id = int(row_get(row, "resource_id", 0) or 0)
      week_from = int(row_get(row, "week_from", 0) or 0)
      week_to = int(row_get(row, "week_to", 0) or 0)
      reason = clean(row_get(row, "reason"))

      if week_to < week_from:
          week_from, week_to = week_to, week_from

      for week in range(week_from, week_to + 1):
          by_resource_week[(resource_id, week)].append(reason or "INDISP")

    return by_resource_week


def build_matrix(projects, resources, demands, allocations, unavailability):
    projects_by_id, overall_old_id, officina_child_ids, visual_project_id = build_project_maps(projects)
    resources_by_id = {int(r["id"]): r for r in resources}
    unavailability_map = load_unavailability_by_resource_week(unavailability)

    demand_by_cell = defaultdict(float)
    source_demand_by_cell = defaultdict(list)

    for d in demands:
        old_project_id = int(row_get(d, "project_id", 0) or 0)
        visual_pid = visual_project_id(old_project_id)
        role = norm_role(row_get(d, "role"))
        week = int(row_get(d, "week", 0) or 0)
        qty = float(row_get(d, "qty", 0) or 0)

        if week < WEEK_FROM or week > WEEK_TO:
            continue

        key = (visual_pid, role, week)
        demand_by_cell[key] += qty

        if qty != 0:
            source_project = projects_by_id.get(old_project_id)
            source_demand_by_cell[key].append(
                {
                    "old_project_id": old_project_id,
                    "old_project_code": clean(row_get(source_project, "code")) if source_project else "",
                    "qty": qty,
                }
            )

    active_alloc_by_cell = defaultdict(list)
    historical_alloc_by_cell = defaultdict(list)
    ext_alloc_by_cell = defaultdict(list)
    no_mans_alloc_by_cell = defaultdict(list)
    indisp_alloc_by_cell = defaultdict(list)

    all_allocation_items = []

    for a in allocations:
        old_project_id = int(row_get(a, "project_id", 0) or 0)
        resource_id = int(row_get(a, "resource_id", 0) or 0)
        visual_pid = visual_project_id(old_project_id)

        project = projects_by_id.get(old_project_id)
        visual_project = projects_by_id.get(visual_pid)
        resource = resources_by_id.get(resource_id)

        if not resource:
            continue

        role = norm_role(row_get(a, "role"))
        week_from = int(row_get(a, "week_from", 0) or 0)
        week_to = int(row_get(a, "week_to", 0) or 0)
        weight = float(row_get(a, "weight", 1) or 1)

        if week_to < week_from:
            week_from, week_to = week_to, week_from

        for week in range(max(WEEK_FROM, week_from), min(WEEK_TO, week_to) + 1):
            status = week_status_for_resource(resource, week)
            is_ext = is_ext_resource(resource)
            resource_role = resource_primary_role(resource)
            is_no_mans = bool(resource_role and role and resource_role != role)
            unavailability_reasons = unavailability_map.get((resource_id, week), [])
            is_indisp = len(unavailability_reasons) > 0

            item = {
                "allocation_id": int(row_get(a, "id", 0) or 0),
                "old_project_id": old_project_id,
                "old_project_code": clean(row_get(project, "code")) if project else "",
                "visual_project_id": visual_pid,
                "visual_project_code": clean(row_get(visual_project, "code")) if visual_project else "",
                "resource_id": resource_id,
                "resource_name": resource_display_name(resource),
                "resource_short": short_resource_name(resource_display_name(resource)),
                "resource_role": resource_role,
                "role": role,
                "week": week,
                "period_key": period_key_from_week(week),
                "weight": weight,
                "load_percent": weight * 100,
                "status": status,
                "is_ext": is_ext,
                "is_no_mans": is_no_mans,
                "is_indisp": is_indisp,
                "unavailability": "; ".join(unavailability_reasons),
                "counts_numeric": False,
                "callout_value": "0:1",
            }

            key = (visual_pid, role, week)

            if is_ext:
                item["counts_numeric"] = True
                item["callout_value"] = "1:1" if weight >= 0.999 else "1/2"
                active_alloc_by_cell[key].append(item)
                ext_alloc_by_cell[key].append(item)
            elif is_indisp:
                item["status"] = "INDISP"
                item["counts_numeric"] = False
                item["callout_value"] = "0:1"
                historical_alloc_by_cell[key].append(item)
                indisp_alloc_by_cell[key].append(item)
            elif status == "VALIDO":
                item["counts_numeric"] = True
                item["callout_value"] = "1:1" if weight >= 0.999 else "1/2"
                active_alloc_by_cell[key].append(item)
                if is_no_mans:
                    no_mans_alloc_by_cell[key].append(item)
            else:
                item["counts_numeric"] = False
                item["callout_value"] = "0:1"
                historical_alloc_by_cell[key].append(item)

            all_allocation_items.append(item)

    all_keys = set(demand_by_cell.keys()) | set(active_alloc_by_cell.keys()) | set(historical_alloc_by_cell.keys())

    rows = []

    for key in sorted(all_keys, key=lambda x: (
        clean(row_get(projects_by_id.get(x[0]), "code")),
        x[1],
        x[2],
    )):
        visual_project_id_value, role, week = key
        project = projects_by_id.get(visual_project_id_value)
        if not project:
            continue

        required = float(demand_by_cell.get(key, 0) or 0)
        active_items = active_alloc_by_cell.get(key, [])
        historical_items = historical_alloc_by_cell.get(key, [])
        ext_items = ext_alloc_by_cell.get(key, [])
        no_mans_items = no_mans_alloc_by_cell.get(key, [])
        indisp_items = indisp_alloc_by_cell.get(key, [])

        allocated = sum(float(item["weight"] or 0) for item in active_items)
        diff = required - allocated

        if required == 0 and allocated == 0:
            cell_state = "EMPTY"
        elif allocated < required:
            cell_state = "MISSING_RED"
        elif allocated == required:
            cell_state = "COVERED_GREEN"
        else:
            cell_state = "SURPLUS_GREEN_RED_BORDER"

        badges = []
        if ext_items:
            badges.append("EXT")
        if no_mans_items:
            badges.append("FM")
        if indisp_items:
            badges.append("IND")
        if allocated > required:
            badges.append("NO_FAB")
        if required == 0 and allocated > 0:
            if "NO_FAB" not in badges:
                badges.append("NO_FAB")
        if historical_items:
            badges.append("STORICO")

        active_callout = []
        for item in active_items:
            mans = "ok mans" if not item["is_no_mans"] else f"no mans {item['resource_role']}"
            if item["is_ext"]:
                label = f"{item['resource_name']} | EXT | {item['callout_value']}"
            else:
                label = f"{item['resource_short']} | {mans} | {item['callout_value']}"
            active_callout.append(label)

        historical_callout = []
        for item in historical_items:
            mans = "ok mans" if not item["is_no_mans"] else f"no mans {item['resource_role']}"
            label = f"{item['resource_short']} | {mans} | {item['callout_value']} | {item['status']}"
            historical_callout.append(label)

        source_demands = source_demand_by_cell.get(key, [])
        source_demand_text = " ; ".join(
            f"{x['old_project_code']}:{x['qty']}"
            for x in source_demands
        )

        rows.append(
            {
                "visual_project_id": visual_project_id_value,
                "visual_project_code": clean(row_get(project, "code")),
                "role": role,
                "week": week,
                "period_key": period_key_from_week(week),
                "required_R": required,
                "allocated_A_numeric": allocated,
                "diff_D": diff,
                "cell_state": cell_state,
                "badges": ",".join(badges),
                "active_resources": " | ".join(active_callout),
                "historical_resources": " | ".join(historical_callout),
                "source_demands": source_demand_text,
                "is_overall_officina": 1 if is_overall_officina(project) else 0,
                "source": "RAW_OLD_FULL_MATRIX",
            }
        )

    return rows, {
        "projects_by_id": projects_by_id,
        "overall_old_id": overall_old_id,
        "officina_child_ids": officina_child_ids,
        "total_alloc_items": len(all_allocation_items),
    }


def write_outputs(rows, meta):
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)

    fieldnames = [
        "visual_project_id",
        "visual_project_code",
        "role",
        "week",
        "period_key",
        "required_R",
        "allocated_A_numeric",
        "diff_D",
        "cell_state",
        "badges",
        "active_resources",
        "historical_resources",
        "source_demands",
        "is_overall_officina",
        "source",
    ]

    with OUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    by_project = defaultdict(lambda: {"rows": 0, "R": 0, "A": 0})
    for r in rows:
        p = r["visual_project_code"]
        by_project[p]["rows"] += 1
        by_project[p]["R"] += float(r["required_R"] or 0)
        by_project[p]["A"] += float(r["allocated_A_numeric"] or 0)

    lines = []
    lines.append("=== OLD PLANNER FULL MATRIX PREVIEW ===")
    lines.append(f"DB: {DB_PATH}")
    lines.append(f"Periodo: W{WEEK_FROM}-W{WEEK_TO} {YEAR}")
    lines.append(f"Import zero: {IMPORT_ZERO_DATE}")
    lines.append("")
    lines.append("=== META ===")
    lines.append(f"Overall officina old id: {meta['overall_old_id']}")
    lines.append(f"Commesse officina figlie rollup: {len(meta['officina_child_ids'])}")
    lines.append(f"Allocazioni settimanali espanse lette: {meta['total_alloc_items']}")
    lines.append(f"Celle non vuote esportate: {len(rows)}")
    lines.append("")
    lines.append("=== TOTALI PER COMMESSA VISUALE ===")

    for project_code in sorted(by_project.keys()):
        item = by_project[project_code]
        lines.append(
            f"{project_code}: celle={item['rows']} | R={item['R']} | A={item['A']} | D={item['R'] - item['A']}"
        )

    lines.append("")
    lines.append("=== PRIME 200 CELLE CON STORICO / EXT / NO_FAB / FM / IND ===")
    interesting = [
        r for r in rows
        if r["badges"]
    ][:200]

    for r in interesting:
        lines.append(
            f"{r['visual_project_code']} | {r['role']} | W{r['week']} | "
            f"R{r['required_R']} A{r['allocated_A_numeric']} D{r['diff_D']} | "
            f"{r['cell_state']} | {r['badges']} | "
            f"ATT=[{r['active_resources']}] | HIST=[{r['historical_resources']}]"
        )

    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")


def main():
    conn = connect(DB_PATH)

    try:
        require_raw(conn)
        projects, resources, demands, allocations, unavailability = load_data(conn)
        rows, meta = build_matrix(projects, resources, demands, allocations, unavailability)
        write_outputs(rows, meta)

        print("PREVIEW COMPLETA GENERATA")
        print(f"CSV: {OUT_CSV}")
        print(f"TXT: {OUT_TXT}")
        print(f"Celle esportate: {len(rows)}")
        print(f"Overall officina old id: {meta['overall_old_id']}")
        print(f"Commesse officina rollup: {len(meta['officina_child_ids'])}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()