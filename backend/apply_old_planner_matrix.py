import csv
import re
import sqlite3
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "planner_v2.sqlite3"
MATRIX_CSV = BASE_DIR / "data" / "old_planner_matrix_preview.csv"
REPORT_TXT = BASE_DIR / "data" / "apply_old_planner_matrix_report.txt"

IMPORT_ZERO_DATE = "07/04/2026"
YEAR_SHORT = 26
OVERALL_OFFICINA_NAME = "OVERALL OFFICINA"


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


def period_key_from_week(week):
    return YEAR_SHORT * 100 + int(week)


def parse_float(value):
    raw = clean(value).replace(",", ".")
    if raw == "":
        return 0.0
    return float(raw)


def table_exists(conn, table):
    row = conn.execute(
        """
        SELECT name
        FROM sqlite_master
        WHERE type='table'
          AND name=?
        """,
        (table,),
    ).fetchone()
    return row is not None


def column_exists(conn, table, column):
    if not table_exists(conn, table):
        return False
    rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    return any(row["name"] == column for row in rows)


def ensure_column(conn, table, column, definition):
    if table_exists(conn, table) and not column_exists(conn, table, column):
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def ensure_schema(conn):
    ensure_column(conn, "projects", "old_id", "INTEGER DEFAULT 0")
    ensure_column(conn, "projects", "is_overall", "INTEGER DEFAULT 0")
    ensure_column(conn, "projects", "parent_overall_id", "INTEGER DEFAULT 0")
    ensure_column(conn, "projects", "workshop_rollup", "INTEGER DEFAULT 0")

    ensure_column(conn, "resources", "old_id", "INTEGER DEFAULT 0")

    ensure_column(conn, "demands", "period_key", "INTEGER DEFAULT 0")
    ensure_column(conn, "allocations", "period_key", "INTEGER DEFAULT 0")
    ensure_column(conn, "demand_history", "period_key", "INTEGER DEFAULT 0")
    ensure_column(conn, "allocation_history", "period_key", "INTEGER DEFAULT 0")


def clear_operational_tables(conn):
    conn.execute("DELETE FROM allocation_history")
    conn.execute("DELETE FROM demand_history")
    conn.execute("DELETE FROM allocations")
    conn.execute("DELETE FROM demands")
    conn.execute("DELETE FROM resources")
    conn.execute("DELETE FROM projects")


def require_raw(conn):
    required = [
        "raw_old_projects",
        "raw_old_resources",
        "raw_old_demands",
        "raw_old_allocations",
        "raw_old_unavailability",
    ]
    missing = [name for name in required if not table_exists(conn, name)]
    if missing:
        raise RuntimeError(
            "Mancano tabelle raw: "
            + ", ".join(missing)
            + ". Esegui prima copy_old_raw.py"
        )


def load_matrix_rows():
    if not MATRIX_CSV.exists():
        raise FileNotFoundError(
            f"Non trovo {MATRIX_CSV}. Esegui prima build_old_planner_matrix.py"
        )

    with MATRIX_CSV.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        return [dict(row) for row in reader]


def import_resources_from_raw(conn):
    rows = conn.execute(
        """
        SELECT *
        FROM raw_old_resources
        ORDER BY id
        """
    ).fetchall()

    old_to_new = {}
    by_upper_name = {}

    for row in rows:
        old_id = int(row["id"])
        name = clean(row["name"])
        role1 = norm_role(row["role1"])
        role2 = norm_role(row["role2"])
        hire_date = clean(row["hire_date"])
        end_date = clean(row["end_date"])
        employer = clean(row["employer"])
        note = clean(row["note"])

        role = role1 or role2

        availability_parts = [
            f"RAW_OLD_ID={old_id}",
            f"ROLE2={role2}" if role2 else "",
            f"HIRE={hire_date}" if hire_date else "",
            f"END={end_date}" if end_date else "",
            f"EMPLOYER={employer}" if employer else "",
            note,
            f"IMPORT_ZERO_{IMPORT_ZERO_DATE}",
        ]

        cursor = conn.execute(
            """
            INSERT INTO resources
            (
                name,
                role,
                availability_note,
                is_active,
                old_id
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                name,
                role,
                " | ".join([part for part in availability_parts if part]),
                1,
                old_id,
            ),
        )

        new_id = cursor.lastrowid
        old_to_new[old_id] = new_id
        by_upper_name[up(name)] = new_id

    return old_to_new, by_upper_name


def load_raw_project_by_id(conn):
    rows = conn.execute("SELECT * FROM raw_old_projects ORDER BY id").fetchall()
    return {int(row["id"]): row for row in rows}


def is_overall_officina_name(name):
    return up(name) == OVERALL_OFFICINA_NAME


def import_projects_from_matrix(conn, matrix_rows):
    project_codes = sorted({
        clean(row["visual_project_code"])
        for row in matrix_rows
        if clean(row.get("visual_project_code"))
    })

    raw_by_id = load_raw_project_by_id(conn)
    raw_by_code = {up(row["code"]): row for row in raw_by_id.values()}

    code_to_new = {}

    for code in project_codes:
        raw = raw_by_code.get(up(code))
        old_id = int(raw["id"]) if raw else 0
        ptype = clean(raw["type"]) if raw else ""
        closed = int(raw["closed"] or 0) if raw else 0
        workshop_rollup = int(raw["workshop_rollup"] or 0) if raw else 0
        activity = clean(raw["activity"]) if raw else ""

        is_overall = 1 if is_overall_officina_name(code) else 0

        note_parts = [
            f"RAW_OLD_ID={old_id}" if old_id else "",
            f"TYPE={ptype}" if ptype else "",
            "OVERALL" if is_overall else "",
            "WORKSHOP_ROLLUP" if workshop_rollup else "",
            "CHIUSA_VECCHIO" if closed else "",
            activity if activity and activity != code else "",
            "MATRIX_CELL_BY_CELL",
            f"IMPORT_ZERO_{IMPORT_ZERO_DATE}",
        ]

        cursor = conn.execute(
            """
            INSERT INTO projects
            (
                name,
                client,
                start_date,
                end_date,
                status,
                note,
                old_id,
                is_overall,
                parent_overall_id,
                workshop_rollup
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                code,
                "",
                "",
                "",
                "chiuso" if closed else "attivo",
                " | ".join([part for part in note_parts if part]),
                old_id,
                is_overall,
                0,
                workshop_rollup,
            ),
        )

        code_to_new[up(code)] = cursor.lastrowid

    return code_to_new


def find_resource_id_by_short(short_name, resources_by_upper_name):
    target = up(short_name)
    if not target:
        return None

    # Se il callout contiene già il nome completo EXT o simile
    if target in resources_by_upper_name:
        return resources_by_upper_name[target]

    # Match su "COGNOME NOM"
    candidates = []
    for full_name, resource_id in resources_by_upper_name.items():
        parts = [p for p in full_name.split(" ") if p]
        if not parts:
            continue

        if len(parts) == 1:
            short = parts[0]
        else:
            surname = " ".join(parts[:-1])
            first3 = parts[-1][:3]
            short = f"{surname} {first3}"

        if up(short) == target:
            candidates.append(resource_id)

    if len(candidates) == 1:
        return candidates[0]

    return None


def get_or_create_virtual_ext_resource(conn, resources_by_upper_name, label, role):
    label_clean = clean(label)
    key = up(label_clean)

    if key in resources_by_upper_name:
        return resources_by_upper_name[key]

    cursor = conn.execute(
        """
        INSERT INTO resources
        (
            name,
            role,
            availability_note,
            is_active,
            old_id
        )
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            label_clean,
            norm_role(role),
            f"EXT | VIRTUAL_FROM_MATRIX | IMPORT_ZERO_{IMPORT_ZERO_DATE}",
            1,
            0,
        ),
    )

    new_id = cursor.lastrowid
    resources_by_upper_name[key] = new_id
    return new_id


def parse_callout_items(text):
    raw = clean(text)
    if not raw:
        return []

    return [clean(part) for part in raw.split(" | ") if clean(part)]


def chunk_active_resources(active_text):
    """
    active_resources è una stringa tipo:
    SEMERARO FRA | ok mans | 1:1 | GENERICO-EXT | EXT | 1:1

    La ricostruiamo in blocchi da 3 campi:
    nome, mansione/EXT, carico.
    """
    parts = parse_callout_items(active_text)
    chunks = []

    i = 0
    while i + 2 < len(parts):
        chunks.append({
            "name": parts[i],
            "mans": parts[i + 1],
            "load": parts[i + 2],
        })
        i += 3

    return chunks


def chunk_historical_resources(history_text):
    """
    historical_resources è tipo:
    DE VITA PAN | ok mans | 0:1 | FUORI_CONTRATTO

    blocchi da 4:
    nome, mansione, valore, motivo.
    """
    parts = parse_callout_items(history_text)
    chunks = []

    i = 0
    while i + 3 < len(parts):
        chunks.append({
            "name": parts[i],
            "mans": parts[i + 1],
            "load": parts[i + 2],
            "reason": parts[i + 3],
        })
        i += 4

    return chunks


def load_percent_from_label(label):
    raw = clean(label)
    if raw == "1:1":
        return 100.0
    if raw == "1/2":
        return 50.0
    if raw == "0:1":
        return 0.0

    try:
        return float(raw.replace(",", ".")) * 100
    except Exception:
        return 100.0


def insert_demands_from_matrix(conn, matrix_rows, project_code_to_id):
    inserted = 0

    for row in matrix_rows:
        required = parse_float(row.get("required_R"))
        if required == 0:
            continue

        project_id = project_code_to_id.get(up(row["visual_project_code"]))
        if not project_id:
            continue

        week = int(row["week"])
        period_key = int(row["period_key"] or period_key_from_week(week))
        role = norm_role(row["role"])

        conn.execute(
            """
            INSERT INTO demands
            (
                project_id,
                week,
                period_key,
                role,
                quantity,
                note
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                project_id,
                week,
                period_key,
                role,
                required,
                f"MATRIX_IMPORT_ZERO_{IMPORT_ZERO_DATE}",
            ),
        )
        inserted += 1

    return inserted


def insert_allocations_from_matrix(conn, matrix_rows, project_code_to_id, resources_by_upper_name):
    inserted_active = 0
    inserted_history = 0
    unresolved_active = []
    unresolved_history = []

    for row in matrix_rows:
        project_id = project_code_to_id.get(up(row["visual_project_code"]))
        if not project_id:
            continue

        project_name = clean(row["visual_project_code"])
        role = norm_role(row["role"])
        week = int(row["week"])
        period_key = int(row["period_key"] or period_key_from_week(week))

        for item in chunk_active_resources(row.get("active_resources", "")):
            name = clean(item["name"])
            mans = clean(item["mans"])
            load_label = clean(item["load"])
            is_ext = up(mans) == "EXT" or "-EXT" in up(name) or up(name).endswith(" EXT")
            load_percent = load_percent_from_label(load_label)

            if is_ext:
                resource_id = get_or_create_virtual_ext_resource(conn, resources_by_upper_name, name, role)
            else:
                resource_id = find_resource_id_by_short(name, resources_by_upper_name)

            if not resource_id:
                unresolved_active.append({
                    "project": project_name,
                    "role": role,
                    "week": week,
                    "name": name,
                    "mans": mans,
                    "load": load_label,
                })
                continue

            conn.execute(
                """
                INSERT INTO allocations
                (
                    resource_id,
                    project_id,
                    week,
                    period_key,
                    role,
                    hours,
                    load_percent,
                    note
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    resource_id,
                    project_id,
                    week,
                    period_key,
                    role,
                    40,
                    load_percent,
                    f"MATRIX_ACTIVE_IMPORT_ZERO_{IMPORT_ZERO_DATE} | {mans}",
                ),
            )
            inserted_active += 1

        for item in chunk_historical_resources(row.get("historical_resources", "")):
            name = clean(item["name"])
            mans = clean(item["mans"])
            load_label = clean(item["load"])
            reason = clean(item["reason"])
            load_percent = load_percent_from_label(load_label)

            resource_id = find_resource_id_by_short(name, resources_by_upper_name)

            if resource_id:
                resource = conn.execute(
                    """
                    SELECT name, role
                    FROM resources
                    WHERE id = ?
                    """,
                    (resource_id,),
                ).fetchone()
                resource_name = resource["name"] if resource else name
                resource_role = resource["role"] if resource else ""
            else:
                unresolved_history.append({
                    "project": project_name,
                    "role": role,
                    "week": week,
                    "name": name,
                    "mans": mans,
                    "load": load_label,
                    "reason": reason,
                })
                resource_name = name
                resource_role = ""

            conn.execute(
                """
                INSERT INTO allocation_history
                (
                    resource_id,
                    resource_name,
                    resource_role,
                    project_id,
                    project_name,
                    week,
                    period_key,
                    role,
                    hours,
                    load_percent,
                    reason,
                    note
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    resource_id,
                    resource_name,
                    resource_role,
                    project_id,
                    project_name,
                    week,
                    period_key,
                    role,
                    40,
                    0,
                    reason or "storico",
                    f"MATRIX_HISTORY_IMPORT_ZERO_{IMPORT_ZERO_DATE} | {mans} | origin_load={load_percent}",
                ),
            )
            inserted_history += 1

    return inserted_active, inserted_history, unresolved_active, unresolved_history


def write_import_meta(conn):
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS import_meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """
    )

    rows = [
        ("last_import_type", "MATRIX_CELL_BY_CELL_IMPORT"),
        ("last_import_date", IMPORT_ZERO_DATE),
        ("last_import_source", str(MATRIX_CSV)),
        ("note", "Import cella-per-cella da matrice vecchio planner. Nessuna demand_history generata."),
    ]

    conn.executemany(
        """
        INSERT OR REPLACE INTO import_meta (key, value)
        VALUES (?, ?)
        """,
        rows,
    )


def compute_operational_totals(conn):
    demand_rows = conn.execute(
        """
        SELECT project_id, role, period_key, SUM(quantity) AS r
        FROM demands
        GROUP BY project_id, role, period_key
        """
    ).fetchall()

    allocation_rows = conn.execute(
        """
        SELECT project_id, role, period_key, SUM(load_percent) / 100.0 AS a
        FROM allocations
        GROUP BY project_id, role, period_key
        """
    ).fetchall()

    dmap = {
        (row["project_id"], up(row["role"]), row["period_key"]): float(row["r"] or 0)
        for row in demand_rows
    }
    amap = {
        (row["project_id"], up(row["role"]), row["period_key"]): float(row["a"] or 0)
        for row in allocation_rows
    }

    keys = set(dmap.keys()) | set(amap.keys())

    total_r = sum(dmap.get(key, 0) for key in keys)
    total_a = sum(amap.get(key, 0) for key in keys)

    return total_r, total_a, len(keys)


def matrix_totals(matrix_rows):
    total_r = sum(parse_float(row.get("required_R")) for row in matrix_rows)
    total_a = sum(parse_float(row.get("allocated_A_numeric")) for row in matrix_rows)
    return total_r, total_a, len(matrix_rows)


def write_report(matrix_rows, demand_count, active_count, history_count, unresolved_active, unresolved_history, operational_totals):
    matrix_r, matrix_a, matrix_cells = matrix_totals(matrix_rows)
    op_r, op_a, op_cells = operational_totals

    lines = []
    lines.append("=== APPLY OLD PLANNER MATRIX REPORT ===")
    lines.append(f"Import zero: {IMPORT_ZERO_DATE}")
    lines.append(f"Matrice: {MATRIX_CSV}")
    lines.append("")
    lines.append("=== CONTEGGI SCRITTURA ===")
    lines.append(f"Demands inseriti: {demand_count}")
    lines.append(f"Allocazioni attive inserite: {active_count}")
    lines.append(f"Allocazioni storiche inserite: {history_count}")
    lines.append(f"Active unresolved: {len(unresolved_active)}")
    lines.append(f"History unresolved: {len(unresolved_history)}")
    lines.append("")
    lines.append("=== TOTALI MATRICE VS DB OPERATIVO ===")
    lines.append(f"Matrice celle: {matrix_cells} | R={matrix_r} | A={matrix_a} | D={matrix_r - matrix_a}")
    lines.append(f"DB celle: {op_cells} | R={op_r} | A={op_a} | D={op_r - op_a}")
    lines.append(f"Delta R: {op_r - matrix_r}")
    lines.append(f"Delta A: {op_a - matrix_a}")
    lines.append("")
    lines.append("=== ACTIVE UNRESOLVED PRIME 100 ===")
    for item in unresolved_active[:100]:
        lines.append(str(item))
    lines.append("")
    lines.append("=== HISTORY UNRESOLVED PRIME 100 ===")
    for item in unresolved_history[:100]:
        lines.append(str(item))

    REPORT_TXT.write_text("\n".join(lines), encoding="utf-8")


def main():
    matrix_rows = load_matrix_rows()
    conn = connect(DB_PATH)

    try:
        require_raw(conn)
        ensure_schema(conn)

        clear_operational_tables(conn)

        _, resources_by_upper_name = import_resources_from_raw(conn)
        project_code_to_id = import_projects_from_matrix(conn, matrix_rows)
        demand_count = insert_demands_from_matrix(conn, matrix_rows, project_code_to_id)
        active_count, history_count, unresolved_active, unresolved_history = insert_allocations_from_matrix(
            conn,
            matrix_rows,
            project_code_to_id,
            resources_by_upper_name,
        )

        write_import_meta(conn)

        operational_totals = compute_operational_totals(conn)

        conn.commit()

        write_report(
            matrix_rows,
            demand_count,
            active_count,
            history_count,
            unresolved_active,
            unresolved_history,
            operational_totals,
        )

        print("IMPORT DA MATRICE COMPLETATO")
        print(f"Demands inseriti: {demand_count}")
        print(f"Allocazioni attive inserite: {active_count}")
        print(f"Allocazioni storiche inserite: {history_count}")
        print(f"Active unresolved: {len(unresolved_active)}")
        print(f"History unresolved: {len(unresolved_history)}")
        print(f"Report: {REPORT_TXT}")
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()