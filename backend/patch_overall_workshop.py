import sqlite3
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "planner_v2.sqlite3"


def rows_to_dicts(cursor):
    columns = [column[0] for column in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def table_exists(conn, name):
    row = conn.execute(
        """
        SELECT name
        FROM sqlite_master
        WHERE type = 'table'
          AND name = ?
        """,
        (name,),
    ).fetchone()
    return row is not None


def get_columns(conn, table):
    rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    return [row[1] for row in rows]


def pick_column(columns, candidates):
    lowered = {column.lower(): column for column in columns}
    for candidate in candidates:
        if candidate.lower() in lowered:
            return lowered[candidate.lower()]
    return None


def normalize(value):
    return str(value or "").strip().upper()


def main():
    conn = sqlite3.connect(DB_PATH)

    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS workshop_rollup_sources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                overall_project_id INTEGER NOT NULL,
                source_old_project_id INTEGER,
                source_project_name TEXT NOT NULL,
                role TEXT NOT NULL,
                week INTEGER NOT NULL,
                period_key INTEGER NOT NULL,
                required REAL NOT NULL DEFAULT 0,
                source TEXT NOT NULL DEFAULT 'RAW_OLD',
                UNIQUE(overall_project_id, source_old_project_id, source_project_name, role, period_key)
            )
            """
        )

        overall = conn.execute(
            """
            SELECT id, name
            FROM projects
            WHERE UPPER(name) LIKE '%OVERALL OFFICINA%'
               OR UPPER(note) LIKE '%OVERALL OFFICINA%'
               OR UPPER(note) LIKE '%WORKSHOP_ROLLUP%'
            ORDER BY id
            LIMIT 1
            """
        ).fetchone()

        if not overall:
            raise RuntimeError("OVERALL OFFICINA non trovato in projects")

        overall_project_id = int(overall[0])

        conn.execute("DELETE FROM workshop_rollup_sources")

        if not table_exists(conn, "raw_old_projects") or not table_exists(conn, "raw_old_demands"):
            raise RuntimeError("raw_old_projects/raw_old_demands non presenti nel DB")

        project_cols = get_columns(conn, "raw_old_projects")
        demand_cols = get_columns(conn, "raw_old_demands")

        p_id = pick_column(project_cols, ["id", "project_id"])
        p_name = pick_column(project_cols, ["name", "project_name", "code", "commessa"])
        p_type = pick_column(project_cols, ["type", "project_type", "kind"])
        p_note = pick_column(project_cols, ["note", "notes", "description"])

        d_project_id = pick_column(demand_cols, ["project_id", "project", "commessa_id"])
        d_role = pick_column(demand_cols, ["role", "mansione"])
        d_week = pick_column(demand_cols, ["week", "week_from"])
        d_period_key = pick_column(demand_cols, ["period_key"])
        d_quantity = pick_column(demand_cols, ["quantity", "qty", "required", "value"])

        missing = []
        for label, column in {
            "raw_old_projects.id": p_id,
            "raw_old_projects.name": p_name,
            "raw_old_demands.project_id": d_project_id,
            "raw_old_demands.role": d_role,
            "raw_old_demands.quantity": d_quantity,
        }.items():
            if not column:
                missing.append(label)

        if missing:
            raise RuntimeError("Colonne mancanti: " + ", ".join(missing))

        old_projects = rows_to_dicts(
            conn.execute(f"""
                SELECT *
                FROM raw_old_projects
            """)
        )

        workshop_project_ids = set()
        project_name_by_old_id = {}

        for project in old_projects:
            old_id = project.get(p_id)
            name = str(project.get(p_name) or "").strip()
            type_value = str(project.get(p_type) or "").strip() if p_type else ""
            note_value = str(project.get(p_note) or "").strip() if p_note else ""

            text = normalize(f"{name} {type_value} {note_value}")

            is_overall = "OVERALL OFFICINA" in text
            is_workshop_child = (
                not is_overall
                and (
                    "OFFICINA" in text
                    or "WORKSHOP" in text
                    or "ROLLUP" in text
                    or "TYPE=WS" in text
                    or type_value.upper() in {"WS", "OFFICINA"}
                )
            )

            if is_workshop_child:
                try:
                    workshop_project_ids.add(int(old_id))
                    project_name_by_old_id[int(old_id)] = name
                except Exception:
                    pass

        if not workshop_project_ids:
            print("ATTENZIONE: nessuna figlia officina riconosciuta da raw_old_projects.")
            print("Provo fallback: progetti raw con nome contenente 401/414/415/416/417/418/419 e OFFICINA non rilevabile.")
            # Nessun fallback cieco: meglio non inventare sorgenti.
            # Verrà lasciata tabella vuota e il report dirà cosa manca.

        inserted = 0

        if workshop_project_ids:
            placeholders = ",".join("?" for _ in workshop_project_ids)
            period_expr = (
                f"COALESCE(NULLIF({d_period_key}, 0), 2600 + {d_week})"
                if d_period_key and d_week
                else f"(2600 + {d_week})"
            )
            week_expr = (
                f"CASE WHEN {d_week} IS NOT NULL THEN {d_week} ELSE ({period_expr} % 100) END"
                if d_week
                else f"({period_expr} % 100)"
            )

            sql = f"""
                SELECT
                    {d_project_id} AS old_project_id,
                    {d_role} AS role,
                    {week_expr} AS week,
                    {period_expr} AS period_key,
                    SUM(COALESCE({d_quantity}, 0)) AS required
                FROM raw_old_demands
                WHERE {d_project_id} IN ({placeholders})
                GROUP BY {d_project_id}, {d_role}, {period_expr}
                HAVING SUM(COALESCE({d_quantity}, 0)) <> 0
                ORDER BY {period_expr}, {d_project_id}, {d_role}
            """

            rows = rows_to_dicts(conn.execute(sql, tuple(workshop_project_ids)))

            for row in rows:
                old_project_id = int(row["old_project_id"])
                source_project_name = project_name_by_old_id.get(old_project_id, f"OLD_PROJECT_{old_project_id}")
                role = str(row["role"] or "").strip().upper()
                week = int(row["week"] or 0)
                period_key = int(row["period_key"] or 0)
                required = float(row["required"] or 0)

                if not role or not period_key:
                    continue

                conn.execute(
                    """
                    INSERT OR REPLACE INTO workshop_rollup_sources (
                        overall_project_id,
                        source_old_project_id,
                        source_project_name,
                        role,
                        week,
                        period_key,
                        required,
                        source
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        overall_project_id,
                        old_project_id,
                        source_project_name,
                        role,
                        week,
                        period_key,
                        required,
                        "RAW_OLD_DEMANDS",
                    ),
                )
                inserted += 1

        conn.commit()

        print("OK: workshop_rollup_sources ricostruita")
        print(f"DB: {DB_PATH}")
        print(f"overall_project_id: {overall_project_id}")
        print(f"figlie officina riconosciute: {len(workshop_project_ids)}")
        print(f"righe inserite: {inserted}")

        print()
        print("ANTEPRIMA:")
        preview = rows_to_dicts(
            conn.execute(
                """
                SELECT
                    source_old_project_id,
                    source_project_name,
                    role,
                    week,
                    period_key,
                    required
                FROM workshop_rollup_sources
                ORDER BY period_key, source_project_name, role
                LIMIT 80
                """
            )
        )
        for row in preview:
            print(row)

    finally:
        conn.close()


if __name__ == "__main__":
    main()