import sqlite3
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "planner_v2.sqlite3"

DEFAULT_YEAR_SHORT = 26


def get_connection():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def table_exists(conn, table_name):
    row = conn.execute(
        """
        SELECT name
        FROM sqlite_master
        WHERE type = 'table'
          AND name = ?
        """,
        (table_name,),
    ).fetchone()
    return row is not None


def column_exists(conn, table_name, column_name):
    if not table_exists(conn, table_name):
        return False
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return any(row["name"] == column_name for row in rows)


def ensure_column(conn, table_name, column_name, column_definition):
    if table_exists(conn, table_name) and not column_exists(conn, table_name, column_name):
        conn.execute(
            f"""
            ALTER TABLE {table_name}
            ADD COLUMN {column_name} {column_definition}
            """
        )


def normalize_period_key_from_week(week, default_year_short=DEFAULT_YEAR_SHORT):
    try:
        week_int = int(week or 0)
    except Exception:
        week_int = 0

    if week_int <= 0:
        return 0

    if week_int >= 1000:
        return week_int

    return int(default_year_short) * 100 + week_int


def backfill_period_key(conn, table_name):
    if not table_exists(conn, table_name):
        return

    if not column_exists(conn, table_name, "period_key"):
        return

    if not column_exists(conn, table_name, "week"):
        return

    rows = conn.execute(
        f"""
        SELECT id, week, period_key
        FROM {table_name}
        WHERE period_key IS NULL
           OR period_key = 0
        """
    ).fetchall()

    for row in rows:
        period_key = normalize_period_key_from_week(row["week"])
        conn.execute(
            f"""
            UPDATE {table_name}
            SET period_key = ?
            WHERE id = ?
            """,
            (period_key, row["id"]),
        )


def init_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = get_connection()

    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS resources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                role TEXT DEFAULT '',
                availability_note TEXT DEFAULT '',
                is_active INTEGER DEFAULT 1,
                old_id INTEGER DEFAULT 0
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                client TEXT DEFAULT '',
                start_date TEXT DEFAULT '',
                end_date TEXT DEFAULT '',
                status TEXT DEFAULT 'attivo',
                note TEXT DEFAULT '',
                old_id INTEGER DEFAULT 0,
                is_overall INTEGER DEFAULT 0,
                parent_overall_id INTEGER DEFAULT 0,
                workshop_rollup INTEGER DEFAULT 0
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS demands (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                week INTEGER NOT NULL,
                period_key INTEGER DEFAULT 0,
                role TEXT NOT NULL,
                quantity REAL DEFAULT 0,
                note TEXT DEFAULT '',
                FOREIGN KEY(project_id) REFERENCES projects(id)
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS allocations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                resource_id INTEGER NOT NULL,
                project_id INTEGER NOT NULL,
                week INTEGER NOT NULL,
                period_key INTEGER DEFAULT 0,
                role TEXT DEFAULT '',
                hours REAL DEFAULT 40,
                load_percent REAL DEFAULT 100,
                note TEXT DEFAULT '',
                FOREIGN KEY(resource_id) REFERENCES resources(id),
                FOREIGN KEY(project_id) REFERENCES projects(id)
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS demand_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                week INTEGER NOT NULL,
                period_key INTEGER DEFAULT 0,
                role TEXT NOT NULL,
                old_quantity REAL DEFAULT 0,
                new_quantity REAL DEFAULT 0,
                note TEXT DEFAULT '',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(project_id) REFERENCES projects(id)
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS allocation_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                resource_id INTEGER,
                resource_name TEXT DEFAULT '',
                resource_role TEXT DEFAULT '',
                project_id INTEGER,
                project_name TEXT DEFAULT '',
                week INTEGER NOT NULL,
                period_key INTEGER DEFAULT 0,
                role TEXT DEFAULT '',
                hours REAL DEFAULT 40,
                load_percent REAL DEFAULT 100,
                reason TEXT DEFAULT '',
                note TEXT DEFAULT '',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

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

        ensure_column(conn, "resources", "old_id", "INTEGER DEFAULT 0")

        ensure_column(conn, "projects", "old_id", "INTEGER DEFAULT 0")
        ensure_column(conn, "projects", "is_overall", "INTEGER DEFAULT 0")
        ensure_column(conn, "projects", "parent_overall_id", "INTEGER DEFAULT 0")
        ensure_column(conn, "projects", "workshop_rollup", "INTEGER DEFAULT 0")

        ensure_column(conn, "demands", "period_key", "INTEGER DEFAULT 0")
        ensure_column(conn, "allocations", "period_key", "INTEGER DEFAULT 0")
        ensure_column(conn, "demand_history", "period_key", "INTEGER DEFAULT 0")
        ensure_column(conn, "allocation_history", "period_key", "INTEGER DEFAULT 0")

        backfill_period_key(conn, "demands")
        backfill_period_key(conn, "allocations")
        backfill_period_key(conn, "demand_history")
        backfill_period_key(conn, "allocation_history")

        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_projects_old_id
            ON projects(old_id)
            """
        )

        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_projects_overall
            ON projects(is_overall, parent_overall_id, workshop_rollup)
            """
        )

        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_resources_old_id
            ON resources(old_id)
            """
        )

        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_demands_project_role_period
            ON demands(project_id, role, period_key)
            """
        )

        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_allocations_resource_period
            ON allocations(resource_id, period_key)
            """
        )

        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_allocations_project_role_period
            ON allocations(project_id, role, period_key)
            """
        )

        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_workshop_rollup_lookup
            ON workshop_rollup_sources(overall_project_id, role, period_key)
            """
        )

        seed_if_empty(conn)
        conn.commit()
    finally:
        conn.close()


def seed_if_empty(conn):
    resource_count = conn.execute("SELECT COUNT(*) AS n FROM resources").fetchone()["n"]
    project_count = conn.execute("SELECT COUNT(*) AS n FROM projects").fetchone()["n"]

    if resource_count == 0:
        conn.executemany(
            """
            INSERT INTO resources (name, role, availability_note, is_active, old_id)
            VALUES (?, ?, ?, ?, ?)
            """,
            [
                ("MARIO ROSSI", "TUBISTA", "", 1, 0),
                ("Luigi Bianchi", "CARPENTIERE", "", 1, 0),
                ("Anna Verdi", "SALDATORE", "", 1, 0),
            ],
        )

    if project_count == 0:
        conn.executemany(
            """
            INSERT INTO projects
            (name, client, start_date, end_date, status, note, old_id, is_overall, parent_overall_id, workshop_rollup)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                ("Commessa Alfa", "Cliente A", "2026-04-01", "2026-06-30", "attivo", "", 0, 0, 0, 0),
                ("Commessa Beta", "Cliente B", "2026-05-01", "2026-07-31", "attivo", "", 0, 0, 0, 0),
            ],
        )

    demand_count = conn.execute("SELECT COUNT(*) AS n FROM demands").fetchone()["n"]

    if demand_count == 0:
        projects = conn.execute("SELECT id, name FROM projects ORDER BY id ASC").fetchall()
        project_by_name = {row["name"]: row["id"] for row in projects}

        seed_demands = [
            (project_by_name.get("Commessa Alfa"), 17, 2617, "TUBISTA", 2, "Seed"),
            (project_by_name.get("Commessa Alfa"), 18, 2618, "TUBISTA", 2, "Seed"),
            (project_by_name.get("Commessa Alfa"), 19, 2619, "TUBISTA", 1, "Seed"),
            (project_by_name.get("Commessa Alfa"), 18, 2618, "CARPENTIERE", 1, "Seed"),
            (project_by_name.get("Commessa Beta"), 19, 2619, "CARPENTIERE", 1, "Seed"),
            (project_by_name.get("Commessa Beta"), 20, 2620, "CARPENTIERE", 2, "Seed"),
        ]

        conn.executemany(
            """
            INSERT INTO demands (project_id, week, period_key, role, quantity, note)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            [row for row in seed_demands if row[0] is not None],
        )