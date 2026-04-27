import sqlite3
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
NEW_DB_PATH = BASE_DIR / "data" / "planner_v2.sqlite3"
OLD_DB_PATH = BASE_DIR / "data" / "old_planner.db"


TABLES_TO_COPY = [
    "projects",
    "resources",
    "demands",
    "allocations",
    "unavailability",
    "activity_log",
    "meta",
    "roles",
    "project_templates",
]


def connect(path):
    conn = sqlite3.connect(path)
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


def get_create_sql(conn, table_name):
    row = conn.execute(
        """
        SELECT sql
        FROM sqlite_master
        WHERE type = 'table'
          AND name = ?
        """,
        (table_name,),
    ).fetchone()
    return row["sql"] if row else None


def get_columns(conn, table_name):
    return [row["name"] for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()]


def copy_table_raw(old_conn, new_conn, table_name):
    if not table_exists(old_conn, table_name):
        print(f"SKIP {table_name}: non esiste nel vecchio DB")
        return 0

    raw_table = f"raw_old_{table_name}"

    old_create_sql = get_create_sql(old_conn, table_name)
    if not old_create_sql:
        print(f"SKIP {table_name}: create SQL non trovato")
        return 0

    raw_create_sql = old_create_sql.replace(
        f"CREATE TABLE {table_name}",
        f"CREATE TABLE {raw_table}",
        1,
    ).replace(
        f"CREATE TABLE IF NOT EXISTS {table_name}",
        f"CREATE TABLE IF NOT EXISTS {raw_table}",
        1,
    )

    new_conn.execute(f"DROP TABLE IF EXISTS {raw_table}")
    new_conn.execute(raw_create_sql)

    cols = get_columns(old_conn, table_name)
    col_list = ", ".join(cols)
    placeholders = ", ".join(["?"] * len(cols))

    rows = old_conn.execute(f"SELECT {col_list} FROM {table_name}").fetchall()

    if rows:
        new_conn.executemany(
            f"""
            INSERT INTO {raw_table} ({col_list})
            VALUES ({placeholders})
            """,
            [[row[col] for col in cols] for row in rows],
        )

    print(f"OK {table_name} -> {raw_table}: {len(rows)} righe")
    return len(rows)


def create_raw_indexes(new_conn):
    index_statements = [
        """
        CREATE INDEX IF NOT EXISTS idx_raw_old_demands_project_role_week
        ON raw_old_demands(project_id, role, week)
        """,
        """
        CREATE INDEX IF NOT EXISTS idx_raw_old_allocations_project_role_range
        ON raw_old_allocations(project_id, role, week_from, week_to)
        """,
        """
        CREATE INDEX IF NOT EXISTS idx_raw_old_allocations_resource_range
        ON raw_old_allocations(resource_id, week_from, week_to)
        """,
        """
        CREATE INDEX IF NOT EXISTS idx_raw_old_projects_code
        ON raw_old_projects(code)
        """,
        """
        CREATE INDEX IF NOT EXISTS idx_raw_old_resources_name
        ON raw_old_resources(name)
        """,
    ]

    for stmt in index_statements:
        try:
            new_conn.execute(stmt)
        except sqlite3.OperationalError as exc:
            print(f"INDEX SKIP: {exc}")


def create_raw_import_meta(new_conn):
    new_conn.execute(
        """
        CREATE TABLE IF NOT EXISTS raw_old_import_meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """
    )

    rows = [
        ("source_file", str(OLD_DB_PATH)),
        ("import_type", "RAW_COPY_NO_INTERPRETATION"),
        ("note", "Copia fedele tabelle vecchio planner. Non usare come modifica utente."),
        ("fabbisogno_import_date", "07/04/2026"),
    ]

    new_conn.executemany(
        """
        INSERT OR REPLACE INTO raw_old_import_meta (key, value)
        VALUES (?, ?)
        """,
        rows,
    )


def main():
    if not OLD_DB_PATH.exists():
        raise FileNotFoundError(f"Non trovo il vecchio DB: {OLD_DB_PATH}")

    if not NEW_DB_PATH.exists():
        raise FileNotFoundError(f"Non trovo il nuovo DB: {NEW_DB_PATH}")

    old_conn = connect(OLD_DB_PATH)
    new_conn = connect(NEW_DB_PATH)

    try:
        print("COPIA RAW VECCHIO PLANNER")
        print(f"OLD: {OLD_DB_PATH}")
        print(f"NEW: {NEW_DB_PATH}")
        print("")

        total = 0
        for table in TABLES_TO_COPY:
            total += copy_table_raw(old_conn, new_conn, table)

        create_raw_indexes(new_conn)
        create_raw_import_meta(new_conn)

        new_conn.commit()

        print("")
        print("RAW COPY COMPLETATA")
        print(f"Totale righe copiate: {total}")
        print("")
        print("Tabelle create nel nuovo DB:")
        for table in TABLES_TO_COPY:
            raw_table = f"raw_old_{table}"
            if table_exists(new_conn, raw_table):
                count = new_conn.execute(f"SELECT COUNT(*) AS n FROM {raw_table}").fetchone()["n"]
                print(f"- {raw_table}: {count}")
        print("- raw_old_import_meta")
    except Exception:
        new_conn.rollback()
        raise
    finally:
        old_conn.close()
        new_conn.close()


if __name__ == "__main__":
    main()