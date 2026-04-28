from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from app.db import init_db, get_connection
from app.old_logic_adapter import build_planner_matrix, get_workshop_breakdown


app = FastAPI(
    title="Local Planner App V2",
    version="0.1.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


DEFAULT_YEAR_SHORT = 26


class ResourceCreate(BaseModel):
    name: str
    role: str = ""
    availability_note: str = ""
    is_active: int = 1


class ProjectCreate(BaseModel):
    name: str
    client: str = ""
    start_date: str = ""
    end_date: str = ""
    status: str = "attivo"
    note: str = ""


class DemandCreate(BaseModel):
    project_id: int
    week: int
    role: str
    quantity: float = 0
    note: str = ""


class DemandRangeUpsert(BaseModel):
    project_id: int
    role: str
    week_from: int
    week_to: int
    quantity: float
    note: str = ""


class AllocationCreate(BaseModel):
    resource_id: int
    project_id: int
    week: int
    role: str = ""
    hours: float = 40
    load_percent: float = 100
    note: str = ""


class AllocationRangePayload(BaseModel):
    resource_id: int
    project_id: int
    role: str
    week_from: int
    week_to: int
    hours: float = 40
    load_percent: float = 100
    note: str = ""


class AllocationResolvePayload(BaseModel):
    resource_id: int
    project_id: int
    role: str
    week: int
    mode: str
    remove_allocation_id: int | None = None
    hours: float = 40
    note: str = ""


@app.on_event("startup")
def on_startup():
    init_db()


@app.get("/")
def root():
    return {"message": "Local Planner App V2 backend attivo"}


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.get("/api/info")
def info():
    return {
        "app": "Local Planner App V2",
        "backend": "FastAPI",
        "database": "SQLite"
    }
@app.get("/api/planner-matrix-old-logic")
def planner_matrix_old_logic(
    show_zero: bool = False,
    project: str | None = None,
    role: str | None = None
):
    return build_planner_matrix(
        show_zero_demand_projects=show_zero,
        project_filter=project,
        role_filter=role
    )


@app.get("/api/workshop-breakdown")
@app.get("/api/workshop-breakdown")
def workshop_breakdown(
    project_id: int | None = None,
    role: str | None = None,
    week: int | None = None,
    period_key: int | None = None
):
    selected_period_key = int(period_key or (2600 + int(week or 0)))
    selected_role = str(role or "").strip().upper()

    conn = get_connection()
    try:
        overall = conn.execute(
            """
            SELECT id, name
            FROM projects
            WHERE id = ?
               OR UPPER(name) LIKE '%OVERALL OFFICINA%'
               OR UPPER(note) LIKE '%WORKSHOP_ROLLUP%'
            ORDER BY CASE WHEN id = ? THEN 0 ELSE 1 END, id
            LIMIT 1
            """,
            (project_id or 0, project_id or 0),
        ).fetchone()

        if not overall:
            return {
                "project_id": project_id,
                "role": selected_role,
                "week": selected_period_key % 100,
                "period_key": selected_period_key,
                "sources": [],
                "total_required": 0,
            }

        overall_project_id = int(overall["id"])

        rows = conn.execute(
            """
            SELECT
                source_old_project_id,
                source_project_name,
                role,
                week,
                period_key,
                required
            FROM workshop_rollup_sources
            WHERE overall_project_id = ?
              AND period_key = ?
              AND UPPER(role) = ?
              AND COALESCE(required, 0) <> 0
            ORDER BY source_project_name
            """,
            (overall_project_id, selected_period_key, selected_role),
        ).fetchall()

        sources = [
            {
                "source_old_project_id": row["source_old_project_id"],
                "project_name": row["source_project_name"],
                "source_project_name": row["source_project_name"],
                "role": row["role"],
                "week": row["week"],
                "period_key": row["period_key"],
                "required": row["required"],
            }
            for row in rows
        ]

        return {
            "project_id": overall_project_id,
            "project_name": overall["name"],
            "role": selected_role,
            "week": selected_period_key % 100,
            "period_key": selected_period_key,
            "sources": sources,
            "total_required": sum(float(source["required"] or 0) for source in sources),
        }
    finally:
        conn.close()

def clean_text(value):
    return str(value or "").strip()


def normalize_role(value):
    return clean_text(value).upper()


def period_key_from_week(week: int, year_short: int = DEFAULT_YEAR_SHORT):
    week_int = int(week or 0)

    if week_int >= 1000:
        return week_int

    return int(year_short) * 100 + week_int


def week_from_period_key(period_key: int):
    value = int(period_key or 0)

    if value >= 1000:
        return value % 100

    return value


def is_external_text(name: str = "", role: str = ""):
    name_text = normalize_role(name)
    role_text = normalize_role(role)

    return (
        "-EXT" in name_text
        or name_text.endswith(" EXT")
        or "-EXT" in role_text
        or role_text.endswith(" EXT")
    )


def get_resource_or_none(conn, resource_id: int):
    return conn.execute(
        """
        SELECT id, name, role, availability_note, is_active
        FROM resources
        WHERE id = ?
        """,
        (resource_id,),
    ).fetchone()


def resource_unavailable_reason(resource):
    if not resource:
        return "resource_not_found"

    if is_external_text(resource["name"], resource["role"]):
        return ""

    if int(resource["is_active"]) != 1:
        return "cessato"

    note_upper = str(resource["availability_note"] or "").upper()

    if "INDISP" in note_upper:
        return "indisp"

    if "NON DISPONIBILE" in note_upper:
        return "indisp"

    return ""


def resource_is_unavailable(resource):
    return resource_unavailable_reason(resource) != ""


def get_resource_week_allocations(conn, resource_id: int, week: int):
    period_key = period_key_from_week(week)

    rows = conn.execute(
        """
        SELECT
            a.id,
            a.resource_id,
            r.name AS resource_name,
            r.role AS resource_role,
            r.availability_note AS resource_availability_note,
            r.is_active AS resource_is_active,
            a.project_id,
            p.name AS project_name,
            a.week,
            a.period_key,
            a.role,
            a.hours,
            a.load_percent,
            a.note
        FROM allocations a
        JOIN resources r ON r.id = a.resource_id
        JOIN projects p ON p.id = a.project_id
        WHERE a.resource_id = ?
          AND COALESCE(NULLIF(a.period_key, 0), 2600 + a.week) = ?
        ORDER BY a.id ASC
        """,
        (resource_id, period_key),
    ).fetchall()

    return [dict(row) for row in rows]


def move_allocation_to_history(conn, allocation_id: int, reason: str, note: str = ""):
    row = conn.execute(
        """
        SELECT
            a.id,
            a.resource_id,
            r.name AS resource_name,
            r.role AS resource_role,
            a.project_id,
            p.name AS project_name,
            a.week,
            COALESCE(NULLIF(a.period_key, 0), 2600 + a.week) AS period_key,
            a.role,
            a.hours,
            a.load_percent,
            a.note
        FROM allocations a
        JOIN resources r ON r.id = a.resource_id
        JOIN projects p ON p.id = a.project_id
        WHERE a.id = ?
        """,
        (allocation_id,),
    ).fetchone()

    if not row:
        return False

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
            row["resource_id"],
            row["resource_name"],
            row["resource_role"],
            row["project_id"],
            row["project_name"],
            row["week"],
            row["period_key"],
            row["role"],
            row["hours"],
            row["load_percent"],
            reason,
            note or row["note"] or "",
        ),
    )

    conn.execute(
        """
        DELETE FROM allocations
        WHERE id = ?
        """,
        (allocation_id,),
    )

    return True


def delete_allocation_to_history_best_effort(conn, allocation_id: int, reason: str, note: str = ""):
    ok = move_allocation_to_history(conn, allocation_id, reason, note)
    if ok:
        return "historized"

    conn.execute(
        """
        DELETE FROM allocations
        WHERE id = ?
        """,
        (allocation_id,),
    )
    return "deleted"


def rebalance_resource_week(conn, resource_id: int, week: int):
    resource = get_resource_or_none(conn, resource_id)
    if resource and is_external_text(resource["name"], resource["role"]):
        return {"count": 0, "load_percent": 100, "external": True}

    period_key = period_key_from_week(week)

    rows = conn.execute(
        """
        SELECT id
        FROM allocations
        WHERE resource_id = ?
          AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
        ORDER BY id ASC
        """,
        (resource_id, period_key),
    ).fetchall()

    ids = [row["id"] for row in rows]

    if len(ids) == 0:
        return {"count": 0, "load_percent": 0}

    if len(ids) == 1:
        conn.execute(
            """
            UPDATE allocations
            SET load_percent = 100
            WHERE id = ?
            """,
            (ids[0],),
        )
        return {"count": 1, "load_percent": 100}

    if len(ids) == 2:
        conn.execute(
            f"""
            UPDATE allocations
            SET load_percent = 50
            WHERE id IN ({",".join(["?"] * len(ids))})
            """,
            ids,
        )
        return {"count": 2, "load_percent": 50}

    keep_ids = ids[:2]
    delete_ids = ids[2:]

    conn.execute(
        f"""
        UPDATE allocations
        SET load_percent = 50
        WHERE id IN ({",".join(["?"] * len(keep_ids))})
        """,
        keep_ids,
    )

    for delete_id in delete_ids:
        delete_allocation_to_history_best_effort(
            conn,
            delete_id,
            "rimosso_extra",
            "Rimosso perché la risorsa aveva più di 2 allocazioni nello stesso periodo",
        )

    return {
        "count": 2,
        "load_percent": 50,
        "deleted_ids": delete_ids,
    }


def insert_allocation(conn, resource_id: int, project_id: int, week: int, role: str, hours: float, load_percent: float, note: str):
    period_key = period_key_from_week(week)

    cursor = conn.execute(
        """
        INSERT INTO allocations (resource_id, project_id, week, period_key, role, hours, load_percent, note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            resource_id,
            project_id,
            week_from_period_key(period_key),
            period_key,
            normalize_role(role),
            hours,
            load_percent,
            clean_text(note),
        ),
    )
    return cursor.lastrowid


def demand_exists_for_project_role(conn, project_id: int, role: str):
    row = conn.execute(
        """
        SELECT id
        FROM demands
        WHERE project_id = ?
          AND UPPER(role) = ?
        LIMIT 1
        """,
        (project_id, normalize_role(role)),
    ).fetchone()

    return row is not None


def get_demand_quantity(conn, project_id: int, role: str, week: int):
    period_key = period_key_from_week(week)

    row = conn.execute(
        """
        SELECT quantity
        FROM demands
        WHERE project_id = ?
          AND UPPER(role) = ?
          AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
        LIMIT 1
        """,
        (project_id, normalize_role(role), period_key),
    ).fetchone()

    if not row:
        return 0

    return float(row["quantity"] or 0)


def release_allocations_for_zero_demand(conn, project_id: int, role: str, week: int):
    quantity = get_demand_quantity(conn, project_id, role, week)
    if quantity > 0:
        return []

    period_key = period_key_from_week(week)

    rows = conn.execute(
        """
        SELECT id, resource_id
        FROM allocations
        WHERE project_id = ?
          AND UPPER(role) = ?
          AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
        """,
        (project_id, normalize_role(role), period_key),
    ).fetchall()

    released = []

    for row in rows:
        allocation_id = row["id"]
        resource_id = row["resource_id"]
        ok = move_allocation_to_history(
            conn,
            allocation_id,
            "fabb 0",
            "Allocazione non più conteggiata perché il fabbisogno della cella è stato portato a 0",
        )
        if ok:
            released.append(allocation_id)
            rebalance_resource_week(conn, resource_id, week)

    return released


def release_unavailable_allocations(conn):
    rows = conn.execute(
        """
        SELECT
            a.id,
            a.resource_id,
            a.week,
            r.name,
            r.role,
            r.is_active,
            r.availability_note
        FROM allocations a
        JOIN resources r ON r.id = a.resource_id
        """
    ).fetchall()

    released = []

    for row in rows:
        resource = {
            "name": row["name"],
            "role": row["role"],
            "is_active": row["is_active"],
            "availability_note": row["availability_note"],
        }
        reason = resource_unavailable_reason(resource)
        if reason in ("cessato", "indisp"):
            ok = move_allocation_to_history(
                conn,
                row["id"],
                reason,
                f"Allocazione non più conteggiata perché la risorsa è {reason}",
            )
            if ok:
                released.append(row["id"])
                rebalance_resource_week(conn, row["resource_id"], row["week"])

    return released


def cleanup_allocations(conn):
    stats = {
        "orphan_resource_deleted": 0,
        "orphan_project_deleted": 0,
        "role_normalized": 0,
        "duplicate_removed": 0,
        "extra_over_2_removed": 0,
        "weeks_rebalanced": 0,
    }

    orphan_resource_rows = conn.execute(
        """
        SELECT a.id
        FROM allocations a
        LEFT JOIN resources r ON r.id = a.resource_id
        WHERE r.id IS NULL
        """
    ).fetchall()

    for row in orphan_resource_rows:
        conn.execute("DELETE FROM allocations WHERE id = ?", (row["id"],))
        stats["orphan_resource_deleted"] += 1

    orphan_project_rows = conn.execute(
        """
        SELECT a.id
        FROM allocations a
        LEFT JOIN projects p ON p.id = a.project_id
        WHERE p.id IS NULL
        """
    ).fetchall()

    for row in orphan_project_rows:
        conn.execute("DELETE FROM allocations WHERE id = ?", (row["id"],))
        stats["orphan_project_deleted"] += 1

    rows = conn.execute(
        """
        SELECT id, role
        FROM allocations
        """
    ).fetchall()

    for row in rows:
        normalized = normalize_role(row["role"])
        if row["role"] != normalized:
            conn.execute(
                """
                UPDATE allocations
                SET role = ?
                WHERE id = ?
                """,
                (normalized, row["id"]),
            )
            stats["role_normalized"] += 1

    duplicate_rows = conn.execute(
        """
        SELECT
            resource_id,
            project_id,
            COALESCE(NULLIF(period_key, 0), 2600 + week) AS period_key,
            UPPER(role) AS role,
            COUNT(*) AS n,
            MIN(id) AS keep_id
        FROM allocations
        GROUP BY resource_id, project_id, COALESCE(NULLIF(period_key, 0), 2600 + week), UPPER(role)
        HAVING COUNT(*) > 1
        """
    ).fetchall()

    for duplicate in duplicate_rows:
        rows_to_remove = conn.execute(
            """
            SELECT id, resource_id, week
            FROM allocations
            WHERE resource_id = ?
              AND project_id = ?
              AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
              AND UPPER(role) = ?
              AND id <> ?
            ORDER BY id ASC
            """,
            (
                duplicate["resource_id"],
                duplicate["project_id"],
                duplicate["period_key"],
                duplicate["role"],
                duplicate["keep_id"],
            ),
        ).fetchall()

        for row in rows_to_remove:
            delete_allocation_to_history_best_effort(
                conn,
                row["id"],
                "duplicato",
                "Allocazione duplicata rimossa da pulizia dati",
            )
            stats["duplicate_removed"] += 1

    resource_period_rows = conn.execute(
        """
        SELECT
            a.resource_id,
            COALESCE(NULLIF(a.period_key, 0), 2600 + a.week) AS period_key,
            COUNT(*) AS n
        FROM allocations a
        JOIN resources r ON r.id = a.resource_id
        WHERE UPPER(r.name) NOT LIKE '%-EXT%'
          AND UPPER(r.name) NOT LIKE '% EXT'
          AND UPPER(r.role) NOT LIKE '%-EXT%'
          AND UPPER(r.role) NOT LIKE '% EXT'
        GROUP BY a.resource_id, COALESCE(NULLIF(a.period_key, 0), 2600 + a.week)
        HAVING COUNT(*) > 2
        """
    ).fetchall()

    for row in resource_period_rows:
        week = week_from_period_key(row["period_key"])

        before_ids = [
            item["id"]
            for item in conn.execute(
                """
                SELECT id
                FROM allocations
                WHERE resource_id = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                ORDER BY id ASC
                """,
                (row["resource_id"], row["period_key"]),
            ).fetchall()
        ]

        rebalance_resource_week(conn, row["resource_id"], week)

        after_ids = [
            item["id"]
            for item in conn.execute(
                """
                SELECT id
                FROM allocations
                WHERE resource_id = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                ORDER BY id ASC
                """,
                (row["resource_id"], row["period_key"]),
            ).fetchall()
        ]

        stats["extra_over_2_removed"] += max(0, len(before_ids) - len(after_ids))

    resource_period_all = conn.execute(
        """
        SELECT
            a.resource_id,
            COALESCE(NULLIF(a.period_key, 0), 2600 + a.week) AS period_key
        FROM allocations a
        JOIN resources r ON r.id = a.resource_id
        WHERE UPPER(r.name) NOT LIKE '%-EXT%'
          AND UPPER(r.name) NOT LIKE '% EXT'
          AND UPPER(r.role) NOT LIKE '%-EXT%'
          AND UPPER(r.role) NOT LIKE '% EXT'
        GROUP BY a.resource_id, COALESCE(NULLIF(a.period_key, 0), 2600 + a.week)
        """
    ).fetchall()

    for row in resource_period_all:
        rebalance_resource_week(conn, row["resource_id"], week_from_period_key(row["period_key"]))
        stats["weeks_rebalanced"] += 1

    return stats


@app.post("/api/admin/cleanup-allocations")
def admin_cleanup_allocations():
    conn = get_connection()
    try:
        stats = cleanup_allocations(conn)
        conn.commit()
        return {
            "ok": True,
            "stats": stats,
        }
    finally:
        conn.close()


@app.get("/api/resources")
def list_resources():
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT id, name, role, availability_note, is_active
            FROM resources
            ORDER BY name ASC
            """
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()


@app.post("/api/resources")
def create_resource(payload: ResourceCreate):
    conn = get_connection()
    try:
        cursor = conn.execute(
            """
            INSERT INTO resources (name, role, availability_note, is_active)
            VALUES (?, ?, ?, ?)
            """,
            (
                clean_text(payload.name),
                normalize_role(payload.role),
                clean_text(payload.availability_note),
                int(payload.is_active),
            )
        )
        conn.commit()

        row = conn.execute(
            """
            SELECT id, name, role, availability_note, is_active
            FROM resources
            WHERE id = ?
            """,
            (cursor.lastrowid,)
        ).fetchone()

        return dict(row)
    finally:
        conn.close()


@app.get("/api/projects")
def list_projects():
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT id, name, client, start_date, end_date, status, note
            FROM projects
            ORDER BY name ASC
            """
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()


@app.post("/api/projects")
def create_project(payload: ProjectCreate):
    conn = get_connection()
    try:
        cursor = conn.execute(
            """
            INSERT INTO projects (name, client, start_date, end_date, status, note)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                clean_text(payload.name),
                clean_text(payload.client),
                clean_text(payload.start_date),
                clean_text(payload.end_date),
                clean_text(payload.status),
                clean_text(payload.note),
            )
        )
        conn.commit()

        row = conn.execute(
            """
            SELECT id, name, client, start_date, end_date, status, note
            FROM projects
            WHERE id = ?
            """,
            (cursor.lastrowid,)
        ).fetchone()

        return dict(row)
    finally:
        conn.close()


@app.get("/api/demands")
def list_demands():
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT
                d.id,
                d.project_id,
                p.name AS project_name,
                d.week,
                COALESCE(NULLIF(d.period_key, 0), 2600 + d.week) AS period_key,
                d.role,
                d.quantity,
                d.note
            FROM demands d
            JOIN projects p ON p.id = d.project_id
            ORDER BY p.name ASC, d.role ASC, period_key ASC, d.week ASC
            """
        ).fetchall()

        result = []
        for row in rows:
            item = dict(row)
            item["period_key"] = int(item.get("period_key") or period_key_from_week(item.get("week")))
            item["week"] = int(item.get("week") or week_from_period_key(item["period_key"]))
            result.append(item)

        return result
    finally:
        conn.close()


@app.post("/api/demands")
def create_demand(payload: DemandCreate):
    conn = get_connection()
    try:
        role = normalize_role(payload.role)
        period_key = period_key_from_week(payload.week)
        week = week_from_period_key(period_key)

        existing = conn.execute(
            """
            SELECT id, quantity
            FROM demands
            WHERE project_id = ?
              AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
              AND UPPER(role) = ?
            """,
            (payload.project_id, period_key, role),
        ).fetchone()

        if existing:
            old_quantity = float(existing["quantity"] or 0)
            new_quantity = float(payload.quantity or 0)

            conn.execute(
                """
                UPDATE demands
                SET quantity = ?, note = ?, role = ?, week = ?, period_key = ?
                WHERE id = ?
                """,
                (
                    new_quantity,
                    clean_text(payload.note),
                    role,
                    week,
                    period_key,
                    existing["id"],
                ),
            )

            if old_quantity != new_quantity:
                conn.execute(
                    """
                    INSERT INTO demand_history
                    (project_id, week, period_key, role, old_quantity, new_quantity, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        payload.project_id,
                        week,
                        period_key,
                        role,
                        old_quantity,
                        new_quantity,
                        "Modifica fabbisogno",
                    ),
                )

            demand_id = existing["id"]
        else:
            cursor = conn.execute(
                """
                INSERT INTO demands (project_id, week, period_key, role, quantity, note)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    payload.project_id,
                    week,
                    period_key,
                    role,
                    payload.quantity,
                    clean_text(payload.note),
                )
            )

            conn.execute(
                """
                INSERT INTO demand_history
                (project_id, week, period_key, role, old_quantity, new_quantity, note)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    payload.project_id,
                    week,
                    period_key,
                    role,
                    0,
                    payload.quantity,
                    "Creazione fabbisogno",
                ),
            )

            demand_id = cursor.lastrowid

        release_allocations_for_zero_demand(conn, payload.project_id, role, week)
        conn.commit()

        row = conn.execute(
            """
            SELECT
                d.id,
                d.project_id,
                p.name AS project_name,
                d.week,
                COALESCE(NULLIF(d.period_key, 0), 2600 + d.week) AS period_key,
                d.role,
                d.quantity,
                d.note
            FROM demands d
            JOIN projects p ON p.id = d.project_id
            WHERE d.id = ?
            """,
            (demand_id,)
        ).fetchone()

        return dict(row)
    finally:
        conn.close()


@app.post("/api/demands/upsert-range")
def upsert_demand_range(payload: DemandRangeUpsert):
    conn = get_connection()
    updated_ids = []
    released_ids = []

    try:
        role = normalize_role(payload.role)
        note = clean_text(payload.note)
        week_from = min(payload.week_from, payload.week_to)
        week_to = max(payload.week_from, payload.week_to)

        for period_key in period_key_range_inclusive(week_from, week_to):
            week = week_from_period_key(period_key)

            existing = conn.execute(
                """
                SELECT id, quantity
                FROM demands
                WHERE project_id = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                  AND UPPER(role) = ?
                """,
                (payload.project_id, period_key, role),
            ).fetchone()

            if existing:
                old_quantity = float(existing["quantity"] or 0)
                new_quantity = float(payload.quantity or 0)

                conn.execute(
                    """
                    UPDATE demands
                    SET quantity = ?, note = ?, role = ?, week = ?, period_key = ?
                    WHERE id = ?
                    """,
                    (new_quantity, note, role, week, period_key, existing["id"]),
                )
                updated_ids.append(existing["id"])

                if old_quantity != new_quantity:
                    conn.execute(
                        """
                        INSERT INTO demand_history
                        (project_id, week, period_key, role, old_quantity, new_quantity, note)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            payload.project_id,
                            week,
                            period_key,
                            role,
                            old_quantity,
                            new_quantity,
                            "Modifica fabbisogno",
                        ),
                    )
            else:
                cursor = conn.execute(
                    """
                    INSERT INTO demands (project_id, week, period_key, role, quantity, note)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (payload.project_id, week, period_key, role, payload.quantity, note),
                )
                updated_ids.append(cursor.lastrowid)

                conn.execute(
                    """
                    INSERT INTO demand_history
                    (project_id, week, period_key, role, old_quantity, new_quantity, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        payload.project_id,
                        week,
                        period_key,
                        role,
                        0,
                        payload.quantity,
                        "Creazione fabbisogno",
                    ),
                )

            released_ids.extend(
                release_allocations_for_zero_demand(conn, payload.project_id, role, week)
            )

        conn.commit()

        rows = []
        if updated_ids:
            rows = conn.execute(
                f"""
                SELECT
                    d.id,
                    d.project_id,
                    p.name AS project_name,
                    d.week,
                    COALESCE(NULLIF(d.period_key, 0), 2600 + d.week) AS period_key,
                    d.role,
                    d.quantity,
                    d.note
                FROM demands d
                JOIN projects p ON p.id = d.project_id
                WHERE d.id IN ({",".join(["?"] * len(updated_ids))})
                ORDER BY period_key ASC, d.week ASC
                """,
                updated_ids,
            ).fetchall()

        return {
            "updated_count": len(updated_ids),
            "released_count": len(released_ids),
            "released_ids": released_ids,
            "rows": [dict(row) for row in rows],
        }
    finally:
        conn.close()


@app.get("/api/demand-history")
def list_demand_history():
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT
                h.id,
                h.project_id,
                p.name AS project_name,
                h.week,
                COALESCE(NULLIF(h.period_key, 0), 2600 + h.week) AS period_key,
                h.role,
                h.old_quantity,
                h.new_quantity,
                h.note,
                h.created_at
            FROM demand_history h
            JOIN projects p ON p.id = h.project_id
            ORDER BY h.created_at DESC, h.id DESC
            """
        ).fetchall()

        result = []
        for row in rows:
            item = dict(row)
            item["period_key"] = int(item.get("period_key") or period_key_from_week(item.get("week")))
            item["week"] = int(item.get("week") or week_from_period_key(item["period_key"]))
            result.append(item)

        return result
    finally:
        conn.close()


@app.get("/api/allocations")
def list_allocations():
    conn = get_connection()
    try:
        release_unavailable_allocations(conn)
        conn.commit()

        rows = conn.execute(
            """
            SELECT
                a.id,
                a.resource_id,
                r.name AS resource_name,
                r.role AS resource_role,
                r.availability_note AS resource_availability_note,
                r.is_active AS resource_is_active,
                a.project_id,
                p.name AS project_name,
                a.week,
                COALESCE(NULLIF(a.period_key, 0), 2600 + a.week) AS period_key,
                a.role,
                a.hours,
                a.load_percent,
                a.note
            FROM allocations a
            JOIN resources r ON r.id = a.resource_id
            JOIN projects p ON p.id = a.project_id
            ORDER BY p.name ASC, period_key ASC, a.week ASC, r.name ASC
            """
        ).fetchall()

        cleaned_rows = []
        for row in rows:
            item = dict(row)
            item["role"] = normalize_role(item.get("role"))
            item["resource_role"] = normalize_role(item.get("resource_role"))
            item["load_percent"] = float(item.get("load_percent") or 0)
            item["period_key"] = int(item.get("period_key") or period_key_from_week(item.get("week")))
            item["week"] = int(item.get("week") or week_from_period_key(item["period_key"]))
            cleaned_rows.append(item)

        return cleaned_rows
    finally:
        conn.close()


@app.get("/api/allocation-history")
def list_allocation_history():
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT
                id,
                resource_id,
                resource_name,
                resource_role,
                project_id,
                project_name,
                week,
                COALESCE(NULLIF(period_key, 0), 2600 + week) AS period_key,
                role,
                hours,
                load_percent,
                reason,
                note,
                created_at
            FROM allocation_history
            ORDER BY created_at DESC, id DESC
            """
        ).fetchall()

        cleaned_rows = []
        for row in rows:
            item = dict(row)
            item["role"] = normalize_role(item.get("role"))
            item["resource_role"] = normalize_role(item.get("resource_role"))
            item["load_percent"] = float(item.get("load_percent") or 0)
            item["period_key"] = int(item.get("period_key") or period_key_from_week(item.get("week")))
            item["week"] = int(item.get("week") or week_from_period_key(item["period_key"]))
            cleaned_rows.append(item)

        return cleaned_rows
    finally:
        conn.close()


@app.post("/api/allocations")
def create_allocation(payload: AllocationCreate):
    conn = get_connection()
    try:
        role = normalize_role(payload.role)

        if not role:
            resource = conn.execute(
                """
                SELECT role
                FROM resources
                WHERE id = ?
                """,
                (payload.resource_id,),
            ).fetchone()
            role = normalize_role(resource["role"]) if resource else ""

        if not demand_exists_for_project_role(conn, payload.project_id, role):
            return {
                "created": False,
                "reason": "demand_row_missing",
                "message": "Prima crea/attiva il fabbisogno per questa commessa e mansione.",
            }

        resource = get_resource_or_none(conn, payload.resource_id)
        is_external = resource and is_external_text(resource["name"], resource["role"])

        existing_week = get_resource_week_allocations(conn, payload.resource_id, payload.week)

        if len(existing_week) > 0 and not is_external:
            return {
                "created": False,
                "reason": "resource_already_allocated_this_week",
                "existing": existing_week,
            }

        new_id = insert_allocation(
            conn,
            payload.resource_id,
            payload.project_id,
            payload.week,
            role,
            payload.hours,
            100,
            payload.note,
        )

        if not is_external:
            rebalance_resource_week(conn, payload.resource_id, payload.week)

        conn.commit()

        return {
            "created": True,
            "created_id": new_id,
        }
    finally:
        conn.close()


@app.post("/api/allocations/assign-range")
def assign_allocation_range(payload: AllocationRangePayload):
    conn = get_connection()
    created_ids = []
    skipped = []
    conflicts = []

    try:
        role = normalize_role(payload.role)

        if not demand_exists_for_project_role(conn, payload.project_id, role):
            return {
                "created_count": 0,
                "created_ids": [],
                "skipped": ["demand_row_missing"],
                "conflicts": [],
                "message": "Prima crea/attiva il fabbisogno per questa commessa e mansione.",
            }

        resource = get_resource_or_none(conn, payload.resource_id)

        if not resource:
            return {
                "created_count": 0,
                "created_ids": [],
                "skipped": ["resource_not_found"],
                "conflicts": [],
            }

        if resource_is_unavailable(resource):
            return {
                "created_count": 0,
                "created_ids": [],
                "skipped": ["resource_unavailable"],
                "conflicts": [],
            }

        is_external = is_external_text(resource["name"], resource["role"])

        week_from = min(payload.week_from, payload.week_to)
        week_to = max(payload.week_from, payload.week_to)

        for period_key in period_key_range_inclusive(week_from, week_to):
            week = week_from_period_key(period_key)

            existing_same_cell = conn.execute(
                """
                SELECT id
                FROM allocations
                WHERE resource_id = ?
                  AND project_id = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                  AND UPPER(role) = ?
                """,
                (payload.resource_id, payload.project_id, period_key, role),
            ).fetchone()

            if existing_same_cell and not is_external:
                skipped.append({"week": week, "period_key": period_key, "reason": "already_assigned_on_this_cell"})
                rebalance_resource_week(conn, payload.resource_id, week)
                continue

            existing_week = get_resource_week_allocations(conn, payload.resource_id, week)

            if existing_week and not is_external:
                conflicts.append({
                    "week": week,
                    "period_key": period_key,
                    "existing": existing_week,
                })
                skipped.append({"week": week, "period_key": period_key, "reason": "resource_already_allocated_this_week"})
                continue

            new_id = insert_allocation(
                conn,
                payload.resource_id,
                payload.project_id,
                week,
                role,
                payload.hours,
                100,
                payload.note,
            )
            created_ids.append(new_id)

            if not is_external:
                rebalance_resource_week(conn, payload.resource_id, week)

        conn.commit()

        return {
            "created_count": len(created_ids),
            "created_ids": created_ids,
            "skipped": skipped,
            "conflicts": conflicts,
        }
    finally:
        conn.close()


@app.post("/api/allocations/resolve-conflict")
def resolve_allocation_conflict(payload: AllocationResolvePayload):
    conn = get_connection()
    try:
        role = normalize_role(payload.role)
        mode = clean_text(payload.mode).lower()
        period_key = period_key_from_week(payload.week)

        if not demand_exists_for_project_role(conn, payload.project_id, role):
            return {
                "ok": False,
                "reason": "demand_row_missing",
                "message": "Prima crea/attiva il fabbisogno per questa commessa e mansione.",
            }

        resource = get_resource_or_none(conn, payload.resource_id)

        if not resource:
            return {"ok": False, "reason": "resource_not_found"}

        if resource_is_unavailable(resource):
            return {"ok": False, "reason": "resource_unavailable"}

        is_external = is_external_text(resource["name"], resource["role"])

        if is_external:
            new_id = insert_allocation(
                conn,
                payload.resource_id,
                payload.project_id,
                payload.week,
                role,
                payload.hours,
                100,
                payload.note,
            )
            conn.commit()
            return {"ok": True, "mode": "external_direct", "created_id": new_id}

        existing_week = get_resource_week_allocations(conn, payload.resource_id, payload.week)

        already_same_cell = any(
            item["project_id"] == payload.project_id
            and normalize_role(item["role"]) == role
            and int(item.get("period_key") or period_key_from_week(item["week"])) == period_key
            for item in existing_week
        )

        if already_same_cell:
            rebalance_resource_week(conn, payload.resource_id, payload.week)
            conn.commit()
            return {"ok": True, "reason": "already_assigned_on_this_cell"}

        if mode == "direct":
            if existing_week:
                return {
                    "ok": False,
                    "reason": "conflict_exists",
                    "existing": existing_week,
                }

            new_id = insert_allocation(
                conn,
                payload.resource_id,
                payload.project_id,
                payload.week,
                role,
                payload.hours,
                100,
                payload.note,
            )

            rebalance_resource_week(conn, payload.resource_id, payload.week)
            conn.commit()
            return {"ok": True, "mode": mode, "created_id": new_id}

        if mode == "replace_all":
            ids_to_delete = [item["id"] for item in existing_week]

            for allocation_id in ids_to_delete:
                move_allocation_to_history(
                    conn,
                    allocation_id,
                    "sostituito",
                    "Allocazione sostituita da nuova assegnazione",
                )

            new_id = insert_allocation(
                conn,
                payload.resource_id,
                payload.project_id,
                payload.week,
                role,
                payload.hours,
                100,
                payload.note,
            )

            rebalance_resource_week(conn, payload.resource_id, payload.week)
            conn.commit()
            return {
                "ok": True,
                "mode": mode,
                "created_id": new_id,
                "removed_ids": ids_to_delete,
            }

        if mode == "replace_one":
            if payload.remove_allocation_id is None:
                return {"ok": False, "reason": "missing_remove_allocation_id"}

            move_allocation_to_history(
                conn,
                payload.remove_allocation_id,
                "sostituito",
                "Allocazione sostituita da nuova assegnazione",
            )

            new_id = insert_allocation(
                conn,
                payload.resource_id,
                payload.project_id,
                payload.week,
                role,
                payload.hours,
                100,
                payload.note,
            )

            rebalance_resource_week(conn, payload.resource_id, payload.week)
            conn.commit()
            return {
                "ok": True,
                "mode": mode,
                "created_id": new_id,
                "removed_id": payload.remove_allocation_id,
            }

        if mode == "split_50":
            if len(existing_week) != 1:
                return {
                    "ok": False,
                    "reason": "split_50_requires_one_existing",
                    "existing": existing_week,
                }

            new_id = insert_allocation(
                conn,
                payload.resource_id,
                payload.project_id,
                payload.week,
                role,
                payload.hours,
                100,
                payload.note,
            )

            rebalance_resource_week(conn, payload.resource_id, payload.week)
            conn.commit()
            return {
                "ok": True,
                "mode": mode,
                "created_id": new_id,
                "updated_id": existing_week[0]["id"],
            }

        return {
            "ok": False,
            "reason": "unknown_mode",
            "mode": mode,
        }
    finally:
        conn.close()


@app.post("/api/allocations/remove-range")
def remove_allocation_range(payload: AllocationRangePayload):
    conn = get_connection()
    removed_ids = []

    try:
        role = normalize_role(payload.role)
        week_from = min(payload.week_from, payload.week_to)
        week_to = max(payload.week_from, payload.week_to)
        touched_weeks = []

        period_from = period_key_from_week(week_from)
        period_to = period_key_from_week(week_to)

        rows = conn.execute(
            """
            SELECT id, week
            FROM allocations
            WHERE resource_id = ?
              AND project_id = ?
              AND UPPER(role) = ?
              AND COALESCE(NULLIF(period_key, 0), 2600 + week) BETWEEN ? AND ?
            """,
            (
                payload.resource_id,
                payload.project_id,
                role,
                period_from,
                period_to,
            ),
        ).fetchall()

        removed_ids = [row["id"] for row in rows]
        touched_weeks = sorted({int(row["week"]) for row in rows})

        for allocation_id in removed_ids:
            move_allocation_to_history(
                conn,
                allocation_id,
                "rimosso",
                "Rimozione manuale da planner V2",
            )

        for week in touched_weeks:
            rebalance_resource_week(conn, payload.resource_id, week)

        conn.commit()

        return {
            "removed_count": len(removed_ids),
            "removed_ids": removed_ids,
            "rebalanced_weeks": touched_weeks,
        }
    finally:
        conn.close()

@app.get("/api/workshop-required-map")
def workshop_required_map():
    conn = get_connection()
    try:
        rows = conn.execute("""
            SELECT
                overall_project_id,
                UPPER(role) AS role,
                period_key,
                week,
                SUM(COALESCE(required, 0)) AS required
            FROM workshop_rollup_sources
            GROUP BY overall_project_id, UPPER(role), period_key, week
            ORDER BY period_key, role
        """).fetchall()

        return [
            {
                "overall_project_id": row["overall_project_id"],
                "project_id": row["overall_project_id"],
                "role": row["role"],
                "period_key": row["period_key"],
                "week": row["week"],
                "quantity": row["required"],
                "required": row["required"],
                "note": "WORKSHOP_ROLLUP_SOURCES"
            }
            for row in rows
        ]
    finally:
        conn.close()

@app.post("/api/demands/upsert-range")
def upsert_demand_range(payload: dict):
    project_id = int(payload.get("project_id") or 0)
    role = str(payload.get("role") or "").strip().upper()
    quantity = float(payload.get("quantity") or 0)
    week_from = int(payload.get("week_from") or 0)
    week_to = int(payload.get("week_to") or week_from or 0)

    if not project_id:
        return {"ok": False, "error": "project_id obbligatorio"}
    if not role:
        return {"ok": False, "error": "role obbligatorio"}
    if week_from <= 0 or week_to <= 0:
        return {"ok": False, "error": "week_from/week_to obbligatori"}
    if week_to < week_from:
        week_from, week_to = week_to, week_from

    conn = get_connection()
    try:
        project = conn.execute(
            "SELECT id, name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()

        if not project:
            return {"ok": False, "error": "Commessa non trovata"}

        changed = 0

        for period_key in period_key_range_inclusive(week_from, week_to):
            week = week_from_period_key(period_key)

            existing = conn.execute(
                """
                SELECT id, quantity
                FROM demands
                WHERE project_id = ?
                  AND UPPER(role) = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                LIMIT 1
                """,
                (project_id, role, period_key),
            ).fetchone()

            if existing:
                old_quantity = float(existing["quantity"] or 0)
                if old_quantity != quantity:
                    conn.execute(
                        """
                        UPDATE demands
                        SET quantity = ?,
                            week = ?,
                            period_key = ?,
                            note = COALESCE(note, '') || ' | V2_RANGE_UPDATE'
                        WHERE id = ?
                        """,
                        (quantity, week, period_key, existing["id"]),
                    )

                    try:
                        conn.execute(
                            """
                            INSERT INTO demand_history(
                                project_id,
                                project_name,
                                role,
                                week,
                                period_key,
                                old_quantity,
                                new_quantity,
                                note
                            )
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                project_id,
                                project["name"],
                                role,
                                week,
                                period_key,
                                old_quantity,
                                quantity,
                                "V2_RANGE_UPDATE",
                            ),
                        )
                    except Exception:
                        pass

                    changed += 1
            else:
                conn.execute(
                    """
                    INSERT INTO demands(project_id, week, period_key, role, quantity, note)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (project_id, week, period_key, role, quantity, "V2_RANGE_INSERT"),
                )

                try:
                    conn.execute(
                        """
                        INSERT INTO demand_history(
                            project_id,
                            project_name,
                            role,
                            week,
                            period_key,
                            old_quantity,
                            new_quantity,
                            note
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            project_id,
                            project["name"],
                            role,
                            week,
                            period_key,
                            0,
                            quantity,
                            "V2_RANGE_INSERT",
                        ),
                    )
                except Exception:
                    pass

                changed += 1

        conn.commit()

        return {
            "ok": True,
            "project_id": project_id,
            "project_name": project["name"],
            "role": role,
            "quantity": quantity,
            "week_from": week_from,
            "week_to": week_to,
            "changed": changed,
        }
    finally:
        conn.close()


@app.post("/api/project-role-row")
def create_project_role_row(payload: dict):
    project_id = int(payload.get("project_id") or 0)
    role = str(payload.get("role") or "").strip().upper()
    quantity = float(payload.get("quantity") or 0)
    week_from = int(payload.get("week_from") or 0)
    week_to = int(payload.get("week_to") or week_from or 0)

    if not project_id:
        return {"ok": False, "error": "project_id obbligatorio"}
    if not role:
        return {"ok": False, "error": "role obbligatorio"}
    if quantity <= 0:
        return {"ok": False, "error": "Per creare una nuova riga mansione, il richiesto deve essere maggiore di 0"}
    if week_from <= 0:
        return {"ok": False, "error": "week_from obbligatoria"}
    if week_to <= 0:
        week_to = week_from
    if week_to < week_from:
        week_from, week_to = week_to, week_from

    # Riusa la stessa logica range: creare riga mansione significa creare demands anche a 0 o quantity scelta.
    return upsert_demand_range(
        {
            "project_id": project_id,
            "role": role,
            "quantity": quantity,
            "week_from": week_from,
            "week_to": week_to,
        }
    )


@app.get("/api/roles")
def list_roles():
    conn = get_connection()
    try:
        roles = set()

        try:
            rows = conn.execute("SELECT role FROM roles ORDER BY role").fetchall()
            for row in rows:
                value = str(row["role"] or "").strip().upper()
                if value:
                    roles.add(value)
        except Exception:
            pass

        rows = conn.execute("""
            SELECT role FROM resources
            UNION
            SELECT role FROM demands
            UNION
            SELECT role FROM allocations
        """).fetchall()

        for row in rows:
            value = str(row["role"] or "").strip().upper()
            if value:
                roles.add(value)

        return sorted(roles)
    finally:
        conn.close()

@app.get("/api/project-demand-matrix/{project_id}")
def project_demand_matrix(project_id: int):
    conn = get_connection()
    try:
        project = conn.execute(
            "SELECT id, name, status, note FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()

        if not project:
            return {"ok": False, "error": "Commessa non trovata"}

        rows = conn.execute(
            """
            SELECT
                role,
                week,
                COALESCE(NULLIF(period_key, 0), 2600 + week) AS period_key,
                quantity,
                note
            FROM demands
            WHERE project_id = ?
            ORDER BY role, period_key
            """,
            (project_id,),
        ).fetchall()

        by_role = {}

        for row in rows:
            role = str(row["role"] or "").strip().upper()
            if not role:
                continue

            if role not in by_role:
                by_role[role] = {
                    "role": role,
                    "weeks": {},
                    "total": 0,
                }

            period_key = int(row["period_key"] or 0)
            week = int(row["week"] or (period_key % 100))
            quantity = float(row["quantity"] or 0)

            by_role[role]["weeks"][str(period_key)] = {
                "week": week,
                "period_key": period_key,
                "quantity": quantity,
                "note": row["note"] or "",
            }
            by_role[role]["total"] += quantity

        return {
            "ok": True,
            "project": dict(project),
            "roles": list(by_role.values()),
        }
    finally:
        conn.close()


@app.post("/api/project-demand-matrix/{project_id}")
def save_project_demand_matrix(project_id: int, payload: dict):
    role = str(payload.get("role") or "").strip().upper()
    quantities = payload.get("quantities") or {}

    if not role:
        return {"ok": False, "error": "role obbligatorio"}

    if not isinstance(quantities, dict):
        return {"ok": False, "error": "quantities non valido"}

    conn = get_connection()
    try:
        project = conn.execute(
            "SELECT id, name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()

        if not project:
            return {"ok": False, "error": "Commessa non trovata"}

        changed = 0

        for raw_period_key, raw_quantity in quantities.items():
            period_key = int(raw_period_key)
            week = period_key % 100
            quantity = float(raw_quantity or 0)

            existing = conn.execute(
                """
                SELECT id, quantity
                FROM demands
                WHERE project_id = ?
                  AND UPPER(role) = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                LIMIT 1
                """,
                (project_id, role, period_key),
            ).fetchone()

            if existing:
                old_quantity = float(existing["quantity"] or 0)

                if old_quantity != quantity:
                    conn.execute(
                        """
                        UPDATE demands
                        SET quantity = ?,
                            week = ?,
                            period_key = ?,
                            note = COALESCE(note, '') || ' | PROJECT_MATRIX_UPDATE'
                        WHERE id = ?
                        """,
                        (quantity, week, period_key, existing["id"]),
                    )

                    try:
                        conn.execute(
                            """
                            INSERT INTO demand_history(
                                project_id,
                                project_name,
                                role,
                                week,
                                period_key,
                                old_quantity,
                                new_quantity,
                                note
                            )
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                project_id,
                                project["name"],
                                role,
                                week,
                                period_key,
                                old_quantity,
                                quantity,
                                "PROJECT_MATRIX_UPDATE",
                            ),
                        )
                    except Exception:
                        pass

                    changed += 1
            else:
                if quantity <= 0:
                    continue

                conn.execute(
                    """
                    INSERT INTO demands(project_id, week, period_key, role, quantity, note)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (project_id, week, period_key, role, quantity, "PROJECT_MATRIX_INSERT"),
                )

                try:
                    conn.execute(
                        """
                        INSERT INTO demand_history(
                            project_id,
                            project_name,
                            role,
                            week,
                            period_key,
                            old_quantity,
                            new_quantity,
                            note
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            project_id,
                            project["name"],
                            role,
                            week,
                            period_key,
                            0,
                            quantity,
                            "PROJECT_MATRIX_INSERT",
                        ),
                    )
                except Exception:
                    pass

                changed += 1

        conn.commit()
        return {"ok": True, "changed": changed}
    finally:
        conn.close()

from datetime import datetime

def _ensure_project_baseline_schema(conn):
    cols = conn.execute("PRAGMA table_info(projects)").fetchall()
    names = {row["name"] if hasattr(row, "keys") else row[1] for row in cols}

    if "baseline_at" not in names:
        conn.execute("ALTER TABLE projects ADD COLUMN baseline_at TEXT")

    # Le commesse gia' presenti derivano dall'import vero del 07/04.
    conn.execute(
        """
        UPDATE projects
        SET baseline_at = '2026-04-07T00:00:00'
        WHERE baseline_at IS NULL OR baseline_at = ''
        """
    )


@app.post("/api/projects/create-baseline")
def create_project_with_baseline(payload: dict):
    name = str(payload.get("name") or "").strip()
    status = str(payload.get("status") or "ACTIVE").strip().upper()
    note = str(payload.get("note") or "").strip()
    rows = payload.get("rows") or []

    if not name:
        return {"ok": False, "error": "Nome/codice commessa obbligatorio"}

    if not isinstance(rows, list):
        return {"ok": False, "error": "rows non valido"}

    baseline_at = datetime.now().isoformat(timespec="seconds")

    conn = get_connection()
    try:
        _ensure_project_baseline_schema(conn)

        existing = conn.execute(
            "SELECT id, name FROM projects WHERE UPPER(name) = UPPER(?) LIMIT 1",
            (name,),
        ).fetchone()

        if existing:
            return {"ok": False, "error": "Commessa gia' esistente"}

        cur = conn.execute(
            """
            INSERT INTO projects(name, status, note, baseline_at)
            VALUES (?, ?, ?, ?)
            """,
            (
                name,
                status or "ACTIVE",
                (note + " | BASELINE_CREATE").strip(" |"),
                baseline_at,
            ),
        )

        project_id = int(cur.lastrowid)
        inserted = 0

        for row in rows:
            role = str(row.get("role") or "").strip().upper()
            quantities = row.get("quantities") or {}

            if not role or not isinstance(quantities, dict):
                continue

            for raw_period_key, raw_quantity in quantities.items():
                quantity = float(raw_quantity or 0)

                if quantity <= 0:
                    continue

                period_key = int(raw_period_key)
                week = period_key % 100

                conn.execute(
                    """
                    INSERT INTO demands(project_id, week, period_key, role, quantity, note)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        project_id,
                        week,
                        period_key,
                        role,
                        quantity,
                        "BASELINE_CREATE",
                    ),
                )
                inserted += 1

        conn.commit()

        return {
            "ok": True,
            "project_id": project_id,
            "baseline_at": baseline_at,
            "inserted": inserted,
        }
    finally:
        conn.close()


@app.post("/api/projects/ensure-baseline")
def ensure_projects_baseline():
    conn = get_connection()
    try:
        _ensure_project_baseline_schema(conn)
        conn.commit()
        return {"ok": True}
    finally:
        conn.close()

@app.get("/api/workshop-breakdown-matrix")
def workshop_breakdown_matrix():
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT
                overall_project_id,
                source_old_project_id,
                source_project_name,
                role,
                week,
                period_key,
                required
            FROM workshop_rollup_sources
            WHERE COALESCE(required, 0) <> 0
            ORDER BY role, source_project_name, period_key
            """
        ).fetchall()

        return {
            "ok": True,
            "rows": [
                {
                    "overall_project_id": row["overall_project_id"],
                    "source_old_project_id": row["source_old_project_id"],
                    "source_project_name": row["source_project_name"],
                    "project_name": row["source_project_name"],
                    "role": row["role"],
                    "week": row["week"],
                    "period_key": row["period_key"],
                    "required": row["required"],
                }
                for row in rows
            ],
        }
    finally:
        conn.close()

# === PROJECTS SHEET V2 BACKEND START ===

from datetime import datetime as _projects_sheet_datetime

def _projects_sheet_ensure_schema(conn):
    cols = conn.execute("PRAGMA table_info(projects)").fetchall()
    names = {row["name"] for row in cols}

    if "baseline_at" not in names:
        conn.execute("ALTER TABLE projects ADD COLUMN baseline_at TEXT")

    conn.execute(
        """
        UPDATE projects
        SET baseline_at = '2026-04-07T00:00:00'
        WHERE baseline_at IS NULL OR baseline_at = ''
        """
    )


def _projects_sheet_period_key(value):
    value = int(value or 0)
    if value >= 1000:
        return value
    return 2600 + value


def _projects_sheet_week_from_period(period_key):
    return int(period_key or 0) % 100


def _projects_sheet_norm(value):
    return str(value or "").strip().upper()


def _projects_sheet_is_overall(project):
    if not project:
        return False
    name = _projects_sheet_norm(project["name"])
    note = _projects_sheet_norm(project["note"])
    return "OVERALL OFFICINA" in name or "OVERALL" in note or "WORKSHOP_ROLLUP" in note


def _projects_sheet_insert_history(conn, project, role, week, period_key, old_quantity, new_quantity, note):
    try:
        conn.execute(
            """
            INSERT INTO demand_history(
                project_id,
                project_name,
                role,
                week,
                period_key,
                old_quantity,
                new_quantity,
                note
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(project["id"]),
                project["name"],
                _projects_sheet_norm(role),
                int(week),
                int(period_key),
                float(old_quantity or 0),
                float(new_quantity or 0),
                note,
            ),
        )
        return
    except Exception:
        pass

    try:
        conn.execute(
            """
            INSERT INTO demand_history(
                project_id,
                role,
                week,
                period_key,
                old_quantity,
                new_quantity,
                note
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(project["id"]),
                _projects_sheet_norm(role),
                int(week),
                int(period_key),
                float(old_quantity or 0),
                float(new_quantity or 0),
                note,
            ),
        )
    except Exception:
        pass


@app.get("/api/projects-sheet/projects")
def projects_sheet_projects():
    conn = get_connection()
    try:
        _projects_sheet_ensure_schema(conn)

        rows = conn.execute(
            """
            SELECT
                id,
                name,
                client,
                start_date,
                end_date,
                status,
                note,
                baseline_at
            FROM projects
            ORDER BY name ASC
            """
        ).fetchall()

        conn.commit()

        result = []
        for row in rows:
            note = _projects_sheet_norm(row["note"])
            name = _projects_sheet_norm(row["name"])

            result.append({
                "id": row["id"],
                "name": row["name"],
                "client": row["client"],
                "start_date": row["start_date"],
                "end_date": row["end_date"],
                "status": row["status"],
                "note": row["note"],
                "baseline_at": row["baseline_at"],
                "is_overall": "OVERALL OFFICINA" in name or "OVERALL" in note or "WORKSHOP_ROLLUP" in note,
                "is_workshop_rollup": "WORKSHOP_ROLLUP" in note,
            })

        return {
            "ok": True,
            "projects": result,
        }
    finally:
        conn.close()


@app.get("/api/projects-sheet/matrix/{project_id}")
def projects_sheet_matrix(project_id: int):
    conn = get_connection()
    try:
        _projects_sheet_ensure_schema(conn)

        project = conn.execute(
            """
            SELECT
                id,
                name,
                client,
                start_date,
                end_date,
                status,
                note,
                baseline_at
            FROM projects
            WHERE id = ?
            """,
            (project_id,),
        ).fetchone()

        if not project:
            return {"ok": False, "error": "Commessa non trovata"}

        if _projects_sheet_is_overall(project):
            rollup_rows = conn.execute(
                """
                SELECT
                    overall_project_id,
                    source_old_project_id,
                    source_project_name,
                    role,
                    week,
                    period_key,
                    required
                FROM workshop_rollup_sources
                WHERE COALESCE(required, 0) <> 0
                ORDER BY role, source_project_name, period_key
                """
            ).fetchall()

            return {
                "ok": True,
                "project": dict(project),
                "is_overall": True,
                "roles": [],
                "workshop_rows": [
                    {
                        "overall_project_id": row["overall_project_id"],
                        "source_old_project_id": row["source_old_project_id"],
                        "source_project_name": row["source_project_name"],
                        "project_name": row["source_project_name"],
                        "role": row["role"],
                        "week": row["week"],
                        "period_key": row["period_key"],
                        "required": row["required"],
                    }
                    for row in rollup_rows
                ],
            }

        rows = conn.execute(
            """
            SELECT
                role,
                week,
                COALESCE(NULLIF(period_key, 0), 2600 + week) AS period_key,
                quantity,
                note
            FROM demands
            WHERE project_id = ?
            ORDER BY role, period_key
            """,
            (project_id,),
        ).fetchall()

        by_role = {}
        for row in rows:
            role = _projects_sheet_norm(row["role"])
            if not role:
                continue

            if role not in by_role:
                by_role[role] = {
                    "role": role,
                    "weeks": {},
                    "total": 0,
                }

            period_key = int(row["period_key"] or 0)
            week = int(row["week"] or _projects_sheet_week_from_period(period_key))
            quantity = float(row["quantity"] or 0)

            by_role[role]["weeks"][str(period_key)] = {
                "week": week,
                "period_key": period_key,
                "quantity": quantity,
                "note": row["note"] or "",
            }
            by_role[role]["total"] += quantity

        return {
            "ok": True,
            "project": dict(project),
            "is_overall": False,
            "roles": list(by_role.values()),
            "workshop_rows": [],
        }
    finally:
        conn.close()


@app.post("/api/projects-sheet/save-project")
def projects_sheet_save_project(payload: dict):
    project_id = int(payload.get("id") or 0)
    name = str(payload.get("name") or "").strip()
    client = str(payload.get("client") or "").strip()
    start_date = str(payload.get("start_date") or "").strip()
    end_date = str(payload.get("end_date") or "").strip()
    status = str(payload.get("status") or "attivo").strip()
    note = str(payload.get("note") or "").strip()
    is_workshop_rollup = bool(payload.get("is_workshop_rollup"))

    if not name:
        return {"ok": False, "error": "Nome/codice commessa obbligatorio"}

    conn = get_connection()
    try:
        _projects_sheet_ensure_schema(conn)

        normalized_note = note
        if is_workshop_rollup and "WORKSHOP_ROLLUP" not in _projects_sheet_norm(normalized_note):
            normalized_note = (normalized_note + " | WORKSHOP_ROLLUP").strip(" |")
        if "OVERALL OFFICINA" in _projects_sheet_norm(name) and "OVERALL" not in _projects_sheet_norm(normalized_note):
            normalized_note = (normalized_note + " | OVERALL | WORKSHOP_ROLLUP").strip(" |")

        if project_id:
            existing = conn.execute(
                "SELECT id FROM projects WHERE id = ?",
                (project_id,),
            ).fetchone()
            if not existing:
                return {"ok": False, "error": "Commessa non trovata"}

            conn.execute(
                """
                UPDATE projects
                SET name = ?,
                    client = ?,
                    start_date = ?,
                    end_date = ?,
                    status = ?,
                    note = ?,
                    baseline_at = COALESCE(NULLIF(baseline_at, ''), '2026-04-07T00:00:00')
                WHERE id = ?
                """,
                (
                    name,
                    client,
                    start_date,
                    end_date,
                    status,
                    normalized_note,
                    project_id,
                ),
            )
        else:
            duplicate = conn.execute(
                "SELECT id FROM projects WHERE UPPER(name) = UPPER(?) LIMIT 1",
                (name,),
            ).fetchone()
            if duplicate:
                return {"ok": False, "error": "Esiste già una commessa con questo nome"}

            cur = conn.execute(
                """
                INSERT INTO projects(
                    name,
                    client,
                    start_date,
                    end_date,
                    status,
                    note,
                    baseline_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    name,
                    client,
                    start_date,
                    end_date,
                    status,
                    (normalized_note + " | BASELINE_CREATE").strip(" |"),
                    _projects_sheet_datetime.now().isoformat(timespec="seconds"),
                ),
            )
            project_id = int(cur.lastrowid)

        conn.commit()

        row = conn.execute(
            """
            SELECT
                id,
                name,
                client,
                start_date,
                end_date,
                status,
                note,
                baseline_at
            FROM projects
            WHERE id = ?
            """,
            (project_id,),
        ).fetchone()

        return {
            "ok": True,
            "project": dict(row),
        }
    finally:
        conn.close()


@app.post("/api/projects-sheet/save-demands")
def projects_sheet_save_demands(payload: dict):
    project_id = int(payload.get("project_id") or 0)
    rows = payload.get("rows") or []
    baseline_create = bool(payload.get("baseline_create"))

    if not project_id:
        return {"ok": False, "error": "project_id obbligatorio"}

    if not isinstance(rows, list):
        return {"ok": False, "error": "rows non valido"}

    conn = get_connection()
    try:
        _projects_sheet_ensure_schema(conn)

        project = conn.execute(
            """
            SELECT
                id,
                name,
                note,
                baseline_at
            FROM projects
            WHERE id = ?
            """,
            (project_id,),
        ).fetchone()

        if not project:
            return {"ok": False, "error": "Commessa non trovata"}

        if _projects_sheet_is_overall(project):
            return {"ok": False, "error": "OVERALL è un rollup: modifica le sottocommesse officina"}

        changed = 0
        inserted = 0
        history = 0

        for row in rows:
            role = _projects_sheet_norm(row.get("role") or "")
            quantities = row.get("quantities") or {}

            if not role or not isinstance(quantities, dict):
                continue

            for raw_period_key, raw_quantity in quantities.items():
                period_key = _projects_sheet_period_key(raw_period_key)
                week = _projects_sheet_week_from_period(period_key)
                quantity = float(raw_quantity or 0)

                existing = conn.execute(
                    """
                    SELECT id, quantity
                    FROM demands
                    WHERE project_id = ?
                      AND UPPER(role) = ?
                      AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                    LIMIT 1
                    """,
                    (project_id, role, period_key),
                ).fetchone()

                if existing:
                    old_quantity = float(existing["quantity"] or 0)
                    if old_quantity == quantity:
                        continue

                    conn.execute(
                        """
                        UPDATE demands
                        SET quantity = ?,
                            week = ?,
                            period_key = ?,
                            role = ?,
                            note = ?
                        WHERE id = ?
                        """,
                        (
                            quantity,
                            week,
                            period_key,
                            role,
                            "PROJECT_SHEET_BASELINE" if baseline_create else "PROJECT_SHEET_UPDATE",
                            existing["id"],
                        ),
                    )
                    changed += 1

                    if not baseline_create:
                        _projects_sheet_insert_history(
                            conn,
                            project,
                            role,
                            week,
                            period_key,
                            old_quantity,
                            quantity,
                            "PROJECT_SHEET_UPDATE",
                        )
                        history += 1

                    try:
                        release_allocations_for_zero_demand(conn, project_id, role, week)
                    except Exception:
                        pass

                else:
                    if quantity <= 0:
                        continue

                    cur = conn.execute(
                        """
                        INSERT INTO demands(project_id, week, period_key, role, quantity, note)
                        VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (
                            project_id,
                            week,
                            period_key,
                            role,
                            quantity,
                            "PROJECT_SHEET_BASELINE" if baseline_create else "PROJECT_SHEET_INSERT",
                        ),
                    )
                    inserted += 1

                    if not baseline_create:
                        _projects_sheet_insert_history(
                            conn,
                            project,
                            role,
                            week,
                            period_key,
                            0,
                            quantity,
                            "PROJECT_SHEET_INSERT",
                        )
                        history += 1

        conn.commit()

        return {
            "ok": True,
            "changed": changed,
            "inserted": inserted,
            "history": history,
        }
    finally:
        conn.close()


@app.get("/api/projects-sheet/test-changes")
def projects_sheet_test_changes():
    conn = get_connection()
    try:
        demand_rows = conn.execute(
            """
            SELECT
                d.id,
                p.name AS project_name,
                d.role,
                d.week,
                COALESCE(NULLIF(d.period_key, 0), 2600 + d.week) AS period_key,
                d.quantity,
                d.note
            FROM demands d
            JOIN projects p ON p.id = d.project_id
            WHERE COALESCE(d.note, '') LIKE '%PROJECT_SHEET%'
               OR COALESCE(d.note, '') LIKE '%PROJECT_MATRIX%'
               OR COALESCE(d.note, '') LIKE '%V2_RANGE%'
               OR COALESCE(d.note, '') LIKE '%PLANNER V2%'
               OR COALESCE(d.note, '') LIKE '%BASELINE_CREATE%'
            ORDER BY d.id DESC
            """
        ).fetchall()

        history_rows = []
        try:
            history_rows = conn.execute(
                """
                SELECT *
                FROM demand_history
                WHERE COALESCE(note, '') LIKE '%PROJECT_SHEET%'
                   OR COALESCE(note, '') LIKE '%PROJECT_MATRIX%'
                   OR COALESCE(note, '') LIKE '%V2_RANGE%'
                   OR COALESCE(note, '') LIKE '%PLANNER V2%'
                ORDER BY id DESC
                """
            ).fetchall()
        except Exception:
            history_rows = []

        return {
            "ok": True,
            "demands": [dict(row) for row in demand_rows],
            "history": [dict(row) for row in history_rows],
        }
    finally:
        conn.close()

# === PROJECTS SHEET V2 BACKEND END ===

# === PERIOD KEY UNIFIED BACKEND START ===

def period_key_range_inclusive(start_value, end_value, default_year_short: int = DEFAULT_YEAR_SHORT):
    """
    Restituisce period_key validi tra start e end.
    Accetta sia week 17 sia period_key 2617.
    Gestisce anche passaggio anno: 2652 -> 2701.
    """
    start_key = period_key_from_week(start_value, default_year_short)
    end_key = period_key_from_week(end_value, default_year_short)

    if start_key <= 0 or end_key <= 0:
        return []

    if end_key < start_key:
        start_key, end_key = end_key, start_key

    start_year = start_key // 100
    start_week = start_key % 100
    end_year = end_key // 100
    end_week = end_key % 100

    keys = []
    year = start_year
    week = start_week

    guard = 0

    while True:
        guard += 1

        if guard > 520:
            break

        if 1 <= week <= 52:
            keys.append(year * 100 + week)

        if year == end_year and week == end_week:
            break

        week += 1

        if week > 52:
            year += 1
            week = 1

    return keys


def period_key_payload_range(payload):
    """
    Utility per endpoint futuri:
    preferisce period_key_from/to, ma accetta week_from/to.
    """
    if hasattr(payload, "dict"):
        data = payload.dict()
    elif isinstance(payload, dict):
        data = payload
    else:
        data = {}

    start = (
        data.get("period_key_from")
        or data.get("periodKeyFrom")
        or data.get("week_from")
        or data.get("weekFrom")
        or data.get("from_week")
        or data.get("fromWeek")
        or 0
    )

    end = (
        data.get("period_key_to")
        or data.get("periodKeyTo")
        or data.get("week_to")
        or data.get("weekTo")
        or data.get("to_week")
        or data.get("toWeek")
        or start
    )

    return period_key_range_inclusive(start, end)

# === PERIOD KEY UNIFIED BACKEND END ===

# === GANTT V2 OLD STYLE BACKEND START ===

def _gantt_to_period_key(value, default_year_short: int = DEFAULT_YEAR_SHORT):
    try:
        value = int(float(value or 0))
    except Exception:
        return 0

    if value <= 0:
        return 0

    if value >= 1000:
        return value

    return int(default_year_short) * 100 + value


def _gantt_period_key_range(start_value, end_value):
    start_key = _gantt_to_period_key(start_value)
    end_key = _gantt_to_period_key(end_value or start_value)

    if start_key <= 0 or end_key <= 0:
        return []

    if end_key < start_key:
        start_key, end_key = end_key, start_key

    start_year = start_key // 100
    start_week = start_key % 100
    end_year = end_key // 100
    end_week = end_key % 100

    result = []
    year = start_year
    week = start_week
    guard = 0

    while True:
        guard += 1
        if guard > 520:
            break

        if 1 <= week <= 52:
            result.append(year * 100 + week)

        if year == end_year and week == end_week:
            break

        week += 1
        if week > 52:
            year += 1
            week = 1

    return result


def _gantt_payload_period_range(payload: dict):
    start_value = (
        payload.get("period_key_from")
        or payload.get("periodKeyFrom")
        or payload.get("week_from")
        or payload.get("weekFrom")
        or payload.get("from")
        or payload.get("period_from")
        or 0
    )

    end_value = (
        payload.get("period_key_to")
        or payload.get("periodKeyTo")
        or payload.get("week_to")
        or payload.get("weekTo")
        or payload.get("to")
        or payload.get("period_to")
        or start_value
    )

    return _gantt_period_key_range(start_value, end_value)


@app.post("/api/gantt/assign-range")
def gantt_assign_range(payload: dict):
    resource_id = int(payload.get("resource_id") or 0)
    project_id = int(payload.get("project_id") or 0)
    role = normalize_role(payload.get("role") or "")
    hours = float(payload.get("hours") or 40)
    note = clean_text(payload.get("note") or "GANTT_V2_ASSIGN")

    if not resource_id:
        return {"ok": False, "error": "resource_id obbligatorio"}

    if not project_id:
        return {"ok": False, "error": "project_id obbligatorio"}

    if not role:
        return {"ok": False, "error": "role obbligatoria"}

    period_keys = _gantt_payload_period_range(payload)

    if not period_keys:
        return {"ok": False, "error": "periodo obbligatorio"}

    conn = get_connection()
    try:
        resource = get_resource_or_none(conn, resource_id)
        if not resource:
            return {"ok": False, "error": "Risorsa non trovata"}

        project = conn.execute(
            "SELECT id, name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()

        if not project:
            return {"ok": False, "error": "Commessa non trovata"}

        inserted = 0
        skipped = 0
        rebalanced = []

        for period_key in period_keys:
            week = week_from_period_key(period_key)

            existing = conn.execute(
                """
                SELECT id
                FROM allocations
                WHERE resource_id = ?
                  AND project_id = ?
                  AND UPPER(role) = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                LIMIT 1
                """,
                (resource_id, project_id, role, period_key),
            ).fetchone()

            if existing:
                skipped += 1
                continue

            conn.execute(
                """
                INSERT INTO allocations(
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
                    hours,
                    100,
                    note,
                ),
            )

            inserted += 1
            rebalanced.append(rebalance_resource_week(conn, resource_id, week))

        conn.commit()

        return {
            "ok": True,
            "inserted": inserted,
            "skipped": skipped,
            "rebalanced": rebalanced,
        }
    finally:
        conn.close()


@app.post("/api/gantt/unassign-range")
def gantt_unassign_range(payload: dict):
    resource_id = int(payload.get("resource_id") or 0)
    project_id = int(payload.get("project_id") or 0)
    role = normalize_role(payload.get("role") or "")
    reason = clean_text(payload.get("reason") or "GANTT_V2_UNASSIGN")
    note = clean_text(payload.get("note") or "Rimozione da Gantt V2")
    period_keys = _gantt_payload_period_range(payload)

    if not resource_id:
        return {"ok": False, "error": "resource_id obbligatorio"}

    if not period_keys:
        return {"ok": False, "error": "periodo obbligatorio"}

    conn = get_connection()
    try:
        removed = 0
        allocation_ids = []

        for period_key in period_keys:
            week = week_from_period_key(period_key)

            params = [resource_id, period_key]
            where = """
                resource_id = ?
                AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
            """

            if project_id:
                where += " AND project_id = ?"
                params.append(project_id)

            if role:
                where += " AND UPPER(role) = ?"
                params.append(role)

            rows = conn.execute(
                f"""
                SELECT id
                FROM allocations
                WHERE {where}
                ORDER BY id ASC
                """,
                tuple(params),
            ).fetchall()

            for row in rows:
                ok = move_allocation_to_history(
                    conn,
                    int(row["id"]),
                    reason,
                    note,
                )
                if ok:
                    removed += 1
                    allocation_ids.append(int(row["id"]))

            rebalance_resource_week(conn, resource_id, week)

        conn.commit()

        return {
            "ok": True,
            "removed": removed,
            "allocation_ids": allocation_ids,
        }
    finally:
        conn.close()

# === GANTT V2 OLD STYLE BACKEND END ===

# === OLD WORKFLOW GANTT RESOURCES BACKEND START ===

def _ow_to_period_key(value, default_year_short: int = DEFAULT_YEAR_SHORT):
    try:
        value = int(float(value or 0))
    except Exception:
        return 0

    if value <= 0:
        return 0

    if value >= 1000:
        return value

    return int(default_year_short) * 100 + value


def _ow_week_from_period_key(value):
    try:
        value = int(float(value or 0))
    except Exception:
        return 0

    if value >= 1000:
        return value % 100

    return value


def _ow_period_range(start_value, end_value):
    start_key = _ow_to_period_key(start_value)
    end_key = _ow_to_period_key(end_value or start_value)

    if start_key <= 0 or end_key <= 0:
        return []

    if end_key < start_key:
        start_key, end_key = end_key, start_key

    start_year = start_key // 100
    start_week = start_key % 100
    end_year = end_key // 100
    end_week = end_key % 100

    keys = []
    year = start_year
    week = start_week
    guard = 0

    while True:
        guard += 1
        if guard > 520:
            break

        if 1 <= week <= 52:
            keys.append(year * 100 + week)

        if year == end_year and week == end_week:
            break

        week += 1
        if week > 52:
            year += 1
            week = 1

    return keys


def _ow_payload_range(payload: dict):
    start = (
        payload.get("period_key_from")
        or payload.get("periodKeyFrom")
        or payload.get("period_from")
        or payload.get("periodFrom")
        or payload.get("week_from")
        or payload.get("weekFrom")
        or payload.get("from")
        or 0
    )

    end = (
        payload.get("period_key_to")
        or payload.get("periodKeyTo")
        or payload.get("period_to")
        or payload.get("periodTo")
        or payload.get("week_to")
        or payload.get("weekTo")
        or payload.get("to")
        or start
    )

    return _ow_period_range(start, end)


@app.get("/api/old-workflow/gantt-state")
def old_workflow_gantt_state():
    conn = get_connection()
    try:
        resources = conn.execute(
            """
            SELECT id, name, role, availability_note, is_active
            FROM resources
            ORDER BY role, name
            """
        ).fetchall()

        projects = conn.execute(
            """
            SELECT id, name, client, start_date, end_date, status, note
            FROM projects
            ORDER BY name
            """
        ).fetchall()

        demands = conn.execute(
            """
            SELECT
                d.id,
                d.project_id,
                p.name AS project_name,
                d.week,
                COALESCE(NULLIF(d.period_key, 0), 2600 + d.week) AS period_key,
                d.role,
                d.quantity,
                d.note
            FROM demands d
            JOIN projects p ON p.id = d.project_id
            ORDER BY p.name, d.role, period_key
            """
        ).fetchall()

        allocations = conn.execute(
            """
            SELECT
                a.id,
                a.resource_id,
                r.name AS resource_name,
                r.role AS resource_role,
                r.availability_note AS resource_availability_note,
                r.is_active AS resource_is_active,
                a.project_id,
                p.name AS project_name,
                a.week,
                COALESCE(NULLIF(a.period_key, 0), 2600 + a.week) AS period_key,
                a.role,
                a.hours,
                a.load_percent,
                a.note
            FROM allocations a
            JOIN resources r ON r.id = a.resource_id
            JOIN projects p ON p.id = a.project_id
            ORDER BY r.name, period_key, p.name
            """
        ).fetchall()

        allocation_history = []
        try:
            allocation_history = conn.execute(
                """
                SELECT *
                FROM allocation_history
                ORDER BY id DESC
                LIMIT 1000
                """
            ).fetchall()
        except Exception:
            allocation_history = []

        demand_history = []
        try:
            demand_history = conn.execute(
                """
                SELECT *
                FROM demand_history
                ORDER BY id DESC
                LIMIT 1000
                """
            ).fetchall()
        except Exception:
            demand_history = []

        return {
            "ok": True,
            "resources": [dict(row) for row in resources],
            "projects": [dict(row) for row in projects],
            "demands": [dict(row) for row in demands],
            "allocations": [dict(row) for row in allocations],
            "allocation_history": [dict(row) for row in allocation_history],
            "demand_history": [dict(row) for row in demand_history],
        }
    finally:
        conn.close()


@app.post("/api/old-workflow/gantt-assign")
def old_workflow_gantt_assign(payload: dict):
    resource_id = int(payload.get("resource_id") or 0)
    project_id = int(payload.get("project_id") or 0)
    role = normalize_role(payload.get("role") or "")
    hours = float(payload.get("hours") or 40)
    note = clean_text(payload.get("note") or "GANTT_OLD_WORKFLOW_ASSIGN")
    period_keys = _ow_payload_range(payload)

    if not resource_id:
        return {"ok": False, "error": "resource_id obbligatorio"}

    if not project_id:
        return {"ok": False, "error": "project_id obbligatorio"}

    if not role:
        return {"ok": False, "error": "mansione obbligatoria"}

    if not period_keys:
        return {"ok": False, "error": "periodo obbligatorio"}

    conn = get_connection()
    try:
        resource = get_resource_or_none(conn, resource_id)
        if not resource:
            return {"ok": False, "error": "Risorsa non trovata"}

        project = conn.execute(
            "SELECT id, name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()
        if not project:
            return {"ok": False, "error": "Commessa non trovata"}

        inserted = 0
        skipped = 0
        conflicts = []

        for period_key in period_keys:
            week = week_from_period_key(period_key)

            duplicate = conn.execute(
                """
                SELECT id
                FROM allocations
                WHERE resource_id = ?
                  AND project_id = ?
                  AND UPPER(role) = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                LIMIT 1
                """,
                (resource_id, project_id, role, period_key),
            ).fetchone()

            if duplicate:
                skipped += 1
                continue

            existing_same_period = conn.execute(
                """
                SELECT
                    a.id,
                    p.name AS project_name,
                    a.role,
                    a.load_percent
                FROM allocations a
                JOIN projects p ON p.id = a.project_id
                WHERE a.resource_id = ?
                  AND COALESCE(NULLIF(a.period_key, 0), 2600 + a.week) = ?
                ORDER BY a.id
                """,
                (resource_id, period_key),
            ).fetchall()

            if len(existing_same_period) >= 2:
                conflicts.append({
                    "period_key": period_key,
                    "reason": "Risorsa già su 2 allocazioni",
                    "rows": [dict(row) for row in existing_same_period],
                })
                skipped += 1
                continue

            conn.execute(
                """
                INSERT INTO allocations(
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
                    hours,
                    100,
                    note,
                ),
            )

            inserted += 1
            rebalance_resource_week(conn, resource_id, week)

        conn.commit()

        return {
            "ok": True,
            "inserted": inserted,
            "skipped": skipped,
            "conflicts": conflicts,
        }
    finally:
        conn.close()


@app.post("/api/old-workflow/gantt-unassign")
def old_workflow_gantt_unassign(payload: dict):
    resource_id = int(payload.get("resource_id") or 0)
    project_id = int(payload.get("project_id") or 0)
    role = normalize_role(payload.get("role") or "")
    period_keys = _ow_payload_range(payload)
    reason = clean_text(payload.get("reason") or "GANTT_OLD_WORKFLOW_UNASSIGN")
    note = clean_text(payload.get("note") or "Svincolo da Gantt")

    if not resource_id:
        return {"ok": False, "error": "resource_id obbligatorio"}

    if not period_keys:
        return {"ok": False, "error": "periodo obbligatorio"}

    conn = get_connection()
    try:
        removed = 0
        ids = []

        for period_key in period_keys:
            week = week_from_period_key(period_key)

            params = [resource_id, period_key]
            where = """
                resource_id = ?
                AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
            """

            if project_id:
                where += " AND project_id = ?"
                params.append(project_id)

            if role:
                where += " AND UPPER(role) = ?"
                params.append(role)

            rows = conn.execute(
                f"""
                SELECT id
                FROM allocations
                WHERE {where}
                ORDER BY id
                """,
                tuple(params),
            ).fetchall()

            for row in rows:
                ok = move_allocation_to_history(conn, int(row["id"]), reason, note)
                if ok:
                    removed += 1
                    ids.append(int(row["id"]))

            rebalance_resource_week(conn, resource_id, week)

        conn.commit()

        return {
            "ok": True,
            "removed": removed,
            "allocation_ids": ids,
        }
    finally:
        conn.close()


@app.post("/api/old-workflow/gantt-demand-upsert")
def old_workflow_gantt_demand_upsert(payload: dict):
    project_id = int(payload.get("project_id") or 0)
    role = normalize_role(payload.get("role") or "")
    quantity = float(payload.get("quantity") or 0)
    period_keys = _ow_payload_range(payload)

    if not project_id:
        return {"ok": False, "error": "project_id obbligatorio"}

    if not role:
        return {"ok": False, "error": "mansione obbligatoria"}

    if not period_keys:
        return {"ok": False, "error": "periodo obbligatorio"}

    conn = get_connection()
    try:
        project = conn.execute(
            "SELECT id, name FROM projects WHERE id = ?",
            (project_id,),
        ).fetchone()

        if not project:
            return {"ok": False, "error": "Commessa non trovata"}

        changed = 0
        inserted = 0
        history = 0

        for period_key in period_keys:
            week = week_from_period_key(period_key)

            existing = conn.execute(
                """
                SELECT id, quantity
                FROM demands
                WHERE project_id = ?
                  AND UPPER(role) = ?
                  AND COALESCE(NULLIF(period_key, 0), 2600 + week) = ?
                LIMIT 1
                """,
                (project_id, role, period_key),
            ).fetchone()

            if existing:
                old_quantity = float(existing["quantity"] or 0)
                if old_quantity == quantity:
                    continue

                conn.execute(
                    """
                    UPDATE demands
                    SET quantity = ?,
                        week = ?,
                        period_key = ?,
                        role = ?,
                        note = ?
                    WHERE id = ?
                    """,
                    (
                        quantity,
                        week,
                        period_key,
                        role,
                        "GANTT_OLD_WORKFLOW_DEMAND_UPDATE",
                        existing["id"],
                    ),
                )
                changed += 1

                conn.execute(
                    """
                    INSERT INTO demand_history(
                        project_id,
                        week,
                        period_key,
                        role,
                        old_quantity,
                        new_quantity,
                        note
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        project_id,
                        week,
                        period_key,
                        role,
                        old_quantity,
                        quantity,
                        "GANTT_OLD_WORKFLOW_DEMAND_UPDATE",
                    ),
                )
                history += 1
            else:
                if quantity <= 0:
                    continue

                conn.execute(
                    """
                    INSERT INTO demands(project_id, week, period_key, role, quantity, note)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        project_id,
                        week,
                        period_key,
                        role,
                        quantity,
                        "GANTT_OLD_WORKFLOW_DEMAND_INSERT",
                    ),
                )
                inserted += 1

                conn.execute(
                    """
                    INSERT INTO demand_history(
                        project_id,
                        week,
                        period_key,
                        role,
                        old_quantity,
                        new_quantity,
                        note
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        project_id,
                        week,
                        period_key,
                        role,
                        0,
                        quantity,
                        "GANTT_OLD_WORKFLOW_DEMAND_INSERT",
                    ),
                )
                history += 1

            try:
                release_allocations_for_zero_demand(conn, project_id, role, week)
            except Exception:
                pass

        conn.commit()

        return {
            "ok": True,
            "changed": changed,
            "inserted": inserted,
            "history": history,
        }
    finally:
        conn.close()


@app.post("/api/old-workflow/resources-save")
def old_workflow_resources_save(payload: dict):
    rows = payload.get("rows") or []

    if not isinstance(rows, list):
        return {"ok": False, "error": "rows non valido"}

    conn = get_connection()
    try:
        changed = 0

        for item in rows:
            resource_id = int(item.get("id") or 0)
            if not resource_id:
                continue

            name = clean_text(item.get("name") or "")
            role = normalize_role(item.get("role") or "")
            availability_note = clean_text(item.get("availability_note") or "")
            is_active = 1 if int(item.get("is_active") or 0) else 0

            conn.execute(
                """
                UPDATE resources
                SET name = ?,
                    role = ?,
                    availability_note = ?,
                    is_active = ?
                WHERE id = ?
                """,
                (
                    name,
                    role,
                    availability_note,
                    is_active,
                    resource_id,
                ),
            )
            changed += 1

        conn.commit()

        return {
            "ok": True,
            "changed": changed,
        }
    finally:
        conn.close()

# === OLD WORKFLOW GANTT RESOURCES BACKEND END ===
