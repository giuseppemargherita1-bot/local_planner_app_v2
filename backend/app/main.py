from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from app.db import init_db, get_connection


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


class AllocationCreate(BaseModel):
    resource_id: int
    project_id: int
    week: int
    role: str = ""
    hours: float = 40
    load_percent: float = 100
    note: str = ""


class DemandRangeUpsert(BaseModel):
    project_id: int
    role: str
    week_from: int
    week_to: int
    quantity: float
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


def get_resource_or_none(conn, resource_id: int):
    return conn.execute(
        """
        SELECT id, name, role, availability_note, is_active
        FROM resources
        WHERE id = ?
        """,
        (resource_id,),
    ).fetchone()


def resource_is_unavailable(resource):
    if not resource:
        return True

    if int(resource["is_active"]) != 1:
        return True

    note_upper = str(resource["availability_note"] or "").upper()

    if "INDISP" in note_upper:
        return True

    if "NON DISPONIBILE" in note_upper:
        return True

    return False


def get_resource_week_allocations(conn, resource_id: int, week: int):
    rows = conn.execute(
        """
        SELECT
            a.id,
            a.resource_id,
            r.name AS resource_name,
            a.project_id,
            p.name AS project_name,
            a.week,
            a.role,
            a.hours,
            a.load_percent,
            a.note
        FROM allocations a
        JOIN resources r ON r.id = a.resource_id
        JOIN projects p ON p.id = a.project_id
        WHERE a.resource_id = ?
          AND a.week = ?
        ORDER BY a.id ASC
        """,
        (resource_id, week),
    ).fetchall()

    return [dict(row) for row in rows]


def rebalance_resource_week(conn, resource_id: int, week: int):
    rows = conn.execute(
        """
        SELECT id
        FROM allocations
        WHERE resource_id = ?
          AND week = ?
        ORDER BY id ASC
        """,
        (resource_id, week),
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

    conn.execute(
        f"""
        DELETE FROM allocations
        WHERE id IN ({",".join(["?"] * len(delete_ids))})
        """,
        delete_ids,
    )

    return {
        "count": 2,
        "load_percent": 50,
        "deleted_ids": delete_ids,
    }


def insert_allocation(conn, resource_id: int, project_id: int, week: int, role: str, hours: float, load_percent: float, note: str):
    cursor = conn.execute(
        """
        INSERT INTO allocations (resource_id, project_id, week, role, hours, load_percent, note)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            resource_id,
            project_id,
            week,
            role.strip().upper(),
            hours,
            load_percent,
            note.strip(),
        ),
    )
    return cursor.lastrowid


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
                payload.name.strip(),
                payload.role.strip().upper(),
                payload.availability_note.strip(),
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
                payload.name.strip(),
                payload.client.strip(),
                payload.start_date.strip(),
                payload.end_date.strip(),
                payload.status.strip(),
                payload.note.strip(),
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
                d.role,
                d.quantity,
                d.note
            FROM demands d
            JOIN projects p ON p.id = d.project_id
            ORDER BY p.name ASC, d.role ASC, d.week ASC
            """
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()


@app.post("/api/demands")
def create_demand(payload: DemandCreate):
    conn = get_connection()
    try:
        role = payload.role.strip().upper()

        cursor = conn.execute(
            """
            INSERT INTO demands (project_id, week, role, quantity, note)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                payload.project_id,
                payload.week,
                role,
                payload.quantity,
                payload.note.strip(),
            )
        )

        conn.execute(
            """
            INSERT INTO demand_history
            (project_id, week, role, old_quantity, new_quantity, note)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                payload.project_id,
                payload.week,
                role,
                0,
                payload.quantity,
                "Creazione fabbisogno",
            ),
        )

        conn.commit()

        row = conn.execute(
            """
            SELECT
                d.id,
                d.project_id,
                p.name AS project_name,
                d.week,
                d.role,
                d.quantity,
                d.note
            FROM demands d
            JOIN projects p ON p.id = d.project_id
            WHERE d.id = ?
            """,
            (cursor.lastrowid,)
        ).fetchone()

        return dict(row)
    finally:
        conn.close()


@app.post("/api/demands/upsert-range")
def upsert_demand_range(payload: DemandRangeUpsert):
    conn = get_connection()
    updated_ids = []

    try:
        role = payload.role.strip().upper()
        note = payload.note.strip()
        week_from = min(payload.week_from, payload.week_to)
        week_to = max(payload.week_from, payload.week_to)

        for week in range(week_from, week_to + 1):
            existing = conn.execute(
                """
                SELECT id, quantity
                FROM demands
                WHERE project_id = ? AND week = ? AND UPPER(role) = ?
                """,
                (payload.project_id, week, role),
            ).fetchone()

            if existing:
                old_quantity = float(existing["quantity"] or 0)
                new_quantity = float(payload.quantity or 0)

                conn.execute(
                    """
                    UPDATE demands
                    SET quantity = ?, note = ?, role = ?
                    WHERE id = ?
                    """,
                    (new_quantity, note, role, existing["id"]),
                )
                updated_ids.append(existing["id"])

                if old_quantity != new_quantity:
                    conn.execute(
                        """
                        INSERT INTO demand_history
                        (project_id, week, role, old_quantity, new_quantity, note)
                        VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (
                            payload.project_id,
                            week,
                            role,
                            old_quantity,
                            new_quantity,
                            "Modifica fabbisogno",
                        ),
                    )
            else:
                cursor = conn.execute(
                    """
                    INSERT INTO demands (project_id, week, role, quantity, note)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (payload.project_id, week, role, payload.quantity, note),
                )
                updated_ids.append(cursor.lastrowid)

                conn.execute(
                    """
                    INSERT INTO demand_history
                    (project_id, week, role, old_quantity, new_quantity, note)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        payload.project_id,
                        week,
                        role,
                        0,
                        payload.quantity,
                        "Creazione fabbisogno",
                    ),
                )

        conn.commit()

        rows = conn.execute(
            f"""
            SELECT
                d.id,
                d.project_id,
                p.name AS project_name,
                d.week,
                d.role,
                d.quantity,
                d.note
            FROM demands d
            JOIN projects p ON p.id = d.project_id
            WHERE d.id IN ({",".join(["?"] * len(updated_ids))})
            ORDER BY d.week ASC
            """,
            updated_ids,
        ).fetchall()

        return {
            "updated_count": len(updated_ids),
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
        return [dict(row) for row in rows]
    finally:
        conn.close()


@app.get("/api/allocations")
def list_allocations():
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT
                a.id,
                a.resource_id,
                r.name AS resource_name,
                a.project_id,
                p.name AS project_name,
                a.week,
                a.role,
                a.hours,
                a.load_percent,
                a.note
            FROM allocations a
            JOIN resources r ON r.id = a.resource_id
            JOIN projects p ON p.id = a.project_id
            ORDER BY p.name ASC, a.week ASC, r.name ASC
            """
        ).fetchall()

        cleaned_rows = []
        for row in rows:
            item = dict(row)
            item["role"] = str(item.get("role") or "").upper()
            item["load_percent"] = float(item.get("load_percent") or 0)
            cleaned_rows.append(item)

        return cleaned_rows
    finally:
        conn.close()


@app.post("/api/allocations")
def create_allocation(payload: AllocationCreate):
    conn = get_connection()
    try:
        role = payload.role.strip().upper()

        if not role:
            resource = conn.execute(
                """
                SELECT role
                FROM resources
                WHERE id = ?
                """,
                (payload.resource_id,),
            ).fetchone()
            role = str(resource["role"] or "").strip().upper() if resource else ""

        existing_week = get_resource_week_allocations(conn, payload.resource_id, payload.week)

        if len(existing_week) > 0:
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
        role = payload.role.strip().upper()

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

        week_from = min(payload.week_from, payload.week_to)
        week_to = max(payload.week_from, payload.week_to)

        for week in range(week_from, week_to + 1):
            existing_same_cell = conn.execute(
                """
                SELECT id
                FROM allocations
                WHERE resource_id = ?
                  AND project_id = ?
                  AND week = ?
                  AND UPPER(role) = ?
                """,
                (payload.resource_id, payload.project_id, week, role),
            ).fetchone()

            if existing_same_cell:
                skipped.append({"week": week, "reason": "already_assigned_on_this_cell"})
                rebalance_resource_week(conn, payload.resource_id, week)
                continue

            existing_week = get_resource_week_allocations(conn, payload.resource_id, week)

            if existing_week:
                conflicts.append({
                    "week": week,
                    "existing": existing_week,
                })
                skipped.append({"week": week, "reason": "resource_already_allocated_this_week"})
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
        role = payload.role.strip().upper()
        mode = payload.mode.strip().lower()

        resource = get_resource_or_none(conn, payload.resource_id)

        if not resource:
            return {"ok": False, "reason": "resource_not_found"}

        if resource_is_unavailable(resource):
            return {"ok": False, "reason": "resource_unavailable"}

        existing_week = get_resource_week_allocations(conn, payload.resource_id, payload.week)

        already_same_cell = any(
            item["project_id"] == payload.project_id and str(item["role"]).upper() == role
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

            if ids_to_delete:
                conn.execute(
                    f"""
                    DELETE FROM allocations
                    WHERE id IN ({",".join(["?"] * len(ids_to_delete))})
                    """,
                    ids_to_delete,
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

            conn.execute(
                """
                DELETE FROM allocations
                WHERE id = ?
                  AND resource_id = ?
                  AND week = ?
                """,
                (
                    payload.remove_allocation_id,
                    payload.resource_id,
                    payload.week,
                ),
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
        role = payload.role.strip().upper()
        week_from = min(payload.week_from, payload.week_to)
        week_to = max(payload.week_from, payload.week_to)
        touched_weeks = []

        rows = conn.execute(
            """
            SELECT id, week
            FROM allocations
            WHERE resource_id = ?
              AND project_id = ?
              AND UPPER(role) = ?
              AND week BETWEEN ? AND ?
            """,
            (
                payload.resource_id,
                payload.project_id,
                role,
                week_from,
                week_to,
            ),
        ).fetchall()

        removed_ids = [row["id"] for row in rows]
        touched_weeks = sorted({int(row["week"]) for row in rows})

        if removed_ids:
            conn.execute(
                f"""
                DELETE FROM allocations
                WHERE id IN ({",".join(["?"] * len(removed_ids))})
                """,
                removed_ids,
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