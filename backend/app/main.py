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
    hours: float = 0
    load_percent: float = 0
    note: str = ""


class DemandRangeUpsert(BaseModel):
    project_id: int
    role: str
    week_from: int
    week_to: int
    quantity: float
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
                payload.role.strip(),
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
        cursor = conn.execute(
            """
            INSERT INTO demands (project_id, week, role, quantity, note)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                payload.project_id,
                payload.week,
                payload.role.strip(),
                payload.quantity,
                payload.note.strip(),
            )
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
                SELECT id
                FROM demands
                WHERE project_id = ? AND week = ? AND UPPER(role) = ?
                """,
                (payload.project_id, week, role),
            ).fetchone()

            if existing:
                conn.execute(
                    """
                    UPDATE demands
                    SET quantity = ?, note = ?, role = ?
                    WHERE id = ?
                    """,
                    (payload.quantity, note, role, existing["id"]),
                )
                updated_ids.append(existing["id"])
            else:
                cursor = conn.execute(
                    """
                    INSERT INTO demands (project_id, week, role, quantity, note)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (payload.project_id, week, role, payload.quantity, note),
                )
                updated_ids.append(cursor.lastrowid)

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
                a.hours,
                a.load_percent,
                a.note
            FROM allocations a
            JOIN resources r ON r.id = a.resource_id
            JOIN projects p ON p.id = a.project_id
            ORDER BY p.name ASC, a.week ASC, r.name ASC
            """
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()


@app.post("/api/allocations")
def create_allocation(payload: AllocationCreate):
    conn = get_connection()
    try:
        cursor = conn.execute(
            """
            INSERT INTO allocations (resource_id, project_id, week, hours, load_percent, note)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                payload.resource_id,
                payload.project_id,
                payload.week,
                payload.hours,
                payload.load_percent,
                payload.note.strip(),
            )
        )
        conn.commit()

        row = conn.execute(
            """
            SELECT
                a.id,
                a.resource_id,
                r.name AS resource_name,
                a.project_id,
                p.name AS project_name,
                a.week,
                a.hours,
                a.load_percent,
                a.note
            FROM allocations a
            JOIN resources r ON r.id = a.resource_id
            JOIN projects p ON p.id = a.project_id
            WHERE a.id = ?
            """,
            (cursor.lastrowid,)
        ).fetchone()

        return dict(row)
    finally:
        conn.close()