
from fastapi import APIRouter
from database.db import get_connection, row_to_dict

router = APIRouter(prefix="/api/history", tags=["History"])


@router.get("/")
def get_history():
    conn = get_connection()

    rows = conn.execute("""
        SELECT * FROM predictions
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return [row_to_dict(row) for row in rows]


@router.delete("/")
def clear_history():
    conn = get_connection()
    conn.execute("DELETE FROM predictions")
    conn.commit()
    conn.close()

    return {"success": True, "message": "History cleared"}
