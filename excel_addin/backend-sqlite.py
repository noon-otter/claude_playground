"""
Excel Add-In Backend - SQLite Implementation
Drop-in replacement for PostgreSQL backend using SQLite

Implements the exact same API specification:
  PUT  /wb/upsert-model         - Create or update a model (with versioning)
  GET  /wb/load-model           - Load model metadata by model_id
  POST /wb/create-model-trace   - Create trace log entry

Install:
    pip install fastapi uvicorn python-multipart

Run:
    python backend-sqlite.py
"""

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional, Any
from datetime import datetime
import logging
import json
import sqlite3
from contextlib import contextmanager

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel Add-In Backend (SQLite)",
    description="SQLite-backed Excel Add-In for Model Tracking",
    version="2.0.0-sqlite"
)

# Enable CORS for local development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =====================================================
# DATABASE CONNECTION
# =====================================================

DB_FILE = "excel_addin.db"


@contextmanager
def get_db():
    """Database connection context manager"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row  # Return rows as dictionaries
    try:
        yield conn
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def init_database():
    """Initialize SQLite database schema"""
    with get_db() as conn:
        cur = conn.cursor()

        # Create workbook_model table
        cur.execute("""
            CREATE TABLE IF NOT EXISTS workbook_model (
                model_id TEXT PRIMARY KEY,
                model_name TEXT NOT NULL,
                version INTEGER NOT NULL DEFAULT 1,
                tracked_ranges TEXT NOT NULL DEFAULT '[]',
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Create workbook_trace table
        cur.execute("""
            CREATE TABLE IF NOT EXISTS workbook_trace (
                trace_id INTEGER PRIMARY KEY AUTOINCREMENT,
                model_id TEXT NOT NULL,
                timestamp TIMESTAMP NOT NULL,
                tracked_range_name TEXT NOT NULL,
                username TEXT NOT NULL,
                value TEXT,
                created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (model_id) REFERENCES workbook_model(model_id) ON DELETE CASCADE
            )
        """)

        # Create indexes
        cur.execute("CREATE INDEX IF NOT EXISTS idx_workbook_model_name ON workbook_model(model_name)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_workbook_model_version ON workbook_model(version)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_workbook_trace_model_id ON workbook_trace(model_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_workbook_trace_timestamp ON workbook_trace(timestamp)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_workbook_trace_range_name ON workbook_trace(tracked_range_name)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_workbook_trace_username ON workbook_trace(username)")

        conn.commit()
        logger.info("✅ SQLite database initialized successfully")


# =====================================================
# PYDANTIC MODELS
# =====================================================

class TrackedRange(BaseModel):
    """Tracked Range: {name, range}"""
    name: str
    range: str


class WorkbookModel(BaseModel):
    """Complete model stored in workbook_model table"""
    model_name: str
    tracked_ranges: List[TrackedRange]
    model_id: str
    version: int


class UpsertModelRequest(BaseModel):
    """Request body for PUT /wb/upsert-model"""
    model_name: str
    tracked_ranges: List[TrackedRange]
    model_id: Optional[str] = None
    version: Optional[int] = None


class CreateTraceRequest(BaseModel):
    """Request body for POST /wb/create-model-trace (single change)"""
    model_id: str
    timestamp: str
    tracked_range_name: str
    username: str
    value: Any


class CreateTraceBatchRequest(BaseModel):
    """Request body for POST /wb/create-model-trace-batch (batch)"""
    model_id: str
    timestamp: str
    changes: List[dict]  # [{"tracked_range_name": str, "value": Any}]
    username: str


class CreateTraceResponse(BaseModel):
    """Response for POST /wb/create-model-trace"""
    success: bool


# =====================================================
# API ENDPOINTS
# =====================================================

@app.get("/")
def root():
    """Health check and service info"""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM workbook_model")
            models_count = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM workbook_trace")
            traces_count = cur.fetchone()[0]

        return {
            "service": "Excel Add-In Backend",
            "version": "2.0.0-sqlite",
            "database": "SQLite",
            "database_file": DB_FILE,
            "status": "healthy",
            "endpoints": {
                "upsert": "PUT /wb/upsert-model",
                "load": "GET /wb/load-model",
                "trace": "POST /wb/create-model-trace"
            },
            "models_count": models_count,
            "traces_count": traces_count
        }
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        return {
            "service": "Excel Add-In Backend",
            "version": "2.0.0-sqlite",
            "database": "SQLite",
            "status": "unhealthy",
            "error": str(e)
        }


@app.put("/wb/upsert-model", response_model=WorkbookModel)
def upsert_model(request: UpsertModelRequest):
    """
    PUT /wb/upsert-model

    Creates new model or updates existing model with version increment.
    """
    with get_db() as conn:
        cur = conn.cursor()
        tracked_ranges_json = json.dumps([r.dict() for r in request.tracked_ranges])

        # Case 1: No model_id provided - create new model
        if not request.model_id:
            model_id = generate_model_id()
            cur.execute("""
                INSERT INTO workbook_model (model_id, model_name, version, tracked_ranges)
                VALUES (?, ?, ?, ?)
            """, (model_id, request.model_name, 1, tracked_ranges_json))

            # Fetch the created row
            cur.execute("""
                SELECT model_id, model_name, version, tracked_ranges
                FROM workbook_model WHERE model_id = ?
            """, (model_id,))
            row = cur.fetchone()

            logger.info(f"✅ Created new model: {request.model_name} ({model_id}) v1")

            return WorkbookModel(
                model_id=row['model_id'],
                model_name=row['model_name'],
                version=row['version'],
                tracked_ranges=[TrackedRange(**r) for r in json.loads(row['tracked_ranges'])]
            )

        # Case 2: model_id provided - check if exists
        cur.execute("SELECT version FROM workbook_model WHERE model_id = ?", (request.model_id,))
        existing = cur.fetchone()

        if existing:
            # Update existing model and increment version
            new_version = existing['version'] + 1
            cur.execute("""
                UPDATE workbook_model
                SET model_name = ?, tracked_ranges = ?, version = ?, updated_at = CURRENT_TIMESTAMP
                WHERE model_id = ?
            """, (request.model_name, tracked_ranges_json, new_version, request.model_id))

            # Fetch the updated row
            cur.execute("""
                SELECT model_id, model_name, version, tracked_ranges
                FROM workbook_model WHERE model_id = ?
            """, (request.model_id,))
            row = cur.fetchone()

            logger.info(f"🔄 Updated model: {request.model_name} ({request.model_id}) v{existing['version']} → v{new_version}")

            return WorkbookModel(
                model_id=row['model_id'],
                model_name=row['model_name'],
                version=row['version'],
                tracked_ranges=[TrackedRange(**r) for r in json.loads(row['tracked_ranges'])]
            )
        else:
            # Create new model with provided ID
            cur.execute("""
                INSERT INTO workbook_model (model_id, model_name, version, tracked_ranges)
                VALUES (?, ?, ?, ?)
            """, (request.model_id, request.model_name, 1, tracked_ranges_json))

            # Fetch the created row
            cur.execute("""
                SELECT model_id, model_name, version, tracked_ranges
                FROM workbook_model WHERE model_id = ?
            """, (request.model_id,))
            row = cur.fetchone()

            logger.info(f"✅ Created model with provided ID: {request.model_name} ({request.model_id}) v1")

            return WorkbookModel(
                model_id=row['model_id'],
                model_name=row['model_name'],
                version=row['version'],
                tracked_ranges=[TrackedRange(**r) for r in json.loads(row['tracked_ranges'])]
            )


@app.get("/wb/load-model", response_model=WorkbookModel)
def load_model(model_id: str = Query(..., description="Model ID to load")):
    """
    GET /wb/load-model?model_id={model_id}

    Load model metadata by model_id.
    """
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT model_id, model_name, version, tracked_ranges
            FROM workbook_model
            WHERE model_id = ?
        """, (model_id,))

        row = cur.fetchone()

        if not row:
            logger.warning(f"❌ Model not found: {model_id}")
            raise HTTPException(status_code=404, detail=f"Model not found: {model_id}")

        logger.info(f"📂 Loaded model: {row['model_name']} ({model_id}) v{row['version']}")

        return WorkbookModel(
            model_id=row['model_id'],
            model_name=row['model_name'],
            version=row['version'],
            tracked_ranges=[TrackedRange(**r) for r in json.loads(row['tracked_ranges'])]
        )


@app.post("/wb/create-model-trace", response_model=CreateTraceResponse)
def create_model_trace(request: CreateTraceRequest):
    """
    POST /wb/create-model-trace

    Create a trace log entry when a tracked range changes.
    """
    with get_db() as conn:
        cur = conn.cursor()

        # Verify model exists
        cur.execute("SELECT model_name FROM workbook_model WHERE model_id = ?", (request.model_id,))
        model = cur.fetchone()

        if not model:
            logger.warning(f"❌ Trace rejected: Model not found: {request.model_id}")
            raise HTTPException(status_code=404, detail=f"Model not found: {request.model_id}")

        # Insert trace
        value_json = json.dumps(request.value) if request.value is not None else None
        cur.execute("""
            INSERT INTO workbook_trace (model_id, timestamp, tracked_range_name, username, value)
            VALUES (?, ?, ?, ?, ?)
        """, (
            request.model_id,
            request.timestamp,
            request.tracked_range_name,
            request.username,
            value_json
        ))

        logger.info(
            f"📝 Trace recorded: {model['model_name']} | "
            f"{request.tracked_range_name} = {request.value} | "
            f"by {request.username}"
        )

        return CreateTraceResponse(success=True)


@app.post("/wb/create-model-trace-batch", response_model=CreateTraceResponse)
def create_model_trace_batch(request: CreateTraceBatchRequest):
    """
    POST /wb/create-model-trace-batch

    Batch version for multiple tracked range changes at once.
    """
    with get_db() as conn:
        cur = conn.cursor()

        # Verify model exists
        cur.execute("SELECT model_name FROM workbook_model WHERE model_id = ?", (request.model_id,))
        model = cur.fetchone()

        if not model:
            logger.warning(f"❌ Batch trace rejected: Model not found: {request.model_id}")
            raise HTTPException(status_code=404, detail=f"Model not found: {request.model_id}")

        # Insert all traces
        for change in request.changes:
            value_json = json.dumps(change.get("value"))
            cur.execute("""
                INSERT INTO workbook_trace (model_id, timestamp, tracked_range_name, username, value)
                VALUES (?, ?, ?, ?, ?)
            """, (
                request.model_id,
                request.timestamp,
                change["tracked_range_name"],
                request.username,
                value_json
            ))

        logger.info(
            f"📦 Batch trace recorded: {model['model_name']} | "
            f"{len(request.changes)} changes | "
            f"by {request.username}"
        )

        return CreateTraceResponse(success=True)


# =====================================================
# UTILITY ENDPOINTS (For debugging/monitoring)
# =====================================================

@app.get("/wb/models")
def get_all_models():
    """Get all registered models (for admin/debugging)"""
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT model_id, model_name, version, tracked_ranges, created_at, updated_at
            FROM workbook_model
            ORDER BY created_at DESC
        """)
        rows = cur.fetchall()
        return [dict(row) for row in rows]


@app.get("/wb/traces/{model_id}")
def get_model_traces(model_id: str, limit: int = Query(50, description="Max traces to return")):
    """Get recent traces for a model (for debugging)"""
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM workbook_model WHERE model_id = ?", (model_id,))
        if not cur.fetchone():
            raise HTTPException(status_code=404, detail=f"Model not found: {model_id}")

        cur.execute("""
            SELECT trace_id, model_id, timestamp, tracked_range_name, username, value, created_at
            FROM workbook_trace
            WHERE model_id = ?
            ORDER BY timestamp DESC
            LIMIT ?
        """, (model_id, limit))

        rows = cur.fetchall()
        return [dict(row) for row in rows]


@app.get("/wb/traces")
def get_all_traces(limit: int = Query(100, description="Max traces to return")):
    """Get all traces (for debugging)"""
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT trace_id, model_id, timestamp, tracked_range_name, username, value, created_at
            FROM workbook_trace
            ORDER BY timestamp DESC
            LIMIT ?
        """, (limit,))

        rows = cur.fetchall()
        return [dict(row) for row in rows]


# =====================================================
# HELPER FUNCTIONS
# =====================================================

def generate_model_id() -> str:
    """Generate unique model ID: excel_<timestamp>_<random>"""
    timestamp = str(int(datetime.utcnow().timestamp() * 1000))
    random = str(abs(hash(datetime.utcnow().isoformat())))[:8]
    return f"excel_{timestamp}_{random}"


# =====================================================
# STARTUP
# =====================================================

@app.on_event("startup")
async def startup_event():
    """Initialize database on startup"""
    try:
        init_database()
        logger.info("✅ SQLite backend ready")
    except Exception as e:
        logger.error(f"❌ Database initialization failed: {e}")


if __name__ == "__main__":
    import uvicorn
    logger.info("🚀 Starting Excel Add-In Backend (SQLite)")
    logger.info(f"📁 Database file: {DB_FILE}")
    logger.info(f"🌐 Server: http://localhost:5000")
    uvicorn.run(app, host="0.0.0.0", port=5000, log_level="info")
