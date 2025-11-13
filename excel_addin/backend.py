"""
Excel Add-In Backend - PostgreSQL Implementation
Implements the exact API specification with PostgreSQL persistence

Endpoints:
  PUT  /wb/upsert-model         - Create or update a model (with versioning)
  GET  /wb/load-model           - Load model metadata by model_id
  POST /wb/create-model-trace   - Create trace log entry

Database Tables:
  - workbook_model: Model metadata and tracked ranges
  - workbook_trace: Trace log entries

Install:
    pip install fastapi uvicorn psycopg[binary] python-multipart

Run:
    python backend.py
"""

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional, Any
from datetime import datetime
import logging
import json
import os
import psycopg
from psycopg.rows import dict_row
from contextlib import contextmanager

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel Add-In Backend",
    description="PostgreSQL-backed Excel Add-In for Model Tracking",
    version="2.0.0"
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

DB_CONFIG = {
    "host": os.getenv("DB_HOST", "localhost"),
    "port": int(os.getenv("DB_PORT", "5432")),
    "database": os.getenv("DB_NAME", "excel_addin"),
    "user": os.getenv("DB_USER", "excel_user"),
    "password": os.getenv("DB_PASSWORD", "excel_pass"),
}


@contextmanager
def get_db():
    """Database connection context manager"""
    conn = psycopg.connect(**DB_CONFIG)
    try:
        yield conn
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


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
            with conn.cursor() as cur:
                cur.execute("SELECT COUNT(*) FROM workbook_model")
                models_count = cur.fetchone()[0]
                cur.execute("SELECT COUNT(*) FROM workbook_trace")
                traces_count = cur.fetchone()[0]

        return {
            "service": "Excel Add-In Backend",
            "version": "2.0.0",
            "database": "PostgreSQL",
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
            "version": "2.0.0",
            "database": "PostgreSQL",
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
        with conn.cursor(row_factory=dict_row) as cur:
            tracked_ranges_json = json.dumps([r.dict() for r in request.tracked_ranges])

            # Case 1: No model_id provided - create new model
            if not request.model_id:
                model_id = generate_model_id()
                cur.execute("""
                    INSERT INTO workbook_model (model_id, model_name, version, tracked_ranges)
                    VALUES (%s, %s, %s, %s::jsonb)
                    RETURNING model_id, model_name, version, tracked_ranges
                """, (model_id, request.model_name, 1, tracked_ranges_json))

                row = cur.fetchone()
                logger.info(f"✅ Created new model: {request.model_name} ({model_id}) v1")

                return WorkbookModel(
                    model_id=row['model_id'],
                    model_name=row['model_name'],
                    version=row['version'],
                    tracked_ranges=[TrackedRange(**r) for r in row['tracked_ranges']]
                )

            # Case 2: model_id provided - check if exists
            cur.execute("SELECT version FROM workbook_model WHERE model_id = %s", (request.model_id,))
            existing = cur.fetchone()

            if existing:
                # Update existing model and increment version
                new_version = existing['version'] + 1
                cur.execute("""
                    UPDATE workbook_model
                    SET model_name = %s, tracked_ranges = %s::jsonb, version = %s
                    WHERE model_id = %s
                    RETURNING model_id, model_name, version, tracked_ranges
                """, (request.model_name, tracked_ranges_json, new_version, request.model_id))

                row = cur.fetchone()
                logger.info(f"🔄 Updated model: {request.model_name} ({request.model_id}) v{existing['version']} → v{new_version}")

                return WorkbookModel(
                    model_id=row['model_id'],
                    model_name=row['model_name'],
                    version=row['version'],
                    tracked_ranges=[TrackedRange(**r) for r in row['tracked_ranges']]
                )
            else:
                # Create new model with provided ID
                cur.execute("""
                    INSERT INTO workbook_model (model_id, model_name, version, tracked_ranges)
                    VALUES (%s, %s, %s, %s::jsonb)
                    RETURNING model_id, model_name, version, tracked_ranges
                """, (request.model_id, request.model_name, 1, tracked_ranges_json))

                row = cur.fetchone()
                logger.info(f"✅ Created model with provided ID: {request.model_name} ({request.model_id}) v1")

                return WorkbookModel(
                    model_id=row['model_id'],
                    model_name=row['model_name'],
                    version=row['version'],
                    tracked_ranges=[TrackedRange(**r) for r in row['tracked_ranges']]
                )


@app.get("/wb/load-model", response_model=WorkbookModel)
def load_model(model_id: str = Query(..., description="Model ID to load")):
    """
    GET /wb/load-model?model_id={model_id}

    Load model metadata by model_id.
    """
    with get_db() as conn:
        with conn.cursor(row_factory=dict_row) as cur:
            cur.execute("""
                SELECT model_id, model_name, version, tracked_ranges
                FROM workbook_model
                WHERE model_id = %s
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
                tracked_ranges=[TrackedRange(**r) for r in row['tracked_ranges']]
            )


@app.post("/wb/create-model-trace", response_model=CreateTraceResponse)
def create_model_trace(request: CreateTraceRequest):
    """
    POST /wb/create-model-trace

    Create a trace log entry when a tracked range changes.
    """
    with get_db() as conn:
        with conn.cursor(row_factory=dict_row) as cur:
            # Verify model exists
            cur.execute("SELECT model_name FROM workbook_model WHERE model_id = %s", (request.model_id,))
            model = cur.fetchone()

            if not model:
                logger.warning(f"❌ Trace rejected: Model not found: {request.model_id}")
                raise HTTPException(status_code=404, detail=f"Model not found: {request.model_id}")

            # Insert trace
            value_json = json.dumps(request.value) if request.value is not None else None
            cur.execute("""
                INSERT INTO workbook_trace (model_id, timestamp, tracked_range_name, username, value)
                VALUES (%s, %s, %s, %s, %s)
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
        with conn.cursor(row_factory=dict_row) as cur:
            # Verify model exists
            cur.execute("SELECT model_name FROM workbook_model WHERE model_id = %s", (request.model_id,))
            model = cur.fetchone()

            if not model:
                logger.warning(f"❌ Batch trace rejected: Model not found: {request.model_id}")
                raise HTTPException(status_code=404, detail=f"Model not found: {request.model_id}")

            # Insert all traces
            for change in request.changes:
                value_json = json.dumps(change.get("value"))
                cur.execute("""
                    INSERT INTO workbook_trace (model_id, timestamp, tracked_range_name, username, value)
                    VALUES (%s, %s, %s, %s, %s)
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
        with conn.cursor(row_factory=dict_row) as cur:
            cur.execute("""
                SELECT model_id, model_name, version, tracked_ranges, created_at, updated_at
                FROM workbook_model
                ORDER BY created_at DESC
            """)
            return cur.fetchall()


@app.get("/wb/traces/{model_id}")
def get_model_traces(model_id: str, limit: int = Query(50, description="Max traces to return")):
    """Get recent traces for a model (for debugging)"""
    with get_db() as conn:
        with conn.cursor(row_factory=dict_row) as cur:
            cur.execute("SELECT 1 FROM workbook_model WHERE model_id = %s", (model_id,))
            if not cur.fetchone():
                raise HTTPException(status_code=404, detail=f"Model not found: {model_id}")

            cur.execute("""
                SELECT trace_id, model_id, timestamp, tracked_range_name, username, value, created_at
                FROM workbook_trace
                WHERE model_id = %s
                ORDER BY timestamp DESC
                LIMIT %s
            """, (model_id, limit))

            return cur.fetchall()


@app.get("/wb/traces")
def get_all_traces(limit: int = Query(100, description="Max traces to return")):
    """Get all traces (for debugging)"""
    with get_db() as conn:
        with conn.cursor(row_factory=dict_row) as cur:
            cur.execute("""
                SELECT trace_id, model_id, timestamp, tracked_range_name, username, value, created_at
                FROM workbook_trace
                ORDER BY timestamp DESC
                LIMIT %s
            """, (limit,))

            return cur.fetchall()


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
    """Test database connection on startup"""
    try:
        with get_db() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
                logger.info("✅ Database connection successful")
    except Exception as e:
        logger.error(f"❌ Database connection failed: {e}")
        logger.error("Make sure PostgreSQL is running: docker-compose up -d")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000)
