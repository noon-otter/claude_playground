"""
Domino Spreadsheet Backend - Architecture-Compliant Implementation
Implements the exact API specification from DEPLOYMENT.md

Endpoints:
  PUT  /wb/upsert-model         - Create or update a model (with versioning)
  GET  /wb/load-model           - Load model metadata by model_id
  POST /wb/create-model-trace   - Create trace log entry

Database Tables:
  - dbo.workbook_model: Model metadata and tracked ranges
  - dbo.workbook_trace: Trace log entries

Install:
    pip install fastapi uvicorn python-multipart

Run:
    uvicorn domino-api-backend:app --reload --port 5000
"""

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional, Any
from datetime import datetime
import logging
import json

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Domino Spreadsheet Backend",
    description="Excel Add-In Backend per DEPLOYMENT.md Architecture",
    version="1.0.0"
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
# IN-MEMORY DATABASE (Replace with SQL Server)
# =====================================================
# In production, replace these with actual SQL queries to:
# - dbo.workbook_model
# - dbo.workbook_trace

workbook_model_db = {}  # Key: model_id, Value: WorkbookModel
workbook_trace_db = []  # List of WorkbookTrace entries


# =====================================================
# PYDANTIC MODELS (Match TypeScript interfaces)
# =====================================================

class TrackedRange(BaseModel):
    """Tracked Range: {name, range}"""
    name: str
    range: str


class WorkbookModel(BaseModel):
    """Complete model stored in dbo.workbook_model"""
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


class LoadModelRequest(BaseModel):
    """Query params for GET /wb/load-model"""
    model_id: str


class CreateTraceRequest(BaseModel):
    """Request body for POST /wb/create-model-trace (single change)"""
    model_id: str
    timestamp: str
    tracked_range_name: str
    username: str
    value: Any


class CreateTraceBatchRequest(BaseModel):
    """Request body for POST /wb/create-model-trace (batch)"""
    model_id: str
    timestamp: str
    changes: List[dict]  # [{"tracked_range_name": str, "value": Any}]
    username: str


class CreateTraceResponse(BaseModel):
    """Response for POST /wb/create-model-trace"""
    success: bool


# =====================================================
# API ENDPOINTS (Per Architecture Specification)
# =====================================================

@app.get("/")
def root():
    """Health check and service info"""
    return {
        "service": "Domino Spreadsheet Backend",
        "version": "1.0.0",
        "architecture": "DEPLOYMENT.md compliant",
        "endpoints": {
            "upsert": "PUT /wb/upsert-model",
            "load": "GET /wb/load-model",
            "trace": "POST /wb/create-model-trace"
        },
        "models_count": len(workbook_model_db),
        "traces_count": len(workbook_trace_db)
    }


@app.put("/wb/upsert-model", response_model=WorkbookModel)
def upsert_model(request: UpsertModelRequest):
    """
    PUT /wb/upsert-model

    Used for:
    - Creating a new model
    - Updating a model's name/tracked ranges
    - Returning versioned model metadata

    Behavior:
    - If no model_id provided → create new model, generate new model_id
    - If model_id exists → update and increment version
    - If provided model_id doesn't exist → create new model

    Request:
        {
          "model_name": string,
          "tracked_ranges": TrackedRange[],
          "model_id": string (optional),
          "version": int (optional)
        }

    Response:
        {
          "model_name": string,
          "tracked_ranges": TrackedRange[],
          "model_id": string,
          "version": int
        }
    """

    # Case 1: Create new model (no model_id provided)
    if not request.model_id:
        model_id = generate_model_id()
        model = WorkbookModel(
            model_name=request.model_name,
            tracked_ranges=request.tracked_ranges,
            model_id=model_id,
            version=1
        )
        workbook_model_db[model_id] = model.dict()

        logger.info(f"✅ Created new model: {model.model_name} ({model_id}) v1")

        return model

    # Case 2: Update existing model
    if request.model_id in workbook_model_db:
        existing = workbook_model_db[request.model_id]
        new_version = existing["version"] + 1

        model = WorkbookModel(
            model_name=request.model_name,
            tracked_ranges=request.tracked_ranges,
            model_id=request.model_id,
            version=new_version
        )
        workbook_model_db[request.model_id] = model.dict()

        logger.info(f"🔄 Updated model: {model.model_name} ({request.model_id}) v{existing['version']} → v{new_version}")

        return model

    # Case 3: model_id provided but doesn't exist → create new
    model = WorkbookModel(
        model_name=request.model_name,
        tracked_ranges=request.tracked_ranges,
        model_id=request.model_id,
        version=1
    )
    workbook_model_db[request.model_id] = model.dict()

    logger.info(f"✅ Created model with provided ID: {model.model_name} ({request.model_id}) v1")

    return model


@app.get("/wb/load-model", response_model=WorkbookModel)
def load_model(model_id: str = Query(..., description="Model ID to load")):
    """
    GET /wb/load-model?model_id={model_id}

    Used for:
    - Loading metadata for an existing model

    Request:
        Query param: model_id

    Response:
        {
          "model_name": string,
          "tracked_ranges": TrackedRange[],
          "model_id": string,
          "version": int
        }
    """

    if model_id not in workbook_model_db:
        logger.warning(f"❌ Model not found: {model_id}")
        raise HTTPException(status_code=404, detail=f"Model not found: {model_id}")

    model = WorkbookModel(**workbook_model_db[model_id])
    logger.info(f"📂 Loaded model: {model.model_name} ({model_id}) v{model.version}")

    return model


@app.post("/wb/create-model-trace", response_model=CreateTraceResponse)
def create_model_trace(request: CreateTraceRequest):
    """
    POST /wb/create-model-trace

    Triggered when a tracked range changes.

    Request (Single Change):
        {
          "model_id": string,
          "timestamp": string,
          "tracked_range_name": string,
          "username": string,
          "value": any
        }

    Response:
        {
          "success": bool
        }

    Note: The backend may also support a batch version:
        {
          "model_id": string,
          "timestamp": string,
          "changes": [{"tracked_range_name": string, "value": any}],
          "username": string
        }
    """

    # Verify model exists
    if request.model_id not in workbook_model_db:
        logger.warning(f"❌ Trace rejected: Model not found: {request.model_id}")
        raise HTTPException(status_code=404, detail=f"Model not found: {request.model_id}")

    # Store trace
    trace = {
        "model_id": request.model_id,
        "timestamp": request.timestamp,
        "tracked_range_name": request.tracked_range_name,
        "username": request.username,
        "value": request.value
    }
    workbook_trace_db.append(trace)

    model = workbook_model_db[request.model_id]
    logger.info(
        f"📝 Trace recorded: {model['model_name']} | "
        f"{request.tracked_range_name} = {request.value} | "
        f"by {request.username}"
    )

    # In production: INSERT INTO dbo.workbook_trace
    # INSERT INTO dbo.workbook_trace (model_id, timestamp, tracked_range_name, username, value)
    # VALUES (@model_id, @timestamp, @tracked_range_name, @username, @value)

    return CreateTraceResponse(success=True)


@app.post("/wb/create-model-trace-batch", response_model=CreateTraceResponse)
def create_model_trace_batch(request: CreateTraceBatchRequest):
    """
    POST /wb/create-model-trace-batch

    Batch version for multiple tracked range changes at once.

    Request:
        {
          "model_id": string,
          "timestamp": string,
          "changes": [
              {"tracked_range_name": string, "value": any},
              {"tracked_range_name": string, "value": any}
          ],
          "username": string
        }

    Response:
        {
          "success": bool
        }
    """

    # Verify model exists
    if request.model_id not in workbook_model_db:
        logger.warning(f"❌ Batch trace rejected: Model not found: {request.model_id}")
        raise HTTPException(status_code=404, detail=f"Model not found: {request.model_id}")

    # Store all traces
    for change in request.changes:
        trace = {
            "model_id": request.model_id,
            "timestamp": request.timestamp,
            "tracked_range_name": change["tracked_range_name"],
            "username": request.username,
            "value": change["value"]
        }
        workbook_trace_db.append(trace)

    model = workbook_model_db[request.model_id]
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
    return list(workbook_model_db.values())


@app.get("/wb/traces/{model_id}")
def get_model_traces(model_id: str, limit: int = Query(50, description="Max traces to return")):
    """Get recent traces for a model (for debugging)"""
    if model_id not in workbook_model_db:
        raise HTTPException(status_code=404, detail=f"Model not found: {model_id}")

    model_traces = [t for t in workbook_trace_db if t["model_id"] == model_id]
    model_traces.reverse()  # Most recent first

    return model_traces[:limit]


@app.get("/wb/traces")
def get_all_traces(limit: int = Query(100, description="Max traces to return")):
    """Get all traces (for debugging)"""
    traces = workbook_trace_db.copy()
    traces.reverse()  # Most recent first
    return traces[:limit]


# =====================================================
# HELPER FUNCTIONS
# =====================================================

def generate_model_id() -> str:
    """Generate unique model ID: excel_<timestamp>_<random>"""
    timestamp = str(int(datetime.utcnow().timestamp() * 1000))
    random = str(abs(hash(datetime.utcnow().isoformat())))[:8]
    return f"excel_{timestamp}_{random}"


# =====================================================
# SQL SERVER INTEGRATION EXAMPLE
# =====================================================
"""
To connect to SQL Server, install:
    pip install pyodbc

Example connection:
    import pyodbc

    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=your-server.database.windows.net;'
        'DATABASE=your-database;'
        'UID=your-username;'
        'PWD=your-password'
    )

    cursor = conn.cursor()

    # Insert model
    cursor.execute('''
        INSERT INTO dbo.workbook_model (model_id, model_name, version, tracked_ranges)
        VALUES (?, ?, ?, ?)
    ''', (model_id, model_name, version, json.dumps([r.dict() for r in tracked_ranges])))

    # Insert trace
    cursor.execute('''
        INSERT INTO dbo.workbook_trace (model_id, timestamp, tracked_range_name, username, value)
        VALUES (?, ?, ?, ?, ?)
    ''', (model_id, timestamp, tracked_range_name, username, json.dumps(value)))

    conn.commit()
"""


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000)
