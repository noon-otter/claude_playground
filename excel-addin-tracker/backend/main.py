"""
Domino Spreadsheet Backend API
Provides endpoints for Excel workbook model tracking
"""
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional, Any
import psycopg
import uuid
import os
from datetime import datetime
import json

app = FastAPI(title="Domino Spreadsheet Backend")

# CORS middleware to allow requests from Excel Add-in
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Database connection
DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres:postgres@localhost:5432/excel_tracker"
)

def get_db_connection():
    """Get database connection"""
    return psycopg.connect(DATABASE_URL)

# Data models
class TrackedRange(BaseModel):
    name: str
    range: str

class WorkbookModel(BaseModel):
    model_name: str
    tracked_ranges: List[TrackedRange]
    model_id: Optional[str] = None
    version: Optional[int] = None

class WorkbookModelResponse(BaseModel):
    model_name: str
    tracked_ranges: List[TrackedRange]
    model_id: str
    version: int

class ModelTraceRequest(BaseModel):
    model_id: str
    timestamp: str
    tracked_range_name: str
    username: str
    value: Any

class ModelTraceResponse(BaseModel):
    success: bool

class BatchModelTraceRequest(BaseModel):
    model_id: str
    timestamp: str
    changes: List[dict]
    username: str

# API Endpoints

@app.get("/")
async def root():
    """Health check endpoint"""
    return {"status": "ok", "service": "Domino Spreadsheet Backend"}

@app.put("/wb/upsert-model", response_model=WorkbookModelResponse)
async def upsert_model(model: WorkbookModel):
    """
    Create or update a workbook model

    - If no model_id provided → create new model, generate new model_id
    - If model_id exists → update and increment version
    - If provided model_id doesn't exist → create new model
    """
    conn = get_db_connection()
    try:
        with conn.cursor() as cur:
            if model.model_id:
                # Check if model exists
                cur.execute(
                    "SELECT version FROM workbook_model WHERE model_id = %s",
                    (model.model_id,)
                )
                result = cur.fetchone()

                if result:
                    # Update existing model
                    new_version = result[0] + 1
                    cur.execute(
                        """
                        UPDATE workbook_model
                        SET model_name = %s,
                            tracked_ranges = %s,
                            version = %s,
                            updated_at = CURRENT_TIMESTAMP
                        WHERE model_id = %s
                        """,
                        (
                            model.model_name,
                            json.dumps([tr.dict() for tr in model.tracked_ranges]),
                            new_version,
                            model.model_id
                        )
                    )
                    conn.commit()

                    return WorkbookModelResponse(
                        model_name=model.model_name,
                        tracked_ranges=model.tracked_ranges,
                        model_id=model.model_id,
                        version=new_version
                    )
                else:
                    # Model ID provided but doesn't exist - create new
                    new_model_id = model.model_id
                    new_version = 1
            else:
                # No model_id provided - create new
                new_model_id = str(uuid.uuid4())
                new_version = 1

            # Insert new model
            cur.execute(
                """
                INSERT INTO workbook_model (model_id, model_name, tracked_ranges, version)
                VALUES (%s, %s, %s, %s)
                """,
                (
                    new_model_id,
                    model.model_name,
                    json.dumps([tr.dict() for tr in model.tracked_ranges]),
                    new_version
                )
            )
            conn.commit()

            return WorkbookModelResponse(
                model_name=model.model_name,
                tracked_ranges=model.tracked_ranges,
                model_id=new_model_id,
                version=new_version
            )
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

@app.get("/wb/load-model", response_model=WorkbookModelResponse)
async def load_model(model_id: str):
    """
    Load model metadata for an existing model
    """
    conn = get_db_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT model_name, tracked_ranges, model_id, version
                FROM workbook_model
                WHERE model_id = %s
                """,
                (model_id,)
            )
            result = cur.fetchone()

            if not result:
                raise HTTPException(status_code=404, detail="Model not found")

            model_name, tracked_ranges_json, model_id, version = result
            tracked_ranges = [
                TrackedRange(**tr) for tr in json.loads(tracked_ranges_json)
            ]

            return WorkbookModelResponse(
                model_name=model_name,
                tracked_ranges=tracked_ranges,
                model_id=model_id,
                version=version
            )
    finally:
        conn.close()

@app.post("/wb/create-model-trace", response_model=ModelTraceResponse)
async def create_model_trace(trace: ModelTraceRequest):
    """
    Create a trace entry when a tracked range changes
    """
    conn = get_db_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO workbook_trace
                (model_id, timestamp, tracked_range_name, username, value)
                VALUES (%s, %s, %s, %s, %s)
                """,
                (
                    trace.model_id,
                    trace.timestamp,
                    trace.tracked_range_name,
                    trace.username,
                    json.dumps(trace.value)
                )
            )
            conn.commit()

            return ModelTraceResponse(success=True)
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

@app.post("/wb/create-model-trace-batch", response_model=ModelTraceResponse)
async def create_model_trace_batch(trace: BatchModelTraceRequest):
    """
    Create multiple trace entries at once (batch operation)
    """
    conn = get_db_connection()
    try:
        with conn.cursor() as cur:
            for change in trace.changes:
                cur.execute(
                    """
                    INSERT INTO workbook_trace
                    (model_id, timestamp, tracked_range_name, username, value)
                    VALUES (%s, %s, %s, %s, %s)
                    """,
                    (
                        trace.model_id,
                        trace.timestamp,
                        change.get("tracked_range_name"),
                        trace.username,
                        json.dumps(change.get("value"))
                    )
                )
            conn.commit()

            return ModelTraceResponse(success=True)
    except Exception as e:
        conn.rollback()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        conn.close()

@app.get("/wb/model-traces/{model_id}")
async def get_model_traces(model_id: str, limit: int = 100):
    """
    Get trace history for a model (useful for debugging)
    """
    conn = get_db_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT timestamp, tracked_range_name, username, value
                FROM workbook_trace
                WHERE model_id = %s
                ORDER BY timestamp DESC
                LIMIT %s
                """,
                (model_id, limit)
            )
            results = cur.fetchall()

            traces = [
                {
                    "timestamp": row[0],
                    "tracked_range_name": row[1],
                    "username": row[2],
                    "value": json.loads(row[3]) if row[3] else None
                }
                for row in results
            ]

            return {"model_id": model_id, "traces": traces}
    finally:
        conn.close()

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
