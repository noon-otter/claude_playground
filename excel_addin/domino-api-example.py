"""
Domino API Backend Example
FastAPI server that receives Excel governance events

Install:
    pip install fastapi uvicorn python-multipart

Run:
    uvicorn domino-api-example:app --reload --port 5000
"""

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional, Any
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Domino Excel Governance API")

# Enable CORS for local development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory storage (replace with real database)
models_db = {}
events_db = []


# Pydantic models
class MonitoredCell(BaseModel):
    range: str
    type: str  # 'input' or 'output'
    addedAt: Optional[str] = None


class ModelRegistration(BaseModel):
    modelId: str
    name: str
    owner: str
    description: Optional[str] = None
    registeredAt: Optional[str] = None
    monitoredCells: Optional[List[MonitoredCell]] = []


class ExcelEvent(BaseModel):
    event: str
    modelId: Optional[str] = None
    timestamp: Optional[str] = None
    cell: Optional[str] = None
    value: Optional[Any] = None
    user: Optional[str] = None
    worksheet: Optional[str] = None
    range: Optional[str] = None


class EventBatch(BaseModel):
    events: List[ExcelEvent]


# Routes

@app.get("/")
def root():
    return {
        "service": "Domino Excel Governance API",
        "version": "1.0",
        "registered_models": len(models_db),
        "total_events": len(events_db)
    }


@app.get("/api/models/{model_id}")
def get_model(model_id: str):
    """Get model by ID - returns config if registered"""
    if model_id not in models_db:
        raise HTTPException(status_code=404, detail="Model not found")

    logger.info(f"Model retrieved: {model_id}")
    return models_db[model_id]


@app.post("/api/models")
def register_model(model: ModelRegistration):
    """Register a new Excel model"""
    if not model.registeredAt:
        model.registeredAt = datetime.utcnow().isoformat()

    if not model.monitoredCells:
        model.monitoredCells = []

    models_db[model.modelId] = model.dict()

    logger.info(f"✅ Model registered: {model.name} ({model.modelId}) by {model.owner}")

    # Store registration event
    events_db.append({
        "event": "model_registered",
        "modelId": model.modelId,
        "timestamp": model.registeredAt,
        "data": model.dict()
    })

    return models_db[model.modelId]


@app.patch("/api/models/{model_id}")
def update_model(model_id: str, updates: dict):
    """Update model configuration"""
    if model_id not in models_db:
        raise HTTPException(status_code=404, detail="Model not found")

    models_db[model_id].update(updates)
    logger.info(f"Model updated: {model_id}")

    return models_db[model_id]


@app.post("/api/models/{model_id}/cells")
def add_monitored_cell(model_id: str, cell: MonitoredCell):
    """Add a cell to the monitored list"""
    if model_id not in models_db:
        raise HTTPException(status_code=404, detail="Model not found")

    if not cell.addedAt:
        cell.addedAt = datetime.utcnow().isoformat()

    models_db[model_id]["monitoredCells"].append(cell.dict())

    logger.info(f"Cell added to {model_id}: {cell.range} ({cell.type})")

    return {"status": "added", "cell": cell.dict()}


@app.delete("/api/models/{model_id}/cells/{cell_range}")
def remove_monitored_cell(model_id: str, cell_range: str):
    """Remove a cell from monitoring"""
    if model_id not in models_db:
        raise HTTPException(status_code=404, detail="Model not found")

    cells = models_db[model_id]["monitoredCells"]
    models_db[model_id]["monitoredCells"] = [
        c for c in cells if c["range"] != cell_range
    ]

    logger.info(f"Cell removed from {model_id}: {cell_range}")

    return {"status": "removed"}


@app.post("/api/excel-events")
def receive_event(event: ExcelEvent):
    """Receive a single Excel event"""
    if not event.timestamp:
        event.timestamp = datetime.utcnow().isoformat()

    event_dict = event.dict()
    events_db.append(event_dict)

    # Log important events
    if event.event in ["cell_changed", "model_opened", "model_saved"]:
        logger.info(f"📊 {event.event}: {event.modelId} - {event.cell or ''}")

    # Store in real database here
    # e.g., postgres, mongodb, etc.

    return {"status": "recorded", "event_id": len(events_db)}


@app.post("/api/excel-events/batch")
def receive_event_batch(batch: EventBatch):
    """Receive multiple events (offline queue flush)"""
    count = len(batch.events)

    for event in batch.events:
        if not event.timestamp:
            event.timestamp = datetime.utcnow().isoformat()
        events_db.append(event.dict())

    logger.info(f"📦 Batch received: {count} events")

    return {"status": "recorded", "count": count}


@app.get("/api/models")
def get_all_models():
    """Get all registered models"""
    return list(models_db.values())


@app.get("/api/models/{model_id}/activity")
def get_model_activity(model_id: str, limit: int = 50):
    """Get recent activity for a model"""
    model_events = [
        e for e in events_db
        if e.get("modelId") == model_id
    ]

    # Return most recent first
    model_events.reverse()

    return model_events[:limit]


@app.get("/api/events")
def get_all_events(limit: int = 100):
    """Get all events (for debugging)"""
    return events_db[-limit:]


# Compliance monitoring example
@app.get("/api/compliance/inactive-models")
def get_inactive_models(days: int = 7):
    """Find models with no activity in X days"""
    from datetime import timedelta

    cutoff = datetime.utcnow() - timedelta(days=days)
    inactive = []

    for model_id, model in models_db.items():
        # Get last event
        model_events = [e for e in events_db if e.get("modelId") == model_id]

        if not model_events:
            inactive.append({
                "modelId": model_id,
                "name": model["name"],
                "owner": model["owner"],
                "reason": "No events recorded"
            })
        else:
            last_event = model_events[-1]
            last_timestamp = datetime.fromisoformat(last_event["timestamp"])

            if last_timestamp < cutoff:
                inactive.append({
                    "modelId": model_id,
                    "name": model["name"],
                    "owner": model["owner"],
                    "lastActivity": last_event["timestamp"],
                    "daysSinceActivity": (datetime.utcnow() - last_timestamp).days
                })

    return inactive


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=5000)
