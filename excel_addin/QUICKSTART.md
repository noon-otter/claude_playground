# Excel Add-In POC - Quick Start Guide

## Overview

This Excel Add-In tracks changes to specified cell ranges in your workbook and stores them in a PostgreSQL database. It's a standalone POC with no external dependencies.

## Prerequisites

- Docker & Docker Compose
- **Python 3.11 or 3.12** (required for stability)
- Node.js 16+
- Excel (Desktop or Microsoft 365)

### Installing Python 3.12 (Recommended)

**macOS:**
```bash
brew install python@3.12
```

**Ubuntu/Debian:**
```bash
sudo apt install python3.12 python3.12-venv
```

**Windows:**
Download from [python.org](https://www.python.org/downloads/) and select version 3.12.x

## 🚀 One-Command Startup

```bash
cd excel_addin
./dev.sh
```

This single script will:
1. Start PostgreSQL database (Docker)
2. Create database tables automatically
3. Start Python backend (FastAPI on port 5000)
4. Start React frontend (Vite on port 3000)

## Manual Setup (Alternative)

If you prefer to run services individually:

### 1. Start Database

```bash
docker-compose up -d
```

### 2. Setup Python Backend

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
python backend.py
```

### 3. Start Frontend

```bash
npm install
npm start
```

## Using the Add-In

### 1. Load the Add-In in Excel

**Mac:**
```bash
bash install_addin_locally.sh
```

**Windows:**
- Copy `manifest.xml` to: `%APPDATA%\Microsoft\Office\Addins\`
- Open Excel, go to Insert > My Add-ins > Shared Folder

### 2. Register Your Workbook

1. Open Excel with your workbook
2. You'll see "Model Tracker" in the Home ribbon
3. Click **Register Model**
4. Enter:
   - Model Name (auto-filled from workbook name)
   - Tracked Ranges (click "Add Range" to define cell ranges to monitor)
     - Example: `Revenue` → `Sheet1!A1:D10`
     - Example: `Costs` → `Sheet1!E1:E20`
5. Click **Save & Register**

### 3. Monitor Changes

- Once registered, any changes to your tracked ranges are automatically logged
- Changes include:
  - Cell value
  - Timestamp
  - Username (from Office 365)
  - Range name

### 4. View Dashboard

- Click **Dashboard** button to see:
  - Model information
  - All tracked ranges
  - Recent change history

## Architecture

```
┌─────────────────┐
│  Excel Add-In   │
│  (JavaScript)   │
└────────┬────────┘
         │
         ├─ Register Model (PUT /wb/upsert-model)
         │
         ├─ Load Model (GET /wb/load-model)
         │
         └─ Track Changes (POST /wb/create-model-trace)
         │
┌────────▼────────┐
│  FastAPI Backend│
│  (Python)       │
│  Port 5000      │
└────────┬────────┘
         │
┌────────▼────────┐
│   PostgreSQL    │
│   (Docker)      │
│   Port 5432     │
│                 │
│  Tables:        │
│  • workbook_model
│  • workbook_trace
└─────────────────┘
```

## Database Schema

### workbook_model
```sql
model_id         VARCHAR(255)  PRIMARY KEY
model_name       VARCHAR(500)
version          INTEGER       (auto-incremented)
tracked_ranges   JSONB         [{name, range}]
created_at       TIMESTAMP
updated_at       TIMESTAMP
```

### workbook_trace
```sql
trace_id            BIGSERIAL  PRIMARY KEY
model_id            VARCHAR(255)  REFERENCES workbook_model
timestamp           TIMESTAMP
tracked_range_name  VARCHAR(255)
username            VARCHAR(500)
value               TEXT
created_at          TIMESTAMP
```

## API Endpoints

### `PUT /wb/upsert-model`
Create or update a model.

**Request:**
```json
{
  "model_name": "Revenue Forecast 2025",
  "tracked_ranges": [
    {"name": "Revenue", "range": "Sheet1!A1:D10"},
    {"name": "Costs", "range": "Sheet1!E1:E20"}
  ],
  "model_id": "excel_abc123_xyz789"  // optional
}
```

**Response:**
```json
{
  "model_id": "excel_abc123_xyz789",
  "model_name": "Revenue Forecast 2025",
  "version": 1,
  "tracked_ranges": [...]
}
```

### `GET /wb/load-model?model_id=<id>`
Load model metadata.

### `POST /wb/create-model-trace`
Log a cell change.

**Request:**
```json
{
  "model_id": "excel_abc123_xyz789",
  "timestamp": "2025-01-15T10:30:00Z",
  "tracked_range_name": "Revenue",
  "username": "user@company.com",
  "value": 12345
}
```

## Development

### View Logs

```bash
# Backend logs
tail -f backend.log

# Frontend logs
tail -f frontend.log

# Database logs
docker-compose logs -f postgres
```

### Access Database

```bash
docker-compose exec postgres psql -U excel_user -d excel_addin

# List models
SELECT * FROM workbook_model;

# List traces
SELECT * FROM workbook_trace ORDER BY timestamp DESC LIMIT 10;
```

### Stop Services

```bash
# Stop dev.sh: Press Ctrl+C

# Stop database
docker-compose down

# Remove database data (clean slate)
docker-compose down -v
```

## Troubleshooting

### Port Already in Use

```bash
# Kill process on port 3000
lsof -ti:3000 | xargs kill -9

# Kill process on port 5000
lsof -ti:5000 | xargs kill -9
```

### Database Connection Failed

```bash
# Check if PostgreSQL is running
docker-compose ps

# Restart database
docker-compose restart postgres

# Check logs
docker-compose logs postgres
```

### Add-In Not Showing in Excel

1. Clear Office cache: `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/`
2. Restart Excel
3. Re-run: `bash install_addin_locally.sh`

### HTTPS Certificate Errors

The frontend uses HTTPS for Office Add-in compatibility. Certs are auto-generated in `~/.office-addin-dev-certs/`.

If you see cert errors:
1. Trust the cert in your system
2. Or restart with: `npm start`

## Next Steps

Once this POC works locally, you can:

1. **Deploy to Production:**
   - Host backend on a server (AWS, Azure, etc.)
   - Use production PostgreSQL (not Docker)
   - Update `manifest.xml` URLs

2. **Add Features:**
   - Email notifications on specific changes
   - Audit reports
   - Data validation rules
   - Integration with other systems

3. **Security:**
   - Add authentication to backend APIs
   - Use HTTPS for backend
   - Implement role-based access control

## Support

For issues or questions, check:
- Backend logs: `backend.log`
- Frontend logs: `frontend.log`
- Browser console (F12)
- Excel add-in diagnostics

---

Built with ❤️ for Excel model governance
