# Testing Guide

This guide will help you test the Excel Model Tracker add-in end-to-end.

## Prerequisites

- ✅ Docker Desktop installed and running
- ✅ Node.js 18+ installed
- ✅ Excel Desktop (Mac or Windows)
- ✅ Terminal/Command Prompt

## Step 1: Start Backend Services

```bash
cd excel-addin-tracker

# Start PostgreSQL and FastAPI backend
docker compose up -d

# Verify services are running
docker compose ps

# Check backend health
curl http://localhost:8000
# Should return: {"status":"ok","service":"Domino Spreadsheet Backend"}

# Check backend logs
docker compose logs -f backend
```

## Step 2: Setup Frontend

```bash
cd frontend

# Install dependencies (first time only)
npm install

# Install SSL certificates for local development (required!)
npx office-addin-dev-certs install --machine

# Verify installation
npm run validate
```

## Step 3: Start Excel Add-in

Open TWO terminal windows:

**Terminal 1 - Start Dev Server:**
```bash
cd frontend
npm run dev-server
```

Wait for "webpack compiled successfully" message. The server will run at https://localhost:3000

**Terminal 2 - Sideload Add-in:**
```bash
cd frontend
npm run start
```

This will:
- Open Excel
- Sideload the add-in
- Show the taskpane automatically

## Step 4: Test Basic Workflow

### 4.1 Register a New Model

1. In Excel, you should see "Show Taskpane" button in the Home ribbon
2. Click it to open the taskpane
3. In the taskpane:
   - Enter model name: "Test Financial Model"
   - Click "Register Model"
4. You should see success message with a Model ID
5. Verify in backend:
   ```bash
   docker compose exec postgres psql -U postgres -d excel_tracker -c "SELECT * FROM workbook_model;"
   ```

### 4.2 Add Tracked Ranges

1. In Excel, create some data:
   - Sheet1, A1:B5 - Input parameters
   - Sheet1, D1:E5 - Output calculations

2. In the taskpane:
   - Range Name: "Inputs"
   - Range Address: "Sheet1!A1:B5"
   - Click "Add Tracked Range"

3. Add another range:
   - Range Name: "Outputs"
   - Range Address: "Sheet1!D1:E5"
   - Click "Add Tracked Range"

4. Click "Update Model"

5. Verify ranges are saved:
   ```bash
   docker compose exec postgres psql -U postgres -d excel_tracker -c "SELECT model_name, tracked_ranges FROM workbook_model;"
   ```

### 4.3 Test Change Tracking

1. In Excel, modify a cell in the tracked range (e.g., A1)
2. Wait a few seconds
3. Check traces were created:
   ```bash
   docker compose exec postgres psql -U postgres -d excel_tracker -c "SELECT * FROM workbook_trace ORDER BY timestamp DESC LIMIT 5;"
   ```

4. You should see trace entries with:
   - model_id
   - tracked_range_name
   - timestamp
   - value (JSON)

### 4.4 Test Model Loading

1. Close Excel
2. Reopen the same workbook
3. In the taskpane, click "Load Model from Workbook"
4. The model info should appear showing:
   - Model name
   - Model ID
   - Version
   - Tracked ranges

## Step 5: Test API Endpoints Directly

### Test Model Creation
```bash
curl -X PUT http://localhost:8000/wb/upsert-model \
  -H "Content-Type: application/json" \
  -d '{
    "model_name": "API Test Model",
    "tracked_ranges": [
      {"name": "TestRange", "range": "Sheet1!A1:A10"}
    ]
  }'
```

### Test Model Loading
```bash
# Use model_id from previous response
curl "http://localhost:8000/wb/load-model?model_id=YOUR_MODEL_ID"
```

### Test Trace Creation
```bash
curl -X POST http://localhost:8000/wb/create-model-trace \
  -H "Content-Type: application/json" \
  -d '{
    "model_id": "YOUR_MODEL_ID",
    "timestamp": "2024-01-15T10:30:00Z",
    "tracked_range_name": "TestRange",
    "username": "testuser",
    "value": [[1, 2, 3]]
  }'
```

### Get Trace History
```bash
curl "http://localhost:8000/wb/model-traces/YOUR_MODEL_ID"
```

## Step 6: Test on Mac

The add-in should work identically on Mac Excel:

1. Follow the same steps as above
2. On Mac, certs are installed in Keychain
3. You may need to trust the localhost certificate manually:
   - Open Keychain Access
   - Find "localhost" certificate
   - Right-click → Get Info → Trust → Always Trust

## Common Issues

### Add-in won't load
- Verify dev server is running at https://localhost:3000
- Check browser console for errors (open taskpane in browser)
- Clear Office cache:
  - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
  - Mac: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`

### SSL Certificate errors
- Re-run: `npx office-addin-dev-certs install --machine`
- On Mac, manually trust certificate in Keychain

### Backend connection errors
- Verify backend is accessible: `curl http://localhost:8000`
- Check CORS headers in browser console
- Verify no firewall blocking localhost:8000

### Database errors
- Check PostgreSQL is running: `docker compose ps`
- View logs: `docker compose logs postgres`
- Restart: `docker compose restart postgres`

### Changes not being tracked
- Check browser console for errors
- Verify tracked ranges are correct (Sheet1!A1:B5 format)
- Make sure you clicked "Update Model" after adding ranges
- Check backend logs: `docker compose logs backend`

## Manual Database Inspection

```bash
# Connect to PostgreSQL
docker compose exec postgres psql -U postgres -d excel_tracker

# List all models
SELECT * FROM workbook_model;

# List all traces
SELECT * FROM workbook_trace ORDER BY timestamp DESC LIMIT 20;

# Get traces for specific model
SELECT * FROM workbook_trace WHERE model_id = 'your-model-id';

# Exit
\q
```

## Performance Testing

### Test with Many Tracked Ranges
1. Add 10+ tracked ranges
2. Verify all are monitored
3. Make changes to multiple ranges
4. Check trace creation performance

### Test with Large Ranges
1. Track a range like "Sheet1!A1:Z1000"
2. Make changes
3. Verify traces are created efficiently

### Test Background Operation
1. Close the taskpane (X button)
2. Make changes to tracked cells
3. Traces should still be created
4. Verify in database

## Cleanup

```bash
# Stop services
docker compose down

# Remove all data (warning: deletes database!)
docker compose down -v

# Remove node_modules
rm -rf frontend/node_modules

# Clear Office cache
# Windows: Delete %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
# Mac: rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
```

## Next Steps

After successful testing:
1. Review the architecture in README.md
2. Customize the UI in `frontend/src/taskpane/taskpane.html`
3. Add authentication (see backend/main.py)
4. Deploy to production (see README.md Production Deployment section)
5. Create proper icons (see frontend/assets/README.md)

## Automated Testing

For CI/CD, you can create automated tests:

```bash
# Backend tests (add pytest to requirements.txt)
cd backend
pytest test_api.py

# Frontend tests (add jest to package.json)
cd frontend
npm test
```

Example backend test (create `backend/test_api.py`):

```python
from fastapi.testclient import TestClient
from main import app

client = TestClient(app)

def test_create_model():
    response = client.put("/wb/upsert-model", json={
        "model_name": "Test Model",
        "tracked_ranges": [{"name": "Test", "range": "A1:B2"}]
    })
    assert response.status_code == 200
    assert "model_id" in response.json()
```
