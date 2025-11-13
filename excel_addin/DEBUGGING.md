# Excel Add-in Debugging Guide

## Quick Start

### Starting the Development Environment

The easiest way to start both servers:

```bash
cd excel_addin
./start-dev.sh
```

Or manually:

```bash
# Terminal 1 - Backend Server
cd excel_addin
python3 backend-sqlite.py

# Terminal 2 - Frontend Dev Server
cd excel_addin
npm run start
```

## Debugging Register Model Issues

### Issue: "Register Model button just spins, nothing in console"

**Root Cause**: The backend server is not running, causing API calls to timeout.

**Solution**: Ensure both servers are running:

1. **Backend Server (Port 5000)**
   ```bash
   python3 backend-sqlite.py
   ```

   Verify it's running:
   ```bash
   curl http://localhost:5000/
   ```

   Expected response:
   ```json
   {
     "service": "Excel Add-In Backend",
     "status": "healthy",
     "database": "SQLite"
   }
   ```

2. **Frontend Server (Port 3000)**
   ```bash
   npm run start
   ```

   Should open https://localhost:3000

### Viewing Console Logs

The Excel add-in has a built-in **DebugPanel** that captures all console.log output:

1. Open the Register Model dialog
2. Look at the bottom of the window
3. You should see a dark panel with console logs
4. Click to expand/collapse
5. Use the "Clear" button to remove old logs

### Common Issues

#### 1. Backend Server Not Running

**Symptoms:**
- Register Model button shows infinite spinner
- No console output in DebugPanel
- Request timeout after 10 seconds

**Fix:**
```bash
# Check if backend is running
ps aux | grep backend-sqlite

# If not, start it
python3 backend-sqlite.py
```

#### 2. Port 5000 Already in Use

**Symptoms:**
- Backend fails to start with "Address already in use" error

**Fix:**
```bash
# Find process using port 5000
lsof -i :5000

# Kill it
kill -9 <PID>

# Restart backend
python3 backend-sqlite.py
```

#### 3. CORS Errors

**Symptoms:**
- Console shows "CORS policy" errors
- API calls fail with network errors

**Fix:**
The backend is already configured for CORS. If you still see errors:
- Ensure frontend is running on https://localhost:3000
- Check backend CORS settings in backend-sqlite.py (lines 44-50)

#### 4. Safari Developer Console (for deeper debugging)

If the DebugPanel isn't showing logs:

1. Open Safari
2. Go to Develop → Show Web Inspector
3. Find the Excel add-in window/iframe
4. Check the Console tab

**Tip**: You'll see TWO Safari processes:
- `index.html` - Main taskpane
- `commands.html` - Background monitoring script

### Enhanced Logging

The latest version includes detailed logging at every step:

**In RegisterModal.jsx:**
- Form submission details
- Validation status
- API call timing
- Success/failure messages

**In domino-api.js:**
- Full URL being called
- Request payload
- Response status and timing
- Detailed error messages

**Example console output:**
```
[RegisterModal] ========================================
[RegisterModal] FORM SUBMISSION STARTED
[RegisterModal] Model ID: excel_1234567890_abcdef
[RegisterModal] Model Name: My Revenue Model
[domino-api] ========================================
[domino-api] PUT /wb/upsert-model
[domino-api] API Base URL: http://localhost:5000
[domino-api] 🚀 Sending request...
[domino-api] 📡 Response received in 45ms
[domino-api] ✅ Upsert successful!
[RegisterModal] ✅ SUCCESS - Parent notified, dialog should close
```

### Database

The backend uses SQLite stored in `excel_addin.db`:

```bash
# View registered models
sqlite3 excel_addin.db "SELECT * FROM workbook_model;"

# View traces
sqlite3 excel_addin.db "SELECT * FROM workbook_trace ORDER BY timestamp DESC LIMIT 10;"

# Reset database
rm excel_addin.db
# Restart backend to recreate
```

### API Endpoints

Test the backend directly:

```bash
# Health check
curl http://localhost:5000/

# List all models
curl http://localhost:5000/wb/models

# Get model by ID
curl "http://localhost:5000/wb/load-model?model_id=test_123"

# Create/update model
curl -X PUT http://localhost:5000/wb/upsert-model \
  -H "Content-Type: application/json" \
  -d '{
    "model_name": "Test Model",
    "tracked_ranges": [{"name": "Revenue", "range": "A1:A10"}],
    "model_id": "test_123"
  }'
```

### Architecture Overview

```
Excel Add-in
├── Frontend (Port 3000)
│   ├── index.html → App.jsx (Main taskpane)
│   ├── commands.html → commands.js (Background script)
│   └── register.html → RegisterModal.jsx (Registration dialog)
│
└── Backend (Port 5000)
    ├── backend-sqlite.py (API server)
    └── excel_addin.db (SQLite database)
```

**Event Flow:**
1. User clicks "Register Model" button in Excel ribbon
2. `showRegisterModal()` in commands.js opens register.html dialog
3. RegisterModal.jsx renders the form with DebugPanel
4. User fills form and clicks "Save & Register"
5. `handleSubmit()` calls `upsertModel()` API
6. Backend saves to SQLite and returns versioned model
7. Success message sent to parent via `messageParent()`
8. commands.js receives message and reloads monitoring

### Troubleshooting Checklist

- [ ] Backend server is running on port 5000
- [ ] Frontend server is running on port 3000
- [ ] DebugPanel is visible at bottom of Register Modal
- [ ] Console logs appear in DebugPanel when clicking Register
- [ ] No CORS errors in Safari Web Inspector
- [ ] API responds to `curl http://localhost:5000/`

### Getting Help

If you're still stuck:

1. Copy the full console output from DebugPanel
2. Check backend.log for server errors
3. Test the API directly with curl
4. Check Safari Web Inspector console for both index.html and commands.html

## Files Modified for Better Debugging

- **RegisterModal.jsx**: Added extensive console.log statements
- **domino-api.js**: Added detailed request/response logging
- **backend-sqlite.py**: Created as SQLite replacement for PostgreSQL
- **start-dev.sh**: One-command startup script
