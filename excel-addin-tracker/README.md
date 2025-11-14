# Excel Model Tracker

A complete Excel Add-in system for tracking workbooks as versioned models with change tracing.

## Architecture

This project implements the Excel Add-In Model Tracking Architecture with:

- **Frontend**: Excel Add-in (TypeScript, Office.js)
- **Backend**: FastAPI REST API (Python)
- **Database**: PostgreSQL

## Features

- ✅ Register Excel workbooks as versioned models
- ✅ Track specific cell ranges for changes
- ✅ Automatic change logging (traces)
- ✅ Model versioning on updates
- ✅ Load model metadata from workbook
- ✅ Works on Mac and Windows Excel
- ✅ Background tasks run even when taskpane is closed

## Project Structure

```
excel-addin-tracker/
├── frontend/              # Excel Add-in
│   ├── src/
│   │   ├── taskpane/     # Main UI
│   │   └── commands/     # Add-in commands
│   ├── assets/           # Icons and resources
│   ├── manifest.xml      # Add-in manifest
│   └── package.json
├── backend/              # FastAPI backend
│   ├── main.py          # API endpoints
│   ├── requirements.txt
│   └── Dockerfile
├── database/            # Database setup
│   └── init.sql        # Schema initialization
└── docker-compose.yml  # Container orchestration
```

---

## 🚀 Local Development Setup (Mac & Windows)

### Prerequisites

**Required:**
- **Node.js** v18 or later ([Download](https://nodejs.org/))
- **Docker Desktop** ([Download Mac](https://www.docker.com/products/docker-desktop/) | [Download Windows](https://www.docker.com/products/docker-desktop/))
- **Excel Desktop** (Microsoft 365 or Office 2019+)

**Verify installations:**
```bash
node --version          # Should be v18+
npm --version          # Should be 9+
docker --version       # Should be 20+
docker compose version # Should be 2+
```

---

## Step-by-Step Setup

### 1️⃣ Start Backend Services

Open a terminal in the project root:

```bash
cd excel-addin-tracker

# Start PostgreSQL and FastAPI backend
docker compose up -d

# Verify services are running
docker compose ps
```

You should see:
```
NAME                     STATUS
excel_tracker_db         Up (healthy)
excel_tracker_backend    Up
```

**Verify backend is responding:**
```bash
curl http://localhost:8000
```

Expected response: `{"status":"ok","service":"Domino Spreadsheet Backend"}`

**View backend logs (optional):**
```bash
docker compose logs -f backend
```

Press `Ctrl+C` to exit logs.

---

### 2️⃣ Setup Frontend (Excel Add-in)

Open a **new terminal** window:

```bash
cd excel-addin-tracker/frontend

# Install dependencies (only needed once)
npm install
```

This will take 2-3 minutes the first time.

**Expected output:** `added XXX packages` with no errors.

---

### 3️⃣ Install SSL Certificates (REQUIRED)

Office Add-ins require HTTPS with trusted certificates:

```bash
# Still in excel-addin-tracker/frontend directory
npx office-addin-dev-certs install
```

**On Mac:** You'll be prompted for your password to install the certificate in the system keychain.

**On Windows:** The certificate will be installed automatically.

**Troubleshooting:**
- If you get "command not found", make sure you ran `npm install` first
- On Mac, if prompted, enter your password to allow certificate installation
- You only need to do this once per machine

**Verify certificate installation:**
```bash
npm run validate
```

Expected: ✅ Manifest is valid

---

### 4️⃣ Create Placeholder Icons

The add-in needs icon files. Create simple placeholders:

```bash
# Still in excel-addin-tracker/frontend directory
cd assets

# On Mac (if you have ImageMagick):
convert -size 16x16 xc:#0078D4 icon-16.png
convert -size 32x32 xc:#0078D4 icon-32.png
convert -size 64x64 xc:#0078D4 icon-64.png
convert -size 80x80 xc:#0078D4 icon-80.png

# If you don't have ImageMagick, download any PNG icon and name them:
# icon-16.png, icon-32.png, icon-64.png, icon-80.png
# Or use online icon generators
```

**Quick alternative:** The add-in will work even with placeholder files, but you'll see broken icons in Excel.

---

### 5️⃣ Start Development Server

**Terminal 1** (keep this running):

```bash
cd excel-addin-tracker/frontend
npm run dev-server
```

Wait for this message:
```
webpack compiled successfully
```

The dev server will run at **https://localhost:3000**

**⚠️ Do NOT close this terminal.** It needs to stay running.

---

### 6️⃣ Sideload Add-in to Excel

Open a **new terminal** (Terminal 2):

```bash
cd excel-addin-tracker/frontend
npm run start
```

This will:
1. ✅ Open Excel
2. ✅ Load a blank workbook
3. ✅ Sideload the add-in
4. ✅ Show the taskpane automatically (or via Home ribbon)

**On Mac:** If Excel doesn't open automatically:
1. Open Excel manually
2. Go to **Insert** → **Add-ins** → **My Add-ins**
3. Click **Excel Model Tracker**

**On Windows:** Excel should open automatically with the add-in loaded.

---

## 🎯 Using the Add-in

### First Time Setup

1. **Open the taskpane:**
   - Look for "Show Taskpane" button in the **Home** ribbon
   - Click it to open the Excel Model Tracker panel

2. **Register a new model:**
   - Enter a model name (e.g., "Test Financial Model")
   - Click **Register Model**
   - You should see: "Model registered successfully! ID: xxxxx"

3. **Add tracked ranges:**
   - First, create some test data in Excel:
     - Sheet1, cells A1:B5 - Enter some numbers
     - Sheet1, cells D1:E5 - Enter some formulas

   - In the taskpane:
     - Range Name: `Inputs`
     - Range Address: `Sheet1!A1:B5`
     - Click **Add Tracked Range**

   - Add another:
     - Range Name: `Outputs`
     - Range Address: `Sheet1!D1:E5`
     - Click **Add Tracked Range**

4. **Update the model:**
   - Click **Update Model**
   - Version should increment to 2

5. **Test change tracking:**
   - Modify a cell in A1:B5 range
   - Changes are automatically logged to the database!

---

## 🐛 Debugging

### Debug Frontend (Excel Add-in)

**Option 1: Browser DevTools (Mac & Windows)**

Right-click anywhere in the taskpane → **Inspect** or **Inspect Element**

This opens Chrome/Edge DevTools where you can:
- View console logs
- Set breakpoints in TypeScript
- Inspect network requests
- Check for JavaScript errors

**Option 2: VS Code Debugging**

Add to `.vscode/launch.json`:

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "type": "office-addin",
      "request": "attach",
      "name": "Attach to Office Add-in",
      "port": 9222,
      "timeout": 30000,
      "webRoot": "${workspaceFolder}/frontend",
      "preLaunchTask": "Start Dev Server"
    }
  ]
}
```

**Mac-Specific Debugging:**

To see Office.js errors on Mac:
```bash
# Enable debugging in Excel for Mac
defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
```

Then right-click taskpane → **Inspect Element** to open Safari Web Inspector.

### Debug Backend (FastAPI)

**View logs in real-time:**
```bash
docker compose logs -f backend
```

You'll see all API requests:
```
INFO:     127.0.0.1:52050 - "PUT /wb/upsert-model HTTP/1.1" 200 OK
INFO:     127.0.0.1:52051 - "GET /wb/load-model?model_id=xxx HTTP/1.1" 200 OK
```

**Run backend locally (outside Docker) for easier debugging:**

```bash
# Terminal 1 - Start just the database
docker compose up -d postgres

# Terminal 2 - Run backend locally
cd backend
pip install -r requirements.txt

# Set database URL and run
export DATABASE_URL=postgresql://postgres:postgres@localhost:5432/excel_tracker
python main.py
```

Now you can:
- Add print statements
- Use Python debugger (pdb)
- See immediate code changes

**VS Code Backend Debugging:**

Add to `.vscode/launch.json`:

```json
{
  "type": "python",
  "request": "launch",
  "name": "Backend API",
  "module": "uvicorn",
  "args": ["main:app", "--reload", "--host", "0.0.0.0", "--port", "8000"],
  "cwd": "${workspaceFolder}/excel-addin-tracker/backend",
  "env": {
    "DATABASE_URL": "postgresql://postgres:postgres@localhost:5432/excel_tracker"
  }
}
```

### Debug Database

**Connect to PostgreSQL:**
```bash
docker compose exec postgres psql -U postgres -d excel_tracker
```

**Useful queries:**
```sql
-- View all models
SELECT * FROM workbook_model;

-- View recent traces
SELECT * FROM workbook_trace ORDER BY timestamp DESC LIMIT 10;

-- View traces for specific model
SELECT * FROM workbook_trace WHERE model_id = 'your-model-id';

-- Exit
\q
```

**Using a GUI (TablePlus, DBeaver, etc.):**
- Host: `localhost`
- Port: `5432`
- User: `postgres`
- Password: `postgres`
- Database: `excel_tracker`

---

## 🔧 Development Workflow

### Making Changes to Frontend

1. Edit files in `frontend/src/taskpane/`
2. Webpack will automatically rebuild
3. Refresh Excel taskpane (close and reopen)

**Hot reload doesn't work in Office Add-ins**, so you need to:
- Close the taskpane
- Reopen it via the ribbon button

### Making Changes to Backend

**If running in Docker:**
```bash
# Restart backend after code changes
docker compose restart backend
```

**If running locally:**
- Backend auto-reloads with uvicorn `--reload` flag

### Making Changes to Database Schema

```bash
# Stop services
docker compose down

# Edit database/init.sql

# Restart (this recreates the database)
docker compose down -v  # -v removes volumes
docker compose up -d
```

**⚠️ Warning:** This deletes all data. For production, use migrations (Alembic).

---

## 🧪 Testing the Complete Workflow

### Test 1: Model Registration

```bash
# API test
curl -X PUT http://localhost:8000/wb/upsert-model \
  -H "Content-Type: application/json" \
  -d '{
    "model_name": "Test Model",
    "tracked_ranges": [{"name": "TestRange", "range": "Sheet1!A1:A10"}]
  }'
```

Expected: JSON response with `model_id` and `version: 1`

### Test 2: Model Loading

```bash
# Replace YOUR_MODEL_ID with actual ID from Test 1
curl "http://localhost:8000/wb/load-model?model_id=YOUR_MODEL_ID"
```

Expected: Same JSON as registration response

### Test 3: Trace Creation

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

Expected: `{"success": true}`

### Test 4: View Traces

```bash
curl "http://localhost:8000/wb/model-traces/YOUR_MODEL_ID"
```

Expected: JSON with array of traces

---

## ⚠️ Troubleshooting

### Excel Add-in Issues

**Problem:** Add-in doesn't appear in Excel
- ✅ Verify dev server is running at https://localhost:3000
- ✅ Check browser console for errors (right-click → Inspect)
- ✅ Clear Office cache (see below)
- ✅ Re-run: `npx office-addin-dev-certs install`

**Problem:** "This add-in is no longer available"
- ✅ Dev server stopped - restart with `npm run dev-server`
- ✅ Certificate expired - reinstall certs

**Problem:** Taskpane shows blank/white screen
- ✅ Open browser DevTools (right-click → Inspect)
- ✅ Check Console tab for errors
- ✅ Verify https://localhost:3000/taskpane.html loads in browser

**Clear Office Cache (Mac):**
```bash
# Remove cached add-ins
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef

# Restart Excel
```

**Clear Office Cache (Windows):**
```cmd
# Delete this folder:
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef

# Restart Excel
```

### Backend Connection Issues

**Problem:** "Failed to fetch" errors in console
- ✅ Verify backend is running: `curl http://localhost:8000`
- ✅ Check CORS settings in `backend/main.py`
- ✅ Ensure no firewall blocking port 8000

**Problem:** Backend won't start
```bash
# Check logs
docker compose logs backend

# Common issues:
# 1. Port 8000 already in use
lsof -i :8000  # Mac/Linux
netstat -ano | findstr :8000  # Windows

# 2. Database not ready
docker compose up -d postgres
sleep 5
docker compose up -d backend
```

### Database Issues

**Problem:** "could not connect to server"
```bash
# Check if PostgreSQL is running
docker compose ps

# Restart PostgreSQL
docker compose restart postgres

# Check logs
docker compose logs postgres
```

**Problem:** "relation does not exist"
```bash
# Database tables not created - reinitialize
docker compose down -v
docker compose up -d
```

### SSL Certificate Issues (Mac-Specific)

**Problem:** "NET::ERR_CERT_AUTHORITY_INVALID"

```bash
# 1. Reinstall certificates
cd frontend
npx office-addin-dev-certs install

# 2. Trust certificate in Keychain
# Open Keychain Access app
# Search for "localhost"
# Double-click → Trust → "Always Trust"

# 3. Restart Excel
```

**Problem:** Password prompt loops

```bash
# Use machine flag
npx office-addin-dev-certs install --machine
```

### Network Issues

**Problem:** Cannot access backend from Excel

On Mac, if localhost doesn't work, try:
```bash
# Get your local IP
ifconfig | grep "inet " | grep -v 127.0.0.1

# Use IP instead of localhost in API calls
# Update API_BASE_URL in taskpane.ts:
# const API_BASE_URL = "http://192.168.1.x:8000";
```

---

## 📊 Monitoring & Logs

### Watch All Logs

```bash
# In project root
docker compose logs -f
```

Shows logs from both backend and database.

### Watch Backend Only

```bash
docker compose logs -f backend
```

### Watch Database Only

```bash
docker compose logs -f postgres
```

### Check Add-in Console

1. Right-click in taskpane → **Inspect**
2. Go to **Console** tab
3. Watch for:
   - API requests
   - Office.js errors
   - JavaScript errors

---

## 🎨 Customization

### Update Icons

Replace placeholder icons in `frontend/assets/`:
- `icon-16.png` (16x16 px)
- `icon-32.png` (32x32 px)
- `icon-64.png` (64x64 px)
- `icon-80.png` (80x80 px)

Use transparent PNG format. Recommended: Use your company logo or a chart/spreadsheet icon.

### Modify UI

Edit `frontend/src/taskpane/taskpane.html` and `taskpane.ts`

After changes:
1. Webpack rebuilds automatically
2. Close and reopen taskpane in Excel

### Add API Endpoints

Edit `backend/main.py`:

```python
@app.get("/wb/my-endpoint")
async def my_endpoint():
    return {"message": "Hello"}
```

Restart backend:
```bash
docker compose restart backend
```

---

## 🚢 Production Deployment

See `TESTING.md` for comprehensive deployment guide.

**Quick checklist:**

1. ✅ Update `manifest.xml` with production URLs
2. ✅ Use proper SSL certificates (Let's Encrypt, etc.)
3. ✅ Set secure `DATABASE_URL` environment variable
4. ✅ Configure production CORS origins in `main.py`
5. ✅ Deploy backend to cloud (AWS, Azure, Heroku)
6. ✅ Host frontend on HTTPS server (Azure Storage, S3+CloudFront)
7. ✅ Use managed PostgreSQL (RDS, Azure Database, etc.)
8. ✅ Optional: Publish to Microsoft AppSource

---

## 📚 API Documentation

### Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| PUT | `/wb/upsert-model` | Create or update model |
| GET | `/wb/load-model?model_id={id}` | Load model metadata |
| POST | `/wb/create-model-trace` | Log single change |
| POST | `/wb/create-model-trace-batch` | Log multiple changes |
| GET | `/wb/model-traces/{model_id}` | Get trace history |

### Interactive API Docs

Visit http://localhost:8000/docs for auto-generated Swagger UI.

---

## 🧰 Useful Commands

```bash
# Start everything
docker compose up -d && cd frontend && npm run dev-server

# Stop everything
docker compose down

# Restart backend only
docker compose restart backend

# View database
docker compose exec postgres psql -U postgres -d excel_tracker

# Clear all data (warning: destructive!)
docker compose down -v

# Rebuild backend after Dockerfile changes
docker compose up -d --build backend

# Check what's running
docker compose ps
lsof -i :3000  # Check if port 3000 is in use
lsof -i :8000  # Check if port 8000 is in use

# Validate manifest
cd frontend && npm run validate
```

---

## 📖 Additional Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Office.js API Reference](https://docs.microsoft.com/en-us/javascript/api/excel)
- [FastAPI Documentation](https://fastapi.tiangolo.com/)
- [PostgreSQL Documentation](https://www.postgresql.org/docs/)

---

## 🤝 Contributing

1. Create a feature branch
2. Make changes
3. Test locally
4. Submit pull request

---

## 📝 License

MIT

---

## 💬 Support

For issues or questions:
- Open an issue on GitHub
- Check TESTING.md for detailed testing guide
- Review troubleshooting section above
