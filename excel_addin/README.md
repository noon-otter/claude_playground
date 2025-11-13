# Excel Model Tracker - Add-In POC

A standalone Excel Add-In that tracks changes to cell ranges and stores them in PostgreSQL.

## ⚠️ Python Version Requirement

**This project requires Python 3.11 or 3.12.**

Other versions are not supported and will be rejected by the startup script.

### Quick Install

**macOS:**
```bash
brew install python@3.12
```

**Ubuntu/Debian:**
```bash
sudo apt install python3.12 python3.12-venv
```

## Quick Start

```bash
cd excel_addin

# The script will automatically detect and use Python 3.11 or 3.12
./dev.sh
```

That's it! The script will:
1. ✅ Verify you have Python 3.11 or 3.12
2. ✅ Start PostgreSQL database (Docker)
3. ✅ Create Python virtual environment
4. ✅ Install all dependencies
5. ✅ Start backend on http://localhost:5000
6. ✅ Start frontend on https://localhost:3000

## What It Does

- **Register workbooks** with a unique model ID
- **Define tracked ranges** (e.g., "Revenue: Sheet1!A1:D10")
- **Monitor changes** to those ranges in real-time
- **Log all changes** to PostgreSQL with:
  - Timestamp
  - User (from Office 365)
  - Range name
  - Cell value

## Documentation

- **[QUICKSTART.md](QUICKSTART.md)** - Detailed setup and usage guide
- **[PYTHON_VERSION.md](PYTHON_VERSION.md)** - Python version troubleshooting
- **[ARCHITECTURE.md](ARCHITECTURE.md)** - System architecture and API specs

## Architecture

```
Excel Add-In (JavaScript)
    ↓
FastAPI Backend (Python)
    ↓
PostgreSQL Database
```

## Project Structure

```
excel_addin/
├── backend.py              # FastAPI backend with PostgreSQL
├── docker-compose.yml      # PostgreSQL container
├── init-db.sql            # Database schema
├── dev.sh                 # Unified startup script
├── requirements.txt       # Python dependencies
├── manifest.xml           # Excel Add-in manifest
├── src/
│   ├── commands/          # Background scripts (Office.js)
│   ├── taskpane/          # React UI components
│   └── utils/             # API client, model ID management
└── public/                # Static assets (icons)
```

## Development

### View Logs
```bash
tail -f backend.log    # Backend API logs
tail -f frontend.log   # React dev server logs
```

### Access Database
```bash
docker compose exec postgres psql -U excel_user -d excel_addin

# Query models
SELECT * FROM workbook_model;

# Query traces
SELECT * FROM workbook_trace ORDER BY timestamp DESC LIMIT 10;
```

### Stop Services
```bash
# Press Ctrl+C to stop dev.sh

# Stop database
docker compose down

# Clean slate (deletes all data)
docker compose down -v
```

## Troubleshooting

### "Python 3.11 or 3.12 required"
The dev.sh script detected the wrong Python version. Install Python 3.12:
```bash
brew install python@3.12
rm -rf venv
./dev.sh
```

### "Port already in use"
```bash
lsof -ti:3000 | xargs kill -9   # Kill frontend
lsof -ti:5000 | xargs kill -9   # Kill backend
```

### Add-in not loading in Excel
1. Clear Excel cache
2. Restart Excel
3. Re-run: `bash install_addin_locally.sh`

## Support

For issues, check:
- `backend.log` - Backend API logs
- `frontend.log` - React logs
- Browser console (F12) - Frontend errors
- Excel add-in diagnostics

---

**Built for Excel model governance and change tracking**
