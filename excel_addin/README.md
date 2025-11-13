# Excel Governance Add-in

Architecture-compliant implementation per [DEPLOYMENT.md](DEPLOYMENT.md).

## 🚀 Quick Start

# Create virtual environment (first time only)
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies (first time only)
pip install fastapi uvicorn

# Start backend
python backend.py
# Runs on http://localhost:5000
# Create virtual environment (first time only)
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies (first time only)
pip install fastapi uvicorn

# Start backend
python backend.py
# Runs on http://localhost:5000
# Create virtual environment (first time only)
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies (first time only)
pip install fastapi uvicorn

# Start backend
python backend.py
# Runs on http://localhost:5000
# Create virtual environment (first time only)
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies (first time only)
pip install fastapi uvicorn

# Start backend
python backend.py
# Runs on http://localhost:5000
# Create virtual environment (first time only)
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies (first time only)
pip install fastapi uvicorn

# Start backend
python backend.py
# Runs on http://localhost:5000
# Create virtual environment (first time only)
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies (first time only)
pip install fastapi uvicorn

# Start backend
python backend.py
# Runs on http://localhost:5000

### 2. Start Frontend
```bash
npm install
npm start
# Runs on https://localhost:3000
```

### 3. Test in Excel
- Open Excel (Windows or Mac)
- Load the add-in from `https://localhost:3000`
- Click "Register Model" to start

## 📁 Project Structure

```
excel_addin/
├── backend.py                    # FastAPI backend
├── database_schema.sql           # SQL DDL
├── src/
│   ├── commands/commands.js      # Background monitoring
│   ├── taskpane/                 # React UI
│   ├── utils/domino-api.js       # API client
│   └── types/model.ts            # TypeScript types
└── docs/
    ├── DEPLOYMENT.md             # Architecture spec
    └── MIGRATION_GUIDE.md        # Reference guide
```

## 🎯 Architecture

### API Endpoints
- `PUT /wb/upsert-model` - Create/update model (with versioning)
- `GET /wb/load-model` - Load model by ID
- `POST /wb/create-model-trace` - Log tracked range change

### Data Model
```typescript
WorkbookModel {
  model_name: string
  tracked_ranges: [{name: string, range: string}]
  model_id: string
  version: int
}
```

## 🗄️ Database

```bash
sqlcmd -S your-server -d your-database -i database_schema.sql
```

Creates:
- `dbo.workbook_model` - Model metadata
- `dbo.workbook_trace` - Trace logs

## 📖 Documentation

- **[DEPLOYMENT.md](DEPLOYMENT.md)** - Architecture specification
- **[MIGRATION_GUIDE.md](MIGRATION_GUIDE.md)** - Reference & testing
- **[ARCHITECTURE_COMPLIANCE.md](ARCHITECTURE_COMPLIANCE.md)** - Compliance report

## ✅ Status

**100% architecture compliant** - All components match the specification exactly.
