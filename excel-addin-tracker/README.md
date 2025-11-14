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

## Quick Start

### Prerequisites

- Node.js (v18 or later)
- Docker and Docker Compose
- Excel (Desktop or Mac)

### 1. Start Backend and Database

```bash
# Start PostgreSQL and FastAPI backend
docker-compose up -d

# Check logs
docker-compose logs -f backend
```

The backend will be available at `http://localhost:8000`

### 2. Install Frontend Dependencies

```bash
cd frontend
npm install
```

### 3. Generate SSL Certificates (Required for Excel Add-ins)

```bash
cd frontend
npx office-addin-dev-certs install
```

### 4. Start the Excel Add-in

```bash
# Start development server
npm run dev-server

# In another terminal, sideload the add-in
npm run start
```

Excel will open with the add-in loaded.

### 5. Use the Add-in

1. Click "Show Taskpane" in the ribbon
2. Enter a model name and click "Register Model"
3. Add tracked ranges (e.g., "Inputs", "Sheet1!A1:B10")
4. Make changes to tracked cells - they'll be logged automatically!

## API Endpoints

### PUT /wb/upsert-model
Create or update a workbook model

**Request:**
```json
{
  "model_name": "Financial Model Q4",
  "tracked_ranges": [
    {"name": "Inputs", "range": "Sheet1!A1:B10"}
  ],
  "model_id": "optional-existing-id",
  "version": 1
}
```

**Response:**
```json
{
  "model_name": "Financial Model Q4",
  "tracked_ranges": [...],
  "model_id": "generated-uuid",
  "version": 1
}
```

### GET /wb/load-model?model_id={id}
Load model metadata

### POST /wb/create-model-trace
Log a tracked range change

### GET /wb/model-traces/{model_id}
Get trace history for a model

## Database Schema

### workbook_model
- `model_id` (PK): Unique model identifier
- `model_name`: Human-readable name
- `tracked_ranges`: JSON array of tracked ranges
- `version`: Integer version number
- `created_at`, `updated_at`: Timestamps

### workbook_trace
- `trace_id` (PK): Auto-increment ID
- `model_id` (FK): Reference to model
- `timestamp`: When change occurred
- `tracked_range_name`: Which range changed
- `username`: Who made the change
- `value`: The new value (JSON)

## Development

### Backend Development

```bash
# Install dependencies
cd backend
pip install -r requirements.txt

# Run locally (without Docker)
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/excel_tracker \
  python main.py
```

### Frontend Development

```bash
cd frontend

# Build for production
npm run build

# Development with hot reload
npm run dev-server

# Validate manifest
npm run validate
```

### Database Access

```bash
# Connect to PostgreSQL
docker exec -it excel_tracker_db psql -U postgres -d excel_tracker

# Run queries
SELECT * FROM workbook_model;
SELECT * FROM workbook_trace ORDER BY timestamp DESC LIMIT 10;
```

## Background Tasks

The add-in runs background event listeners even when the taskpane is closed:

- **Workbook Load**: Automatically loads model metadata
- **Cell Changes**: Monitors tracked ranges and logs changes
- **Auto-save**: Stores model metadata in workbook custom properties

## Troubleshooting

### Add-in won't load
- Ensure SSL certificates are installed: `npx office-addin-dev-certs install`
- Check that dev server is running on `https://localhost:3000`
- Clear Office cache: Delete `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (Mac)

### Backend connection errors
- Verify backend is running: `curl http://localhost:8000`
- Check CORS settings in `main.py`
- Ensure Excel allows connections to localhost

### Database issues
- Check PostgreSQL is running: `docker-compose ps`
- View logs: `docker-compose logs postgres`
- Reset database: `docker-compose down -v && docker-compose up -d`

## Production Deployment

For production:

1. Update manifest.xml with production URLs
2. Use proper SSL certificates (not self-signed)
3. Set secure DATABASE_URL environment variable
4. Configure proper CORS origins
5. Deploy backend to cloud service (e.g., AWS, Azure)
6. Host add-in files on HTTPS server
7. Publish to Microsoft AppSource (optional)

## License

MIT

## Support

For issues or questions, please open an issue on GitHub.
