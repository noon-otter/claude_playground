# Domino Excel Governance Add-in

A modern Excel Office JS add-in that provides always-on governance tracking and monitoring for Excel models. Built with React and Office.js.

## Overview

This add-in enables:
- **Always-on monitoring** - Runs in the background, tracks all user interactions
- **Persistent model IDs** - Unique identifiers that survive Save As, renames, and versions
- **Live event streaming** - All changes sent to Domino API in real-time
- **Cell-level tracking** - Mark specific cells as inputs/outputs for focused monitoring
- **Ribbon integration** - One-click actions directly in Excel
- **Offline queue** - Events queued locally when Domino API is unreachable

## Architecture

```
manifest.xml (defines add-in structure)
  ↓
commands.js (ALWAYS RUNNING - background script)
  ├─ Monitors all worksheet events
  ├─ Checks registration on file open
  └─ Streams events to Domino
  ↓
Ribbon Buttons
  ├─ Register Model → Opens modal
  ├─ Mark Input → Tag selected cells
  ├─ Mark Output → Tag selected cells
  └─ Dashboard → Show taskpane
  ↓
React UI (optional taskpane)
  └─ View monitored cells and activity
```

## Features

### 1. Persistent Model Identification
- Uses Excel's Custom Document Properties
- Model ID survives Save As, renames, cloud sync
- Format: `excel_<timestamp>_<random>`

### 2. Background Monitoring
- Runs independently of taskpane being open
- Tracks:
  - Cell changes (monitored and unmonitored)
  - Selection changes (user interactions)
  - Worksheet activations
  - Save events
  - Model open/close

### 3. Event Streaming
- Real-time POST to Domino API
- Offline queue (up to 100 events)
- Batch flush when back online
- `keepalive` flag for reliability

### 4. Ribbon Integration
- Native Excel ribbon buttons
- No need to open taskpane for common actions
- Visual feedback (cell highlighting)

## Setup

### Prerequisites

- **Node.js** (v18 or higher)
- **Excel** (Mac, Windows, or Excel Online)
- **Office 365** account
- **Domino API** endpoint

### Installation

```bash
# Clone or navigate to project
cd excel-governance

# Install dependencies
npm install

# Generate SSL certificates for local development
npx office-addin-dev-certs install

# Start development server
npm start
```

The dev server will start on `https://localhost:3000`

### First-time Excel Setup

1. Open Excel
2. Go to **Insert** → **Get Add-ins**
3. Click **My Add-ins** tab
4. Choose **Upload My Add-in**
5. Browse to `manifest.xml` in this project
6. Click **Upload**

The add-in will appear in the **Home** ribbon under "Domino" group.

## Configuration

### Update API Endpoint

Edit the following files to point to your Domino API:

**src/commands/commands.js:**
```javascript
const DOMINO_API_BASE = 'https://your-domino.com/api';
```

**src/utils/domino-api.js:**
```javascript
const DOMINO_API_BASE = process.env.VITE_DOMINO_API_URL || 'https://your-domino.com/api';
```

Or use environment variables:

```bash
# .env.local
VITE_DOMINO_API_URL=https://your-domino.com/api
```

## Usage

### For End Users

#### 1. Install Add-in (one-time)
- IT sends manifest link
- Click "Add" in Excel
- Done

#### 2. Register a Model
1. Open your Excel file
2. Click **Register Model** in ribbon
3. Fill in:
   - Model name
   - Owner email
   - Description (optional)
4. Click **Register**

Model is now being monitored!

#### 3. Mark Important Cells
1. Select cell(s) in Excel
2. Click **Mark Input** or **Mark Output**
3. Cells are highlighted and tracked

#### 4. Work Normally
- The add-in runs silently in background
- All changes streamed to Domino
- No performance impact

#### 5. View Dashboard (optional)
- Click **Dashboard** button to see:
  - Monitored cells
  - Recent activity
  - Model info

### For Developers

#### Project Structure

```
excel-governance/
├── manifest.xml              # Add-in configuration
├── package.json
├── vite.config.js
├── index.html                # Main taskpane entry
├── commands.html             # Background script entry
├── register.html             # Registration modal entry
├── src/
│   ├── main.jsx              # React mount for taskpane
│   ├── register.jsx          # React mount for modal
│   ├── commands/
│   │   └── commands.js       # Background monitoring (CORE)
│   ├── taskpane/
│   │   ├── App.jsx           # Main taskpane component
│   │   ├── RegisterModal.jsx # Registration form
│   │   └── MonitorView.jsx   # Activity dashboard
│   └── utils/
│       ├── domino-api.js     # API client
│       └── model-id.js       # Model ID management
└── assets/
    └── icon-*.png            # Add-in icons
```

#### Key Files

**commands.js** - The heart of the add-in
- Runs in background (not dependent on taskpane)
- Registers event handlers on Office.onReady
- Streams events to Domino
- Handles offline queue

**model-id.js** - Persistent identification
- Uses Custom Document Properties
- Survives file operations
- Provides utility functions for model metadata

**domino-api.js** - API integration
- All HTTP calls to Domino backend
- Timeout handling
- Error recovery

## Domino API Requirements

The add-in expects the following endpoints:

### GET /api/models/:modelId
Check if model is registered
```json
{
  "id": "excel_abc123",
  "name": "Revenue Forecast",
  "owner": "nick@company.com",
  "monitoredCells": [
    { "range": "A1", "type": "input" }
  ]
}
```

### POST /api/models
Register new model
```json
{
  "modelId": "excel_abc123",
  "name": "Revenue Forecast",
  "owner": "nick@company.com",
  "description": "Q1 revenue model"
}
```

### POST /api/excel-events
Stream individual event
```json
{
  "event": "cell_changed",
  "modelId": "excel_abc123",
  "cell": "A1",
  "value": 1000,
  "timestamp": "2025-11-12T10:00:00Z"
}
```

### POST /api/excel-events/batch
Stream multiple events (offline queue flush)
```json
{
  "events": [
    { "event": "cell_changed", ... },
    { "event": "model_saved", ... }
  ]
}
```

### POST /api/models/:modelId/cells
Add monitored cell
```json
{
  "range": "B2:B10",
  "type": "output"
}
```

## Deployment

### For IT Distribution

#### Option 1: Centralized Deployment (Enterprise)
1. Admin uploads `manifest.xml` to SharePoint or Azure
2. Add-in deployed via Office 365 Admin Center
3. Appears automatically for all users

#### Option 2: Self-service
1. Host `manifest.xml` on internal server
2. Send link to users
3. Users click "Add" in Excel

### Production Build

```bash
# Build for production
npm run build

# Output in /dist folder
# Deploy to web server (e.g., Azure Static Web Apps)
```

Update `manifest.xml` URLs from `localhost:3000` to production URL:

```xml
<bt:Url id="CommandsFile.Url" DefaultValue="https://your-domain.com/commands.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://your-domain.com/index.html"/>
```

## Development Tips

### Live Reload
- Changes to React components hot-reload automatically
- Changes to `commands.js` require Excel restart
- Changes to `manifest.xml` require re-uploading add-in

### Debugging

**Browser DevTools:**
- Right-click taskpane → Inspect
- Console logs appear here
- React DevTools work

**Background Script (commands.js):**
- Logs appear in browser console when add-in loads
- Use `console.log()` liberally
- No visible UI to debug from

**Office.js Issues:**
- Check browser console for errors
- Verify Office.js loaded: `typeof Office !== 'undefined'`
- Use `Office.onReady()` for all initialization

### Testing

1. **Local testing**: `npm start` → Load in Excel
2. **Test registration**: Use mock Domino API (see below)
3. **Test monitoring**: Change cells, check console logs
4. **Test offline queue**: Disconnect network, make changes, reconnect

### Mock Domino API (for development)

Create a simple Express server:

```javascript
// mock-api.js
const express = require('express');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(express.json());

const models = {};

app.get('/api/models/:id', (req, res) => {
  const model = models[req.params.id];
  if (model) {
    res.json(model);
  } else {
    res.status(404).json({ error: 'Not found' });
  }
});

app.post('/api/models', (req, res) => {
  const model = req.body;
  models[model.modelId] = model;
  console.log('Model registered:', model);
  res.json(model);
});

app.post('/api/excel-events', (req, res) => {
  console.log('Event received:', req.body);
  res.json({ status: 'ok' });
});

app.listen(5000, () => console.log('Mock API on http://localhost:5000'));
```

## Monitoring Non-compliance

On Domino side, run daily checks:

```python
def check_model_compliance():
    models = get_registered_excel_models()
    for model in models:
        last_ping = get_last_event(model.id)
        if last_ping is None or days_since(last_ping) > 7:
            alert_manager(model.owner, f"No telemetry from {model.name}")
```

## Troubleshooting

### Add-in doesn't load
- Check manifest.xml is valid XML
- Verify URLs in manifest point to running server
- Clear Office cache: `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`

### Background monitoring not working
- Check browser console for errors in commands.js
- Verify Office.js loaded: F12 → Console → `typeof Office`
- Restart Excel completely

### Events not reaching Domino
- Check network tab in DevTools
- Verify CORS configured on Domino API
- Check API endpoint URL in code

### Modal doesn't open
- Verify popup blockers disabled
- Check browser console for dialog errors
- Ensure HTTPS (not HTTP) for production

## Security Considerations

- **HTTPS required** for production (Office.js requirement)
- **CORS** must be configured on Domino API
- **No authentication in MVP** - add OAuth/JWT for production
- **Custom properties** visible to anyone with file access

## Future Enhancements

Not in MVP, but could add:

- Formula tracking and dependency analysis
- Data lineage visualization
- Rollback capability
- Conflict detection for multi-user scenarios
- Advanced compliance rules
- Integration with Git for version control
- Automated testing framework

## License

Proprietary - Domino Data Lab

## Support

For issues or questions:
- Internal: #excel-governance Slack channel
- Email: governance@domino.com
