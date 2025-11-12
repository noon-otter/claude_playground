# Quick Start Guide

Get the Domino Excel Governance Add-in running in 5 minutes.

## Prerequisites

- Node.js 18+ installed
- Excel (Mac, Windows, or Online)
- Terminal/Command Line

## Steps

### 1. Install Dependencies

```bash
cd excel-governance
npm install
```

### 2. Generate SSL Certificates

Office Add-ins require HTTPS, even for local development:

```bash
npx office-addin-dev-certs install
```

This creates self-signed certificates in the `certs/` folder.

### 3. Start Dev Server

```bash
npm start
```

Server starts on `https://localhost:3000`

### 4. Load in Excel

**Option A: Mac/Windows Excel**
1. Open Excel
2. Insert → Get Add-ins → My Add-ins
3. Upload My Add-in
4. Browse to `manifest.xml` in this folder
5. Click Upload

**Option B: Excel Online**
1. Go to Office.com → Excel
2. Open any workbook
3. Insert → Office Add-ins → Upload My Add-in
4. Upload `manifest.xml`

### 5. Start Mock API (optional)

In a new terminal:

```bash
pip install fastapi uvicorn
python domino-api-example.py
```

Mock API runs on `http://localhost:5000`

Update `src/commands/commands.js` and `src/utils/domino-api.js`:
```javascript
const DOMINO_API_BASE = 'http://localhost:5000/api';
```

### 6. Test It Out

1. In Excel, look for "Domino" group in Home ribbon
2. Click **Register Model**
3. Fill in the form and submit
4. Click **Mark Input** after selecting a cell
5. Change that cell's value
6. Check the mock API logs - you should see events!

### 7. View Dashboard

Click **Dashboard** button to see:
- Model info
- Monitored cells
- Recent activity

## Next Steps

- Read [README.md](README.md) for full documentation
- Configure your real Domino API endpoint
- Deploy to production
- Distribute to users

## Common Issues

**"Add-in won't load"**
- Check console for errors (F12)
- Verify dev server is running on https://localhost:3000
- Clear Excel cache: Delete `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (Mac)

**"Certificate errors"**
- Re-run `npx office-addin-dev-certs install`
- Restart browser/Excel

**"Events not showing in API"**
- Check CORS is enabled
- Verify API endpoint URL in code
- Look at browser Network tab (F12)

## Development Workflow

1. Make changes to React components → Hot reload (automatic)
2. Make changes to `commands.js` → Restart Excel
3. Make changes to `manifest.xml` → Re-upload add-in
4. Check console logs → F12 → Console
5. Check network calls → F12 → Network

## What Gets Monitored?

Once a model is registered, these events stream to Domino:

- ✅ Model opened/closed
- ✅ Model saved
- ✅ Cell changed (monitored cells)
- ✅ Cell changed (unmonitored - for audit trail)
- ✅ Selection changed (user interactions)
- ✅ Worksheet activated

All events include:
- Model ID
- Timestamp
- User (if available)
- Cell address / value (if applicable)

## Tips

- Use Chrome DevTools to debug React components
- `console.log()` in `commands.js` for background debugging
- Test offline mode: disable network, make changes, re-enable
- Check `events_db` in mock API to see all captured events
