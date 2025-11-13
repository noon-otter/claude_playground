# Quick Fixes Applied

## Issue: "API you are trying to use could not be found"

### Root Cause
1. Excel APIs (`workbook.properties.custom`) were called before Office.js was fully initialized
2. Requirement set was too high (1.7) for older Excel versions
3. Missing fallback handling for unsupported APIs
4. Missing icon assets

### Fixes Applied

#### 1. Added API Availability Checks (`src/utils/model-id.js`)
```javascript
// Now checks if Excel API is available before using it
if (typeof Excel === 'undefined' || !Excel.run) {
  return fallbackValue;
}
```

#### 2. Added Fallback Logic
- Falls back to localStorage if Excel properties unavailable
- Graceful degradation for unsupported Excel versions

#### 3. Lowered API Requirements (`manifest.xml`)
Changed from:
```xml
<Set Name="ExcelApi" MinVersion="1.7"/>
```
To:
```xml
<Set Name="ExcelApi" MinVersion="1.1"/>
```

#### 4. Created Placeholder Icons
- `public/assets/icon-16.png`
- `public/assets/icon-32.png`
- `public/assets/icon-64.png`
- `public/assets/icon-80.png`

Note: These are placeholders. Replace with proper icons later.

## Testing

### 1. Restart Everything
```bash
# Kill all processes
lsof -ti:3000 | xargs kill -9
lsof -ti:5000 | xargs kill -9

# Restart backend
npm run backend

# In new terminal:
npm start
```

### 2. Reinstall Add-in
```bash
./install_addin_locally.sh
```

### 3. Restart Excel
- Quit Excel completely
- Reopen Excel
- Go to: Insert → Add-ins → Shared Folder Add-ins
- Click on "Domino Governance"

## Finding Logs in Excel

### Method 1: Browser DevTools (Recommended)
1. Right-click anywhere in the add-in task pane
2. Select "Inspect" or "Inspect Element"
3. DevTools opens showing Console tab
4. Look for errors in red

### Method 2: Safari Web Inspector (Mac Only)
1. Open Safari → Preferences → Advanced
2. Enable "Show Develop menu"
3. Safari → Develop → localhost
4. Select your add-in page
5. Console shows all logs

### Method 3: Office.js Logging
Add this to your code to see Office.js info:
```javascript
Office.onReady((info) => {
  console.log('Host:', info.host);
  console.log('Platform:', info.platform);
  console.log('Version:', Office.context.diagnostics);
});
```

## Common Issues

### Add-in Still Not Loading
1. Check HTTPS certificate: `npm run dev-certs`
2. Clear Excel cache: Delete `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches`
3. Check manifest is in correct location:
   ```bash
   ls -la ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```

### API Still Not Found
Your Excel version might be too old. Check version:
1. Excel → About Excel
2. Look for version number
3. Minimum required: Excel 2016 or later

### Backend Connection Issues
1. Check backend is running: `curl http://localhost:5000/`
2. Check CORS: Look for CORS errors in console
3. Update API URL in code if needed

## Next Steps

1. **Test the fixes**: Restart Excel and try loading the add-in
2. **Check console logs**: Open DevTools (right-click → Inspect)
3. **Report errors**: If still failing, share the console output
4. **Create proper icons**: Replace placeholder icons with real ones

## References

- [Excel API Requirement Sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)
- [Debug Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-overview)
- [Office.js API Reference](https://docs.microsoft.com/javascript/api/office)
