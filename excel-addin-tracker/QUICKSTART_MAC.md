# Quick Start Guide for Mac

Get the Excel Model Tracker running on your Mac in 5 minutes.

## ✅ Prerequisites Check

```bash
# Check Node.js (need v18+)
node --version

# Check Docker
docker --version

# Check Excel
# Open Excel manually to verify it's installed
```

If anything is missing:
- **Node.js**: Download from https://nodejs.org/ (get LTS version)
- **Docker Desktop**: Download from https://www.docker.com/products/docker-desktop/
- **Excel**: Use Microsoft 365 or Office 2019+

---

## 🚀 One-Command Setup

Open Terminal and run:

```bash
cd excel-addin-tracker
./start-dev.sh
```

This will:
- ✅ Start Docker containers (backend + database)
- ✅ Install npm dependencies
- ✅ Generate placeholder icons
- ✅ Install SSL certificates
- ✅ Show you next steps

**Note:** You'll be prompted for your Mac password to install SSL certificates.

---

## 📱 Start the Add-in

After running `start-dev.sh`, you need **two more terminal windows**:

### Terminal 1: Dev Server

```bash
cd excel-addin-tracker/frontend
npm run dev-server
```

Wait for: `✔ webpack compiled successfully`

**Keep this terminal running!**

### Terminal 2: Sideload to Excel

```bash
cd excel-addin-tracker/frontend
npm run start
```

This opens Excel with the add-in loaded.

---

## 🎯 Using the Add-in

1. **Find the ribbon button:**
   - Look in the **Home** ribbon
   - Click "Show Taskpane" button

2. **Register your first model:**
   - Enter name: "Test Model"
   - Click "Register Model"

3. **Add tracked ranges:**
   - Create some data in Excel (Sheet1, A1:B5)
   - Range Name: `Inputs`
   - Range Address: `Sheet1!A1:B5`
   - Click "Add Tracked Range"

4. **Test it works:**
   - Modify a cell in A1:B5
   - Changes are automatically logged!

---

## 🐛 Mac-Specific Debugging

### Enable Web Inspector

```bash
defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
```

Then right-click in taskpane → **Inspect Element**

This opens Safari Web Inspector for debugging.

### View Console Logs

Right-click anywhere in taskpane → **Inspect Element** → **Console** tab

### Check Backend Logs

```bash
docker compose logs -f backend
```

Press `Ctrl+C` to exit.

---

## ⚠️ Troubleshooting (Mac)

### "Add-in is no longer available"

**Fix 1: Restart dev server**
```bash
# In Terminal 1 (where dev-server is running)
# Press Ctrl+C to stop
npm run dev-server
```

**Fix 2: Clear Office cache**
```bash
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
# Then restart Excel
```

### SSL Certificate Issues

If you see certificate warnings:

```bash
# Reinstall certificates
cd frontend
npx office-addin-dev-certs install

# Then trust in Keychain Access:
# 1. Open "Keychain Access" app
# 2. Search for "localhost"
# 3. Double-click certificate
# 4. Expand "Trust"
# 5. Set "When using this certificate" to "Always Trust"
# 6. Close (enter password when prompted)
# 7. Restart Excel
```

### Backend Not Responding

```bash
# Check if backend is running
curl http://localhost:8000

# If no response, check Docker
docker compose ps

# Restart backend
docker compose restart backend

# Check logs
docker compose logs backend
```

### Port Already in Use

```bash
# Check what's using port 3000
lsof -i :3000

# Check what's using port 8000
lsof -i :8000

# Kill process if needed
kill -9 <PID>
```

### Excel Doesn't Open Automatically

If `npm run start` doesn't open Excel:

1. Open Excel manually
2. Go to **Insert** → **Add-ins** → **My Add-ins**
3. Look for "Excel Model Tracker"
4. Click it

If you don't see it there:
```bash
# Re-sideload
cd frontend
npm run stop
npm run start
```

---

## 🔄 Daily Development Workflow

### Starting Work

```bash
# Terminal 1
cd excel-addin-tracker
docker compose up -d

# Terminal 2
cd excel-addin-tracker/frontend
npm run dev-server

# Terminal 3
cd excel-addin-tracker/frontend
npm run start
```

### Making Changes

1. Edit code in `frontend/src/taskpane/`
2. Webpack auto-rebuilds
3. **Close and reopen taskpane** in Excel (hot reload doesn't work in Office)

### Stopping Work

```bash
# Stop Docker services
docker compose down

# Stop dev server (Ctrl+C in terminal)
# Stop Excel add-in debugging (Ctrl+C in terminal or close Excel)
```

---

## 📊 Testing Your Setup

### Test Backend

```bash
curl http://localhost:8000
# Should return: {"status":"ok","service":"Domino Spreadsheet Backend"}
```

### Test Database

```bash
docker compose exec postgres psql -U postgres -d excel_tracker -c "SELECT 1;"
# Should return: ?column?
#                1
```

### Test API Endpoints

```bash
# Create a model
curl -X PUT http://localhost:8000/wb/upsert-model \
  -H "Content-Type: application/json" \
  -d '{
    "model_name": "Test",
    "tracked_ranges": [{"name": "Test", "range": "Sheet1!A1:A10"}]
  }'

# Should return JSON with model_id and version
```

---

## 🎨 Customize Icons

Replace the placeholder icons:

```bash
cd excel-addin-tracker/frontend/assets

# Option 1: Use ImageMagick (if installed)
brew install imagemagick
convert -size 16x16 canvas:#0078D4 icon-16.png
convert -size 32x32 canvas:#0078D4 icon-32.png
convert -size 64x64 canvas:#0078D4 icon-64.png
convert -size 80x80 canvas:#0078D4 icon-80.png

# Option 2: Download from icon sites
# - flaticon.com
# - icons8.com
# Save as icon-16.png, icon-32.png, etc.

# Option 3: Use your own logo
# Just resize to 16x16, 32x32, 64x64, 80x80 pixels
```

---

## 🧰 Useful Mac Commands

```bash
# View all logs
docker compose logs -f

# Check running containers
docker compose ps

# Restart everything
docker compose restart

# Stop everything
docker compose down

# Clean slate (deletes all data!)
docker compose down -v

# Check ports in use
lsof -i :3000
lsof -i :8000
lsof -i :5432

# Check Office cache location
ls -la ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef

# Check SSL certificates
ls -la ~/.office-addin-dev-certs
```

---

## 📚 Next Steps

- Read full documentation: `README.md`
- See comprehensive testing guide: `TESTING.md`
- Customize the UI: Edit `frontend/src/taskpane/taskpane.html`
- Add API endpoints: Edit `backend/main.py`

---

## 💡 Pro Tips for Mac

1. **Use iTerm2** instead of Terminal for better experience
2. **Enable Web Inspector** (command shown above) for easier debugging
3. **Use VS Code** with Office Add-in extension for better development
4. **Trust localhost cert** in Keychain to avoid SSL warnings
5. **Use TablePlus** for a nice PostgreSQL GUI (free tier available)

---

## 🆘 Still Having Issues?

1. Check `README.md` Troubleshooting section
2. Check `TESTING.md` for detailed testing steps
3. Enable debug mode in Excel:
   ```bash
   defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
   ```
4. View console in browser DevTools (right-click → Inspect)
5. Check Docker logs: `docker compose logs`

---

## ✅ Success Checklist

After setup, you should have:

- [ ] Backend running at http://localhost:8000
- [ ] Database running (check with `docker compose ps`)
- [ ] Dev server running at https://localhost:3000
- [ ] Excel open with add-in taskpane visible
- [ ] Can register a model successfully
- [ ] Can add tracked ranges
- [ ] Changes are logged to database

If all checked, you're ready to develop! 🎉
