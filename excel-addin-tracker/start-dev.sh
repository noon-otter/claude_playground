#!/bin/bash

echo "🚀 Starting Excel Model Tracker Development Environment"
echo ""

# Check if Docker is running
if ! docker info > /dev/null 2>&1; then
    echo "❌ Docker is not running. Please start Docker Desktop and try again."
    exit 1
fi

echo "✅ Docker is running"
echo ""

# Navigate to project root
cd "$(dirname "$0")"

# Start backend and database
echo "📦 Starting backend and database..."
docker compose up -d

# Wait for services to be healthy
echo "⏳ Waiting for services to be ready..."
for i in {1..30}; do
    if curl -s http://localhost:8000 > /dev/null 2>&1; then
        echo "✅ Backend is running at http://localhost:8000"
        break
    fi
    if [ $i -eq 30 ]; then
        echo "⚠️  Backend did not start in time. Check logs with: docker compose logs backend"
        exit 1
    fi
    sleep 1
done

# Check database
if docker compose ps | grep -q "excel_tracker_db.*Up"; then
    echo "✅ Database is running"
else
    echo "⚠️  Database may not be ready. Check logs with: docker compose logs postgres"
fi

echo ""

# Check if frontend dependencies are installed
if [ ! -d "frontend/node_modules" ]; then
    echo "📥 Installing frontend dependencies (this will take 2-3 minutes)..."
    cd frontend && npm install && cd ..
    echo "✅ Dependencies installed"
else
    echo "✅ Frontend dependencies already installed"
fi

# Check if SSL certs are installed
echo ""
echo "🔒 Checking SSL certificates..."
if [ ! -f "$HOME/.office-addin-dev-certs/localhost.crt" ]; then
    echo "📜 SSL certificates not found. Installing..."
    cd frontend
    npx office-addin-dev-certs install
    cd ..
    echo "✅ SSL certificates installed"
else
    echo "✅ SSL certificates already installed"
fi

echo ""
echo "✨ Development environment is ready!"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "📋 Next steps:"
echo ""
echo "1. Open a NEW terminal and run:"
echo "   cd excel-addin-tracker/frontend"
echo "   npm run dev-server"
echo ""
echo "2. Wait for 'webpack compiled successfully', then open ANOTHER terminal:"
echo "   cd excel-addin-tracker/frontend"
echo "   npm run start"
echo ""
echo "   This will open Excel with the add-in loaded!"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "🔧 Useful commands:"
echo "   docker compose logs -f backend    # View backend logs"
echo "   docker compose logs -f postgres   # View database logs"
echo "   docker compose down               # Stop all services"
echo "   docker compose down -v            # Stop and remove data"
echo ""
echo "🐛 Debugging:"
echo "   - Right-click in taskpane → Inspect (opens DevTools)"
echo "   - Check console for errors"
echo "   - View API calls in Network tab"
echo ""
echo "📚 Full documentation: README.md"
echo ""
