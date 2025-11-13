#!/bin/bash

# Excel Add-in Development Startup Script
# This script starts both the backend server and frontend dev server

set -e

echo "================================================="
echo "🚀 Starting Excel Add-in Development Environment"
echo "================================================="
echo ""

# Kill any existing processes on ports 3000 and 5000
echo "🧹 Cleaning up existing processes..."
lsof -ti:3000 | xargs kill -9 2>/dev/null || true
lsof -ti:5000 | xargs kill -9 2>/dev/null || true
sleep 1

# Start backend server
echo ""
echo "📦 Starting Backend Server (SQLite)..."
echo "   URL: http://localhost:5000"
python3 backend-sqlite.py > backend.log 2>&1 &
BACKEND_PID=$!
echo "   PID: $BACKEND_PID"

# Wait for backend to start
sleep 2

# Check if backend is running
if ! curl -s http://localhost:5000/ > /dev/null; then
    echo "❌ Backend failed to start. Check backend.log for details"
    cat backend.log
    exit 1
fi

echo "   ✅ Backend is running"

# Start frontend dev server
echo ""
echo "🎨 Starting Frontend Dev Server (Vite)..."
echo "   URL: https://localhost:3000"
npm run start

# This script will stay running until you press Ctrl+C
# When you exit, it will clean up the backend process

trap "echo 'Stopping servers...'; kill $BACKEND_PID 2>/dev/null; exit 0" INT TERM EXIT
