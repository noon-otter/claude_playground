#!/bin/bash
# Auto-start backend with port cleanup

# Kill any process using port 5000
echo "Checking for processes on port 5000..."
lsof -ti:5000 | xargs kill -9 2>/dev/null || true

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    echo "Activating virtual environment..."
    source venv/bin/activate
else
    echo "Warning: No virtual environment found. Run: python3 -m venv venv"
    exit 1
fi

# Start backend
echo "Starting backend on http://localhost:5000..."
python backend.py
