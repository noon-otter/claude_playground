#!/bin/bash
# Unified Development Environment Startup Script
# Starts: PostgreSQL Database + Python Backend + React Frontend

set -e

echo "🚀 Starting Excel Add-In Development Environment"
echo "================================================"

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Function to check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Determine docker compose command
if command_exists docker-compose; then
    DOCKER_COMPOSE="docker-compose"
elif docker compose version >/dev/null 2>&1; then
    DOCKER_COMPOSE="docker compose"
else
    echo -e "${RED}Error: Neither 'docker-compose' nor 'docker compose' is available${NC}"
    exit 1
fi

# Check required tools
echo -e "${BLUE}Checking required tools...${NC}"

if ! command_exists docker; then
    echo -e "${RED}Error: Docker is not installed${NC}"
    exit 1
fi

if ! command_exists python3; then
    echo -e "${RED}Error: python3 is not installed${NC}"
    exit 1
fi

if ! command_exists npm; then
    echo -e "${RED}Error: npm is not installed${NC}"
    exit 1
fi

echo -e "${GREEN}✓ All required tools found${NC}"
echo ""

# Step 1: Start PostgreSQL
echo -e "${BLUE}Step 1: Starting PostgreSQL database...${NC}"
$DOCKER_COMPOSE up -d

# Wait for PostgreSQL to be healthy
echo -e "${YELLOW}Waiting for PostgreSQL to be ready...${NC}"
for i in {1..30}; do
    if $DOCKER_COMPOSE exec -T postgres pg_isready -U excel_user -d excel_addin >/dev/null 2>&1; then
        echo -e "${GREEN}✓ PostgreSQL is ready${NC}"
        break
    fi
    if [ $i -eq 30 ]; then
        echo -e "${RED}Error: PostgreSQL failed to start${NC}"
        exit 1
    fi
    sleep 1
    echo -n "."
done
echo ""

# Step 2: Setup Python Virtual Environment
echo -e "${BLUE}Step 2: Setting up Python backend...${NC}"

if [ ! -d "venv" ]; then
    echo "Creating Python virtual environment..."
    python3 -m venv venv
fi

echo "Activating virtual environment..."
source venv/bin/activate

echo "Installing Python dependencies..."
pip install -q --upgrade pip
pip install -q -r requirements.txt

echo -e "${GREEN}✓ Python environment ready${NC}"
echo ""

# Step 3: Kill existing processes on ports 3000 and 5000
echo -e "${BLUE}Step 3: Cleaning up ports...${NC}"
lsof -ti:3000 | xargs kill -9 2>/dev/null || true
lsof -ti:5000 | xargs kill -9 2>/dev/null || true
echo -e "${GREEN}✓ Ports cleaned${NC}"
echo ""

# Step 4: Install npm dependencies if needed
if [ ! -d "node_modules" ]; then
    echo -e "${BLUE}Installing npm dependencies...${NC}"
    npm install
    echo -e "${GREEN}✓ npm dependencies installed${NC}"
    echo ""
fi

# Step 5: Start all services
echo -e "${BLUE}Step 4: Starting services...${NC}"
echo ""

# Function to cleanup on exit
cleanup() {
    echo ""
    echo -e "${YELLOW}Shutting down...${NC}"
    kill $BACKEND_PID 2>/dev/null || true
    kill $FRONTEND_PID 2>/dev/null || true
    echo -e "${GREEN}✓ Services stopped${NC}"
    exit 0
}

trap cleanup SIGINT SIGTERM

# Start backend in background
echo -e "${BLUE}Starting Python backend on http://localhost:5000...${NC}"
python backend.py > backend.log 2>&1 &
BACKEND_PID=$!

# Wait a moment for backend to start
sleep 2

if ps -p $BACKEND_PID > /dev/null; then
    echo -e "${GREEN}✓ Backend started (PID: $BACKEND_PID)${NC}"
else
    echo -e "${RED}Error: Backend failed to start. Check backend.log${NC}"
    exit 1
fi

# Start frontend in background
echo -e "${BLUE}Starting React frontend on https://localhost:3000...${NC}"
npm start > frontend.log 2>&1 &
FRONTEND_PID=$!

sleep 2

if ps -p $FRONTEND_PID > /dev/null; then
    echo -e "${GREEN}✓ Frontend started (PID: $FRONTEND_PID)${NC}"
else
    echo -e "${RED}Error: Frontend failed to start. Check frontend.log${NC}"
    kill $BACKEND_PID
    exit 1
fi

echo ""
echo -e "${GREEN}================================================${NC}"
echo -e "${GREEN}🎉 Development environment is ready!${NC}"
echo -e "${GREEN}================================================${NC}"
echo ""
echo -e "${BLUE}Services running:${NC}"
echo "  • PostgreSQL:  localhost:5432 (Docker)"
echo "  • Backend API: http://localhost:5000"
echo "  • Frontend:    https://localhost:3000"
echo ""
echo -e "${BLUE}Logs:${NC}"
echo "  • Backend:  tail -f backend.log"
echo "  • Frontend: tail -f frontend.log"
echo ""
echo -e "${YELLOW}Press Ctrl+C to stop all services${NC}"
echo ""

# Wait for processes
wait $BACKEND_PID $FRONTEND_PID
