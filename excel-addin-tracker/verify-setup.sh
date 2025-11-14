#!/bin/bash

echo "🔍 Excel Model Tracker - Setup Verification"
echo "==========================================="
echo ""

ERRORS=0
WARNINGS=0

# Function to check command
check_command() {
    if command -v "$1" &> /dev/null; then
        echo "✅ $2 is installed"
        if [ ! -z "$3" ]; then
            VERSION=$($1 --version 2>&1 | head -n 1)
            echo "   Version: $VERSION"
        fi
    else
        echo "❌ $2 is NOT installed"
        echo "   Install from: $3"
        ERRORS=$((ERRORS + 1))
    fi
}

# Function to check port
check_port() {
    if lsof -Pi :$1 -sTCP:LISTEN -t >/dev/null 2>&1; then
        echo "✅ Port $1 is in use (service running)"
    else
        echo "⚠️  Port $1 is NOT in use (service may not be running)"
        WARNINGS=$((WARNINGS + 1))
    fi
}

# Check Node.js
echo "📦 Checking Node.js..."
if command -v node &> /dev/null; then
    NODE_VERSION=$(node --version | sed 's/v//')
    MAJOR_VERSION=$(echo $NODE_VERSION | cut -d. -f1)
    if [ "$MAJOR_VERSION" -ge 18 ]; then
        echo "✅ Node.js is installed (v$NODE_VERSION)"
    else
        echo "⚠️  Node.js version is too old (v$NODE_VERSION)"
        echo "   Recommended: v18 or later"
        WARNINGS=$((WARNINGS + 1))
    fi
else
    echo "❌ Node.js is NOT installed"
    echo "   Install from: https://nodejs.org/"
    ERRORS=$((ERRORS + 1))
fi
echo ""

# Check npm
echo "📦 Checking npm..."
check_command "npm" "npm" "true"
echo ""

# Check Docker
echo "🐳 Checking Docker..."
if command -v docker &> /dev/null; then
    if docker info &> /dev/null; then
        echo "✅ Docker is installed and running"
        DOCKER_VERSION=$(docker --version)
        echo "   Version: $DOCKER_VERSION"
    else
        echo "⚠️  Docker is installed but NOT running"
        echo "   Start Docker Desktop and try again"
        WARNINGS=$((WARNINGS + 1))
    fi
else
    echo "❌ Docker is NOT installed"
    echo "   Install from: https://www.docker.com/products/docker-desktop/"
    ERRORS=$((ERRORS + 1))
fi
echo ""

# Check Docker Compose
echo "🐳 Checking Docker Compose..."
if docker compose version &> /dev/null; then
    COMPOSE_VERSION=$(docker compose version)
    echo "✅ Docker Compose is available"
    echo "   Version: $COMPOSE_VERSION"
else
    echo "❌ Docker Compose is NOT available"
    echo "   Usually comes with Docker Desktop"
    ERRORS=$((ERRORS + 1))
fi
echo ""

# Check if in correct directory
echo "📁 Checking project structure..."
if [ -f "docker-compose.yml" ] && [ -d "frontend" ] && [ -d "backend" ]; then
    echo "✅ Project structure looks correct"
else
    echo "⚠️  Project structure may be incomplete"
    echo "   Make sure you're in the excel-addin-tracker directory"
    WARNINGS=$((WARNINGS + 1))
fi
echo ""

# Check Docker services
echo "🔧 Checking Docker services..."
if docker compose ps &> /dev/null; then
    if docker compose ps | grep -q "Up"; then
        echo "✅ Docker services are running"
        docker compose ps
    else
        echo "⚠️  Docker services are not running"
        echo "   Run: docker compose up -d"
        WARNINGS=$((WARNINGS + 1))
    fi
else
    echo "⚠️  Cannot check Docker services"
    WARNINGS=$((WARNINGS + 1))
fi
echo ""

# Check backend
echo "🌐 Checking backend API..."
if curl -s http://localhost:8000 &> /dev/null; then
    RESPONSE=$(curl -s http://localhost:8000)
    echo "✅ Backend API is responding"
    echo "   Response: $RESPONSE"
else
    echo "⚠️  Backend API is NOT responding"
    echo "   Run: docker compose up -d"
    WARNINGS=$((WARNINGS + 1))
fi
echo ""

# Check ports
echo "🔌 Checking ports..."
check_port 3000  # Frontend dev server
check_port 8000  # Backend API
check_port 5432  # PostgreSQL
echo ""

# Check frontend dependencies
echo "📦 Checking frontend dependencies..."
if [ -d "frontend/node_modules" ]; then
    echo "✅ Frontend dependencies are installed"
else
    echo "⚠️  Frontend dependencies are NOT installed"
    echo "   Run: cd frontend && npm install"
    WARNINGS=$((WARNINGS + 1))
fi
echo ""

# Check SSL certificates
echo "🔒 Checking SSL certificates..."
if [ -f "$HOME/.office-addin-dev-certs/localhost.crt" ]; then
    echo "✅ SSL certificates are installed"
else
    echo "⚠️  SSL certificates are NOT installed"
    echo "   Run: cd frontend && npx office-addin-dev-certs install"
    WARNINGS=$((WARNINGS + 1))
fi
echo ""

# Check icons
echo "🎨 Checking icons..."
ICON_COUNT=0
for size in 16 32 64 80; do
    if [ -f "frontend/assets/icon-${size}.png" ]; then
        # Check if it's a real PNG (starts with PNG signature)
        if file "frontend/assets/icon-${size}.png" | grep -q "PNG"; then
            ICON_COUNT=$((ICON_COUNT + 1))
        fi
    fi
done

if [ $ICON_COUNT -eq 4 ]; then
    echo "✅ All icon files are present and valid"
else
    echo "⚠️  Some icons are missing or invalid ($ICON_COUNT/4)"
    echo "   Run: cd frontend && npm run generate-icons"
    WARNINGS=$((WARNINGS + 1))
fi
echo ""

# Summary
echo "==========================================="
echo "📊 Verification Summary"
echo "==========================================="
echo ""

if [ $ERRORS -eq 0 ] && [ $WARNINGS -eq 0 ]; then
    echo "🎉 Everything looks good! Your setup is complete."
    echo ""
    echo "Next steps:"
    echo "1. cd frontend && npm run dev-server    (Terminal 1)"
    echo "2. cd frontend && npm run start         (Terminal 2)"
    echo ""
elif [ $ERRORS -eq 0 ]; then
    echo "⚠️  Setup is mostly complete with $WARNINGS warning(s)"
    echo ""
    echo "Review the warnings above and follow the suggested fixes."
    echo ""
else
    echo "❌ Setup is incomplete with $ERRORS error(s) and $WARNINGS warning(s)"
    echo ""
    echo "Please install the missing prerequisites before continuing."
    echo ""
fi

echo "For detailed setup instructions, see:"
echo "  - README.md (comprehensive guide)"
echo "  - QUICKSTART_MAC.md (Mac-specific quick start)"
echo ""

exit $ERRORS
