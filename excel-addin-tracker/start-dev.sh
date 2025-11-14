#!/bin/bash

echo "🚀 Starting Excel Model Tracker Development Environment"
echo ""

# Check if Docker is running
if ! docker info > /dev/null 2>&1; then
    echo "❌ Docker is not running. Please start Docker and try again."
    exit 1
fi

# Start backend and database
echo "📦 Starting backend and database..."
docker-compose up -d

# Wait for services to be healthy
echo "⏳ Waiting for services to be ready..."
sleep 5

# Check backend health
if curl -s http://localhost:8000 > /dev/null; then
    echo "✅ Backend is running at http://localhost:8000"
else
    echo "⚠️  Backend may not be ready yet. Check logs with: docker-compose logs backend"
fi

# Check if frontend dependencies are installed
if [ ! -d "frontend/node_modules" ]; then
    echo "📥 Installing frontend dependencies..."
    cd frontend && npm install && cd ..
fi

echo ""
echo "✨ Development environment is ready!"
echo ""
echo "Next steps:"
echo "1. cd frontend"
echo "2. npm run dev-server    (in one terminal)"
echo "3. npm run start         (in another terminal to sideload Excel)"
echo ""
echo "Useful commands:"
echo "  docker-compose logs -f backend    # View backend logs"
echo "  docker-compose down               # Stop all services"
echo "  docker-compose down -v            # Stop and remove volumes"
