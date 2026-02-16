#!/bin/bash

echo "Setting up Donation App..."
echo ""

echo "Creating backend .env file..."
if [ ! -f backend/.env ]; then
    cp backend/env.example backend/.env
    echo "Backend .env file created."
else
    echo "Backend .env file already exists."
fi

echo ""
echo "Installing root dependencies..."
npm install

echo ""
echo "Installing backend dependencies..."
cd backend
npm install
cd ..

echo ""
echo "Installing frontend dependencies..."
cd frontend
npm install
cd ..

echo ""
echo "Setup complete!"
echo ""
echo "Make sure MongoDB is running on your local machine."
echo "Then run: npm run dev"
