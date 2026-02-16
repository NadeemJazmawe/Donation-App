@echo off
echo Setting up Donation App...
echo.

echo Creating backend .env file...
if not exist backend\.env (
    copy backend\env.example backend\.env
    echo Backend .env file created.
) else (
    echo Backend .env file already exists.
)

echo.
echo Installing root dependencies...
call npm install

echo.
echo Installing backend dependencies...
cd backend
call npm install
cd ..

echo.
echo Installing frontend dependencies...
cd frontend
call npm install
cd ..

echo.
echo Setup complete!
echo.
echo Make sure MongoDB is running on your local machine.
echo Then run: npm run dev
pause
