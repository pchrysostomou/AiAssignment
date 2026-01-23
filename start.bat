@echo off
echo ========================================
echo AI Assignment Architect - PRO SaaS
echo ========================================
echo.

REM Check if virtual environment exists
if not exist "venv\" (
    echo Creating virtual environment...
    python -m venv venv
    echo.
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate
echo.

REM Check if requirements are installed
if not exist "venv\Lib\site-packages\flask\" (
    echo Installing dependencies...
    pip install -r requirements.txt
    echo.
)

REM Check if .env exists
if not exist ".env" (
    echo WARNING: .env file not found!
    echo Please copy .env.template to .env and add your API keys.
    echo.
    pause
    exit /b 1
)

REM Start the application
echo Starting application...
echo.
echo Application will be available at: http://localhost:5000
echo Demo login - Username: demo, Password: demo123
echo.
echo Press Ctrl+C to stop the server
echo.
python app.py