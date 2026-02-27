@echo off
REM ADGSentinel Quick Start Setup Script (Windows)
REM This script sets up the development environment for the Outlook Web Add-in

setlocal enabledelayedexpansion

echo.
echo ==================================================
echo   ADGSentinel Office.js Add-in Setup
echo ==================================================
echo.

REM Check Node.js installation
where node >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Node.js is not installed. Please install Node.js 12+ first.
    echo Visit: https://nodejs.org/
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('node --version') do set NODE_VERSION=%%i
echo [OK] Node.js !NODE_VERSION! detected

REM Check npm installation
where npm >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: npm is not installed. Please install npm first.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('npm --version') do set NPM_VERSION=%%i
echo [OK] npm !NPM_VERSION! detected

REM Create necessary directories
echo.
echo Creating directory structure...
if not exist "public\assets" mkdir "public\assets"
echo [OK] Directories created

REM Install dependencies
echo.
echo Installing npm dependencies...
call npm install
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)
echo [OK] Dependencies installed

REM Generate SSL certificates
echo.
echo Generating SSL certificates for HTTPS...

if not exist "certs\" (
    mkdir "certs"
    
    REM Check if OpenSSL is installed
    where openssl >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        echo Generating self-signed certificate...
        openssl req -x509 -newkey rsa:2048 -keyout certs\key.pem -out certs\cert.pem ^
            -days 365 -nodes -subj "/C=US/ST=State/L=City/O=Organization/CN=localhost"
        echo [OK] SSL certificates generated in .\certs\
    ) else (
        echo WARNING: OpenSSL not found. You need to generate certificates manually.
        echo Download OpenSSL from: https://slproweb.com/products/Win32OpenSSL.html
        echo.
        echo Or download pre-generated certs from:
        echo https://github.com/adg-tech/outlook-web-addin/releases/download/v1.0.0/certs.zip
        echo And extract to the ./certs directory
        pause
    )
) else (
    echo [OK] Certificates directory exists
)

REM Create .env file if it doesn't exist
echo.
echo Checking configuration...

if not exist ".env" (
    echo Creating .env file from template...
    copy .env.example .env
    echo [WARNING] Please update .env with your email configuration:
    echo   - INFOSEC_EMAIL
    echo   - SPAM_REPORT_EMAIL
    echo   - SUPPORT_EMAIL
    echo   - GOPHISH_URL (optional)
) else (
    echo [OK] .env file already exists
)

REM Summary
echo.
echo ==================================================
echo   Setup Complete!
echo ==================================================
echo.
echo Next steps:
echo 1. Update your email addresses in .env file
echo 2. Update configuration in function-file.js (CONFIG object)
echo 3. Create icons in public\assets\ directory (optional)
echo 4. Run 'npm start' to start the development server
echo.
echo Server will be available at: https://localhost:3000
echo Manifest URL: https://localhost:3000/manifest.xml
echo.
echo To upload to Outlook:
echo 1. Open Outlook Web Settings ^> Get Add-ins
echo 2. Click 'My Add-ins' ^> 'Upload My Add-in'
echo 3. Choose 'Upload from URL' or upload manifest.xml directly
echo.
echo ==================================================
echo.
pause
