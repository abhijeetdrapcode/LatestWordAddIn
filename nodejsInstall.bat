@echo off
setlocal EnableDelayedExpansion

:: Check for administrator privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo This script requires administrator privileges.
    echo Please right-click and select "Run as administrator"
    pause
    exit /b 1
)

echo Installing Node.js...

:: Create a temporary directory for downloads
set "INSTALL_DIR=%TEMP%\nodejs-installer"
mkdir "%INSTALL_DIR%" 2>nul
cd /d "%INSTALL_DIR%"

:: Download Node.js LTS installer
echo Downloading Node.js installer...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('https://nodejs.org/dist/v20.11.1/node-v20.11.1-x64.msi', 'nodejs.msi')"

if not exist "nodejs.msi" (
    echo Failed to download Node.js installer.
    goto :error
)

:: Install Node.js
echo Installing Node.js...
msiexec /i nodejs.msi /qn /norestart ADDLOCAL=ALL
if %errorlevel% neq 0 (
    echo Failed to install Node.js.
    goto :error
)

:: Wait for installation to complete
echo Waiting for installation to complete...
timeout /t 20 /nobreak

:: Clean up
cd /d "%USERPROFILE%"
rmdir /s /q "%INSTALL_DIR%"

echo Installation completed successfully!
echo Please restart your terminal to start using Node.js.
pause
exit /b 0

:error
echo.
echo Installation failed! Please try the following:
echo 1. Restart your computer and try again
echo 2. Try installing Node.js manually from https://nodejs.org
echo 3. Check if your antivirus is blocking the installation
cd /d "%USERPROFILE%"
rmdir /s /q "%INSTALL_DIR%" 2>nul
pause
exit /b 1