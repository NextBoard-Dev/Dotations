@echo off
setlocal

set "ROOT=%~dp0"
set "APP_DIR=%ROOT%smartphone"
set "LOCAL_IP="

if not exist "%APP_DIR%\package.json" (
  echo [ERREUR] Dossier smartphone introuvable: "%APP_DIR%"
  pause
  exit /b 1
)

where node >nul 2>nul
if errorlevel 1 (
  echo [ERREUR] Node.js n'est pas installe ou absent du PATH.
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo [ERREUR] npm n'est pas disponible.
  pause
  exit /b 1
)

cd /d "%APP_DIR%"

if not exist "node_modules" (
  echo [INFO] Installation des dependances...
  call npm install
  if errorlevel 1 (
    echo [ERREUR] npm install a echoue.
    pause
    exit /b 1
  )
)

for /f "usebackq delims=" %%I in (`powershell -NoProfile -Command "$ip=(Get-NetIPAddress -AddressFamily IPv4 ^| Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' -and $_.PrefixOrigin -ne 'WellKnown' } ^| Select-Object -ExpandProperty IPAddress -First 1); if(-not $ip){$ip='IP_DU_PC'}; Write-Output $ip"`) do (
  set "LOCAL_IP=%%I"
)
if "%LOCAL_IP%"=="" set "LOCAL_IP=IP_DU_PC"

echo [INFO] Lancement smartphone en mode telephone...
echo [INFO] Le telephone doit etre sur le meme Wi-Fi que le PC.
echo [INFO] URL telephone: http://%LOCAL_IP%:5173
call npm run dev -- --host 0.0.0.0 --port 5173

endlocal
