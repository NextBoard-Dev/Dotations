@echo off
setlocal EnableDelayedExpansion

set "ROOT=%~dp0"
set "APP_DIR=%ROOT%smartphone"
set "LOCAL_IP=IP_DU_PC"

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

for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /R /C:"IPv4" /C:"Adresse IPv4"') do (
  for /f "tokens=* delims= " %%B in ("%%A") do (
    set "CANDIDATE=%%B"
    if not "!CANDIDATE:~0,4!"=="127." if not "!CANDIDATE:~0,8!"=="169.254." (
      set "LOCAL_IP=!CANDIDATE!"
      goto :ip_found
    )
  )
)
:ip_found

echo [INFO] Lancement smartphone en mode telephone...
echo [INFO] Le telephone doit etre sur le meme Wi-Fi que le PC.
echo [INFO] URL telephone: http://%LOCAL_IP%:5173
call npm run dev -- --host 0.0.0.0 --port 5173

endlocal
