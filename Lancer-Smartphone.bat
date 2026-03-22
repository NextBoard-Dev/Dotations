@echo off
setlocal

set "ROOT=%~dp0"
set "APP_DIR=%ROOT%smartphone"

if not exist "%APP_DIR%\package.json" (
  echo [ERREUR] Dossier smartphone introuvable: "%APP_DIR%"
  pause
  exit /b 1
)

where node >nul 2>nul
if errorlevel 1 (
  echo [ERREUR] Node.js n'est pas installe ou absent du PATH.
  echo Installe Node.js puis relance ce script.
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

echo [INFO] Lancement smartphone...
echo [INFO] URL: http://127.0.0.1:5173
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://127.0.0.1:5173"
call npm run dev -- --host 127.0.0.1 --port 5173

endlocal
