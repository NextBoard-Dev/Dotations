@echo off
setlocal
pushd "%~dp0" >nul
echo ========================================
echo CHECK REPO - DOTATIONS
echo ========================================
echo DOSSIER: "%~dp0"
for /f "delims=" %%i in ('git rev-parse --show-toplevel') do set ROOT=%%i
for /f "delims=" %%i in ('git branch --show-current') do set BRANCH=%%i
for /f "delims=" %%i in ('git remote get-url origin') do set ORIGIN=%%i
echo ROOT   : "%ROOT%"
echo BRANCHE: %BRANCH%
echo ORIGIN : %ORIGIN%
if /I not "%ORIGIN%"=="https://github.com/NextBoard-Dev/Dotations.git" (
  echo.
  echo [ALERTE] REMOTE INATTENDU POUR DOTATIONS
  popd >nul
  exit /b 1
)
echo.
echo OK - CONTEXTE DOTATIONS VALIDE
echo.
echo CONTROLES CRITIQUES...
node --check app.js || goto :fail
node --test tests/*.mjs || goto :fail
echo.
echo OK - CONTROLES CRITIQUES VALIDES
popd >nul
exit /b 0

:fail
echo.
echo [ECHEC] CONTROLES CRITIQUES DOTATIONS
popd >nul
exit /b 1
