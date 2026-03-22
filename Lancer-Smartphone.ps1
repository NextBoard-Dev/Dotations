$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$appDir = Join-Path $root "smartphone"

if (-not (Test-Path (Join-Path $appDir "package.json"))) {
  Write-Host "[ERREUR] Dossier smartphone introuvable: $appDir" -ForegroundColor Red
  Read-Host "Appuie sur Entree pour fermer"
  exit 1
}

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  Write-Host "[ERREUR] Node.js n'est pas installe ou absent du PATH." -ForegroundColor Red
  Read-Host "Appuie sur Entree pour fermer"
  exit 1
}

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
  Write-Host "[ERREUR] npm n'est pas disponible." -ForegroundColor Red
  Read-Host "Appuie sur Entree pour fermer"
  exit 1
}

Set-Location $appDir

if (-not (Test-Path (Join-Path $appDir "node_modules"))) {
  Write-Host "[INFO] Installation des dependances..." -ForegroundColor Yellow
  npm install
}

Write-Host "[INFO] Lancement smartphone..." -ForegroundColor Cyan
Write-Host "[INFO] URL: http://127.0.0.1:5173" -ForegroundColor Cyan
Start-Job -ScriptBlock {
  Start-Sleep -Seconds 3
  Start-Process "http://127.0.0.1:5173"
} | Out-Null
npm run dev -- --host 127.0.0.1 --port 5173
