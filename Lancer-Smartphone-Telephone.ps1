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

$localIp = (Get-NetIPAddress -AddressFamily IPv4 |
  Where-Object {
    $_.IPAddress -notlike "127.*" -and
    $_.IPAddress -notlike "169.254.*" -and
    $_.PrefixOrigin -ne "WellKnown"
  } |
  Select-Object -ExpandProperty IPAddress -First 1)

if (-not $localIp) {
  $localIp = "IP_DU_PC"
}

Write-Host "[INFO] Lancement smartphone en mode telephone..." -ForegroundColor Cyan
Write-Host "[INFO] PC et telephone doivent etre sur le meme Wi-Fi." -ForegroundColor Cyan
Write-Host "[INFO] URL telephone: http://$localIp`:5173" -ForegroundColor Green
npm run dev -- --host 0.0.0.0 --port 5173
