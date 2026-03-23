$ErrorActionPreference = "Stop"

try {
  $root = Split-Path -Parent $MyInvocation.MyCommand.Path
  $appDir = Join-Path $root "smartphone"

  Write-Host "[INFO] Script: Lancer-Smartphone-Telephone.ps1" -ForegroundColor Cyan
  Write-Host "[INFO] Dossier app: $appDir" -ForegroundColor Cyan

  if (-not (Test-Path (Join-Path $appDir "package.json"))) {
    throw "Dossier smartphone introuvable: $appDir"
  }

  if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    throw "Node.js n'est pas installe ou absent du PATH."
  }

  if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    throw "npm n'est pas disponible."
  }

  Set-Location $appDir

  if (-not (Test-Path (Join-Path $appDir "node_modules"))) {
    Write-Host "[INFO] Installation des dependances..." -ForegroundColor Yellow
    npm install
  }

  $localIp = ""
  $ipconfigLines = ipconfig | Select-String -Pattern "IPv4|Adresse IPv4"
  foreach ($line in $ipconfigLines) {
    if ($line -match "(\d{1,3}(?:\.\d{1,3}){3})") {
      $candidate = $Matches[1]
      if ($candidate -notlike "127.*" -and $candidate -notlike "169.254.*") {
        $localIp = $candidate
        break
      }
    }
  }
  if (-not $localIp) {
    $localIp = "IP_DU_PC"
  }

  Write-Host "[INFO] Lancement smartphone en mode telephone..." -ForegroundColor Cyan
  Write-Host "[INFO] PC et telephone doivent etre sur le meme Wi-Fi." -ForegroundColor Cyan
  Write-Host "[INFO] URL telephone: http://$localIp`:5173" -ForegroundColor Green
  npm run dev -- --host 0.0.0.0 --port 5173
}
catch {
  Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "[AIDE] Si le script est bloque par la policy, lance cette commande :" -ForegroundColor Yellow
  Write-Host "powershell -NoProfile -ExecutionPolicy Bypass -File ""$PSScriptRoot\Lancer-Smartphone-Telephone.ps1""" -ForegroundColor Yellow
  Read-Host "Appuie sur Entree pour fermer"
  exit 1
}
