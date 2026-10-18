param(
  [Parameter(Mandatory = $true)]
  [string]$Email,
  [Parameter(Mandatory = $true)]
  [string]$Password,
  [string]$OutDir = "C:\Users\sebastien.duc\CLOUD\02_ARCHIVAGE PERSONNEL\DASHBOARDS\80. SNAPSHOTS ET SAUVEGARDES\03. DOTATIONS\DOTATIONS SNAPSHOTS"
)

$ErrorActionPreference = "Stop"

$projectUrl = "https://dphrvdhqhgycmllietuk.supabase.co"
$publishableKey = "sb_publishable_2wYXnIDj4-c8daQZW8D5hA_2Py6k7z6"
$edgeDataUrl = "$projectUrl/functions/v1/dotations-api/data"

if (-not (Test-Path -LiteralPath $OutDir)) {
  New-Item -ItemType Directory -Path $OutDir | Out-Null
}

$loginBody = @{
  email = $Email
  password = $Password
} | ConvertTo-Json -Depth 4

$loginHeaders = @{
  apikey = $publishableKey
  "Content-Type" = "application/json"
}

$loginRes = Invoke-RestMethod `
  -Method Post `
  -Uri "$projectUrl/auth/v1/token?grant_type=password" `
  -Headers $loginHeaders `
  -Body $loginBody

$token = [string]$loginRes.access_token
if ([string]::IsNullOrWhiteSpace($token)) {
  throw "Login Supabase impossible (token vide)."
}

$dataHeaders = @{
  apikey = $publishableKey
  Authorization = "Bearer $token"
  "Content-Type" = "application/json"
}

$payload = Invoke-RestMethod -Method Get -Uri $edgeDataUrl -Headers $dataHeaders

$stamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$outFile = Join-Path $OutDir ("backup_dotations_" + $stamp + ".json")

($payload | ConvertTo-Json -Depth 100) | Set-Content -LiteralPath $outFile -Encoding UTF8
Write-Output "BACKUP_OK:$outFile"
