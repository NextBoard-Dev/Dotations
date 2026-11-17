$ErrorActionPreference = 'Stop'
try {
    $root = Split-Path $PSScriptRoot -Parent
    $data = Get-Content -LiteralPath (Join-Path $root 'data/local_state.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $data.personnes) { throw 'Liste personnes absente' }
    $ids = @{}
    foreach ($person in $data.personnes) {
        if (-not $person.id -or $ids.ContainsKey([string]$person.id)) { throw 'Identifiant personne absent ou duplique' }
        $ids[[string]$person.id] = $true
    }
    foreach ($file in @('server.js', 'app.js')) {
        & node --check (Join-Path $root $file)
        if ($LASTEXITCODE -ne 0) { throw "Syntaxe invalide : $file" }
    }
    Write-Output 'STATE=OK; SYNTAX=OK'
    exit 0
} catch {
    Write-Output "STATE=KO : $($_.Exception.Message)"
    exit 1
}
