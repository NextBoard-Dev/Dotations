param([switch]$Json)
$ErrorActionPreference = 'Stop'
$errors = [System.Collections.Generic.List[string]]::new()
$checked = 0
try {
    $root = Split-Path $PSScriptRoot -Parent
    $pdfRoot = [IO.Path]::GetFullPath((Join-Path $root 'data/pdf')) + [IO.Path]::DirectorySeparatorChar
    $data = Get-Content -LiteralPath (Join-Path $root 'data/local_state.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $data.documentsArchives) { throw 'Liste archives absente' }
    foreach ($archive in $data.documentsArchives) {
        $relative = [string]$archive.localPath
        if (-not $relative) { $relative = [string]$archive.pdfPath }
        if (-not $relative -or $relative -match '^(https?://|storage://)') {
            if (-not $archive.storagePath -and -not $archive.publicUrl) { $errors.Add("$($archive.id): aucun PDF reference") }
            continue
        }
        $file = [IO.Path]::GetFullPath((Join-Path $root $relative))
        if (-not $file.StartsWith($pdfRoot, [StringComparison]::OrdinalIgnoreCase)) { $errors.Add("$($archive.id): chemin PDF hors dossier autorise"); continue }
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { $errors.Add("$($archive.id): PDF local absent"); continue }
        $bytes = [IO.File]::ReadAllBytes($file)
        if ($bytes.Length -lt 8 -or [Text.Encoding]::ASCII.GetString($bytes,0,5) -ne '%PDF-') { $errors.Add("$($archive.id): entete PDF invalide"); continue }
        $tail = [Text.Encoding]::ASCII.GetString($bytes,[Math]::Max(0,$bytes.Length-2048),[Math]::Min(2048,$bytes.Length))
        if ($tail -notmatch '%%EOF') { $errors.Add("$($archive.id): PDF incomplet") }
        $checked++
    }
} catch { $errors.Add($_.Exception.Message) }
$result = @{ok=($errors.Count -eq 0); checked=$checked; errors=@($errors.ToArray())}
$result | ConvertTo-Json -Depth 4 -Compress
if ($errors.Count) { exit 1 }
exit 0
