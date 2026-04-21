param(
    [string]$EntryPoint = "le_visage.py",
    [string]$ExeName = "DATACOLISA"
)

$ErrorActionPreference = "Stop"
$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$CodeDir = Join-Path $RepoRoot "code"

Set-Location $CodeDir

Write-Host ""
Write-Host "=== BUILD DATACOLISA PORTABLE ===" -ForegroundColor Cyan
Write-Host ""

python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --onefile `
  --name $ExeName `
  --icon "assets\colisa_fr.ico" `
  --add-data "assets;assets" `
  --collect-submodules openpyxl `
  --collect-submodules PySide6 `
  $EntryPoint

$distDir = Join-Path (Get-Location) "dist"
$exePath = Join-Path $distDir "$ExeName.exe"
$flagPath = Join-Path $distDir "portable.flag"
$dataDir = Join-Path $distDir "data"
$exportsDir = Join-Path $distDir "exports"
$releaseDir = Join-Path $distDir "${ExeName}_A_NE_PAS_TOUCHER"

if (-not (Test-Path $exePath)) {
    throw "Executable introuvable : $exePath"
}

"" | Out-File -FilePath $flagPath -Encoding utf8 -Force
New-Item -ItemType Directory -Force -Path $dataDir | Out-Null
New-Item -ItemType Directory -Force -Path $exportsDir | Out-Null

if (Test-Path $releaseDir) {
    Remove-Item -LiteralPath $releaseDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
Copy-Item -LiteralPath $exePath -Destination (Join-Path $releaseDir "$ExeName.exe") -Force
Copy-Item -LiteralPath $flagPath -Destination (Join-Path $releaseDir "portable.flag") -Force
New-Item -ItemType Directory -Force -Path (Join-Path $releaseDir "data") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $releaseDir "exports") | Out-Null

@"
DATACOLISA - version portable

Utilisation :
1. Ouvrir DATACOLISA.exe
2. Laisser portable.flag dans le meme dossier que l'exe
3. Le logiciel ecrit ses fichiers dans data\ et exports\

Distribution :
- copier tout le dossier DATACOLISA_A_NE_PAS_TOUCHER
- Python n'est pas necessaire sur le PC utilisateur
"@ | Out-File -FilePath (Join-Path $releaseDir "LIRE_MOI.txt") -Encoding utf8 -Force

Write-Host "portable.flag cree dans $distDir" -ForegroundColor Green
Write-Host "Dossiers data/ et exports/ crees" -ForegroundColor Green
Write-Host ""
Write-Host "=== BUILD TERMINE ===" -ForegroundColor Cyan
Write-Host "Executable : $exePath"
Write-Host "Dossier final a livrer : $releaseDir"
Write-Host ""
