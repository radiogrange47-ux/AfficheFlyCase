$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$PythonCommand = Get-Command python -ErrorAction SilentlyContinue
$Python = $null
if ($PythonCommand -and $PythonCommand.Source -notmatch "\\WindowsApps\\") {
    $Python = $PythonCommand.Source
} else {
    $PythonRoot = Join-Path $env:LOCALAPPDATA "Programs\Python"
    $PythonInstallations = Get-ChildItem $PythonRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending
    foreach ($Installation in $PythonInstallations) {
        $Candidate = Join-Path $Installation.FullName "python.exe"
        if (Test-Path $Candidate) {
            $Python = $Candidate
            break
        }
    }
}

if (-not $Python) {
    throw "Python 3.10 ou plus récent est requis pour construire AfficheFlyCase."
}

$VenvPython = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    & $Python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "Impossible de créer l'environnement Python." }
}
$Python = $VenvPython

& $Python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "Échec de mise à jour de pip." }
& $Python -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "Échec d'installation des dépendances." }
& $Python -m PyInstaller --noconfirm --onefile --windowed --name AfficheFlyCase `
    --collect-all reportlab --collect-all openpyxl `
    --add-data "assets\battery_warning.png;assets" app.py
if ($LASTEXITCODE -ne 0) { throw "Échec de création de l'exécutable." }

Write-Host "Exécutable créé : $Root\dist\AfficheFlyCase.exe"
