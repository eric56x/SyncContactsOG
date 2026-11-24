$ErrorActionPreference = "Stop"

$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir     = Split-Path -Parent $scriptDir
$docDir      = Join-Path $rootDir "doc"
$releasesDir = Join-Path $rootDir "releases"
$pyFile      = Join-Path $scriptDir "SyncContactsOG.py"
$iconFile    = Join-Path $scriptDir "SyncContactsOG.ico"

$pyContent = Get-Content $pyFile -Raw -Encoding UTF8
if ($pyContent -match '__version__\s*=\s*["'']([^"'']+)["'']') {
    $version = $Matches[1]
} else {
    throw "Impossible de détecter __version__ dans SyncContactsOG.py"
}

Write-Host "=== Compilation et packaging de SyncContactsOG v$version ===" -ForegroundColor Cyan

Get-Process -Name "SyncContactsOG" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 300

Get-ChildItem -Path $rootDir -Filter "SyncContactsOG_*_NoCreds.exe" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

$buildTemp = Join-Path "C:\Temp" "Build_SyncContactsOG_v$version"
if (Test-Path $buildTemp) { Remove-Item $buildTemp -Recurse -Force }
$distDir = Join-Path $buildTemp "dist"
$workDir = Join-Path $buildTemp "work"
$specDir = Join-Path $buildTemp "spec"
New-Item -ItemType Directory -Force -Path $distDir, $workDir, $specDir, $releasesDir | Out-Null

$pyInstallerArgs = @(
    "--noconfirm",
    "--onefile",
    "--windowed",
    "--name", "SyncContactsOG",
    "--distpath", $distDir,
    "--workpath", $workDir,
    "--specpath", $specDir,
    "--collect-all", "customtkinter"
)

if (Test-Path $iconFile) {
    $pyInstallerArgs += @("--icon", $iconFile, "--add-data", "$iconFile;.")
}

$pyInstallerArgs += $pyFile

& pyinstaller @pyInstallerArgs
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller a échoué (code $LASTEXITCODE)."
}

# 1. Copier l'exécutable principal à la racine et dans .\releases\
$builtExe   = Join-Path $distDir "SyncContactsOG.exe"
$finalExe   = Join-Path $rootDir "SyncContactsOG.exe"
$releaseExe = Join-Path $releasesDir "SyncContactsOG_v$version.exe"
Copy-Item -Path $builtExe -Destination $finalExe -Force
Copy-Item -Path $builtExe -Destination $releaseExe -Force

# 2. Créer l'archive source ZIP dans .\releases\
$zipPath = Join-Path $releasesDir "SyncContactsOG_v${version}_src.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $scriptDir "*") -DestinationPath $zipPath -Force

# 3. Exporter la documentation et le changelog dans .\releases\
$docFiles = @(
    @{ Src = Join-Path $docDir "SPECIFICATIONS_FR.md";   Dst = "SyncContactsOG_v${version}_SPECIFICATIONS_FR.md" },
    @{ Src = Join-Path $docDir "SPECIFICATIONS_FR.html"; Dst = "SyncContactsOG_v${version}_SPECIFICATIONS_FR.html" },
    @{ Src = Join-Path $docDir "SPECIFICATIONS_EN.md";   Dst = "SyncContactsOG_v${version}_SPECIFICATIONS_EN.md" },
    @{ Src = Join-Path $docDir "SPECIFICATIONS_EN.html"; Dst = "SyncContactsOG_v${version}_SPECIFICATIONS_EN.html" },
    @{ Src = Join-Path $docDir "BEST_PRACTICES.html";    Dst = "SyncContactsOG_v${version}_BEST_PRACTICES.html" },
    @{ Src = Join-Path $scriptDir "CHANGELOG.md";        Dst = "SyncContactsOG_v${version}_CHANGELOG.md" },
    @{ Src = Join-Path $docDir "CHANGELOG.html";         Dst = "SyncContactsOG_v${version}_CHANGELOG.html" }
)

foreach ($item in $docFiles) {
    if (Test-Path $item.Src) {
        Copy-Item -Path $item.Src -Destination (Join-Path $releasesDir $item.Dst) -Force
    }
}

# 4. Vérifier le dossier portable .\data
$dataDir = Join-Path $rootDir "data"
New-Item -ItemType Directory -Force -Path $dataDir | Out-Null

Remove-Item $buildTemp -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Compilation et release v$version terminées : $releaseExe" -ForegroundColor Green

