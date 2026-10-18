# Builds the deployable AppSail folder for Zoho Catalyst (managed Python runtime).
#
# Catalyst does NOT run "pip install" for managed runtimes: every dependency
# must be present in the build path. It also runs Linux, so the vendored wheels
# have to be Linux wheels even though we are building on Windows.
#
# If a compiled Flutter web app is present it is copied into build/web, so this
# single AppSail service serves both the UI and the API from one origin. That
# removes cross-origin requests, which the Catalyst gateway does not allow
# through for preflighted (JSON POST) calls.
#
# Run from the attendance_backend directory:
#   .\build_appsail.ps1
#   .\build_appsail.ps1 -FlutterWebDir ..\attendance_app\build\web

param(
    [string]$FlutterWebDir = ""
)

$ErrorActionPreference = "Stop"

# Must match the AppSail stack (python_3_12).
$PythonVersion = "3.12"
$Platform = "manylinux2014_x86_64"
$BuildDir = "build"

$python = Join-Path $PSScriptRoot "venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    $python = "python"
}

Write-Host "==> Cleaning $BuildDir" -ForegroundColor Cyan
if (Test-Path $BuildDir) {
    Remove-Item $BuildDir -Recurse -Force
}
New-Item -ItemType Directory -Path $BuildDir | Out-Null

Write-Host "==> Vendoring Linux wheels for Python $PythonVersion" -ForegroundColor Cyan
& $python -m pip install `
    --target $BuildDir `
    --platform $Platform `
    --python-version $PythonVersion `
    --implementation cp `
    --only-binary=:all: `
    --upgrade `
    -r requirements.txt

if ($LASTEXITCODE -ne 0) {
    throw "pip install failed"
}

Write-Host "==> Copying app source" -ForegroundColor Cyan
Copy-Item main.py (Join-Path $BuildDir "main.py") -Force
Copy-Item requirements.txt (Join-Path $BuildDir "requirements.txt") -Force

# Bundle the compiled Flutter app as build/web. Default to the sibling app
# repo's release output.
if (-not $FlutterWebDir) {
    $FlutterWebDir = Join-Path (Split-Path $PSScriptRoot -Parent) "attendance_app\build\web"
}

if (Test-Path (Join-Path $FlutterWebDir "index.html")) {
    Write-Host "==> Bundling Flutter web app from $FlutterWebDir" -ForegroundColor Cyan
    $webDest = Join-Path $BuildDir "web"
    New-Item -ItemType Directory -Path $webDest | Out-Null
    Copy-Item (Join-Path $FlutterWebDir "*") $webDest -Recurse -Force
}
else {
    Write-Host "==> No Flutter web build at $FlutterWebDir - deploying the API only" -ForegroundColor Yellow
    Write-Host "    Build it with: flutter build web --release --dart-define=API_BASE_URL=<url>" -ForegroundColor Yellow
}

Write-Host "==> Sanity checks" -ForegroundColor Cyan
$so = Get-ChildItem $BuildDir -Recurse -Include *.so |
    Select-Object -First 1
if (-not $so) {
    throw "No .so files found - the Linux wheels were not vendored."
}
$pyd = Get-ChildItem $BuildDir -Recurse -Include *.pyd
if ($pyd) {
    throw "Found Windows .pyd binaries - the app would fail on Catalyst: $($pyd[0].Name)"
}
if (-not (Test-Path (Join-Path $BuildDir "main.py"))) {
    throw "main.py is missing from the build folder."
}
if (Test-Path (Join-Path $BuildDir "web")) {
    if (-not (Test-Path (Join-Path $BuildDir "web\index.html"))) {
        throw "build\web exists but index.html is missing - the Flutter bundle is incomplete."
    }
    $webFiles = (Get-ChildItem (Join-Path $BuildDir "web") -Recurse -File).Count
    Write-Host "    Flutter web bundle: $webFiles files" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "Build folder ready: $BuildDir" -ForegroundColor Green
Write-Host "Set this as the AppSail buildPath and deploy." -ForegroundColor Green
