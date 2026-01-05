# Stop on error
$ErrorActionPreference = "Stop"

# Paths
$projectRoot = Get-Location
$distExe = Join-Path $projectRoot "dist\Report Writer.exe"
$finalExe = Join-Path $projectRoot "Report Writer.exe"
$signtool = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe"

# Clean old builds
Write-Host "Cleaning old build artifacts..."
if (Test-Path build) { Remove-Item build -Recurse -Force }
if (Test-Path dist)  { Remove-Item dist  -Recurse -Force }

# Build using spec file (includes manifest for drag-and-drop fix)
Write-Host "Building executable with .spec file..."
pyinstaller ReportWriter.spec

# Copy EXE to project root
Write-Host "Copying executable to project root..."
Copy-Item $distExe $finalExe -Force

# Sign with timestamp
Write-Host "Signing executable..."
& $signtool sign `
  /fd SHA256 `
  /tr http://timestamp.digicert.com `
  /td SHA256 `
  /a `
  $finalExe

Write-Host "Build complete."
Write-Host "Executable ready at: $finalExe"