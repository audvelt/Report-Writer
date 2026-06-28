# Build Script - Corrected Signing

Write-Host "=== BUILD REPORT WRITER ===" -ForegroundColor Cyan
Write-Host ""

# Step 1: Clean
Write-Host "Step 1: Cleaning..." -ForegroundColor Yellow
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue _build
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue dist
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue build
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue __pycache__
Remove-Item -Force -ErrorAction SilentlyContinue version_info.txt
Write-Host "  Done" -ForegroundColor Green
Write-Host ""

# Step 2: Check Python
Write-Host "Step 2: Checking Python..." -ForegroundColor Yellow
python --version
Write-Host ""

# Step 3: Check version file
Write-Host "Step 3: Reading version..." -ForegroundColor Yellow
$versionPath = Join-Path "_assets" "version.txt"
if (Test-Path $versionPath) {
    $version = (Get-Content $versionPath).Trim()
    Write-Host "  Version: $version" -ForegroundColor Green
} else {
    Write-Host "  ERROR: _assets/version.txt not found!" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Step 4: Build
Write-Host "Step 4: Building..." -ForegroundColor Yellow
pyinstaller ReportWriter.spec
Write-Host ""

# Step 5: Check result
Write-Host "Step 5: Checking result..." -ForegroundColor Yellow
$finalExe = Join-Path "dist" "Report Writer.exe"
if (Test-Path $finalExe) {
    $sizeInMB = [math]::Round((Get-Item $finalExe).Length / 1MB, 2)
    Write-Host "  SUCCESS!" -ForegroundColor Green
    Write-Host "  Size: $sizeInMB MB" -ForegroundColor Green
    Write-Host "  Location: $finalExe" -ForegroundColor Green
} else {
    Write-Host "  FAILED!" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Step 6: Sign with certificate from Windows store
Write-Host "Step 6: Signing executable..." -ForegroundColor Yellow

# Find signtool
$signtool = $null
$paths = @(
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe",
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe",
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.18362.0\x64\signtool.exe"
)

foreach ($p in $paths) {
    if (Test-Path $p) {
        $signtool = $p
        break
    }
}

if ($signtool) {
    Write-Host "  Using: $signtool" -ForegroundColor Cyan
    
    # Sign with timestamp (your original command)
    & $signtool sign `
      /fd SHA256 `
      /tr http://timestamp.digicert.com `
      /td SHA256 `
      /a `
      $finalExe
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Signed successfully!" -ForegroundColor Green
    } else {
        Write-Host "  Signature failed!" -ForegroundColor Red
        Write-Host "  (Build is still usable)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  signtool not found, skipping signature" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== BUILD COMPLETE ===" -ForegroundColor Green
Write-Host ""
Write-Host "Version: $version" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next: Test with 'cd dist' then '.\Report Writer.exe'" -ForegroundColor Cyan
Write-Host ""