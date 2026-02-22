# removebg/deps.ps1
# Installs the rembg Python package (with GPU support if available).
# Idempotent: checks if rembg is already importable before installing.

Write-Host "  [removebg] Checking dependencies..." -ForegroundColor Cyan

$installed = python -c "import rembg; print('ok')" 2>$null
if ($installed -eq "ok") {
    Write-Host "    OK  rembg is already installed" -ForegroundColor Green
    return
}

Write-Host "    Installing rembg[gpu] via pip..." -ForegroundColor Yellow
pip install "rembg[gpu]"

if ($LASTEXITCODE -ne 0) {
    Write-Host "    rembg[gpu] install failed; trying rembg (CPU only)..." -ForegroundColor Yellow
    pip install rembg
    if ($LASTEXITCODE -ne 0) {
        Write-Host "    ERROR: rembg installation failed. Make sure Python and pip are on your PATH." -ForegroundColor Red
    } else {
        Write-Host "    OK  rembg installed (CPU only)" -ForegroundColor Green
    }
} else {
    Write-Host "    OK  rembg[gpu] installed" -ForegroundColor Green
}
