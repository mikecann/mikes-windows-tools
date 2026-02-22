# backup-phone/deps.ps1
# Installs Python packages needed for HEIC -> WebP conversion.
# Idempotent: checks each package before installing.

Write-Host "  [backup-phone] Checking dependencies..." -ForegroundColor Cyan

$packages = @(
    @{ Import = "PIL";          Pip = "Pillow"       },
    @{ Import = "pillow_heif";  Pip = "pillow-heif"  }
)

foreach ($pkg in $packages) {
    $ok = python -c "import $($pkg.Import); print('ok')" 2>$null
    if ($ok -eq "ok") {
        Write-Host "    OK  $($pkg.Pip) is already installed" -ForegroundColor Green
    } else {
        Write-Host "    Installing $($pkg.Pip) via pip..." -ForegroundColor Yellow
        pip install $pkg.Pip
        if ($LASTEXITCODE -ne 0) {
            Write-Host "    ERROR: Failed to install $($pkg.Pip). Make sure Python and pip are on your PATH." -ForegroundColor Red
        } else {
            Write-Host "    OK  $($pkg.Pip) installed" -ForegroundColor Green
        }
    }
}
