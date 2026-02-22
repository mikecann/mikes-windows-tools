# transcribe/deps.ps1
# Checks that the large binaries required by transcribe are present in C:\dev\tools.
# These cannot be auto-downloaded; this script just tells you what is missing.

$ToolsDir = "C:\dev\tools"
$ok = $true

Write-Host "  [transcribe] Checking dependencies..." -ForegroundColor Cyan

foreach ($bin in @("ffmpeg.exe", "faster-whisper-xxl.exe")) {
    $path = Join-Path $ToolsDir $bin
    if (Test-Path $path) {
        Write-Host "    OK  $bin" -ForegroundColor Green
    } else {
        Write-Host "    MISSING  $bin  (expected at $path)" -ForegroundColor Red
        $ok = $false
    }
}

$modelsDir = Join-Path $ToolsDir "_models"
if (Test-Path $modelsDir) {
    Write-Host "    OK  _models\" -ForegroundColor Green
} else {
    Write-Host "    MISSING  _models\  (expected at $modelsDir)" -ForegroundColor Red
    $ok = $false
}

if (-not $ok) {
    Write-Host ""
    Write-Host "    transcribe needs large binaries that must be downloaded manually:" -ForegroundColor Yellow
    Write-Host "      ffmpeg.exe         https://ffmpeg.org/download.html" -ForegroundColor Yellow
    Write-Host "      faster-whisper-xxl.exe  https://github.com/Purfview/whisper-standalone-win/releases" -ForegroundColor Yellow
    Write-Host "    Place them in $ToolsDir and create a $modelsDir folder." -ForegroundColor Yellow
}
