# deps.ps1 -- dependency check + first-time build for taskmon

Write-Host "  [taskmon] Checking dependencies..." -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# nvml.dll -- ships with NVIDIA drivers, needed for GPU monitoring
# ---------------------------------------------------------------------------
$nvml = 'C:\Windows\System32\nvml.dll'
if (Test-Path $nvml) {
    Write-Host "  [taskmon] nvml.dll found -- NVIDIA GPU monitoring available." -ForegroundColor Green
} else {
    Write-Host "  [taskmon] WARN: nvml.dll not found at $nvml" -ForegroundColor Yellow
    Write-Host "            GPU monitoring will be unavailable (install NVIDIA drivers to enable)." -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# MSBuild -- required to compile taskmon.csproj
# ---------------------------------------------------------------------------
$msbuild = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe'
if (-not (Test-Path $msbuild)) {
    Write-Host "  [taskmon] ERROR: MSBuild.exe not found at $msbuild" -ForegroundColor Red
    Write-Host "            .NET Framework 4 is required. It ships with Windows 10/11." -ForegroundColor DarkGray
    exit 1
}
Write-Host "  [taskmon] MSBuild found." -ForegroundColor Green

# ---------------------------------------------------------------------------
# .NET 4.8 Developer Pack -- optional, silences the MSB3644 build warning.
# Without it, MSBuild falls back to GAC DLLs (works fine, just noisy).
# ---------------------------------------------------------------------------
$refPath = 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework'
$hasRefAsm = (Test-Path "$refPath\v4.8") -or (Test-Path "$refPath\v4.8.1")
if (-not $hasRefAsm) {
    Write-Host "  [taskmon] NOTE: .NET Framework targeting pack not installed." -ForegroundColor DarkGray
    Write-Host "            Build will show warning MSB3644 (harmless -- MSBuild uses GAC fallback)." -ForegroundColor DarkGray
    Write-Host "            To silence it: install the .NET 4.8 Developer Pack from https://aka.ms/net48devpack" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Build taskmon.dll if it doesn't exist yet (first install on a fresh clone)
# ---------------------------------------------------------------------------
$dll = Join-Path $env:LOCALAPPDATA 'taskmon\taskmon.dll'
if (Test-Path $dll) {
    Write-Host "  [taskmon] taskmon.dll already built -- skipping build." -ForegroundColor Green
} else {
    Write-Host "  [taskmon] taskmon.dll not found -- building now..." -ForegroundColor Yellow
    $proj = Join-Path $PSScriptRoot 'taskmon.csproj'
    & $msbuild $proj /nologo /v:minimal /p:Configuration=Release
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  [taskmon] ERROR: Build failed. See errors above." -ForegroundColor Red
        exit 1
    }
    Write-Host "  [taskmon] Build succeeded: $dll" -ForegroundColor Green
}
