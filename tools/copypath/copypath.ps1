param(
    [string]$Path = $PWD.Path
)

# Resolve path if it exists to get the actual provider path 
# (handles Registry paths, relative paths, etc.)
$resolvedPath = $Path
if (Test-Path -Path $Path -IsValid) {
    try {
        $resolvedPath = (Resolve-Path $Path -ErrorAction Stop).ProviderPath
    } catch {
        # Fallback to just joining the path if it doesn't exist
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    }
}

Set-Clipboard -Value $resolvedPath

Write-Host "Copied: " -NoNewline
Write-Host $resolvedPath -ForegroundColor Green
