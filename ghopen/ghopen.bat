@echo off
setlocal

git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo Not a git repository.
    exit /b 1
)

where gh >nul 2>&1
if errorlevel 1 goto :no_gh

:: Try to open the PR page first; if there is no PR on this branch, gh exits non-zero
echo Checking for PR...
gh pr view --web 2>nul
if not errorlevel 1 exit /b 0

:: No PR - open the repo at the current path and branch
echo Opening repo...
gh browse
exit /b 0

:no_gh
:: gh CLI not installed - parse origin remote and open manually
echo Opening repo...
for /f "tokens=*" %%i in ('git remote get-url origin 2^>nul') do set "REMOTE=%%i"
if "%REMOTE%"=="" (
    echo No origin remote found.
    echo Tip: install the GitHub CLI ^(gh^) for smarter GitHub navigation.
    exit /b 1
)

powershell -NoProfile -Command ^
  "$url = $env:REMOTE;" ^
  "$url = $url -replace '^git@github\.com:', 'https://github.com/';" ^
  "$url = $url -replace '\.git$', '';" ^
  "if ($url -match 'github\.com') { Start-Process $url; exit 0 }" ^
  "else { Write-Error ('Not a GitHub remote: ' + $url); exit 1 }"
exit /b %errorlevel%
