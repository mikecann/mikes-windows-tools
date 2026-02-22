@echo off
powershell -NoProfile -Command ^
    "Get-WmiObject Win32_Process | Where-Object { $_.Name -eq 'powershell.exe' -and $_.CommandLine -like '*taskmon.ps1*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; Write-Host ('Stopped PID ' + $_.ProcessId) }"
