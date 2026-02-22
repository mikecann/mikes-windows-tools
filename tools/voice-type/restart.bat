@echo off
echo Stopping any running voice-type instances...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*voice-type.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
echo Launching voice-type...
wscript.exe "%~dp0voice-type.vbs"
echo Done.
