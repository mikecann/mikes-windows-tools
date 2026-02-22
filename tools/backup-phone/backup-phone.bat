@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0backup-phone.ps1" %*
