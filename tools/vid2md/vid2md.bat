@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Sta -File "%~dp0vid2md.ps1" %*
