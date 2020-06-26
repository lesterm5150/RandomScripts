@echo off
setlocal


taskkill /F /FI "WINDOWTITLE eq Administrator: T*" 
rem /IM cmd.exe
REM call breaktimer.bat