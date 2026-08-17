@echo off
echo ==========================================
echo   ANNUAL LEAVE - DAILY SYNC TEST
echo ==========================================
cscript.exe //nologo "%~dp0RunAnnualLeaveSync.vbs"
set "RC=%ERRORLEVEL%"
if "%RC%"=="0" (echo DONG BO THANH CONG.) else (echo DONG BO THAT BAI. Ma loi: %RC%)
pause
exit /b %RC%
