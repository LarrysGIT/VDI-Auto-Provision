@echo off
cd /d "%~dp0"

reg add "HKEY_CURRENT_USER\Software\Sysinternals\SDelete" /v "EulaAccepted" /t REG_DWORD /d 1 /f
sdelete.exe -q -z C:

echo %errorlevel%>sdelete.txt
