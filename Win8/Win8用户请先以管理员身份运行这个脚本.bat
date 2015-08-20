@echo off
copy /Y .\mscomctl.ocx %SystemRoot%\SysWOW64\
regsvr32 /s %SystemRoot%\SysWOW64\mscomctl.ocx
echo =========================================
echo Works finished.
echo Press any key to exit.
echo =========================================
pause