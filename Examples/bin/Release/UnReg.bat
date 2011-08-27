@echo off

REM This two lines will ensure we move to the drive / folder that contains
REM this batch file, so we don't need to worry about the path to the dll
%~d0
cd "%~p0"

echo.
%windir%\Microsoft.NET\Framework\v2.0.50727\regasm /nologo /u Examples.dll
echo.

REM In case this run from Windows Explorer, ensure we get to view the
REM completion message
pause
