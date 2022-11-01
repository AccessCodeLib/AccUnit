@echo off

SET BaseDir=%0
SET BaseDir=%BaseDir:"=%
SET BaseDir=%BaseDir:\SubWCRev.cmd=%

%BaseDir%\SubWCRev.exe %1 %2 %3
:: IF /I %PROCESSOR_ARCHITECTURE% EQU x86 (%BaseDir%\SubWCRev_x86.exe %1 %2 %3) ELSE %BaseDir%\SubWCRev.exe %1 %2 %3

pause
