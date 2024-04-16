@ECHO OFF
setlocal enabledelayedexpansion

:: ###################################################################
:: Config ClassGuit, ProgId, ...
:: -------------------------------------------------------------------
::
:: -------------------------------------------------------------------
:: VBE Add-in:
:: -----------

SET AddInClassGuid=F15F18C3-CA43-421E-9585-6A04F51C5786
SET AddInProgId=AccUnit.VbeAddIn.Connect

:: -------------------------------------------------------------------
:: VBE ControlHost (VBE window)
:: ----------------------------

SET ControlHostClassGuid=030A1F2F-4E0B-4041-A7F5-C4C0B94BAF07
SET ControlHostProgId=AccLib.VbeUserControlHost

:: ###################################################################
:: Find Office bitness
:: -------------------

set "key64=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
set "key32=HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
set "bitness="

rem Prüfen für 64-Bit Office unter 64-Bit und 32-Bit Windows
reg query "%key64%" /v "Platform" >nul 2>&1
if !errorlevel! equ 0 (
    for /f "tokens=3" %%a in ('reg query "%key64%" /v "Platform"') do set "bitness=%%a"
)

rem Prüfen für 32-Bit Office unter 64-Bit Windows
if not defined bitness (
    reg query "%key32%" /v "Platform" >nul 2>&1
    if !errorlevel! equ 0 (
        for /f "tokens=3" %%a in ('reg query "%key32%" /v "Platform"') do set "bitness=%%a"
    )
)

if defined bitness (
    set "bitness=!bitness:x86=32!"
    set "bitness=!bitness:x=!"
) else (
    set /p bitness="Please insert Office bitness (32 or 64):"
)
echo ------------------------
echo  Office bitness: !bitness! bit
echo ------------------------

:: ###################################################################
:: set registry data (bitness)
:: ---------------------------

if %bitness%==64 (
    SET AddInsFolder=Addins64
) else (
    SET AddInsFolder=Addins
)


:: ###################################################################
:: set registry data VBE Add-in
:: ----------------------------

if %bitness%==64 (
    SET ClsIdFolder=CLSID
    SET AddInsFolder=Addins64
) else (
    SET ClsIdFolder=CLSID
    SET AddInsFolder=Addins
)

SET ClassGuid=%AddInClassGuid%
SET ProgId=%AddInProgId%

echo.
echo.
echo Remove registry data for %ProgId% (%AddInClassGuid%) ...
echo.
%windir%\system32\REG DELETE "HKCU\Software\Classes\%ProgId%" /f
%windir%\system32\REG DELETE "HKCU\Software\Classes\CLSID\{%ClassGuid%}" /reg:%bitness% /f
%windir%\system32\REG DELETE "HKCU\Software\Microsoft\VBA\VBE\6.0\%AddInsFolder%\%ProgId%" /f


:: ###################################################################
:: set registry data VBE Window Host
:: ---------------------------------

SET ClassGuid=%ControlHostClassGuid%
SET ProgId=%ControlHostProgId%

echo.
echo.
echo Remove registry data for %ProgId% (%AddInClassGuid%) ...
echo.
%windir%\system32\REG DELETE "HKCU\Software\Classes\%ProgId%" /f
%windir%\system32\REG DELETE "HKCU\Software\Classes\CLSID\{%ClassGuid%}" /reg:%bitness% /f

endlocal
timeout 5