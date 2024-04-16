@ECHO OFF
setlocal enabledelayedexpansion

:: ###################################################################
:: Config ClassGuit, ProgId, ...
:: -------------------------------------------------------------------

SET TargetRuntimeVersion=v4.0.30319

:: -------------------------------------------------------------------
:: VBE Add-in:
:: -----------

SET AddInAssemblyName=AccUnit.VbeAddIn
SET AddInAssemblyFile=%AddInAssemblyName%.dll

SET AddInClassGuid=F15F18C3-CA43-421E-9585-6A04F51C5786
SET AddInProgId=AccUnit.VbeAddIn.Connect
SET AddInFullClassName=AccessCodeLib.AccUnit.VbeAddIn.Connect
SET AddInProgVersion=0.2

:: -------------------------------------------------------------------
:: VBE ControlHost (VBE window)
:: ----------------------------

SET ControlHostAssemblyName=AccessCodeLib.Common.VbeUserControlHost
SET ControlHostAssemblyFile=%ControlHostAssemblyName%.dll

SET ControlHostClassGuid=030A1F2F-4E0B-4041-A7F5-C4C0B94BAF07
SET ControlHostProgId=AccLib.VbeUserControlHost
SET ControlHostFullClassName=AccessCodeLib.Common.VBIDETools.VbeUserControlHost
SET ControlHostProgVersion=1.0

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

SET AssemblyDir=%~dp0
SET AssemblyDir=%AssemblyDir:\=/%

SET AssemblyName=%AddInAssemblyName%
SET AssemblyFile=%AddInAssemblyFile%

SET ClassGuid=%AddInClassGuid%
SET ProgId=%AddInProgId%
SET FullClassName=%AddInFullClassName%
SET ProgVersion=%AddInProgVersion%

echo.
echo.
echo Insert registry data for %ProgId% (%AddInClassGuid%) ...
echo.
%windir%\system32\REG ADD "HKCU\Software\Classes\%ProgId%" /ve /t REG_SZ /d %FullClassName% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\%ProgId%\CLSID" /ve /t REG_SZ /d "{%ClassGuid%}" /f

%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}" /ve /t REG_SZ /d %FullClassName% /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\ProgId" /ve /t REG_SZ /d %ProgId% /reg:%bitness% /f

%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /ve /t REG_SZ /d "mscoree.dll" /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "ThreadingModel" /t REG_SZ /d "Both" /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "Class" /t REG_SZ /d %FullClassName% /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "Assembly" /t ^
								REG_SZ /d "%AssemblyName%, Version=%ProgVersion%, Culture=neutral, PublicKeyToken=null" /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "RuntimeVersion" /t ^
								REG_SZ /d %TargetRuntimeVersion% /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "CodeBase" /t ^
								REG_SZ /d "file:///%AssemblyDir%%AssemblyFile%" /reg:%bitness% /f

%windir%\system32\REG ADD "HKCU\Software\Microsoft\VBA\VBE\6.0\%AddInsFolder%\%ProgId%" /v "LoadBehavior" /t REG_DWORD /d "0x00000003" /f
%windir%\system32\REG ADD "HKCU\Software\Microsoft\VBA\VBE\6.0\%AddInsFolder%\%ProgId%" /v "FriendlyName" /t REG_SZ /d "AccUnit VBE Add-in" /f
%windir%\system32\REG ADD "HKCU\Software\Microsoft\VBA\VBE\6.0\%AddInsFolder%\%ProgId%" /v "Description" /t REG_SZ /d "AccUnit VBE Add-in" /f



:: ###################################################################
:: set registry data VBE Window Host
:: ---------------------------------

SET AssemblyName=%ControlHostAssemblyName%
SET AssemblyFile=%ControlHostAssemblyFile%

SET ClassGuid=%ControlHostClassGuid%
SET ProgId=%ControlHostProgId%
SET FullClassName=%ControlHostFullClassName%
SET ProgVersion=%ControlHostProgVersion%

echo.
echo.
echo Insert registry data for %ProgId% (%AddInClassGuid%)
echo.
%windir%\system32\REG ADD "HKCU\Software\Classes\%ProgId%" /ve /t REG_SZ /d %FullClassName% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\%ProgId%\CLSID" /ve /t REG_SZ /d "{%ClassGuid%}" /reg:%bitness% /f

%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}" /ve /t REG_SZ /d %FullClassName% /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\ProgId" /ve /t REG_SZ /d %ProgId% /reg:%bitness% /f

%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /ve /t REG_SZ /d "mscoree.dll" /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "ThreadingModel" /t REG_SZ /d "Both" /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "Class" /t REG_SZ /d %FullClassName% /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "Assembly" /t ^
								REG_SZ /d "%AssemblyName%, Version=%ProgVersion%, Culture=neutral, PublicKeyToken=null" /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "RuntimeVersion" /t ^
								REG_SZ /d %TargetRuntimeVersion% /reg:%bitness% /f
%windir%\system32\REG ADD "HKCU\Software\Classes\CLSID\{%ClassGuid%}\InprocServer32" /v "CodeBase" /t ^
								REG_SZ /d "file:///%AssemblyDir%%AssemblyFile%" /reg:%bitness% /f


endlocal
timeout 5
