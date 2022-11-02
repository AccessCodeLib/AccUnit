@ECHO OFF

REM Generate Interop.TLI.dll by forcing it to reference the .NET 2.0 runtime. Current versions of Visual Studio would use the .NET 4 runtime, causing problems with compilation down stream.
REM Error message otherwise:
REM The primary reference "D:\Projects\DotNetCommons\branches\accunit\src\Common.VBIDETools\bin\Debug\AccessCodeLib.Common.VBIDETools.dll" could not be resolved because it has an indirect dependency on the .NET Framework assembly "mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" which has a higher version "4.0.0.0" than the version "2.0.0.0" in the current target framework.

SET CurrDir=%~dp0
SET WinSdkPath=C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin

REM IMPORTANT
REM Make sure to have tied TlbImp.exe to use the .NET 2.0 runtime by adding <startup><supportedRuntime version="v2.0.50727"/></startup> to the TlbImp.exe.config!

"%WinSdkPath%\TlbImp.exe" "C:\Windows\System\TLBINF32.DLL" /out:%CurrDir%Interop.TLI.64.dll /namespace:TLI

PAUSE