if exist .\lib\ (
  del /Q .\lib\*
) else (
  mkdir .\lib\
)

:: Framework:
copy .\..\source\AccUnit\bin\Release\AccUnit.dll* .\lib\
copy .\..\source\AccUnit\bin\Release\AccUnit.tlb .\lib\
copy .\..\source\AccUnit\bin\Release\AccessCodeLib.*.dll .\lib\

:: VBE Add-in:
copy .\..\vbe-add-In\AccUnit.VbeAddIn\bin\Release\AccUnit.VbeAddIn.dll* .\lib\
copy .\..\vbe-add-In\AccUnit.VbeAddIn\bin\Release\AccessCodeLib.Common.VbeUserControlHost.dll .\lib\

timeout 3