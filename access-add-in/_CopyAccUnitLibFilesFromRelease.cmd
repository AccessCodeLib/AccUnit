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
copy .\..\vbe-add-In\AccUnit.VbeAddIn\bin\Release\AccessCodeLib.AccUnit.Extension.OpenAI.dll .\lib\
copy .\..\vbe-add-In\AccUnit.VbeAddIn\bin\Release\AccessCodeLib.AccUnit.Extension.OpenAI.dll.config .\lib\

copy .\..\vbe-add-In\AccUnit.VbeAddIn\bin\Release\Microsoft.Extensions.Configuration.dll .\lib\
copy .\..\vbe-add-In\AccUnit.VbeAddIn\bin\Release\Microsoft.Extensions.Configuration.Abstractions.dll .\lib\
copy .\..\vbe-add-In\AccUnit.VbeAddIn\bin\Release\Microsoft.Extensions.Configuration.UserSecrets.dll .\lib\

Newtonsoft.Json.dll

timeout 3