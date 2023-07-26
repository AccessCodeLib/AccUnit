
if exist .\lib\ (
 echo.
) else (
  mkdir .\lib
)

if exist .\lib\x64\ (
  del /Q .\lib\x64\*
) else (
  mkdir .\lib\x64\
)

copy .\..\source\AccUnit\bin\x64\Release\AccessCodeLib.*.tlb .\lib\x64\
copy .\..\source\AccUnit\bin\x64\Release\AccessCodeLib.*.dll .\lib\x64\
copy .\..\source\AccUnit\bin\x64\Release\*Interop*.dll .\lib\x64\

timeout 3