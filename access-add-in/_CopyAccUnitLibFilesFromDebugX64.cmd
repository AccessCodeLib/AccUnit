
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

copy .\..\source\AccUnit\bin\x64\Debug\AccessCodeLib.*.tlb .\lib\x64\
copy .\..\source\AccUnit\bin\x64\Debug\AccessCodeLib.*.dll .\lib\x64\
copy .\..\source\AccUnit\bin\x64\Debug\*Interop*.dll .\lib\x64\

timeout 3