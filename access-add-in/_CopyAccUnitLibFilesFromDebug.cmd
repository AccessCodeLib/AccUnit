
if exist .\lib\ (
  del /Q .\lib\*
) else (
  mkdir .\lib
)

copy .\..\source\AccUnit\bin\Debug\AccessCodeLib.*.tlb .\lib\
copy .\..\source\AccUnit\bin\Debug\AccessCodeLib.*.dll .\lib\
copy .\..\source\AccUnit\bin\Debug\*Interop*.dll .\lib\

timeout 3