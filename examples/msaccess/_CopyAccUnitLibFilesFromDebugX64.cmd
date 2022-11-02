
if exist .\lib\ (
  del /Q .\lib\*
) else (
  mkdir .\lib
)

copy .\..\..\source\AccUnit\bin\X64\Debug\AccessCodeLib.*.tlb .\lib\
copy .\..\..\source\AccUnit\bin\X64\Debug\AccessCodeLib.*.dll .\lib\
copy .\..\..\source\AccUnit\bin\X64\Debug\*Interop*.dll .\lib\

timeout 3