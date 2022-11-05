if exist .\accunit\x64\ (
  del /Q .\accunit\x64\*
) else (
  mkdir .\accunit\x64\
)

copy .\..\source\AccUnit\bin\x64\Debug\AccessCodeLib.*.tlb .\accunit\x64\
copy .\..\source\AccUnit\bin\x64\Debug\AccessCodeLib.*.dll .\accunit\x64\
copy .\..\source\AccUnit\bin\x64\Debug\*Interop*.dll .\accunit\x64\

timeout 3