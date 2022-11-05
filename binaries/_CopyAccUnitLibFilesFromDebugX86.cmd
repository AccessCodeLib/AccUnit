
if exist .\accunit\x86\ (
  del /Q .\accunit\x86\*
) else (
  mkdir .\accunit\x86\
)
copy .\..\source\AccUnit\bin\x86\Debug\AccessCodeLib.*.tlb .\accunit\x86\
copy .\..\source\AccUnit\bin\x86\Debug\AccessCodeLib.*.dll .\accunit\x86\
copy .\..\source\AccUnit\bin\x86\Debug\*Interop*.dll .\accunit\x86\

timeout 3