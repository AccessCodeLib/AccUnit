
if exist .\accunit\x86\ (
  del /Q .\accunit\x86\*
) else (
  mkdir .\accunit\x86\
)
copy .\..\source\AccUnit\bin\x86\Release\AccessCodeLib.*.tlb .\accunit\x86\
copy .\..\source\AccUnit\bin\x86\Release\AccessCodeLib.*.dll .\accunit\x86\
copy .\..\source\AccUnit\bin\x86\Release\*Interop*.dll .\accunit\x86\

timeout 3