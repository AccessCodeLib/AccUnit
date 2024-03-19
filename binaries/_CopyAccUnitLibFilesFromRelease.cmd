
if exist .\accunit\ (
  del /Q .\accunit\*
) else (
  mkdir .\accunit\
)
copy .\..\source\AccUnit\bin\Release\AccessCodeLib.*.tlb .\accunit\
copy .\..\source\AccUnit\bin\Release\AccessCodeLib.*.dll .\accunit\

timeout 3