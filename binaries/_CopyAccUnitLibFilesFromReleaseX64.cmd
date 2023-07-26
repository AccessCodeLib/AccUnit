if exist .\accunit\x64\ (
  del /Q .\accunit\x64\*
) else (
  mkdir .\accunit\x64\
)

copy .\..\source\AccUnit\bin\x64\Release\AccessCodeLib.*.tlb .\accunit\x64\
copy .\..\source\AccUnit\bin\x64\Release\AccessCodeLib.*.dll .\accunit\x64\
copy .\..\source\AccUnit\bin\x64\Release\*Interop*.dll .\accunit\x64\

timeout 3