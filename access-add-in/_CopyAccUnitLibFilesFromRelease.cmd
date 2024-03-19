if exist .\lib\ (
  del /Q .\lib\*
) else (
  mkdir .\lib\
)
copy .\..\source\AccUnit\bin\Release\AccessCodeLib.*.tlb .\lib\
copy .\..\source\AccUnit\bin\Release\AccessCodeLib.*.dll .\lib\

timeout 3