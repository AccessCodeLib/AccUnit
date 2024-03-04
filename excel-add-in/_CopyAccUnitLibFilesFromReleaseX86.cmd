
if exist .\lib\ (
 echo.
) else (
  mkdir .\lib
)

if exist .\lib\x86\ (
  del /Q .\lib\x86\*
) else (
  mkdir .\lib\x86\
)
copy .\..\source\AccUnit\bin\x86\Release\AccessCodeLib.*.tlb .\lib\x86\
copy .\..\source\AccUnit\bin\x86\Release\AccessCodeLib.*.dll .\lib\x86\
copy .\..\source\AccUnit\bin\x86\Release\*Interop*.dll .\lib\x86\

timeout 3