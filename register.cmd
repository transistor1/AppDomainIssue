set REGASM=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe
set COM1=App_Domain_Issue_1\bin\Debug\App_Domain_Issue_1.dll
set COM2=App_Domain_Issue_2\bin\Debug\App_Domain_Issue_2.dll

echo %COM1%
echo %COM2%

if "%1"=="/u" goto :unregister

"%REGASM%"  /tlb /codebase "%COM1%"
"%REGASM%"  /tlb /codebase "%COM2%"

exit /b


:unregister
"%REGASM%"  /u /tlb /codebase "%COM1%"
"%REGASM%"  /u /tlb /codebase "%COM2%"