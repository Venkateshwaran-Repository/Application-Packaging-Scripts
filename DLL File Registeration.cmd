CD /d "%~dp0" 

for /f "tokens=*" %%i in ('dir *.dll;*.ocx /b /s') do call :REGIT "%%i" 

goto :eof 

:REGIT 

echo register %1 

regsvr32 /s %1 

goto:eof 
