
Option Explicit
on error resume next
Dim oShell, oFilesys, oSleep
Dim iReturn, iReturn1
Dim Killexe

Set oShell = Createobject("wscript.shell")
Set oFileSys = CreateObject("Scripting.FileSystemObject")


Sub sleep(strSeconds) 
        Dim dteWait : dteWait = DateAdd("s", strSeconds, Now()) 
        Do Until (Now() > dteWait) 
        Loop 
End Sub


Killexe="AMSWindowsService.exe"


iReturn1=oShell.run("taskkill /F /IM " & Killexe & chr(34), 0, True)
