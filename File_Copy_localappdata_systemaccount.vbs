On Error Resume Next
Dim objFSO
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
Set ObjFso = CreateObject("Scripting.FileSystemObject")  
Set WshProcEnv = WshShell.Environment("Process")
Appdata = WshShell.ExpandEnvironmentStrings("%appdata%")
LocalAppdata = WshShell.ExpandEnvironmentStrings("%localappdata%")


Dim StrSourceLocation
Dim StrDestinationLocation
Dim StrSourceFileName
Dim StrDestinationFileName 

  
StrSourceLocation = "C:\ProgramData\SAS\EnterpriseGuide"

StrSourceFileName = "EGOptions.xml"  

StrDestinationLocation = "Appdata\SAS\EnterpriseGuide\8"

ObjFso.CreateFolder (StrDestinationLocation & "\File")

ObjFso.CopyFile StrSourceLocation & "\" & StrSourceFileName, StrDestinationLocation & "\EGOptions.xml"

