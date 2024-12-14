On Error Resume Next

Dim WshShell

Set WshShell = WScript.CreateObject("WScript.Shell")

Dim objFSO

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
overallRet = 0
strLTEMP = WshShell.ExpandEnvironmentStrings("%Temp%")
strProgFiles = WshShell.ExpandEnvironmentStrings("%SystemDrive%")

strFolderPath = strProgFiles & "\temp"

'Log File Creation

StrScriptName = Replace(WScript.ScriptName, "Folder_Permission.vbs", "")

strLogFile = strLTEMP & "\" & StrScriptName

Set LogFile = objFSO.CreateTextFile( strLogFile & ".log", True)

Logfile.WriteLine( Now() & " : Logging Started" & vbnewline)

Function CompleteInstall 

On Error Resume Next

If objFSO.FolderExists (strFolderPath) Then

 LogFile.WriteLine(Time() & " Folder Exists : Applying Permissions")

 intRet = WshShell.Run ("cacls """&strFolderPath&""" /E /T /G USERS:F",0,True)

  If True = ProcessRet(overallRet, intRet, "Folder Permission") Then

             Exit Function

  End If
  Else

   LogFile.WriteLine(Time() & " Folder Does not Exist")

End If


End Function



Call CompleteInstall

ProcessRet overallRet,overallRet,StrScriptName

Logfile.WriteLine(vbnewline & Now() & " : Logging Stopped")

Logfile.Close 



