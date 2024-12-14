

strPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\shortcut"

DeleteFolder strPath

Function DeleteFolder(strFolderPath)
Dim objFSO, objFolder
Set objFSO = CreateObject ("Scripting.FileSystemObject")
If objFSO.FolderExists(strFolderPath) Then
	objFSO.DeleteFolder strFolderPath, True
End If
Set objFSO = Nothing
End Function
