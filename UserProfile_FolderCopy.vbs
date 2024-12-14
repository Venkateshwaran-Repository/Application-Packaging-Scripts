Const LocalDocumentsFolder = "C:\Users\"
Set ObjShellApp = CreateObject("Shell.Application")
Set ObjShell = CreateObject("WScript.Shell")
Set ObjFSO = Createobject("Scripting.FileSystemObject")
Set ObjFolder = ObjFSO.GetFolder(localdocumentsfolder)

On Error Resume Next

Dim StrSourceLocation1, StrDestinationLocation1

StrSourceLocation1 = "C:\ProgramData\Genesys_Composer_8o1o54_601_EN\.eclipse\org.eclipse.platform_4.5.0_920333535_win32_win32_x86_64"

For each fldr in objFolder.SubFolders

   If not isexception(fldr.name) Then
		 
		 StrDestinationLocation1 = fldr & "\.eclipse"
		 ObjFSO.CopyFolder StrSourceLocation1, StrDestinationLocation1 & "\" 
	
	If not (objFSO.FolderExists(StrDestinationLocation1)) Then
		 
		 ObjFSO.CreateFolder (fldr & "\.eclipse")
		 StrDestinationLocation1 = fldr & "\.eclipse"
		 ObjFSO.CopyFolder StrSourceLocation1, StrDestinationLocation1 & "\" 
	End If	      		
   
   End If

Next     	 


Function isException(byval foldername)
 Select Case FolderName
  Case "All Users"
   isException = True
  Case "Default User"
   isException = True
  Case "LocalService"
   isException = True
  Case "NetworkService"
   isException = True
  Case "Administrator"
   isException = True
  Case Else
   isException = False
 End Select
End Function
