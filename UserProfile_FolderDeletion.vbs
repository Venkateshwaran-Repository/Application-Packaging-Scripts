
Const LocalDocumentsFolder = "C:\users\"
Set ObjShellApp = CreateObject("Shell.Application")
Set ObjShell = CreateObject("WScript.Shell")
Set ObjFSO = Createobject("Scripting.FileSystemObject")
Set ObjFolder = ObjFSO.GetFolder(localdocumentsfolder)

On Error Resume Next

Dim StrSourceFolder1

for each fldr in ObjFolder.SubFolders
  if not isexception(fldr.name) then

	StrSourceFolder1 =  fldr & "\Documents\PowerShell\Modules\Az.SecurityInsights"
	

	if(ObjFSO.FolderExists(StrSourceFolder1)) Then	 
  
		ObjFso.DeleteFolder(StrSourceFolder1)

	end if     		
	
	if(ObjFSO.FolderExists(StrSourceFolder2)) Then	 
  
		ObjFso.DeleteFolder(StrSourceFolder2)

	end if
	
	if(ObjFSO.FolderExists(StrSourceFolder3)) Then	 
  
		ObjFso.DeleteFolder(StrSourceFolder3)

	end if     		
	
	if(ObjFSO.FolderExists(StrSourceFolder4)) Then	 
  
		ObjFso.DeleteFolder(StrSourceFolder4)

	end if     		
	
  
  end if

next


Function isException(byval foldername)
 Select Case foldername
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
