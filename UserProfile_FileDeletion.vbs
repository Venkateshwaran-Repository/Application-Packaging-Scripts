
Const LocalDocumentsFolder = "C:\users\"
Set objShellApp = CreateObject("Shell.Application")
Set objShell = CreateObject("WScript.Shell")
set objFSO = createobject("Scripting.FileSystemObject")
set objFolder = objFSO.GetFolder(localdocumentsfolder)
on error resume next
Dim StrSourceFile

for each fldr in objFolder.SubFolders
  if not isexception(fldr.name) then

	StrSourceFile =  fldr & "\Documents\Micro Focus\Reflection\AddvantageVP_PRD.rdox"

	if(objFSO.FileExists(StrSourceFile)) Then	 
  
		objFso.DeleteFile(StrSourceFile)

	end if     		

  end if

next


Function isException(byval foldername)
 select case foldername
  case "All Users"
   isException = True
  case "Default User"
   isException = True
  case "LocalService"
   isException = True
  case "NetworkService"
   isException = True
  case "Administrator"
   isException = True
  case Else
   isException = False
 End Select
End Function
