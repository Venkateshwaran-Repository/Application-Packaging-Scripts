Const LocalDocumentsFolder = "C:\users\"
Set objShellApp = CreateObject("Shell.Application")
Set objShell = CreateObject("WScript.Shell")
set objFSO = createobject("Scripting.FileSystemObject")

set objFolder = objFSO.GetFolder(localdocumentsfolder)
on error resume next
Dim StrSourceLocation, StrSourceLocation1
Dim StrDestinationLocation, StrDestinationLocation1
Dim StrSourceFileName
Dim StrDestinationFileName 
Dim StrDestinationFileName1
  
for each fldr in objFolder.SubFolders
  if not isexception(fldr.name) then
		
		 StrDestinationLocation = fldr & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
		check= StrDestinationLocation &"\Google Chrome.lnk"

		 If (objFSO.FileExists(check)) Then

		ObjFso.DeleteFile StrDestinationLocation &"\Google Chrome.lnk"

		End If
	
                            		
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

