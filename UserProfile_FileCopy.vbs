Const LocalDocumentsFolder = "C:\Users\"
Set objShellApp = CreateObject("Shell.Application")
Set objShell = CreateObject("WScript.Shell")
set objFSO = createobject("Scripting.FileSystemObject")
set objFolder = objFSO.GetFolder(localdocumentsfolder)
on error resume next
Dim StrSourceLocation1
Dim StrDestinationLocation1, StrDestinationLocation2
Dim StrSourceFileName1, StrSourceFileName2

StrSourceLocation1 = "C:\Program Files (x86)\SunGard\AddVantageBackoffice"
StrSourceFileName1 = "AddvantageVP_PRD.rdox"

for each fldr in objFolder.SubFolders

   If not isexception(fldr.name) then
		 
		 StrDestinationLocation1 = fldr & "\Documents\Micro Focus\Reflection"
		 ObjFSO.CopyFile StrSourceLocation1 & "\" & StrSourceFileName1, StrDestinationLocation1 & "\"
	
	If not (objFSO.FolderExists(StrDestinationLocation1)) Then
		 
		 ObjFSO.CreateFolder (fldr & "\Documents\Micro Focus")
		 ObjFSO.CreateFolder (fldr & "\Documents\Micro Focus\Reflection")
   		 StrDestinationLocation1 = fldr & "\Documents\Micro Focus\Reflection"
		 ObjFSO.CopyFile StrSourceLocation1 & "\" & StrSourceFileName1, StrDestinationLocation1 & "\"
	
	end If	      		
   
   end If

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
