on error resume next

StrDestinationLocation = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Python 3.6\"
set objFSO = createobject("Scripting.FileSystemObject")

S1 = StrDestinationLocation & "\Python 3.6 Manuals (32-bit).lnk"

If objFSO.FileExists(S1)Then
 objFSO.DeleteFile S1, True
End If
 

                 

