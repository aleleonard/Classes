Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists("C:\Data\DataFile.txt") Then
  fso.DeleteFile "C:\Data\DataFile.txt"
 End If