
' ************* WEEK 3 - iLabs ***************

' VBScript: NetShareServer.vbs
' Written by: Alejandro Jaque
' Date: Jan 18, 2016
' Class: COMP230
' Professor: Ray Blankenship

Set fso = CreateObject("Scripting.FileSystemObject")
Set fileServ = GetObject("WinNT://vlab-PC1/LanmanServer,FileService")

If fso.folderexists("C:\public") then
   fso.deletefolder("C:\public")
   WScript.echo "Public folder deleted"
   End if
fso.CreateFolder("C:\Public")
fso.CopyFile "C:\Windows\Cursors\w*.*","C:\Public"

WScript.Echo vbCrLf & "Current Network Shares"
For Each sh In fileServ
    WScript.Echo sh.name
Next

Set share = fileServ.Create("FileShare", "PublicData")
share.path = "C:\Public"
share.MaxUserCount = 10
share.SetInfo

WScript.Echo vbCrLf
WScript.Echo vbCrLf & "New Network Shares"
For Each sh In fileServ
    WScript.Echo sh.name
Next

WScript.Echo vbCrLf & "\\vlab-PC1\PublicData Share will be Available for 60 seconds!!"
WScript.Sleep(60000)

fileServ.Delete "FileShare","PublicData"
fso.DeleteFolder "C:\Public",True

WScript.Echo vbCrLf & "End of Program"


' VBScript: NetShareClient.vbs
' Written by: Alejandro Jaque
' Date: Jan 18, 2016
' Class: COMP230
' Professor: Ray Blankenship

Set fso = CreateObject("Scripting.FileSystemObject")
Set networkObj = WScript.CreateObject("WScript.Network")

networkObj.MapNetworkDrive "X:","\\vlab-PC1\PublicData"
Set folder = fso.GetFolder("X:\")
Set files = folder.Files

WScript.Echo vbCrLf
WScript.Echo "Contents of Mapped Drive X:"
For Each item In files
 WScript.Echo item.Name
Next

networkObj.RemoveNetworkDrive "X:",True
WScript.Echo vbCrLf & "End of Program"


