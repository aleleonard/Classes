
' ******** WEEK 6 - Lesson **********


'  =============== FILE MANAGEMENT ===============

' To access a computer's file system object, we use the File System Object (FSO)
dim objFSO
set objFSO = CreateObject("Scripting.FileSystemObject")

' Checks to see if Notepad's file exists in the window folder
if objFSO.FileExists("c:\windows\notepad.exe") then
   wscript.echo "File exists"
 else
   wscript.echo "File does not exists"
end if

' Retrieves a file's version information
wscript.echo
objFSO.GetFileVersion("c:\windows\systems32\scrrun.dll")

' Copies all the .log files in the windows folder into Testfolder folder
const OverWriteFiles = True
set objFSO = CreateObject("Scripting.FileSystemObject")

if not objFSO.FolderExists("c:\Testfolder") then
   objFSO.CreateFolder("c:\TestFolder")
end if

call objFSO.CopyFile("c:\windows\*.log" , "c:\Testfolder" , OverWriteFiles)
wscript.echo "Copy complete"

' Moves log files from the testfolder to testfolder 2
set objFSO = CreateObject("Scripting.FileSystemObject")

if not objFSO.FolderExists("c:\Testfolder\Testfolder2") then
objFSO.CreateFolder("c:\Testfolder\Testfolder2")
end if

call objFSO.MoveFile("c:\Testfolder\*.log" , "c:\Testfolder\Testfolder2")
wscript.echo "Move complete"

' Deletes all the .log files in the Tesfolder2 folder
set objFSO = CreateObject("Scripting.FileSystemObject")

call objFSO.DeleteFile("c:\Testfolder\Testfolder2\*.log")
wscript.echo "Delete complete"

' Lists key file properties
set objFSO = CreateObject("Scripting.FileSystemObject")

if objFSO.FolderExists("c:\Testfolder") then

Set objFile = objFSO.GetFile("c:\windows\systems32\scrrun.dll")
wscript.echo "Date Created: " & objFile.DateCreated
wscript.echo "Date Modified: " & objFile.DateLastModified
wscript.echo "Drive: " & objFile.Drive
wscript.echo "Name: " & objFile.Name
wscript.echo "Path: " & objFile.Path
wscript.echo "Size: " & objFile.Size
wscript.echo "Type: " & objFile.Type
else
wscript.echo "File can't be found"
end if

'  =============== FOLDER MANAGEMENT ===============

' Checks to see if Folder exists in the window folder
set objFSO = CreateObject("Scripting.FileSystemObject")

if objFSO.FolderExists("c:\windows") then
   wscript.echo "Folder exists"
 else
   wscript.echo "File does not exists"
end if

' Creates a Folder
set objFSO = CreateObject("Scripting.FileSystemObject")

if objFSO.FolderExists("c:\testfolder") then
   wscript.echo "Folder already exists"
 else
 call objFSO.CreateFolder("c:\testfolder")
 wscript.echo "Folder was created"
end if

' Copies a Folder and its contents
const OverWriteFiles = True
set objFSO = CreateObject("Scripting.FileSystemObject")

if objFSO.FolderExists("c:\testfolder") then
call objFSO.CopyFolder("c:\testfolder", "c:\testfolder2",OverWriteFiles)
wscript.echo "New Folder was created"
else
wscript.echo "No folder to copy"
end If

' Moves or Renames a folder
set objFSO = CreateObject("Scripting.FileSystemObject")

if objFSO.FolderExists("c:\testfolder") then
if objFSO.FolderExists("c:\testfolder2") then
call objFSO.MoveFolder("c:\testfolder2", "c:\testfolder")
wscript.echo "Folder was moved"
else
wscript.echo "Move from folder missing"
else
wscript.echo "No move-to folder"
end If

' Deletes a folder and its contents
set objFSO = CreateObject("Scripting.FileSystemObject")

if objFSO.FolderExists("c:\testfolder") then
call objFSO.DeleteFolder("c:\testfolder\testfolder2")
wscript.echo "folder deleted"
else
wscript.echo "folder did not exist"
end if

' Folder properties
set objFSO = CreateObject("Scripting.FileSystemObject")

if objFSO.FolderExists("c:\testfolder") then
set objFolder = objFSO.GetFolder("c:\testfolder")
wscript.echo "date created" & objFolder.DateCreated
wscript.echo "date modified" & objFolder.DateLastModified
wscript.echo "drive" & objFolder.Drive
wscript.echo "Name" & objFolder.Name
wscript.echo "path" & objFolder.Path
wscript.echo "size" & objFolder.Size
wscript.echo "Type" & objFolder.Type
else
wscript.echo "folder can't be found"
end If


'  =============== DRIVES MANAGEMENT ===============

' List the drives on a computer
set objFSO = CreateObject("Scripting.FileSystemObject")
set colDrives = objFSO.Drives
for each objDrive in colDrives
wscript.echo "drive: " & objDrive.driveLetter & "Size: " & objDrive.totalSize/1024*2 & "MB"
next

' Drives properties
set objFSO = CreateObject("Scripting.FileSystemObject")
set objDrive = objFSO.Drives("c:")
wscript.echo "Size: " & objDrive.totalSize/1024*2 & "MB"
wscript.echo "Available Space: " & objDrive.availableSpace/1024*2 & "MB"
wscript.echo "Drive Type: " & objDrive.driveType
wscript.echo "Free Space: " & objDrive.freeSpace/1024*2 & "MB"
wscript.echo "File System: " & objDrive.fileSystem
wscript.echo "Is Ready: " & objDrive.isReady
wscript.echo "Path: " & objDrive.Path
wscript.echo "Root Folder: " & objDrive.rootFolder
wscript.echo "Serial Number: " & objDrive.serialNumber
wscript.echo "Share Name: " & objDrive.shareName
wscript.echo "Volume Name: " & objDrive.volumeName


'  =============== READING AND WRITING FILES ===============

' Create a text file
const OverWrite = True
set objFSO = CreateObject("Scripting.FileSystemObject")

set objFile = objFSO.CreateTextFile("c:\testfolder\scriptTest.txt",OverWrite)


' Write to a text file
const ForWriting = 2
set objFSO = CreateObject("Scripting.FileSystemObject")

set objFile = objFSO.CreateTextFile("c:\testfolder\scriptTest.txt",ForWriting)
objFile.Write("This is on line 1")
objFile.Write("This is on line 1 also")
objFile.WriteBlankLines(1)
objFile.WriteLine("This is on line 2")
objFile.WriteLine("This is on line 3")
objFile.WriteBlankLines(2)
objFile.WriteLine("This is on line 6")
objFile.Close

' Appends data to an existing text file
const ForAppending = 8
set objFSO = CreateObject("Scripting.FileSystemObject")

set objFile = objFSO.OpenTextFile("c:\testfolder\scriptTest.txt",ForAppending)
objFile.WriteLine("I'm appended on line 7")
objFile.WriteLine("I'm appended on line 8")
objFile.Close

' Reads data from a text file
const ForReading = 1
set objFSO = CreateObject("Scripting.FileSystemObject")
' Ensure file has data
set objFile = objFSO.GetFile("c:\testfolder\scriptTest.txt")
if objFile.size > 0 Then
' open the file and begining reading
set objReadFile = objFSO.OpenTextFile("c:\testfolder\scriptTest.txt",ForReading)
do until objReadFile.AtEndofStream
strLine = objReadFile.ReadLine
wscript.echo strLine
loop
objReadFile.Close
end If

' Read characters within a line
strLine = objReadFile.Read(characters)












