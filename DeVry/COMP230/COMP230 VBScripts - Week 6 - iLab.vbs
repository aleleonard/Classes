
' ******** WEEK 6 - iLab **********





' VBScript: IP_FileWrite.vbs
' Written by: Alejandro Jaque
' Date: Feb 10, 2016
' Class: COMP230
' Professor: Ray Blankenship
' ===================================
' This initializes a 2-dimension array
' of IP Address. The first index +100
' is the room# and the second index+1
' is the computer# in the room.
dim ipAddress(5,3)
  ipAddress(0,0)="192.168.10.11"
  ipAddress(0,1)="192.168.10.12"
  ipAddress(0,2)="192.168.10.13"
  ipAddress(0,3)="192.168.10.14"
  ipAddress(1,0)="192.168.10.19"
  ipAddress(1,1)="192.168.10.20"
  ipAddress(1,2)="192.168.10.21"
  ipAddress(1,3)="192.168.10.22"
  ipAddress(2,0)="192.168.10.27"
  ipAddress(2,1)="192.168.10.28"
  ipAddress(2,2)="192.168.10.29"
  ipAddress(2,3)="192.168.10.30"
  ipAddress(3,0)="192.168.10.35"
  ipAddress(3,1)="192.168.10.36"
  ipAddress(3,2)="192.168.10.37"
  ipAddress(3,3)="192.168.10.38"
  ipAddress(4,0)="192.168.10.43"
  ipAddress(4,1)="192.168.10.44"
  ipAddress(4,2)="192.168.10.45"
  ipAddress(4,3)="192.168.10.46"
  ipAddress(5,0)="192.168.10.51"
  ipAddress(5,1)="192.168.10.52"
  ipAddress(5,2)="192.168.10.53"
  ipAddress(5,3)="192.168.10.54" 

' Defining Constants
CONST READ = 1, WRITE = 2, APPEND = 8, ASSCII = 0
' Defining Variables
dim fileName
fileName = "c:\comp230\IP_Addresses.csv"
ipAddrStr = ""
' Creating file for writing
Set fso = CreateObject("Scripting.FileSystemObject")
Set ipFileObj = fso.CreateTextFile(fileName,True,ASCII)
' Listing and Writing Array into file
For room = 0 to 5 
   For computer = 0 to 3
      ipAddrStr = CStr(room+100) & "," & CStr(computer+1) _
      & "," & ipAddress(room,computer)
      ipFileObj.Writeline ipAddrStr
   Next
Next
' Closing file
ipFileObj.Close

' Opening and closing file for reading
Set ipFileObj = fso.OpenTextFile(fileName,Read,ASCII)
WScript.Echo ipFileObj.ReadAll
ipFileObj.Close




' VBScript: IP_AppendRead.vbs
' Written by: Alejandro Jaque
' Date: Feb 10, 2016
' Class: COMP230
' Professor: Ray Blankenship
' ===================================

CONST READ = 1, WRITE = 2, APPEND = 8, ASCII = 0
dim fileName
fileName = "c:\comp230\IP_Addresses.csv"
ipAddrStr = ""
' Defines newroom value
newRoom = "106"
' Defines values for comp's
comp1_IP = "192.168.10.59"
comp2_IP = "192.168.10.60"
comp3_IP = "192.168.10.61"
comp4_IP = "192.168.10.62"

Set fso = CreateObject("Scripting.FileSystemObject")
Set ipFileObj = fso.OpenTextFile(fileName,APPEND,ASCII)
' Creates a single data for room computers IP's
ipAddrStr = _
  newRoom & ",1," & comp1_IP & vbCrLf & _
  newRoom & ",2," & comp2_IP & vbCrLf & _
  newRoom & ",3," & comp3_IP & vbCrLf & _
  newRoom & ",4," & comp4_IP
' Checks if file already exists
If not fso.FileExists(fileName) Then
  wscript.Echo Chr(7) & Chr(7) & "File Does Not Exist!" & vbCrLf & _
  "You Must Create the File Before You can Read the File !!"
  WScript.Quit
End If
' Writes new room data into file
ipFileObj.Writeline ipAddrStr
ipFileObj.Close

Set ipFileObj = fso.OpenTextFile(fileName,READ,ASCII)
' Read text data from file and assigns them to variables
Do Until ipFileObj.AtEndofStream
  room = ipFileObj.Read(3)
  ipFileObj.Skip(1)
  computer = ipFileObj.Read(1)
  ipFileObj.Skip(1)
  ipAddress = ipFileObj.Read(13)
  ipFileObj.SkipLine
WScript.Echo "The IP Address in Room " & room & " for Computer " _
& computer & " is " & ipAddress
Loop
' Closes text file
ipFileObj.Close
































'===== Not modular version ======

:: Run PC_Tests.vbs
echo off
:start
cls
echo        Computer System Analysis
echo. & echo.
echo [1] Check System Information
echo [2] Check System Memory
echo [3] Check Operating System Version
echo [4] Check Printers Status
echo [5] Check Logical Drive Information
echo [x] Exit Program"
echo.
set /p choice="Enter the Number of your Choice .... "
if %choice% equ x exit
if %choice% equ X exit
cscript //nologo PC_Tests.vbs %choice%
echo.
pause
goto start



' VBScript: PC_Tests.vbs
' Written by: Alejandro Jaque
' Date: Feb 1, 2016
' Class: COMP230
' Professor: Ray Blankenship
' Menu Driven Computer / Network Tests
' This VBScript program is run using the PC_Tests.cmd Batch Script
Set args = WScript.Arguments
WScript.Echo vbCrLf

Select Case args.Item(0)
  Case "1"
    Set WshShell = WScript.CreateObject("WScript.Shell")
    WScript.Echo "The computer name is ............ " & _
      WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    WScript.Echo "The Num of CPUs is .............. " & _
      WshShell.ExpandEnvironmentStrings("%NUMBER_OF_PROCESSORS%")
    WScript.Echo "The Processor Architecture is ... " & _
      WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
  Case "2"
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colComputer = objWMIService.ExecQuery _
      ("Select * from Win32_ComputerSystem")
    For Each objComputer in colComputer
      intRamMB = int((objComputer.TotalPhysicalMemory) /1048576)+1
      Wscript.Echo "System Name ...... " & objComputer.Name _
      & vbCrLf & "Total RAM ........ " & intRamMB & " MBytes."
    Next
  Case "3"
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colOperatingSystems = objWMIService.ExecQuery _
      ("Select * from Win32_OperatingSystem")
    WScript.Echo "The Operating System Detected is Shown Below:" & vbCrLf
    For Each objOperatingSystem in colOperatingSystems
      WScript.Echo objOperatingSystem.Caption & "Version: " & _
        objOperatingSystem.Version
    Next
  Case "4"
    strComputer ="."
    intPrinters = 1
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery _
      ("SELECT * FROM Win32_Printer")
    WScript.Sleep(1000)
    For Each objItem In colItems
      WScript.Echo _
        "Printer: " & objItem.DeviceID & vbCrLf & _
        "===============================================" & vbCrLf & _
        "Driver Name ............. " & objItem.DriverName & vbCrLf & _
        "Port Name ............... " & objItem.PortName & vbCrLf & _
        "Printer State ........... " & objItem.PrinterState & vbCrLf & _
        "Printer Status .......... " & objItem.PrinterStatus & vbCrLf & _
        "Print Processor ......... " & objItem.PrintProcessor & vbCrLf & _
        "Spool Enabled ........... " & objItem.SpoolEnabled & vbCrLf & _
        "Shared .................. " & objItem.Shared & vbCrLf & _
        "ShareName ............... " & objItem.ShareName & vbCrLf & _
        "Horizontal Res .......... " & objItem.HorizontalResolution & vbCrLf & _
        "Vertical Res ............ " & objItem.VerticalResolution & vbCrLf
      intPrinters = intPrinters + 1
    Next
  case "5"
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery _
      ("Select * from Win32_LogicalDisk Where FreeSpace > 0")
    For Each objItem in colItems
      WScript.Echo vbCrLf & _
        "========================================" & vbCrLf & _
        "Drive Letter ......... " & objItem.Name & vbCrLf & _
        "Description .......... " & objItem.Description & vbCrLf & _
        "Volume Name .......... " & objItem.VolumeName & vbCrLf & _
        "Drive Type ........... " & objItem.DriveType & vbCrLf & _
        "Media Type ........... " & objItem.MediaType & vbCrLf & _
        "VolumeSerialNumber ... " & objItem.VolumeSerialNumber & vbCrLf & _
        "Size ................. " & Int(objItem.Size /1073741824) & " GB" & vbCrLf & _
        "Free Space ........... " & Int(objItem.FreeSpace /1073741824) & " GB"
    Next
End Select



'===== Modular version ======

:: Run Mod1_PCTests.vbs
echo off
:start
cls
echo        Computer System Analysis
echo. & echo.
echo [1] Check System Information
echo [2] Check System Memory
echo [3] Check Operating System Version
echo [4] Check Printers Status
echo [5] Check Logical Drive Information
echo [x] Exit Program"
echo.
set /p choice="Enter the Number of your Choice .... "
if %choice% equ x exit
if %choice% equ X exit
cscript //nologo Mod1_PCTests.vbs %choice%
echo.
pause
goto start

' VBScript: Mod1_PCTests.vbs
' Written by: Alejandro Jaque
' Date: Feb 1, 2016
' Class: COMP230
' Professor: Ray Blankenship
' Menu Driven Computer / Network Tests
' This VBScript program is run using the Mod1_PCTests.cmd Batch Script
Set args = WScript.Arguments
WScript.Echo vbCrLf

Select Case args.Item(0)
  Case "1"
    Call System_Information
  Case "2"
    Call System_Memory_Size
  Case "3"
    Call OS_Version
  Case "4"
    Call Printers_Status
  case "5"
    Call Logical_HDD_Information
  case Else
    WScript.Echo chr(7) & chr(7) & "Error, Choices are 1..5 or x!!!"
End Select

Sub System_Information
    Set WshShell = WScript.CreateObject("WScript.Shell")
    WScript.Echo "The computer name is ............ " & _
      WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    WScript.Echo "The Num of CPUs is .............. " & _
      WshShell.ExpandEnvironmentStrings("%NUMBER_OF_PROCESSORS%")
    WScript.Echo "The Processor Architecture is ... " & _
      WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
End Sub

Sub System_Memory_Size
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colComputer = objWMIService.ExecQuery _
      ("Select * from Win32_ComputerSystem")
    For Each objComputer in colComputer
      intRamMB = int((objComputer.TotalPhysicalMemory) /1048576)+1
      Wscript.Echo "System Name ...... " & objComputer.Name _
      & vbCrLf & "Total RAM ........ " & intRamMB & " MBytes."
    Next
End Sub

Sub OS_Version
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colOperatingSystems = objWMIService.ExecQuery _
      ("Select * from Win32_OperatingSystem")
    WScript.Echo "The Operating System Detected is Shown Below:" & vbCrLf
    For Each objOperatingSystem in colOperatingSystems
      WScript.Echo objOperatingSystem.Caption & "Version: " & _
        objOperatingSystem.Version
    Next
End Sub
    
Sub Printers_Status
    strComputer ="."
    intPrinters = 1
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery _
      ("SELECT * FROM Win32_Printer")
    WScript.Sleep(1000)
    For Each objItem In colItems
      WScript.Echo _
        "Printer: " & objItem.DeviceID & vbCrLf & _
        "===============================================" & vbCrLf & _
        "Driver Name ............. " & objItem.DriverName & vbCrLf & _
        "Port Name ............... " & objItem.PortName & vbCrLf & _
        "Printer State ........... " & objItem.PrinterState & vbCrLf & _
        "Printer Status .......... " & objItem.PrinterStatus & vbCrLf & _
        "Print Processor ......... " & objItem.PrintProcessor & vbCrLf & _
        "Spool Enabled ........... " & objItem.SpoolEnabled & vbCrLf & _
        "Shared .................. " & objItem.Shared & vbCrLf & _
        "ShareName ............... " & objItem.ShareName & vbCrLf & _
        "Horizontal Res .......... " & objItem.HorizontalResolution & vbCrLf & _
        "Vertical Res ............ " & objItem.VerticalResolution & vbCrLf
      intPrinters = intPrinters + 1
    Next
End Sub

Sub Logical_HDD_Information
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery _
      ("Select * from Win32_LogicalDisk Where FreeSpace > 0")
    For Each objItem in colItems
      WScript.Echo vbCrLf & _
        "========================================" & vbCrLf & _
        "Drive Letter ......... " & objItem.Name & vbCrLf & _
        "Description .......... " & objItem.Description & vbCrLf & _
        "Volume Name .......... " & objItem.VolumeName & vbCrLf & _
        "Drive Type ........... " & objItem.DriveType & vbCrLf & _
        "Media Type ........... " & objItem.MediaType & vbCrLf & _
        "VolumeSerialNumber ... " & objItem.VolumeSerialNumber & vbCrLf & _
        "Size ................. " & Int(objItem.Size /1073741824) & " GB" & vbCrLf & _
        "Free Space ........... " & Int(objItem.FreeSpace /1073741824) & " GB"
    Next
End Sub
    



' ====== Library used version ======

:: Run Mod2_PCTests.vbs
echo off
:start
cls
echo        Computer System Analysis
echo. & echo.
echo [1] Check System Information
echo [2] Check System Memory
echo [3] Check Operating System Version
echo [4] Check Printers Status
echo [5] Check Logical Drive Information
echo [x] Exit Program"
echo.
set /p choice="Enter the Number of your Choice .... "
if %choice% equ x exit
if %choice% equ X exit
cscript //nologo Mod2_PCTests.vbs %choice%
echo.
pause
goto start



' VBScript: Mod2_PCTests.vbs
' Written by: Alejandro Jaque
' Date: Feb 1, 2016
' Class: COMP230
' Professor: Ray Blankenship
' Menu Driven Computer / Network Tests
' This VBScript program is run using the PCT_Library.vbs Script
Set args = WScript.Arguments
WScript.Echo vbCrLf
Set fso = CreateObject("Scripting.FileSystemObject")
Set vbsLib = fso.OpenTextFile("C:\Comp230\PCT_Library.vbs",1,False)
librarySubs = vbsLib.ReadAll
vbsLib.Close
Set vbsLib=Nothing
Set fso=Nothing
ExecuteGlobal librarySubs
Select Case args.Item(0)
  Case "1"
    Call System_Information
  Case "2"
    Call System_Memory_Size
  Case "3"
    Call OS_Version
  Case "4"
    Call Printers_Status
  case "5"
    Call Logical_HDD_Information
  case Else
    WScript.Echo chr(7) & chr(7) & "Error, Choices are 1..5 or x!!!"
End Select



' VBScript: PCT_Library.vbs
' Written by: Alejandro Jaque
' Date: Feb 1, 2016
' Class: COMP230
' Professor: Ray Blankenship

Sub System_Information
    Set WshShell = WScript.CreateObject("WScript.Shell")
    WScript.Echo "The computer name is ............ " & _
      WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    WScript.Echo "The Num of CPUs is .............. " & _
      WshShell.ExpandEnvironmentStrings("%NUMBER_OF_PROCESSORS%")
    WScript.Echo "The Processor Architecture is ... " & _
      WshShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
End Sub

Sub System_Memory_Size
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colComputer = objWMIService.ExecQuery _
      ("Select * from Win32_ComputerSystem")
    For Each objComputer in colComputer
      intRamMB = int((objComputer.TotalPhysicalMemory) /1048576)+1
      Wscript.Echo "System Name ...... " & objComputer.Name _
      & vbCrLf & "Total RAM ........ " & intRamMB & " MBytes."
    Next
End Sub

Sub OS_Version
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colOperatingSystems = objWMIService.ExecQuery _
      ("Select * from Win32_OperatingSystem")
    WScript.Echo "The Operating System Detected is Shown Below:" & vbCrLf
    For Each objOperatingSystem in colOperatingSystems
      WScript.Echo objOperatingSystem.Caption & "Version: " & _
        objOperatingSystem.Version
    Next
End Sub
    
Sub Printers_Status
    strComputer ="."
    intPrinters = 1
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery _
      ("SELECT * FROM Win32_Printer")
    WScript.Sleep(1000)
    For Each objItem In colItems
      WScript.Echo _
        "Printer: " & objItem.DeviceID & vbCrLf & _
        "===============================================" & vbCrLf & _
        "Driver Name ............. " & objItem.DriverName & vbCrLf & _
        "Port Name ............... " & objItem.PortName & vbCrLf & _
        "Printer State ........... " & objItem.PrinterState & vbCrLf & _
        "Printer Status .......... " & objItem.PrinterStatus & vbCrLf & _
        "Print Processor ......... " & objItem.PrintProcessor & vbCrLf & _
        "Spool Enabled ........... " & objItem.SpoolEnabled & vbCrLf & _
        "Shared .................. " & objItem.Shared & vbCrLf & _
        "ShareName ............... " & objItem.ShareName & vbCrLf & _
        "Horizontal Res .......... " & objItem.HorizontalResolution & vbCrLf & _
        "Vertical Res ............ " & objItem.VerticalResolution & vbCrLf
      intPrinters = intPrinters + 1
    Next
End Sub

Sub Logical_HDD_Information
    strComputer = "."
    Set objWMIService = GetObject _
      ("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery _
      ("Select * from Win32_LogicalDisk Where FreeSpace > 0")
    For Each objItem in colItems
      WScript.Echo vbCrLf & _
        "========================================" & vbCrLf & _
        "Drive Letter ......... " & objItem.Name & vbCrLf & _
        "Description .......... " & objItem.Description & vbCrLf & _
        "Volume Name .......... " & objItem.VolumeName & vbCrLf & _
        "Drive Type ........... " & objItem.DriveType & vbCrLf & _
        "Media Type ........... " & objItem.MediaType & vbCrLf & _
        "VolumeSerialNumber ... " & objItem.VolumeSerialNumber & vbCrLf & _
        "Size ................. " & Int(objItem.Size /1073741824) & " GB" & vbCrLf & _
        "Free Space ........... " & Int(objItem.FreeSpace /1073741824) & " GB"
    Next
End Sub
    



