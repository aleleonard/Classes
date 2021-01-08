
' ******** WEEK 5 - iLab **********




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
    



