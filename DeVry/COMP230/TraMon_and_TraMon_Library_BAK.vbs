'============================================================
' ********** WEEK 6 - Final Project Proposal ***********
'
'   This script provides Live Monitoring of Network Traffic
' and saves snapshots of the counters reading into a text
' file with the name of the counter monitored.
'   This Script is useful to an Systems Administrator
' interested in knowing the machines' network activities.
'
' It only works during maintenance times (M-F 6-10PM).
'
'   VBScript features used in this program:
'     - Variables, Constants, and Data Types	✓
'     - Output & Input Methods			✓
'     - Decision-Making Statements		✓
'     - Loop Structures and Arrays		✓
'     - Procedures, Functions, and Libraries	✓
'     - File Input/Output Methods			✓
'     - Batch Commands and File Saving		✓
'============================================================

' ============== Main Script Definition ===============

' Setting Constants and Variable
CONST_WEEKDAY = weekday(Date) ' Sets date constant
CONST_TIMEDAY = hour(Time) ' Sets time constant
protName = "" ' Sets protocol name variable

' Calls Schedule Hours Library
Set fso = CreateObject("Scripting.FileSystemObject")
Set vbsLib = fso.OpenTextFile("TraMon_Library.vbs",1,False)
scheduleLibSub = vbsLib.ReadAll
vbsLib.Close
Set vbsLib=Nothing
Set fso=Nothing
ExecuteGlobal scheduleLibSub

' Decision-Making routine to validate date and time
If (CONST_WEEKDAY = 2 OR CONST_WEEKDAY = 3 OR CONST_WEEKDAY = 4 OR CONST_WEEKDAY = 5 OR CONST_WEEKDAY = 6) AND (CONST_TIMEDAY >= 18 AND CONST_TIMEDAY <= 22) Then

' Sets CMD Shell environment for monitoring commands use
 Set ws = Wscript.CreateObject("Wscript.Shell")

' Calls Function with input from user to assign to protocol variable
 protName = UserInput()

' Decision-Making routine to execute batch commands depending on user input
  Select Case protName
   Case "ip"
	 Call Save_Logfile
     Call Monitor_IPProc
     Call Print_Filename
   Case "tcp"
	 Call Save_Logfile
     Call Monitor_TCPProc
     Call Print_Filename
   Case "udp"
   	 Call Save_Logfile
     Call Monitor_UDPProc
     Call Print_Filename
   Case "tcpcon"
   	 Call Save_Logfile
     Call Monitor_TCPConProc
     Call Print_Filename
   Case "udpcon" 
   	 Call Save_Logfile
     Call Monitor_UDPConProc
     Call Print_Filename
   Case else
      Call Print_InvalidProt
  End Select

' Decision-Making routine to validate date and time (cont'd)
 ElseIf (CONST_WEEKDAY = 2 OR CONST_WEEKDAY = 3 OR CONST_WEEKDAY = 4 OR CONST_WEEKDAY = 5 OR CONST_WEEKDAY = 6) AND (CONST_TIMEDAY < 18 AND CONST_TIMEDAY > 22) Then
  Call Print_Schedule
 ElseIf CONST_WEEKDAY = 7 OR CONST_WEEKDAY = 1 Then
  Call Print_Schedule
End If





' ============== Array Definition ===============

' Sets array with maintenance schedule values for posterior printout
  dim dayarrayVar
scheduleVar = Array("( 6PM - 10PM )","( No Maintenance Allowed)","Monday","Tuesday","Wednesday","Thursday","Friday","Satuday","Sunday")


' ============== Subroutines Definition ===============

' Sets Subroutine to echo message with maintenance schedule
Sub Print_Schedule
WSCript.Echo "ERROR: PROCESS ABORTED!" & vbCrLf & _
"Maintenance is only permitted during:"  & vbCrLf & vbCrLf & _ 
scheduleVar(2) & vbTab & vbTab & scheduleVar(0) & vbCrLf & _ 
scheduleVar(3) & vbTab & vbTab & scheduleVar(0) & vbCrLf & _
scheduleVar(4) & vbTab & scheduleVar(0) & vbCrLf & _
scheduleVar(5) & vbTab & vbTab & scheduleVar(0) & vbCrLf & _
scheduleVar(6) & vbTab & vbTab & scheduleVar(0) & vbCrLf & _
scheduleVar(7) & vbTab & vbTab & scheduleVar(1) & vbCrLf & _
scheduleVar(8) & vbTab & vbTab & scheduleVar(1) & vbCrLf & vbCrLf & _
"Please resume work during maintenance hours." & vbCrLf & "Thanks."
End Sub

' Pops up message displaying filename with saved commands output
Sub Print_Filename
	 Wscript.Echo "A snapshot of the counters were saved on file: tramon_" & protName & "_results.txt"
End Sub

' Pops up message if input by user is invalid
Sub Print_InvalidProt
     WScript.Echo "Invalid protocol entered. Run program again"
End Sub

' Saves logfile
Sub Save_Logfile
    Set WshShell = WScript.CreateObject("WScript.Shell")

    syscompInfo = WshShell.ExpandEnvironmentStrings("Computer Name: " & "%COMPUTERNAME%" & vbCrLf & "Number of Processors: " & "%NUMBER_OF_PROCESSORS%" & vbCrLf & "Processor Architecture: " & "%PROCESSOR_ARCHITECTURE%" & vbCrLf & vbCrLf)
	
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile("tramon_logfile.txt",8,True)
    f.Write syscompInfo
    f.Close
End Sub



' ============== Functions Definition ===============

' Sets Function with input from the user
Function UserInput()
UserInput = InputBox("Checking Internet Traffic from local machine" & vbCrLf & _
"" & vbCrLf & _
"Enter Protocol to Monitor:" & vbCrLf & _
"ip" & vbTab & vbTab & "(Monitors IP Packets Traffic)" & vbCrLf & _
"tcp" & vbTab & vbTab & "(Monitors TCP Packets Traffic)" & vbCrLf & _
"udp" & vbTab & vbTab & "(Monitors UDP Packets Traffic)" & vbCrLf & _
"tcpcon" & vbTab & vbTab & "(Monitors TCP Connections)" & vbCrLf & _
"udpcon" & vbTab & vbTab & "(Monitors UDP Connections)" & vbCrLf & _
"","Internet Protocol Traffic Monitor")
End Function




' Executes if user input is "ip" for IP Protocol
Function Monitor_IPProc
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_ip_results.txt & netsh interface ip show ipstats >> tramon_ip_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting IP Traffic Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show ipstats >> tramon_ip_results.txt & netsh interface ip show ipstats rr=1"
End Function

' Executes if user input is "tcp" for TCP Protocol
Function Monitor_TCPProc
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_tcp_results.txt & netsh interface ip show tcpstats >> tramon_tcp_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting TCP Traffic Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show tcpstats >> tramon_tcp_results.txt & netsh interface ip show tcpstats rr=1"
End Function

' Executes if user input is "udp" for UDP Protocol
Function Monitor_UDPProc
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_udp_results.txt & netsh interface ip show udpstats >> tramon_udp_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting UDP Traffic Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show udpstats >> tramon_udp_results.txt & netsh interface ip show udpstats rr=1"
End Function

' Executes if user input is "tcpcon" for TCP Connections
Function Monitor_TCPConProc
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_tcpcon_results.txt & netsh interface ip show tcpconnections >> tramon_tcpcon_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting TCP Connections Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show tcpconnections >> tramon_tcpcon_results.txt & netsh interface ip show tcpconnections rr=1"
End Function

' Executes if user input is "udpcon" for UDP Connections
Function Monitor_UDPConProc
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_udpcon_results.txt & netsh interface ip show udpconnections >> tramon_udpcon_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting UDP Connections Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show udpconnections >> tramon_udpcon_results.txt & netsh interface ip show udpconnections rr=1"
End Function
