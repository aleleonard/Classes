


' ************ WEEK 3 - Project Proposal *************
'-----------------------------------------------------------
' This VBScript provides Live Monitoring of Network Traffic
' and saves snapshots of the counters reading into a text
' file with the name of the counter monitored.
' This Script can be useful to an Systems Administrator
' who is interest in knowing the machines' activities
'-----------------------------------------------------------

' Sets variable for protocol selected by user
protName = ""

' Sets CMD Shell environment for use
Set ws = Wscript.CreateObject("Wscript.Shell")

' Sets variable with input from the user in the InputBox
protName = InputBox("Checking Internet Traffic from local machine" & vbCrLf & _
"" & vbCrLf & _
"Enter Protocol to Monitor:" & vbCrLf & _
"ip" & vbTab & vbTab & "(Monitors IP Packets Traffic)" & vbCrLf & _
"tcp" & vbTab & vbTab & "(Monitors TCP Packets Traffic)" & vbCrLf & _
"udp" & vbTab & vbTab & "(Monitors UDP Packets Traffic)" & vbCrLf & _
"tcpcon" & vbTab & vbTab & "(Monitors TCP Connections)" & vbCrLf & _
"udpcon" & vbTab & vbTab & "(Monitors UDP Connections)" & vbCrLf & _
"","Internet Protocol Traffic Monitor")

' Decision routine to classify commands to run depending on user input

' Executes if user input is "ip" for IP Protocol
if protName = "ip" then
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_ip_results.txt & netsh interface ip show ipstats >> tramon_ip_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting IP Traffic Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show ipstats >> tramon_ip_results.txt & netsh interface ip show ipstats rr=1"
' Pops up message displaying filename with saved commands output
	 Wscript.Echo "A snapshot of the counters were saved on file: tramon_" & protName & "_results.txt"

' Executes if user input is "tcp" for TCP Protocol
 elseif protName = "tcp" then
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_tcp_results.txt & netsh interface ip show tcpstats >> tramon_tcp_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting TCP Traffic Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show tcpstats >> tramon_tcp_results.txt & netsh interface ip show tcpstats rr=1"
' Pops up message displaying filename with saved commands output
	 Wscript.Echo "A snapshot of the counters were saved on file: tramon_" & protName & "_results.txt"

' Executes if user input is "udp" for UDP Protocol
 elseif protName = "udp" then
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_udp_results.txt & netsh interface ip show udpstats >> tramon_udp_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting UDP Traffic Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show udpstats >> tramon_udp_results.txt & netsh interface ip show udpstats rr=1"
' Pops up message displaying filename with saved commands output
	 Wscript.Echo "A snapshot of the counters were saved on file: tramon_" & protName & "_results.txt"

' Executes if user input is "tcpcon" for TCP Connections
 elseif protName = "tcpcon" then
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_tcpcon_results.txt & netsh interface ip show tcpconnections >> tramon_tcpcon_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting TCP Connections Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show tcpconnections >> tramon_tcpcon_results.txt & netsh interface ip show tcpconnections rr=1"
' Pops up message displaying filename with saved commands output
	 Wscript.Echo "A snapshot of the counters were saved on file: tramon_" & protName & "_results.txt"

' Executes if user input is "udpcon" for UDP Connections
 elseif protName = "udpcon" then
     ws.Run "cmd /c echo Paired Snapshost Time: %date% at %time% >> tramon_udpcon_results.txt & netsh interface ip show udpconnections >> tramon_udpcon_results.txt & echo. & echo. & echo Saving logfile...... done. & echo. & echo. & echo Starting UDP Connections Live Monitor... & echo. & echo. & timeout /t 10 & netsh interface ip show udpconnections >> tramon_udpcon_results.txt & netsh interface ip show udpconnections rr=1"
' Pops up message displaying filename with saved commands output
	 Wscript.Echo "A snapshot of the counters were saved on file: tramon_" & protName & "_results.txt"

' Pops up error message when no valid option is entered in InputBox
 else WScript.Echo "Invalid protocol entered. Run program again"
 
end if







' ************ WEEK 6 - Final Project Proposal *************
'-----------------------------------------------------------
' This VBScript provides Live Monitoring of Network Traffic
' and saves snapshots of the counters reading into a text
' file with the name of the counter monitored.
' This Script can be useful to an Systems Administrator
' who is interest in knowing the machines' activities
'-----------------------------------------------------------

' Set Variables
weekdayVar = weekday(Date) ' Sets date variable
timedayVar = hour(Time) ' Sets time variable
protName = "" ' Sets protocol variable


' Sets array with maintenance schedule values for posterior printout
  dim dayarrayVar
scheduleVar = Array("( 6PM - 10PM )","( No Maintenance Allowed)","Monday","Tuesday","Wednesday","Thursday","Friday","Satuday","Sunday")

' Decision routine to validate date and time to execute program
If (weekdayVar = 2 OR weekdayVar = 3 OR weekdayVar = 4 OR weekdayVar = 5 OR weekdayVar = 6) AND (timedayVar >= 18 AND timedayVar <= 20) Then

 ' Sets CMD Shell environment for use
 Set ws = Wscript.CreateObject("Wscript.Shell")

 ' Calls Function with input from user to assign to protocol variable
 protName = UserInput()

  ' Decision routine to execute commands depending on user input
  Select Case protName
   Case "ip"
     Call Monitor_IPProc
     Call Print_Filename
   Case "tcp"
     Call Monitor_TCPProc
     Call Print_Filename
   Case "udp"
     Call Monitor_UDPProc
     Call Print_Filename
   Case "tcpcon"
     Call Monitor_TCPConProc
     Call Print_Filename
   Case "udpcon" 
     Call Monitor_UDPConProc
     Call Print_Filename
   Case else
      Call Print_InvalidProt
  End Select

 ElseIf (weekdayVar = 2 OR weekdayVar = 3 OR weekdayVar = 4 OR weekdayVar = 5 OR weekdayVar = 6) AND (timedayVar < 18 AND timedayVar > 20) Then
  Call Print_Schedule
 ElseIf weekdayVar = 7 OR weekdayVar = 1 Then
  Call Print_Schedule
End If

' ============== Functions ===============

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
Function Monintor_TCPProc
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

' ============== Subroutines ===============

' Pops up message displaying filename with saved commands output
Sub Print_Filename
	 Wscript.Echo "A snapshot of the counters were saved on file: tramon_" & protName & "_results.txt"
End Sub

' Pops up message if input by user is invalid
Sub Print_InvalidProt
     WScript.Echo "Invalid protocol entered. Run program again"
End Sub

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




