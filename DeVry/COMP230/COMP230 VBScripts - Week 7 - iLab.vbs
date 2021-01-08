'==========================================================================
' NAME: ComputerDatabase.vbs
'
' AUTHOR: jlmorgan , 
' DATE  : 10/22/2011
'
' COMMENT: Use 32 bit ODBC Microsoft Access Driver
'
'==========================================================================
recordsStr = ""
sqlStr = "SELECT * FROM Computers"
dataSource = "provider=Microsoft.ACE.OLEDB.12.0;" _
& "data source=Computers.accdb"

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open dataSource
Set objRecordSet = CreateObject("ADODB.Recordset")
objRecordSet.Open sqlStr , objConnection
objRecordSet.MoveFirst
' Display Headers
recordsStr = "Computer                 HostName           Room_Num" & _
     "   CPU_Type  Speed   Num_CPUs   Bit_Size          OS_Type   " & _
     "      Memory       HDD_Size" & vbCrLf & _
     "============================================================" & _
     "=============================" & vbCrLf
Do Until objRecordSet.EOF
   recordsStr = recordsStr & objRecordSet.Fields.Item("Computer") & _
        vbTab & pad(objRecordSet.Fields.Item("HostName"),12) & _
        vbTab & pad(objRecordSet.Fields.Item("Room_Num"),14) & _
        vbTab & objRecordSet.Fields.Item("CPU_Type") & _
        vbTab & objRecordSet.Fields.Item("Speed") & _
        vbTab & objRecordSet.Fields.Item("Num_CPUs") & _
        vbTab & objRecordSet.Fields.Item("Bit_Size") & _
        vbTab & pad(objRecordSet.Fields.Item("OS_Type"),12) & _
        vbTab & objRecordSet.Fields.Item("Memory") & _
        vbTab & objRecordSet.Fields.Item("HDD_Size") & vbCrLf
    objRecordSet.MoveNext
Loop
objRecordSet.Close
objConnection.Close

WScript.Echo recordsStr

function pad(ByVal strText, ByVal len)
	pad = Left(strText & Space(len), len)
end Function






'==========================================================================
' NAME: ComputerReplace.vbs
' AUTHOR: Alejandro Jaque , 
' DATE  : 02/16/2016
' Class: COMP230
' Professor: Ray Blankenship
' COMMENT: Use 32 bit ODBC Microsoft Access Driver
'
'==========================================================================
recordsStr = ""
sqlStr = "SELECT Computer, Room_Num, Speed, Num_CPUs, OS_Type, HDD_Size FROM Computers WHERE Num_CPUs = 1 OR Speed < 2.1 OR HDD_Size < 300 ORDER BY Room_Num"
dataSource = "provider=Microsoft.ACE.OLEDB.12.0;" _
& "data source=Computers.accdb"

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open dataSource
Set objRecordSet = CreateObject("ADODB.Recordset")
objRecordSet.Open sqlStr , objConnection
objRecordSet.MoveFirst
' Display Headers
recordsStr = "Computer          Room_Num" & _
     "   Speed   Num_CPUs   OS_Type   " & _
     "      HDD_Size" & vbCrLf & _
     "==============================================" & vbCrLf
Do Until objRecordSet.EOF
   recordsStr = recordsStr & objRecordSet.Fields.Item("Computer") & _
        vbTab & pad(objRecordSet.Fields.Item("Room_Num"),12) & _
        vbTab & pad(objRecordSet.Fields.Item("Speed"),5) & _
        vbTab & objRecordSet.Fields.Item("Num_CPUs") & _
        vbTab & pad(objRecordSet.Fields.Item("OS_Type"),12) & _
        vbTab & objRecordSet.Fields.Item("HDD_Size") & vbCrLf
    objRecordSet.MoveNext
Loop
objRecordSet.Close
objConnection.Close

WScript.Echo recordsStr

function pad(ByVal strText, ByVal len)
	pad = Left(strText & Space(len), len)
end Function





'==========================================================================
' NAME: ComputerUpgrade.vbs
' AUTHOR: Alejandro Jaque , 
' DATE  : 02/16/2016
' Class: COMP230
' Professor: Ray Blankenship
' COMMENT: Use 32 bit ODBC Microsoft Access Driver
'
'==========================================================================
recordsStr = ""
sqlStr = "SELECT Computer, HostName, Room_Num, OS_Type, Memory FROM Computers WHERE OS_Type = 'Fedora 10' OR OS_Type = 'Windows XP' OR Memory = 2 ORDER BY OS_Type"
dataSource = "provider=Microsoft.ACE.OLEDB.12.0;" _
& "data source=Computers.accdb"

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open dataSource
Set objRecordSet = CreateObject("ADODB.Recordset")
objRecordSet.Open sqlStr , objConnection
objRecordSet.MoveFirst
' Display Headers
recordsStr = "Computer                 HostName           Room_Num" & _
     "   OS_Type   " & _
     "      Memory" & vbCrLf & _
     "==============================================" & vbCrLf
Do Until objRecordSet.EOF
   recordsStr = recordsStr & objRecordSet.Fields.Item("Computer") & _
        vbTab & pad(objRecordSet.Fields.Item("HostName"),12) & _
        vbTab & pad(objRecordSet.Fields.Item("Room_Num"),10) & _
        vbTab & pad(objRecordSet.Fields.Item("OS_Type"),10) & _
        vbTab & objRecordSet.Fields.Item("Memory") & vbCrLf
    objRecordSet.MoveNext
Loop
objRecordSet.Close
objConnection.Close

WScript.Echo recordsStr

function pad(ByVal strText, ByVal len)
	pad = Left(strText & Space(len), len)
end Function

