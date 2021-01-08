
' ******** WEEK 7 - Lesson **********


'  =============== Connection Object ===============

' Create an Active Directory Object Database (ADODB) object
Set objConnection = CreateObject("ADODB.Connection")
ConnectionStr = "Provider = Microsoft.Jet.OLEDB.4.0;" & "Data Source = c:\Database\inventory.mdb"
objConnection.Open ConnectionStr

' Connection string used to connect to an SQL Server database
Set ConWSH=WScript.CreateObject("ADODB.Connection")
' Define our connection string
strConnection = "Driver={SQL Server};Server=SERVER;User ID=USERID;Password=PASSWORD;Database=TEST;"

' Connection string used to connect to a MySQL Server database
Set oCn = CreateObject("ADODB.Connection")
ConnectionStr = "Driver={MySQL ODBC 3.51 Driver};SERVER=mysqlserver.mydomain.com;" & _

    "DATABASE=testdatabase;" &_
	"USER=user;" & _
	"PASSWORD=userpassword;" & _
	"OPTION=3;"



' Create an Active Directory Object Database (ADODB) object for MS Access from 2003 and earlier
Set objConnection = CreateObject("ADODB.Connection")
ConnectionStr = "Provider = Microsoft.Jet.OLEDB.4.0;" & "Data Source = c:\Database\inventory.mdb"
objConnection.Open ConnectionStr
	
' Create an Active Directory Object Database (ADODB) object for MS Access 2010
Set objConnection = CreateObject("ADODB.Connection")
dataSource = "Provider = Microsoft.ACE.OLEDB.12.0;" & "Data Source = c:\VBScripts\Pres_DB.mdb"
objConnection.Open dataSource



' Creates Connection object
Set objConnection = CreateObject("ADODB.Connection")
' Specify provider and data source
dataSource = "Provider = Microsoft.ACE.OLEDB.12.0;" & "Data Source = c:\VBScripts\Pres_DB.mdb"
' Makes connection to database
objConnection.Open dataSource
' Create recordsSet object
Set objRecordSet = CreateObject("ADODB.Recordset")
' Specifiy SQP queries examples (all records or filtered)
SQLstr = "SELECT * FROM Presidents"
SQLstr = "SELECT * FROM Presidents WHERE YearElected < 1864"
SQLstr = "SELECT * FROM Presidents WHERE YearElected < 1865 ORDER BY PresName"
SQLstr = "SELECT YearElected,PresName FROM Presidents WHERE YearElected < 1865 ORDER BY PresName"
' Submit query to database
objRecordSet.Open sqlStr , objConnection
' Move to the begining of database
objRecordSet.MoveFirst
' Retrive records that matches with SQL query
recordsStr = ""
Do Until objRecordSet.EOF
   recordsStr = recordsStr & objRecordSet.Fields.Item("PresNumber") & _
        vbTab & objRecordSet.Fields.Item("YearElected") & _
   recordsStr = recordsStr & objRecordSet.Fields.Item("YearElected") & vbTab & pad(objRecordSet.Fields.Item("PresName"),22) & vbCrlf & vbTab & pad(objRecordSet.Fields.Item("VPresName"),22) & vbCrlf
   objRecordSet.MoveNext
Loop
objRecordSet.Close
objConnection.Close


















