
' *** Week 3 - Discussions ***
' Deleting Resource by asking the user the name of it
resourceName = ""
Set fileServ = GetObject("WinNt://[servername]/LanmanServer,FileService")
resourceName = InputBox("Please enter the Misspelled Resource:","Misspelled Resource to Delete")
WScript.StdOut.Write("Please enter the Misspelled Resource: ")
resourceName = WScript.StdIn.ReadLine()
fileServ.Delete "FileShare","[misspelled_resource_name]"
WScript.Echo vbCrLf & "Resource " & "'" & resourceName & "'" & " deleted!" & vbCrLf & vbCrLf & "End of Program"



' If Else vs ElseIf
Dim userNum
userNum = InputBox("Enter Number:","Number Comparison")
If userNum < 10 Then
   WScript.Echo "You typed a number under 10"
 Else
   WScript.Echo "You typed a number equal or over 10"
End If

userNum = InputBox("Enter Number:","Number Comparison")
If userNum < 10 Then
   WScript.Echo "You typed a number under 10"
 ElseIf userNum = 10 Then
   WScript.Echo "You typed a number equal to 10"
 Else
   WScript.Echo "You typed a number over 10"
End If


' Nested If
yourAge = inputBox("Type Your Age")
 If IsNumeric(yourAge) Then
    If yourAge < 0 Then
     yourAge = inputBox("Type Your Age Over 0")
	 End If
	 If yourAge >= 0 AND yourAge <= 14 Then 
	 WSCript.Echo("You are a Kid")
	 End If
	 If yourAge > 14 AND yourAge <= 25 Then 
	 WSCript.Echo("You are a Teen")
	 End If
	 If yourAge > 25 AND yourAge <= 80 Then 
	 WSCript.Echo("You are an Adult")
     End If	 
	 If yourAge > 80 Then 
	 WSCript.Echo("You should be dead")
	 End If
 Else 
     WSCript.Echo("You need to enter your Age")
End If



