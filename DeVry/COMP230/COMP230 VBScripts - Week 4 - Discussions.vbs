
' *** Week 4 - Discussions ***

' Validates if input is numeric or string
' Run with >cscript 
WScript.StdOut.Write("Type Anything ...... ")
anyVar = WScript.StdIn.ReadLine()
 If IsNumeric(anyVar) Then
     WSCript.StdOut.Write("You Entered a Numeric Value")
 ElseIf anyVar = "" Then
        WSCript.StdOut.Write("You Entered a Blank Value")
 Else 
     WSCript.StdOut.Write("You Entered a String Value")
End If


' Validates if input is numeric or string
' Run with >wscript 
anyVar = inputBox("Type Anything:")
 If IsNumeric(anyVar) Then
     WSCript.Echo("You Entered a Numeric Value")
 ElseIf anyVar = "" Then
        WSCript.Echo("You Entered a Blank Value")
 Else 
     WSCript.Echo("You Entered a String Value")
End If




' Select Case
d=weekday(date)
Select Case d
  Case 1
    WSCript.Echo("Sleepy Sunday")
  Case 2
    WSCript.Echo("Monday again!")
  Case 3
    WSCript.Echo("Just Tuesday!")
  Case 4
    WSCript.Echo("Wednesday!")
  Case 5
    WSCript.Echo("Thursday...")
  Case 6
    WSCript.Echo("Finally Friday!")
  Case else
    WSCript.Echo("Super Saturday!!!!")
End Select
m=month(date)
Select Case m
  Case 1
    WSCript.Echo("January")
  Case 2
    WSCript.Echo("February")
  Case 3
    WSCript.Echo("March")
  Case 4
    WSCript.Echo("April")
  Case 5
    WSCript.Echo("May")
  Case 6
    WSCript.Echo("June")
  Case else
    WSCript.Echo("Another Month")
End Select
y=year(date)
Select Case y
  Case 2015
    WSCript.Echo("2015")
  Case 2016
    WSCript.Echo("2016")
  Case 2017
    WSCript.Echo("2017")
  Case 2018
    WSCript.Echo("2018")
  Case else
    WSCript.Echo("Another Year")
End Select
WSCript.Echo(d & " " & m & " " & y)



devryClass = ""
devryClass = InputBox(″Enter the DeVry Class you are:″)
Select Case devryClass
  Case "COMP230"
     MsgBox ″You are correct!″
  Case Else
     MsgBox ″You might be lost!″
End Select


myVar = ""
myVar = InputBox(″Enter The Expected Options:″)
Select Case myVar
  Case "Good_Option_1"
     MsgBox ″You entered an expected option #1!″
  Case "Good_Option_2"
     MsgBox ″You entered an expected option #2!″
  Case "Good_Option_3"
     MsgBox ″You entered an expected option #3!″
  Case "Good_Option_4"
     MsgBox ″You entered an expected option #4!″
  Case "Good_Option_5"
     MsgBox ″You entered an expected option #5!″
  Case "Good_Option_6"
     MsgBox ″You entered an expected option #6!″
  Case "Good_Option_7"
     MsgBox ″You entered an expected option #7!″
  Case Else
     MsgBox ″You entered an UNEXPECTED option″
End Select




' For Next Loop (Counter)
WSCript.Echo("For Next Counter")
For n = 1 to 10
WSCript.Echo(n)
Next

WSCript.Echo("For Next Counter")
For n = 10 to 1 step -1
WSCript.Echo(n)
Next


WSCript.Echo("For Next Counter")
dim numbers
numbers = Array(100,200,300,400,500)
For n = 0 to 4
    WSCript.Echo "numbers(" & n & ")" & vbTab & numbers(n)
Next


WSCript.Echo("For Next Counter")
dim groceries
groceries = Array("Bread","Fruit","Veggies","Milk","Beer","Steak","Cat Food")
For Each item In groceries
WSCript.Echo item
Next

' Same as previous without For Next loop
WSCript.Echo("For Next Counter")
dim groceries
groceries = Array("Bread","Fruit","Veggies","Milk","Beer","Steak","Cat Food")
WSCript.Echo groceries(0)
WSCript.Echo groceries(1)
WSCript.Echo groceries(2)
WSCript.Echo groceries(3)
WSCript.Echo groceries(4)
WSCript.Echo groceries(5)
WSCript.Echo groceries(6)




' Do Until Loop
countNumber = 0
Do Until countNumber = 5
countNumber = countNumber + 1
yourAge = inputBox("Count # " & countNumber & "Type Your Age:")
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
'	 WSCript.Echo("Count # " & countNumber)
Loop




' Do While Loop
countNumber = 0
Do While countNumber < 5
countNumber = countNumber + 1
yourAge = inputBox("Count # " & countNumber & "Type Your Age:")
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
'	 WSCript.Echo("Count # " & countNumber)
Loop




