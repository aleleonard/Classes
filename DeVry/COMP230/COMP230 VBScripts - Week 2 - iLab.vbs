
' ************ WEEK 2 - iLab *****************

' VBScript: NameAge.vbs
' Written by: Alejandro Jaque
' Date: Jan 11, 2016
' Class: COMP230
' Professor: Ray Blankenship

' Create name and age variables
name = ""
ageStr = ""

' Prompt Uer for Name and Age
WScript.StdOut.Write("Please Enter your Full Name ............ ")
name = WScript.StdIn.ReadLine()
WScript.StdOut.WriteLine()  'Skip 1 line
WScript.StdOut.Write("Please Enter you age ................... ")
ageStr = WScript.StdIn.ReadLine()

' Calculate age+10 and assign to ageStr10
ageStr10 = CStr( CInt(ageStr)+10 )

' Display Name and Age Values
WScript.StdOut.WriteBlankLines(2) 'Skip 2 lines
WScript.StdOut.WriteLine("Your Name is " & vbTab & vbTab & name)
WScript.StdOut.WriteLine("Your Age  is " & vbTab & vbTab & ageStr)
WScript.StdOut.WriteLine(vbCrLf & "Your Age in 10 years is ...... " & _
  ageStr10 & vbCrLf)
WScript.StdOut.WriteLine("End of Program")



' VBScript: PopUpWindow.vbs
' Written by: Alejandro Jaque
' Date: Jan 11, 2016
' Class: COMP230
' Professor: Ray Blankenship

' Create name and age variables
name = "John Doe"
ageStr = "50"

' Calculate age+10 and assign to ageStr10
ageStr10 = CStr( CInt(ageStr)+10 )

' Option 1
' Display Name and Age Values
WScript.Echo "Your Name is " & vbTab & vbTab & name
WScript.Echo "Your Age  is " & vbTab & vbTab & ageStr
WScript.Echo vbCrLf & "Your Age in 10 years is ...... " & _
  ageStr10 & vbCrLf
WScript.Echo "End of Program"

' Option 2
' Build output as a single string msgStr
msgStr = "Your Name is " & vbTab & vbTab & name & _
vbCrLf & "Your Age  is  " & vbTab & vbTab & ageStr & _
vbCrLf & vbCrLf & "Your Age in 10 years is ...... " & _
ageStr10 & vbCrLf & vbCrLf & "End of Program"

'One Echo Statement for all output
Wscript.Echo msgStr



' VBScript: CmdArgs.vbs
' Written by: Alejandro Jaque
' Date: Jan 11, 2016
' Class: COMP230
' Professor: Ray Blankenship

' Check for Command Line Arguments
Set args = WScript.Arguments
If args.Count < 2 then
WScript.Echo "You must enter the name and age as Command Line Arguments!!"
WScript.Sleep(5000)
WScript.Quit
end If
' Assign name and age variables to Cmd Line Args
name = args.item(0)
ageStr = args.item(1)
' calculate Age+10 and assign to ageStr10
ageStr10 = CStr( CInt(ageStr)+10 )
' Build output as a single string msgStr
msgStr = "Your Name is " & vbTab & vbTab & name & _
vbCrLf & "Your Age  is  " & vbTab & vbTab & ageStr & _
vbCrLf & vbCrLf & "Your Age in 10 years is ...... " & _
ageStr10 & vbCrLf & vbCrLf & "End of Program"
'One Echo Statement for all output
Wscript.Echo msgStr




