' ************ Final Exam Exercises *************
'-----------------------------------------------
' GrossPay Calculation
'-----------------------------------------------
payRate = InputBox("Please enter your hourly rate")
hoursWorks = InputBox("Please enter amount of hours worked")
if hoursWorks <= 40 Then
 regularPay = hoursWorks * payRate
 WScript.Echo hoursWorks & " hours were entered at USD$ " & payRate & " per hour" & vbCrlf & vbCrlf & _
 hoursWorks & " hours will be paid at the regular hourly rate. Total of USD$ " & regularPay
else
 regularPay = 40 * payRate
 extraHours = hoursWorks - 40
 extraPay = extraHours * (payRate * 1.5)
 grossPay = regularPay + extraPay
 WScript.Echo hoursWorks & " hours were entered at USD$ " & payRate & " per hour" & vbCrlf & vbCrlf & _
 "40 hours will be paid at the regular hourly rate." & vbTab & "USD$ " & regularPay & vbCrlf & _
 extraHours & " hours will be paid at the 1.5 hourly rate." & vbTab & vbTab & "USD$ " & extraPay & vbCrlf & _
 "Total Gross Pay" & vbTab & vbTab & vbTab & vbTab & "USD$ " & grossPay
end If


'-----------------------------------------------
' GrossPay Modular Calculation
'-----------------------------------------------
payRate = payRateUI()
hoursWorks = hoursWorksUI()

if hoursWorks <= 40 Then
call Calculates_Regulartime
else
call Calculates_Overtime
end If

Function payRateUI()
payRateUI = InputBox("Please enter your hourly rate")
End Function

Function hoursWorksUI()
hoursWorksUI = InputBox("Please enter amount of hours worked")
End Function

Sub Calculates_Regulartime
 regularPay = hoursWorks * payRate
 WScript.Echo hoursWorks & " hours were entered at USD$ " & payRate & " per hour" & vbCrlf & vbCrlf & _
 hoursWorks & " hours will be paid at the regular hourly rate. Total of USD$ " & regularPay
End Sub

Sub Calculates_Overtime
 regularPay = 40 * payRate
 extraHours = hoursWorks - 40
 extraPay = extraHours * (payRate * 1.5)
 grossPay = regularPay + extraPay
 WScript.Echo hoursWorks & " hours were entered at USD$ " & payRate & " per hour" & vbCrlf & vbCrlf & _
 "40 hours will be paid at the regular hourly rate." & vbTab & "USD$ " & regularPay & vbCrlf & _
 extraHours & " hours will be paid at the 1.5 hourly rate." & vbTab & vbTab & "USD$ " & extraPay & vbCrlf & _
 "Total Gross Pay" & vbTab & vbTab & vbTab & vbTab & "USD$ " & grossPay
End Sub




'-----------------------------------------------
' GrossPay Calculation as all hrs 1.5
'-----------------------------------------------
basePay = InputBox("Please enter your hourly rate")
hoursWorked = InputBox("Please enter amount of hours worked")
grossPay = hoursWorked * (basePay * 1.5)
WScript.Echo hoursWorked & " hours were entered at USD$ " & basePay & " per hour paid at 1.5" & vbCrlf & vbCrlf & _
 "Total Gross Pay of USD$ " & grossPay

 
 
 
 
'-----------------------------------------------
' TAXRATE Calculation
'-----------------------------------------------
CONST_TAXRATE = 0.25
basePay = 1000
bonusPay = 500
taxPercentage = CONST_TAXRATE * 100
grossPay = basePay + bonusPay
taxWithheld = grossPay * CONST_TAXRATE
netPay = grossPay - (grossPay * CONST_TAXRATE)
WScript.Echo "basePay :" & vbTab & vbTab & "USD$ " & basePay & vbCrlf & _
"bonusPay :" & vbTab & "USD$ " & bonusPay & vbCrlf & _
"grossPay :" & vbTab & vbTab & "USD$ " & grossPay & vbTab & vbCrlf & _
"TAXRATE :" & vbTab & "% " & taxPercentage & vbCrlf & _
"taxWithheld :" & vbTab & "USD$ " & taxWithheld & vbCrlf & _
"netPay :" & vbTab & vbTab & "USD$ " & netPay




'-----------------------------------------------
' Do While Loop to display all intergers in numArray(100)
'-----------------------------------------------
arrayIndex = 0
arrayValue = 1
do Until arrayIndex = 101
WSCript.Echo "Array numArray(" & arrayIndex & ") = " & arrayValue
arrayIndex = arrayIndex + 1
arrayValue = arrayValue + 1
loop

