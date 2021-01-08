' payRate = InputBox("Please enter your hourly rate")
' hoursWorks = InputBox("Please enter amount of hours worked")

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


