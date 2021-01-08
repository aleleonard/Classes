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


