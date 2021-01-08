'-----------------------------------------------
' GrossPay Calculation as all hrs 1.5
'-----------------------------------------------
basePay = InputBox("Please enter your hourly rate")
hoursWorked = InputBox("Please enter amount of hours worked")
grossPay = hoursWorked * (basePay * 1.5)
WScript.Echo hoursWorked & " hours were entered at USD$ " & basePay & " per hour paid at 1.5" & vbCrlf & vbCrlf & _
 "Total Gross Pay of USD$ " & grossPay
