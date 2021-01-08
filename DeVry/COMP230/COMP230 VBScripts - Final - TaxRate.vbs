'-----------------------------------------------
' Tax Rate Calculation
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