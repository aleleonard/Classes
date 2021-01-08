' Function MaxNum accepts numbers and return value
firstNum = InputBox("First Number?")
secondNum = InputBox("Second Number?")

funcValue = MaxNum(firstNum, secondNum)
WScript.Echo "Function 'MaxNum' returned " & funcValue & " as a value."

Function MaxNum(firstNum, secondNum)
if firstNum > secondNum Then
  WScript.Echo "First Number was : " & firstNum & vbCrlf & _
  "Second Number was : " & secondNum & vbCrlf & vbCrlf & _
  "FIRST NUMBER is bigger!" & vbCrlf 
  funcReturn = firstNum
  elseif firstNum < secondNum Then
  WScript.Echo "First Number was : " & firstNum & vbCrlf & _
  "Second Number was : " & secondNum & vbCrlf & vbCrlf &  _
  "SECOND NUMBER is bigger!" & vbCrlf 
  funcReturn = secondNum
  elseif firstNum = secondNum Then
  WScript.Echo "First Number was : " & firstNum & vbCrlf & _
  "Second Number was : " & secondNum & vbCrlf & vbCrlf & _
  "BOTH NUMBERS are the same!" & vbCrlf 
  funcReturn = CStr(firstNum & " = " & secondNum)
End If
 MaxNum = funcReturn
End Function

' References
'
' Microsoft. (2015). Function statement. VBScript Language Reference. Retrieved from https://msdn.microsoft.com/en-us/library/x7hbf8fa(v=vs.84).aspx



' Function MaxNum just process calculation
firstNum = InputBox("First Number?")
secondNum = InputBox("Second Number?")
call MaxNum

Function MaxNum
if firstNum > secondNum Then
  WScript.Echo "First Number was :" & vbTab & vbTab & firstNum & vbCrlf & _
  "Second Number was :" & vbTab & secondNum & vbCrlf & vbCrlf & _
"First Number is bigger!"
  elseif firstNum < secondNum Then
  WScript.Echo "First Number was :" & vbTab & vbTab & firstNum & vbCrlf & _
  "Second Number was :" & vbTab & secondNum & vbCrlf & vbCrlf & _
  "Second Number is bigger!"
  elseif firstNum = secondNum Then
  WScript.Echo "First Number was :" & vbTab & firstNum & vbCrlf & _
  "Second Number was :" & vbTab & secondNum & vbCrlf & vbCrlf & _
  "Both Numbers are the same!"
End If
End Function




Dim firstNum
Dim secondNum

firstNum = InputBox("Enter First Number")
secondNum = InputBox("Enter Second Number")

funcValue = MaxNum(firstNum, secondNum)

WScript.Echo "Function 'MaxNum' returned " & funcValue & " as a value."

Function MaxNum(firstNum, secondNum)
  if firstNum > secondNum Then
      WScript.Echo "First Number was : " & firstNum & vbCrlf & _
      "Second Number was : " & secondNum & vbCrlf & vbCrlf & _
      "FIRST NUMBER is bigger!" & vbCrlf 
    funcReturn = firstNum
  elseif firstNum < secondNum Then
      WScript.Echo "First Number was : " & firstNum & vbCrlf & _
      "Second Number was : " & secondNum & vbCrlf & vbCrlf &  _
      "SECOND NUMBER is bigger!" & vbCrlf 
    funcReturn = secondNum
  elseif firstNum = secondNum Then
      WScript.Echo "First Number was : " & firstNum & vbCrlf & _
      "Second Number was : " & secondNum & vbCrlf & vbCrlf & _
      "BOTH NUMBERS are the same!" & vbCrlf 
    funcReturn = CStr(firstNum & " = " & secondNum)
End If
 MaxNum = funcReturn
End Function


' References
'
' Microsoft. (2015). Function statement. VBScript Language Reference. Retrieved from https://msdn.microsoft.com/en-us/library/x7hbf8fa(v=vs.84).aspx

