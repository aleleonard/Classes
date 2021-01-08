' Pseudocode
processRecords( )
    read record
     while not eof
      if contactState = "ca" then
     print contactName  ‘Returns a value
  endif
  read record
 endwhile
return 


num1 = GetNumber()
num2 = GetNumber()

call DisplayNumbers(num1,num2)
call SwapNumbers(num1,num2)
call DisplayNumbers(num1,num2)

WScript.StdOut.WriteLine(vbCrlf & "End of Program")
' ===================================================
Function GetNumber()
 WScript.StdOut.Write("Please enter a number value.....")
 GetNumber=CInt(WScript.StdIn.ReadLine())
 WScript.StdOut.WriteBlankLines(1)
End Function

Sub DisplayNumbers(number1,number2)
 WScript.StdOut.WriteLine("The Value of num1 is " & CStr(number1) & _
 " and the value of num2 is " & CStr(number2) & vbCrlf)
end Sub

Sub SwapNumbers(ByRef number1, ByRef number 2)
 temp = number1
 number1 = number2
 number2 = temp
end Sub