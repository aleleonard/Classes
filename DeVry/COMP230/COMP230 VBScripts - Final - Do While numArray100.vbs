'-----------------------------------------------
' DO While Loop to display all intergers in numArray(100)
'-----------------------------------------------
arrayIndex = 0
arrayValue = 1
do Until arrayIndex = 101
WSCript.Echo "Array numArray(" & arrayIndex & ") = " & arrayValue
arrayIndex = arrayIndex + 1
arrayValue = arrayValue + 1
loop
