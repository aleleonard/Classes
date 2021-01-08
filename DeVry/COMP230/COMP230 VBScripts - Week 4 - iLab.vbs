
' ******** WEEK 4 - iLab **********

' VBScript: IP_Array_start.vbs
' Written by: Alejandro Jaque
' Date: Jan 25, 2016
' Class: COMP230
' Professor: Ray Blankenship
' ===================================
' Below is an initialize a 2-dimension 
' array of IP Address. The first index 
' +100 is the room# and the second index
' +1 is the computer# in the room.

dim ipAddress(5,3)
  ipAddress(0,0)="192.168.10.11"
  ipAddress(0,1)="192.168.10.12"
  ipAddress(0,2)="192.168.10.13"
  ipAddress(0,3)="192.168.10.14"
  ipAddress(1,0)="192.168.10.19"
  ipAddress(1,1)="192.168.10.20"
  ipAddress(1,2)="192.168.10.21"
  ipAddress(1,3)="192.168.10.22"
  ipAddress(2,0)="192.168.10.27"
  ipAddress(2,1)="192.168.10.28"
  ipAddress(2,2)="192.168.10.29"
  ipAddress(2,3)="192.168.10.30"
  ipAddress(3,0)="192.168.10.35"
  ipAddress(3,1)="192.168.10.36"
  ipAddress(3,2)="192.168.10.37"
  ipAddress(3,3)="192.168.10.38"
  ipAddress(4,0)="192.168.10.43"
  ipAddress(4,1)="192.168.10.44"
  ipAddress(4,2)="192.168.10.45"
  ipAddress(4,3)="192.168.10.46"
  ipAddress(5,0)="192.168.10.51"
  ipAddress(5,1)="192.168.10.52"
  ipAddress(5,2)="192.168.10.53"
  ipAddress(5,3)="192.168.10.54" 

' Define Script Variables
roomStr = ""
compStr = ""
room = 0
computer = 0

' Displays Prompt to User asking for Room #
WScript.StdOut.Write("Please Enter the Room Number (100-105) ...... ")
roomStr = WScript.StdIn.ReadLine()
room = CInt(roomStr)
' Validates Room #
While room < 100 OR room > 105
WScript.StdOut.WriteLine(chr(7) & chr(7) & "Error, 100 to 105 Only!!! ") ' Beeps (chr(7)) twice + Error Message
WScript.StdOut.Write("Please Enter the Room Number (100-105) ...... ") ' Asks User again if wrong value entered
roomStr = WScript.StdIn.ReadLine()
       room = CInt(roomStr)
Wend

' Displays Prompt to User asking for Computer #
WScript.StdOut.Write("Please Enter the Computer Number (1-4) ...... ")
compStr = WScript.StdIn.ReadLine()
computer = CInt(compStr)
' Validates Computer #
While computer < 1 OR computer > 4
WScript.StdOut.WriteLine(chr(7) & chr(7) & "Error, 1 to 4 Only!!! ") ' Beeps (chr(7)) twice + Error Message
WScript.StdOut.Write("Please Enter the Computer Number (1-4) ...... ") ' Asks User again if wrong value entered
compStr = WScript.StdIn.ReadLine()
       computer = CInt(compStr)
Wend

' Prints response from IPAddresses Array
WScript.StdOut.WriteLine()
WScript.StdOut.WriteLine("The IP Address in Room " & roomStr & _
 " for Computer " & compStr & " is " & ipAddress(room-100,computer-1))
WScript.StdOut.WriteLine()

' Display All IP Address Y/N?
WScript.StdOut.Write("Do you wish to Display all of the P Addresses (Y/N) ..... ")
ans = WScript.StdIn.ReadLine()
' Validates User's reply + Prints out full array
If ans = "Y" OR ans = "y" Then
WScript.StdOut.WriteLine()
   For room = 0 to 5
    For computer = 0 to 3
      WScript.StdOut.WriteLine("The IP Address in Room " & (room+100) & _
 " for Computer " & (computer+1) & " is " & ipAddress(room,computer))
    Next
   Next
End If





