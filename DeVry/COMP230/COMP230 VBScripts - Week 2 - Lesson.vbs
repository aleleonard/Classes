
>wscript {scriptname.vbs}
>cscript {scriptname.vbs}

' ============= OUTPUT METHODS ====================
' methods that can be used to send text output to the Window CLI or the desktop.
WScript.Echo "Output Message" & {var} & {CONST} & Chr7 & vbTab & vbCrlf

' method we will use for console output uses
WScript.StdOut.Write("Output Message" & {var} & {CONST} & Chr(7) & vbTab & vbCrlf)
WScript.StdOut.WriteBlankLines(4)
WScript.StdOut.WriteLine("Output Message with CrLf" & {var} & {CONST} & Chr7 & vbTab)

' method creates a pop-up window on the desktop that displays text information
Set WshShell = Wscript.CreateObject("Wscript.Shell")
WshShell.Popup "Output Message"

' Define the message you want to see inside the message box. 
Dim msg = "Output Message" 

' Display a simple message box.
MsgBox(msg)

' ============= INPUT METHODS ====================

' method we will use for console input uses
WScript.StdOut.Write("Enter input.......")
input = WScript.StdIn.ReadLine()




