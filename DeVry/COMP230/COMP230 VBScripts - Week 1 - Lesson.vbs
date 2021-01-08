
' *** Week 1 - Lesson ***

' clear screen
cls

' turn off command echo
echo off

'  cmd in the Open textbox, where username is the username of a user who is a member of the Administrators' group.
runas /user:{username}

' gives very basic information (i.e., IP address, subnet mask, and default gateway).
ipconfig

' gives additional information about the active NICs
ipconfig /all 

'  display only the basic IP configuration settings on single interface "Wired" on Windows 7
netsh interface ip show config "Wired"

' display basic IP configuration for all the interfaces without IPv6 tunnel interfaces
netsh interface ipv4 show config

' display more complete IP configuration information for your NIC interfaces
ipconfig /all | more

' assigning a static IP configuration to an NIC
netsh interface ip set address "Wired" source=static {IP address} {mask} {GW} {metric}

' assigning a dynamic DHCP configuration to an NIC
netsh interface ip set address "Wired" source=dhcp

' Windows 7 CLI command sequence that changes the IP address from DHCP to static and back to DHCP. On a Windows XP, the only difference is the use of ipconfig instead of 
netsh interface ip show config "Wired"

' Windows 7 PC, the local computer will be shut down in 60 seconds.
shutdown /s /t 60 /c "Shutdown in 60 seconds"

' computer named jlmW7laptop will be forcibly restarted in 120 seconds (or 2 minutes) with the message Restart in 2 Minutes!!
shutdown -r -f -m \\jlmW7laptop /t 120 -c "Shutdown in 2 minutes"

' Start spooler
net start spooler

' Stop spooler
net stop spooler

' To view the print queue of MyPrinter1 on print server W2KPRN1:
net print \\W2KPRN1\MyPrinter1

' To delete print job number 2 on MyPrinter1 on print server W2KPRN1:
net print \\W2KPRN1 2/delete

' To redirect print output for the LPT1 port to MyPrinter1 on print server W2KPRN1:
net use LPT1: \\W2KPRN1\MyPrinter1

' To display information about the LPT1 port:
net use LPT1:

' To list the contents of the Dotmatrix print queue on the \\Production computer, type:
net print \\production\dotmatrix

' To assign the disk-drive device name E: to the Letters shared directory on the \\Financial server, type:
net use e: \\financial\letters

' To connect the user identifier Dan as if the connection were made from the Accounts domain, type:
net use d:\\server\share /user:Accounts\Dan

' To assign the disk-drive device name F: to a file cabinet in an MSN Internet Access community called TargetName using the Passport account UserName@passport.com, type:
net use f: http://www.msnusers.com/ TargetName /user: UserName @passport.com

' To access share as user...
net use f: http://www.msnusers.com/ TargetName /user: UserName @passport.com

' To disconnect from the \\Financial\Public directory, type:
net use f: \\financial\public /delete

' To connect to the resource memos shared on the \\Financial 2 server, type:
net use k: "\\financial 2" \memos

' To restore the current connections at each logon, regardless of future changes, type:
net use /persistent:yes




