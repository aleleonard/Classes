' ************ WEEK 1 - iLab *****************
cls
@echo off
rem Name: Alejandro Jaque
rem DSI#: D01242011
rem Date: Monday, January 4, 2016
echo "MyCommands.cmd Script is now running"
echo.
netsh interface ip show config "NIC"
netsh interface ip set address "NIC" source=static 150.100.10.10 255.255.255.0 150.100.10.1 0
timeout /t 6 > nul
netsh interface ip show config "NIC"
netsh interface ip set address "NIC" source=dhcp
timeout /t 6 > nul
netsh interface ip show config "NIC"
echo.
shutdown /s /t 60 /c "Local Shutdown in 1 minute!!"
timeout /t 4 > nul
shutdown /a
echo "Shutdown has been Aborted"
echo.
echo "End of Script"


