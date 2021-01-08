
' *** Week 1 - Discussions ***
' release my IP address, flush the DNS, renew the IP address.

rem This Batch Script releases the IP Address, 
rem flushes the DNS, 
rem and renews the IP Address
@echo off
echo Releasing the IP Address
ipconfig /release
timeout /t 10
echo Flushing DNS
ipconfig /flushdns
timeout /t 10
echo Renewing the IP Address
ipconfig /renew
timeout /t 10
echo You are good now! Smile!

