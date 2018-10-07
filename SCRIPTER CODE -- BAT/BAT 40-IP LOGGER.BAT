@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- --- 24-May-2018 04:02400 Created Time
@REM __ TO   ---- Sun 07-Oct-2018 04:45:37 Last Editor _ Update Time+Header Info
@REM ----------------------------------------------------------------


CD\

REM nslookup myip.opendns.com. resolver1.opendns.com

REM >%temp%\MyIP.TXT

REM wget -q -O %temp%\MyIP.TXT http://www.whatismyip.com/automation/n09230945.asp

for /f "delims= " %%G in ('"C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 23-MY IP.VBS"') do set PublicIP=%%G

echo. & echo Your public IP address is %PublicIP% & echo.

PAUSE