echo off
CD\
D:


C:\PStart\Progs\#_SysinternalsSuite\Psexec \\4-ASUS-GL522VW -c -f @C:\TEMP\NET_SHARE_ALL.BAT C:\TEMP\NET_SHARE_ALL.BAT

PAUSE
exit /b

C:\PStart\Progs\#_SysinternalsSuite\Psexec -c -f @c:\temp\complist.txt "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT"

PAUSE
exit /b


REM C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 

NET START RasMan
NET Start RpcLocator





echo oFF

REM C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -accepteula \\2-ASUS-EEE "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 
REM TIMEOUT /T 7

REM C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -s -accepteula \\1-ASUS-X5DIJ "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 
REM TIMEOUT /T 7

REM C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -accepteula -s \\3-LINDA-PC -u "MATT 01" -p " " "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT"
REM TIMEOUT /T 7

echo oN

C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -accepteula -s  -u "MATT 01" -p " " \\4-ASUS-GL522VW "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" C:\TEMP\NET_SHARE_ALL.BAT
TIMEOUT /T 7

GOTO JUMP

C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -accepteula -s \\4-ASUS-GL522VW -u "MATT 01" -p " " "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT"
TIMEOUT /T 7

C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -accepteula -s \\5-ASUS-P2520LA -u "MATT 01" -p " " "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT"
TIMEOUT /T 7

C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -accepteula -s \\8-MSI-GP62M-7RD -u "MATT 01" -p " " "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT"
TIMEOUT /T 7

:JUMP


echo on
C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -accepteula \\4-ASUS-GL522VW "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 
C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -h -i -accepteula \\5-ASUS-P2520LA "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 

@echo off
REM 2-ASUS-EEE
REM 1-ASUS-X5DIJ
REM 3-LINDA-PC
REM 4-ASUS-GL522VW
REM 5-ASUS-P2520LA
REM 7-ASUS-GL522VW
REM 8-MSI-GP62M-7RD


REM C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -e -h -i -s -accepteula \\* "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 

echo on
PAUSE
exit /b


C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -s -h /accepteula \\4-ASUS-GL522VW "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 


C:\PStart\Progs\#_SysinternalsSuite\PsExec.exe -c -f -s -h /accepteula \\8-MSI-GP62M-7RD "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 

	   
PAUSE

exit /b

