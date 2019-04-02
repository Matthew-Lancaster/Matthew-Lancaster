@echo off
if exist c:\temp\sssssr1.txt del c:\temp\sssssr1.txt
echo hello >c:\temp\startbup.txt
verify on
cd c:\
cd d:\
cd e:\
cd f:\
cd g:\
cd h:\

ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   e:\ drive f:
ECHO ----------------------------------------------- 

xCOPY e:\.         f:         /s /m /f /h /k /r /c /x /y
rem xCOPY e:\.         f:         /s /d /f /h /k /r /c /x /y

if exist c:\bat\shutdown.tmp copy c:\bat\shutdown.tmp c:\temp\sssssr1.txt
pause
