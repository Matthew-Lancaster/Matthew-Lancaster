@echo off
if exist c:\temp\sssssr1.txt del c:\temp\sssssr1.txt
echo hello >c:\temp\startbup.txt
verify on
cd c:\
cd d:\
cd e:\
cd f:\

ECHO --------------------------------------------------------- 
ECHO  STARTING TO COPY   F:\DOCUME~1\*.* TO drive E:\DOCUME~1
ECHO ---------------------------------------------------------

REM xCOPY F:\DOCUME~1\.         E:\DOCUME~1     /s /m /f /h /k /r /c /x /y
xCOPY F:\DOCUME~1\.         E:\DOCUME~1         /s /d /f /h /k /r /c /x /y
REM xCOPY F:\.         k:         /s /d /f /h /k /r /c /x /y

if exist c:\bat\shutdown.tmp copy c:\bat\shutdown.tmp c:\temp\sssssr1.txt
pause
