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
ECHO STARTING TO COPY   c:\ drive to d: 
ECHO ----------------------------------------------- 

xCOPY C:\.         D:         /s /d /f /h /k /r /c /x /y
xCOPY C:\.         D:         /s /m /f /h /k /r /c /x /y

rem pause

ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   e:\ drive f:
ECHO ----------------------------------------------- 

xCOPY e:\.         f:         /s /d /f /h /k /r /c /x /y
xCOPY e:\.         f:         /s /m /f /h /k /r /c /x /y

ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   g:\ drive to H:
ECHO ----------------------------------------------- 

xCOPY g:\.         H:         /s /d /f /h /k /r /c /x /y
xCOPY g:\.         H:         /s /m /f /h /k /r /c /x /y

ECHO -----------------------------------------------
ECHO ----Now Creating Lists of files to delete------ 
ECHO ----------------------------------------------- 

xCOPY d:\.         c:         /l /s /d /f /h /k /r /y >C:\list-c-todel.txt
xCOPY c:\list-c-todel.txt d:\ /y /f

xCOPY f:\.         e:         /l /s /d /f /h /k /r /y >C:\list-e-todel.txt
xCOPY c:\list-e-todel.txt f:\ /y /f
xCOPY c:\list-e-todel.txt e:\ /y /f

xCOPY H:\.         g:         /l /s /d /f /h /k /r /y >C:\list-g-todel.txt
xCOPY c:\list-g-todel.txt h:\ /y /f
xCOPY c:\list-g-todel.txt g:\ /y /f


ECHO ----------------------------------------------- 
echo Finished.........................
ECHO ----------------------------------------------- 
dir c:\temp\startbup.txt
dir c:\list-g-todel.txt
ECHO ----------------------------------------------- >>c:\coolcid\backup.txt
ECHO ----------------------------------------------- Start  Backup Time >>c:\coolcid\backup.txt
dir c:\temp\startbup.txt >>c:\coolcid\backup.txt
dir c:\list-g-todel.txt >>c:\coolcid\backup.txt
ECHO ----------------------------------------------- Finish Backup Time >>c:\coolcid\backup.txt
ECHO ----------------------------------------------- >>c:\coolcid\backup.txt
if exist c:\bat\shutdown.tmp copy c:\bat\shutdown.tmp c:\temp\sssssr1.txt

sleep
