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
cd i:\
cd j:\
cd k:\
cd l:\
cd m:\
cd n:\
cd O:\
cd P:\

ECHO -----------------------------------------------
ECHO STARTING TO COPY   c:\ drive to d: 
ECHO ----------------------------------------------- 

xCOPY C:\.         D:         /s /m /f /h /k /r /c /x /y
xCOPY C:\.         D:         /s /d /f /h /k /r /c /x /y

rem pause

ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   E:\ drive F: Newer files Only
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 

xCOPY E:\.         F:       /c /u /s /d /f /h /k /r /y 

ECHO -----------------------------------------------
ECHO -----------------------------------------------
ECHO -----------------------------------------------
ECHO -----------------------------------------------
ECHO STARTING TO COPY   F:\ drive E: BACKWARDS COPY NOW
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 

xCOPY F:\.         E:         /s /m /f /h /k /r /c /x /y
xCOPY F:\.         E:         /s /d /f /h /k /r /c /x /y
REM xCOPY F:\.         k:         /s /d /f /h /k /r /c /x /y

ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   g:\ drive to H: BACKWARDS COPY NOW also
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 

xCOPY h:\.         G:         /s /m /f /h /k /r /c /x /y
xCOPY h:\.         G:         /s /d /f /h /k /r /c /x /y

ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   I:\ drive to M:
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 

xCOPY I:\.         M:         /s /m /f /h /k /r /c /x /y
xCOPY I:\.         M:         /s /d /f /h /k /r /c /x /y

ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   J:\ drive to N:
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 

xCOPY J:\.         N:         /s /m /f /h /k /r /c /x /y
xCOPY J:\.         N:         /s /d /f /h /k /r /c /x /y

ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   K:\ drive to O:
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 

xCOPY K:\.         O:         /s /m /f /h /k /r /c /x /y
xCOPY K:\.         O:         /s /d /f /h /k /r /c /x /y

ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO STARTING TO COPY   L:\ drive to P:
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 

xCOPY L:\.         P:         /s /m /f /h /k /r /c /x /y
xCOPY L:\.         P:         /s /d /f /h /k /r /c /x /y

ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO -----------------------------------------------
ECHO ----Now Creating Lists of files to delete------ 
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
echo C:\ drive....
ECHO -------------------------------------- 
xCOPY d:\.         c:         /l /s /d /f /h /k /r /y >C:\drvdata\c-todel.txt
xCOPY c:\drvdata\c-todel.txt d:\ /y /f

ECHO -------------------------------------- 
echo F:\ drive.... Backwards copy
ECHO -------------------------------------- 
xCOPY E:\.         F:         /l /s /d /f /h /k /r /y >C:\drvdata\F-todel.txt
xCOPY c:\drvdata\F-todel.txt f:\ /y /f
xCOPY c:\drvdata\F-todel.txt e:\ /y /f

ECHO -------------------------------------- 
echo H:\ drive.... Backwards copy 
ECHO -------------------------------------- 
xCOPY g:\.         h:         /l /s /d /f /h /k /r /y >C:\drvdata\H-todel.txt
xCOPY c:\drvdata\H-todel.txt h:\ /y /f
xCOPY c:\drvdata\H-todel.txt g:\ /y /f

ECHO -------------------------------------- 
echo I:\ drive....
ECHO -------------------------------------- 
xCOPY M:\.         I:         /l /s /d /f /h /k /r /y >C:\drvdata\I-todel.txt
xCOPY c:\drvdata\I-todel.txt M:\ /y /f
xCOPY c:\drvdata\I-todel.txt I:\ /y /f

ECHO -------------------------------------- 
echo J:\ drive....
ECHO -------------------------------------- 
xCOPY N:\.         J:         /l /s /d /f /h /k /r /y >C:\drvdata\J-todel.txt
xCOPY c:\drvdata\J-todel.txt N:\ /y /f
xCOPY c:\drvdata\J-todel.txt j:\ /y /f

ECHO -------------------------------------- 
echo K:\ drive....
ECHO -------------------------------------- 
xCOPY O:\.         K:         /l /s /d /f /h /k /r /y >C:\drvdata\K-todel.txt
xCOPY c:\drvdata\K-todel.txt O:\ /y /f
xCOPY c:\drvdata\K-todel.txt k:\ /y /f

ECHO -------------------------------------- 
echo L:\ drive....
ECHO -------------------------------------- 
xCOPY P:\.         L:         /l /s /d /f /h /k /r /y >C:\drvdata\L-todel.txt
xCOPY c:\drvdata\L-todel.txt P:\ /y /f
xCOPY c:\drvdata\L-todel.txt L:\ /y /f

ECHO ----------------------------------------------- 
ECHO ----------------------------------------------- 
echo Finished.........................
ECHO ----------------------------------------------- 
dir c:\temp\startbup.txt
dir c:\drvdata\L-todel.txt
ECHO ----------------------------------------------- >>c:\coolcid\backup.txt
ECHO ----------------------------------------------- Start  Backup Time >>c:\coolcid\backup.txt
dir c:\temp\startbup.txt >>c:\coolcid\backup.txt
dir c:\drvdata\L-todel.txt >>c:\coolcid\backup.txt
ECHO ----------------------------------------------- Finish Backup Time >>c:\coolcid\backup.txt
ECHO ----------------------------------------------- >>c:\coolcid\backup.txt
if exist c:\bat\shutdown.tmp copy c:\bat\shutdown.tmp c:\temp\sssssr1.txt

REM SLEEP 40 SECONDS
pause
rem sleep
