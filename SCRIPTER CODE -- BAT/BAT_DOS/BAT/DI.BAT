@echo off
c:\callerid\2diaryz

call C:\HISTORY\BACKDIAR

call c:\callerid\editpro2

if errorlevel 2 goto winedit

if errorlevel 1 goto proedit

goto ende2

:proedit
if exist c:\temp\editlock.txt del c:\temp\editlock.txt
cALL c:\callerid\editpro

:loop2
if not exist c:\temp\editlock.txt goto loop2
del c:\temp\editlock.txt
goto ende


:winedit
CALL winedit c:\callerid\2diary.txt
goto ende

:ende
call C:\HISTORY\BACKDIAR

:ende2
