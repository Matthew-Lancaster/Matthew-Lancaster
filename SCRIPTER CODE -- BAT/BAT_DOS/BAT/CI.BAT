@echo off
c:
cd \coolcid
rem c:\coolcid\cid-ru~1.exe
rem atzdial

if "%1"=="edit" goto edit
if "%1"=="EDIT" goto edit
if "%1"=="run" goto cid
if "%1"=="RUN" goto cid

if exist c:\coolcid\ci.exe ci.exe
if exist c:\coolcid\ci.exe goto ende
if "%1%!"=="!" goto cid
goto end
:cid
if "%1"=="run" vbdos /ah /RUN c:\coolcid\ci /Es /S:100 /lc:\vbdos\vbdos.qlb           
if "%1"=="RUN" vbdos /ah /RUN c:\coolcid\ci /Es /S:100 /lc:\vbdos\vbdos.qlb           
if "%1!"=="!" vbdos /ah /RUN c:\coolcid\ci /Es /S:100 /lc:\vbdos\vbdos.qlb           
echo -------------------------------------------------------------------------------
if "%1"=="run" echo Been Running CI.BAS in RUN mode.
if "%1"=="RUN" echo Been Running CI.BAS in RUN mode.
if "%1!"=="!" echo Been Running CI.BAS in Run mode.
goto end
:edit
rem basic deletes .exe now if in edit/program mode
echo Deleting ci.EXE's coz in edit mode
if exist c:\coolcid\ci.exe del c:\coolcid\ci.exe
if exist c:\ci.exe del c:\ci.exe
vbdos /ah c:\coolcid\ci /Es /S:100 /lc:\vbdos\vbdos.qlb
call cilink
echo -------------------------------------------------------------------------------
echo Been Running CI.BAS in Edit then compile mode.
goto end2
:ende
echo -------------------------------------------------------------------------------
:end
if exist c:\coolcid\ci.exe echo Been Running CI.EXE.
:end2
echo Return from \coolcid to root level so that ci.bat runs rather than ci.exe neat.
echo Been Running CI.BAT.
cd\
if exist c:\sssssr.txt goto resta
if "%config%"=="WIN" pause
:resta
