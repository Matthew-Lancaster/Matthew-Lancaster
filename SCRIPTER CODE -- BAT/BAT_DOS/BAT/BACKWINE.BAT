@echo off
ncc /50
cd f:\
beep
e:
cd\
beep
echo Creating List of files on e:\listfiles-e-drive.txt
beep
dir e:/s/Oa >e:\listfiles-e-drive.txt
time >>c:\list-of-files-copied-history.txt <c:\bat\e-to-f.dat
date >>c:\list-of-files-copied-history.txt <c:\bat\e-to-f.dat
echo ----copying e: to f:--STAGE 1/5---------------------------------------------
echo ----copying e: to f:--STAGE 1/5--------------------------------------------- >>c:\list-of-files-copied-history.txt
rem echo Press control c or break if you want to continue later.......
rem pause
rem format f:/q<c:\bat\backwin.dat
rem label f:<c:\bat\backwin2.dat
xcopy e:\. f: /L/d/e/c/h/k >>c:\list-of-files-copied-history.txt
xcopy e:\. f: /d/e/c/h/k 
beep
echo **********************************************************************
echo *******************************Finished*******************************
echo **********************************************************************
echo ********************************************************************** >>c:\list-of-files-copied-history.txt
echo *******************************Finished******************************* >>c:\list-of-files-copied-history.txt
echo ********************************************************************** >>c:\list-of-files-copied-history.txt
beep
rem restart
