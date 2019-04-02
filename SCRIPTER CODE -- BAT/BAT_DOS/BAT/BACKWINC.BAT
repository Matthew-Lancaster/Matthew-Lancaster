@echo off
ncc /50
cd d:\
beep
c:
cd\
beep
echo Creating List of files on C:\listfiles-C-drive.txt
beep
dir c:/s/Oa >C:\listfiles-C-drive.txt
time >>c:\list-of-files-copied-history.txt <c:\bat\e-to-f.dat
date >>c:\list-of-files-copied-history.txt <c:\bat\e-to-f.dat
echo ----copying c: to d:--STAGE 5/5---------------------------------------------
echo ----copying c: to d:--STAGE 5/5--------------------------------------------- >>c:\list-of-files-copied-history.txt
rem echo Press control c or break if you want to continue later.......
rem pause
rem format d:/q<c:\bat\backwin.dat
rem label d:<c:\bat\backwin2.dat
xcopy c:\. d: /L/d/e/c/h/k >>c:\list-of-files-copied-history.txt
xcopy c:\. d: /d/e/c/h/k 
beep
echo **********************************************************************
echo *******************************Finished*******************************
echo **********************************************************************
echo ********************************************************************** >>c:\list-of-files-copied-history.txt
echo *******************************Finished******************************* >>c:\list-of-files-copied-history.txt
echo ********************************************************************** >>c:\list-of-files-copied-history.txt
echo Will run my-history\list.exe to get a comparison of c & d drives
pause
c:
cd\my-his~1
list
rem pause
beep
rem restart
