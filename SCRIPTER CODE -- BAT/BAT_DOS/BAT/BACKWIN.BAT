@echo off
ncc /50
cd d:\
beep
cd e:\
beep
cd f:\
beep
cd g:\
beep
cd h:\
beep
cd i:\
beep
cd k:\
beep
cd l:\
beep
c:
beep
cd\
beep
echo Creating List of files on C:\listfiles-C-drive.txt
beep
dir c:/s/Oa >C:\listfiles-C-drive.txt
beep
echo Creating List of files on E:\listfiles-E-drive.txt
beep
dir e:/s/Oa >E:\listfiles-E-drive.txt
copy E:\listfiles-E-drive.txt c:\
beep
echo Creating List of files on G:\listfiles-G-drive.txt
beep
dir g:/s/Oa >G:\listfiles-G-drive.txt
copy g:\listfiles-g-drive.txt c:\
beep
echo Creating List of files on H:\listfiles-H-drive.txt
beep
dir h:/s/Oa >H:\listfiles-H-drive.txt
copy h:\listfiles-h-drive.txt c:\
beep
echo Creating List of files on I:\listfiles-I-drive.txt
beep
dir i:/s/Oa >I:\listfiles-I-drive.txt
copy i:\listfiles-i-drive.txt c:\
beep
echo ----copying E: to F:--STAGE 1/5---------------------------------------------
echo ----copying E: to F:--STAGE 1/5--------------------------------------------- >c:\list-of-files-copied-history.txt
rem format f:/q<c:\bat\backwin.dat
beep
time >>c:\list-of-files-copied-history.txt <c:\bat\e-to-f.dat
date >>c:\list-of-files-copied-history.txt <c:\bat\e-to-f.dat
xcopy e:\. f: /L/d/e/c/h/k >>c:\list-of-files-copied-history.txt
beep
xcopy e:\. f: /d/e/c/h/k
echo ----copying G: to J:--STAGE 2/5---------------------------------------------
echo ----copying G: to J:--STAGE 2/5--------------------------------------------- >>c:\list-of-files-copied-history.txt
beep
rem format j:/q<c:\bat\backwin.dat
beep
xcopy g:\. j: /L/d/e/c/h/k >>c:\list-of-files-copied-history.txt
beep
xcopy g:\. j: /d/e/c/h/k
echo ----copying H: to K:--STAGE 3/5---------------------------------------------
echo ----copying H: to K:--STAGE 3/5--------------------------------------------- >>c:\list-of-files-copied-history.txt
beep
rem format k:/q<c:\bat\backwin.dat
beep
xcopy h:\. k: /L/d/e/c/h/k >>c:\list-of-files-copied-history.txt
beep
xcopy h:\. k: /d/e/c/h/k
echo ----copying I: to L:--STAGE 4/5---------------------------------------------
echo ----copying I: to L:--STAGE 4/5--------------------------------------------- >>c:\list-of-files-copied-history.txt
beep
rem format l:/q<c:\bat\backwin.dat
beep
xcopy i:\. l: /l/d/e/c/h/k >>c:\list-of-files-copied-history.txt
beep
xcopy i:\. l: /d/e/c/h/k
echo ----copying c: to d:--STAGE 5/5---------------------------------------------
echo ----copying c: to d:--STAGE 5/5--------------------------------------------- >>c:\list-of-files-copied-history.txt
echo Press control c or break if you want to continue later.......
rem pause
REM format d:/q<c:\bat\backwin.dat
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
rem pause
c:
cd\my-his~1
list
rem pause
beep
rem restart
