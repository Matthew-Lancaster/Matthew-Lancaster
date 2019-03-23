@echo off
if exist c:\vb\pulse\pulse2.obj del c:\vb\pulse\pulse2.obj
c:\vb\pulse\pulse2 4
call edit c:\vb\pulse\pulsem.txt
c:\vb\pulse\pulse2 3
pause
rem restart
