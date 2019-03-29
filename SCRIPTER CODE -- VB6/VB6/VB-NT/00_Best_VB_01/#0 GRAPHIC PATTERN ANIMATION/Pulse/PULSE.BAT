@echo off
if exist c:\vb\pulse\pulse2.obj del c:\vb\pulse\pulse2.obj
c:\vb\pulse\pulse2 1
call edit c:\vb\pulse\pulse.txt
c:\vb\pulse\pulse2 2
rem pause
rem restart
