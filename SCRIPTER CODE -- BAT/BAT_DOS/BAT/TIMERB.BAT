@ECHO OFF
rem ECHO WAITING TWO MINUTES BEFORE LAUNCHING scanregw.EXE
rem timer3
rem scanregw /autorun
ECHO WAITING FIVE MINUTE BEFORE LAUNCHING WINAMP.EXE
timer2
if errorlevel 1 e:\Progra~1\Winamp\winamp.exe
rem if errorlevel 1 e:\Progra~1\Winamp\winamp.exe e:\mymusic\mp3\1listxxx.m3u
rem was added to end of line above e:\mymusic\mp3\1listxxx.m3u
rem ECHO WAITING ONE MINUTE BEFORE LAUNCHING CI.BAT
rem timer7
rem CD\
rem ci
