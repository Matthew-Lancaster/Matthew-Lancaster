@ECHO OFF
ECHO WAITING TWO MINUTES BEFORE LAUNCHING scanregw.EXE
timer3
scanregw /autorun
c:\utils\restart1 caller
rem ECHO WAITING ONE MINUTE BEFORE LAUNCHING WINAMP.EXE
rem timer2
rem if errorlevel 1 e:\Progra~1\Winamp\winamp.exe
rem if errorlevel 1 e:\Progra~1\Winamp\winamp.exe e:\mymusic\mp3\1listxxx.m3u
rem was added to end of line above e:\mymusic\mp3\1listxxx.m3u
rem ECHO WAITING TWO MINUTE BEFORE LAUNCHING CI.BAT
rem timer7
rem CD\
rem ci
