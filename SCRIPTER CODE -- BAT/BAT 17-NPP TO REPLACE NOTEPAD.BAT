@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- --- 16-Dec-2017 04:56:00 Created Time
@REM __ TO   ---- Sun 07-Oct-2018 04:45:37 Last Editor _ Update Time+Header Info
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM -- Located in Sync Location 1.. 2..
@REM ----------------------------------------------------------------
@REM -- 1..
@REM ----------------------------------------------------------------
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 17-NPP TO REPLACE NOTEPAD.BAT
@REM ----------------------------------------------------------------
@REM -- 2..
@REM ----------------------------------------------------------------
@REM -- C:\SCRIPTER\SCRIPTER CODE -- REG KEY SETTER\REG KEY\# Windows 7 REG SCRIPT\NOTEPAD++ WORK\notepad++ replace notepad\BAT 17-NPP TO REPLACE NOTEPAD.BAT
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- SYNCED BY
@REM ----------------------------------------------------------------
@REM -- C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 32-COPIER_SYNC.VBS
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 002 SESSSION INFO DO
@REM ----------------------------------------------------------------
@REM __ FROM ---- Sun 07-Oct-2018 09:24:00
@REM __ TO   ---- Sun 07-Oct-2018 09:28:00 Update & SYNC INFO & Time+Header
@REM ----------------------------------------------------------------


CD\

IF NOT EXIST=="%~dp0Notepad++ParamProxy.exe" GOTO JUMP
COPY "%~dp0Notepad++ParamProxy.exe" "C:\Program Files (x86)\Notepad++\Notepad++ParamProxy.exe"
:JUMP

IF NOT EXIST=="E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\NOTEPAD++ WORK\notepad++ replace notepad\Notepad++ParamProxy.exe" GOTO JUMP
COPY "E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\NOTEPAD++ WORK\notepad++ replace notepad\Notepad++ParamProxy.exe" "C:\Program Files (x86)\Notepad++\Notepad++ParamProxy.exe"
:JUMP

IF NOT EXIST=="%~dp0notepad64.exe" GOTO JUMP
COPY "%~dp0notepad64.exe" "C:\Program Files (x86)\Notepad++\notepad64.exe"
:JUMP

IF NOT EXIST=="E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\NOTEPAD++ WORK\notepad++ replace notepad\notepad64.exe" GOTO JUMP
COPY "E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\NOTEPAD++ WORK\notepad++ replace notepad\notepad64.exe" "C:\Program Files (x86)\Notepad++\notepad64.exe"
:JUMP

mklink /j “C:\Program Files\Notepad++” “C:\Program Files (x86)\Notepad++”
REM RMDIR CAN REMOVE

REM NPP REPLACE NOTEPAD
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" /v "Debugger" /t REG_SZ /d "\"C:\Program Files\Notepad++\Notepad++ParamProxy.exe\"" /f


REM ---- NOT WORKER
REM Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" /v "Debugger" /t REG_SZ /d "\"C:\Program Files\Notepad++\notepad64.exe\"" /f


REM RESTORE ORGINAL NOTEPAD
Reg.exe delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" /f


TIMEOUT /T 4
