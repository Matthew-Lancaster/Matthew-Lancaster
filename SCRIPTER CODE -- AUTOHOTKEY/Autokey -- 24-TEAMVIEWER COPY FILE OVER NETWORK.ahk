;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 24-TEAMVIEWER COPY FILE OVER NETWORK.ahk
;# __ 
;# __ Autokey -- 24-TEAMVIEWER COPY FILE OVER NETWORK.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com
;# START     TIME [ Sunday 00:14:00 Am_22 April 2018 ]
;# END       TIME [ Sunday 03:25:10 Am_22 April 2018 ]
;# __ 
;  =============================================================

#Warn
#NoEnv
#SingleInstance Force
; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp CALLS ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; SCRIPT ============================================================
DetectHiddenWindows, on

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

setTimer TIMER_SUB_1,1

setTimer TIMER_PREVIOUS_INSTANCE, 1
; -------------------------------------------------------------------


;--------------------------------------------------------------------
TIMER_SUB_1:
setTimer TIMER_SUB_1,OFF
	
COMP_N=FALSE
IF (A_ComputerName="4-ASUS-GL522VW") 
	COMP_N=TRUE
IF (A_ComputerName="5-ASUS-P2520LA") 
	COMP_N=TRUE
IF (A_ComputerName="7-ASUS-GL522VW") 
	COMP_N=TRUE
IF (A_ComputerName="8-MSI-GP62M-7RD") 
	COMP_N=TRUE
	
IF COMP_N=FALSE
	RETURN
	
COMPUTER_NAME_11=4-ASUS-GL522VW
COMPUTER_NAME_21=5-ASUS-P2520LA
COMPUTER_NAME_31=7-ASUS-GL522VW
COMPUTER_NAME_41=8-MSI-GP62M-7RD

StringReplace,COMPUTER_NAME_12,COMPUTER_NAME_11,-,_,All
StringReplace,COMPUTER_NAME_22,COMPUTER_NAME_21,-,_,All
StringReplace,COMPUTER_NAME_32,COMPUTER_NAME_31,-,_,All
StringReplace,COMPUTER_NAME_42,COMPUTER_NAME_41,-,_,All


NETPATH_TAB_1=_02_d_drive\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\
NETPATH_TAB_2=\TEAMVIEWER.TXT

NETPATH_1=\\%COMPUTER_NAME_11%\%COMPUTER_NAME_12%%NETPATH_TAB_1%%COMPUTER_NAME_11%%NETPATH_TAB_2%
NETPATH_2=\\%COMPUTER_NAME_21%\%COMPUTER_NAME_22%%NETPATH_TAB_1%%COMPUTER_NAME_21%%NETPATH_TAB_2%
NETPATH_3=\\%COMPUTER_NAME_31%\%COMPUTER_NAME_32%%NETPATH_TAB_1%%COMPUTER_NAME_31%%NETPATH_TAB_2%
NETPATH_4=\\%COMPUTER_NAME_41%\%COMPUTER_NAME_42%%NETPATH_TAB_1%%COMPUTER_NAME_41%%NETPATH_TAB_2%


;MSGBOX %NETPATH_3%`r`n%NETPATH_4%

DF_1=0
DF_2=0
DF_3=0
DF_4=0

if FileExist(NETPATH_1) 
	{
	FileGetTime, DF_1, %NETPATH_1%, M
	}
if FileExist(NETPATH_2) 
	{
	FileGetTime, DF_2, %NETPATH_2%, M
	}
ifExist %NETPATH_3%
	{
	FileGetTime, DF_3, %NETPATH_3%, M
	}
if FileExist(NETPATH_4) 
	{
	FileGetTime, DF_4, %NETPATH_4%, M
	}


DF_8=A_NOW
IF DF_1<%DF_8% 
	DF_8=%DF_1%
IF DF_2<%DF_8% 
	DF_8=%DF_2%
IF DF_3<%DF_8% 
	DF_8=%DF_3%
IF DF_4<%DF_8% 
	DF_8=%DF_4%

NETPATH_COPY=
IF DF_1>%DF_8% 
	{
	DF_8=%DF_1%
	NETPATH_COPY=%NETPATH_1%
	}
IF DF_2>%DF_8%
	{
	DF_8=DF_2
	NETPATH_COPY=%NETPATH_2%
	}
IF DF_3>%DF_8% 
	{
	DF_8=%DF_3%
	NETPATH_COPY=%NETPATH_3%
	}
IF DF_4>%DF_8% 
	{
	DF_8=DF_4
	NETPATH_COPY=%NETPATH_4%
	}

IF NOT (!NETPATH_COPY)
{	
	IF NETPATH_COPY<>%NETPATH_1%
		FileCopy, %NETPATH_COPY%, %NETPATH_1%,1 
	IF NETPATH_COPY<>%NETPATH_2%
		FileCopy, %NETPATH_COPY%, %NETPATH_2%,1 
	IF NETPATH_COPY<>%NETPATH_3%
		FileCopy, %NETPATH_COPY%, %NETPATH_3%,1 
	IF NETPATH_COPY<>%NETPATH_4%
		FileCopy, %NETPATH_COPY%, %NETPATH_4%,1 
}

Exitapp

Return


ShowFileExist(file)
{
  If (FileExist(file) && !InStr(FileExist(file), "D"))
    MsgBox, file: %file% exists.
  Else
    MsgBox, file: %file% does NOT exist.
  Return
}



;----------------------------------------
TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000

if ScriptInstanceExist()
{
	Exitapp
}
return

ScriptInstanceExist() {
	static title := " - AutoHotkey v" A_AhkVersion
	DHW_2 := A_DetectHiddenWindows
	DetectHiddenWindows, On
	WinGet, match, List, % A_ScriptFullPath . title
	DetectHiddenWindows, % DHW_2
	return (match > 1)
	}
Return




;# ----------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ----------------------------------------------------------------------

;# ----------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode)
{
    if ExitReason not in Logoff,Shutdown
    {
        ;MsgBox, 4, , Are you sure you want to exit?
        ;IfMsgBox, No
        ;    return 1  ; OnExit functions must return non-zero to prevent exit.
    }
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}

class MyObject
{
    Exiting()
    {
        ;
        ;MsgBox, MyObject is cleaning up prior to exiting...
        /*
        this.SayGoodbye()
        this.CloseNetworkConnections()
        */
    }
}
;# ------------------------------------------------------------------------------
; exit the app


;GOOD SCRIPT EXAMPLE PAGE HELPER
;----
;1 Hour Software by Skrommel - DonationCoder.com
;http://www.donationcoder.com/Software/Skrommel/index.html#GoneIn60s
;----

;StringMid,now,A_Now,9,4
