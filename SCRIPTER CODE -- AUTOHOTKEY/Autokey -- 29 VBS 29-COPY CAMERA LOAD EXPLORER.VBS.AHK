;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\
;# __ 
;# __ Autokey -- 29 VBS 29-COPY CAMERA LOAD EXPLORER.VBS.AHK
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; HERE USER CALL FROM
; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 29-COPY CAMERA PHOTO IMAGES.VBS
; -------------------------------------------------------------------
; OPERATION DO AR LOAD EXPLORER WITH THE PATH
; BY PARAMETER ARGUE GIVER -- 1ST AHK SCRIPT LIKE THIS ONE
; WHEN UP TO 84 FOLDER
; ONLY IF ANOTHER EXPLORER NOT ALREADY LOAD OF SAME
; -------------------------------------------------------------------
; SAVE THE QUESTION HAS TO ASK FROM VBSCRIPT 
; IF WANT EXPLORER WHEN ANOTHER ALREADY LOAD ABOUT COUNT
; 
; -------------------------------------------------------------------
; FROM __ Tue 14-Jan-2020 22:39:59
; TO   __ Tue 14-Jan-2020 00:04:00 -- 1 HOUR 20 MINUTE
; -------------------------------------------------------------------


#noEnv
; #persistent
#singleInstance force
detectHiddenWindows, on
setWorkingDir %a_scriptDir%
; #NoTrayIcon

; -------------------------------------------------------------------
; CODE INITIALIZE
; -------------------------------------------------------------------

GOSUB EXPLORER_ROUTINE

EXITAPP

EXPLORER_ROUTINE:

for n, GivenPath in A_Args  ; For each parameter (or file dropped onto a script):
{
    Loop Files, %GivenPath%, FD  ; Include files and directories.
        LongPath_1 := A_LoopFileFullPath
}

; -------------------------------------------------------------------
; TESTER DEBUG TRY
; -------------------------------------------------------------------
; LongPath_1=D:\DSC\2015+SONY\2020 CyberShot HX60V\DCIM\2020 03 10\2020-03-10 13-37-09 - SONY DSC-HX60V - DSC01700.JPG
; -------------------------------------------------------------------

LongPath_2="%LongPath_1%"

; STRIP QUOTES -- LongPath_2:=StrReplace(LongPath_2, """" , "")
LongPath_2:=StrReplace(LongPath_2, """" , "")

IF !LongPath_1
	RETURN

WinGetTITLE, TITLE_VAR_BT_MAIL_A, %LongPath_1%
WinGet Path, ProcessPath, %LongPath_1%
	
; IF INSTR(PATH,"C:\Windows\explorer.exe")>0
	; RETURN

; IF !TITLE_VAR_BT_MAIL_A

FN_VAR_EXE=C:\Windows\explorer.exe
IF FileExist(FN_VAR_EXE)
{
	FN_VAR_EXE:="C:\Windows\explorer.exe "
	if InStr(FileExist(LongPath_2), "D")
	{
		MSGBOX % FN_VAR_EXE LongPath_2
		Run, %FN_VAR_EXE% %LongPath_2%
		
		SoundBeep , 2500 , 100
		Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
		RETURN
	}
	
	FN_VAR_EXE:="C:\Windows\explorer.exe /select, "
	if FileExist(LongPath_2)
	{

		Run, %FN_VAR_EXE% %LongPath_2%
		
		SoundBeep , 2500 , 100
		Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
	}
}
RETURN
