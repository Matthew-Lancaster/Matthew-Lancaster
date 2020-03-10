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

; for n, GivenPath in A_Args  ; For each parameter (or file dropped onto a script):
; {
    ; Loop Files, %GivenPath%, FD  ; Include files and directories.
        ; PATH_TARGET := A_LoopFileFullPath
; }

FOR n, param in A_Args  ; For each parameter:
{
	PARAM_SET=%PARAM_SET%%A_Space%%PARAM%
}	

PARAM_SET:=StrReplace(PARAM_SET, "*", " ")
PATH_TARGET=%PARAM_SET%

; -------------------------------------------------------------------
; -------------------------------------------------------------------
PATH_TARGET=D:\DSC\2015+SONY\2020 CyberShot HX60V\DCIM\2020 03 10\GGG\DDD
; PATH_TARGET=D:\DSC\2015+SONY\2020 CyberShot HX60V\DCIM\2020 03 10\2020-03-10 13-37-09 - SONY DSC-HX60V - DSC01700.JPG
; -------------------------------------------------------------------

; STRIP QUOTES -- PATH_TARGET:=StrReplace(PATH_TARGET, """" , "")
PATH_TARGET:=StrReplace(PATH_TARGET, """" , "")

IF !PATH_TARGET
	RETURN

WinGetTITLE, TITLE_VAR_BT_MAIL_A, %PATH_TARGET%
WinGet Path, ProcessPath, %PATH_TARGET%
	
; IF INSTR(PATH,"C:\Windows\explorer.exe")>0
	; RETURN

; IF !TITLE_VAR_BT_MAIL_A


If !FileExist(PATH_TARGET)
{
	oText := StrSplit(PATH_TARGET, "\")
	Loop, % oText.MaxIndex()
	{
		
		VAR_ARR:=oText[A_Index]
		PATH_NAME=%PATH_NAME%%VAR_ARR%\
		PATH_NAME_2=%PATH_NAME_2%%PATH_NAME%`n
		
		; MSGBOX % PATH_NAME "`n" PATH_NAME_2
		
		
	}

	Sort, PATH_NAME_2, ReverseDirection
	MSGBOX % PATH_NAME_2

	
}

; IF SUBSTR(oText[A_Index],1,1)<>"#"
; IF TRIM(oText[A_Index])<>""


FN_VAR_EXE=C:\Windows\explorer.exe
IF FileExist(FN_VAR_EXE)
{
	FN_VAR_EXE:="C:\Windows\explorer.exe "
	if FileExist(PATH_TARGET)
		FN_VAR_EXE:="C:\Windows\explorer.exe /select, "
	if InStr(FileExist(PATH_TARGET), "D")
		FN_VAR_EXE:="C:\Windows\explorer.exe "
	
	Run, %FN_VAR_EXE% %PATH_TARGET%
	
	SoundBeep , 2500 , 100
	Soundplay, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav

}
RETURN
