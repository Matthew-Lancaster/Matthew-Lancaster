











;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
;# __ 
;# __ Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START     TIME [ Sun 28-Apr-2019 14:38:04 ]
;# LAST EDIT TIME [ Tue 20-Aug-2019 15:40:00 ]
;# __ 
;  =============================================================

; ------------------------------------------------------------------
; Location Internet
;---------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOHOTKEY/Autokey%20--%2019-SCRIPT_TIMER_UTIL_2.ahk
; ----
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; SESSION 001
; -------------------------------------------------------------------
; -------------------------------------------------------------------


; -------------------------------------------------------------------
; SESSION 002
; -------------------------------------------------------------------
; -------------------------------------------------------------------


#Warn
#NoEnv
#SingleInstance Force
; #NoTrayIcon

; -------------------------------------------------------------------
#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp HAVE AR CALL ONTO ExitFunc
; -------------------------------------------------------------------

SetTitleMatchMode, 3 
; SetTitleMatchMode, slow

DetectHiddenText, On
DetectHiddenWindows, On

; -------------------------------------------------------------------
; Register a function to be called on exit:
; OnExit("ExitFunc")

; Register an object to be called on exit:
; OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

; ---------------------------------------------------------------
; ---------------------------------------------------------------



; DetectHiddenWindows, on
; SetStoreCapslockMode, off

; SoundBeep , 2000 , 100
; SoundBeep , 2500 , 100


 
; OSVER_N_VAR:=a_osversion
; IF INSTR(a_osversion,".")>0
	; OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
; IF OSVER_N_VAR=WIN_XP
	; OSVER_N_VAR=5
; IF OSVER_N_VAR=WIN_7
	; OSVER_N_VAR=6
; IF OSVER_N_VAR=10
; 	WANT_GO=TRUE

; -------------------------------------------------------------------
; WRITE CODER TIME
; -------------------------------------------------------------------
; Sat 11-Jan-2020 12:18:21
; Sat 11-Jan-2020 15:30:00 -- 3 HOUR 15 MINUTE
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; WANT_GO=
; IF (A_ComputerName<>"7-ASUS-GL522VW")
	; WANT_GO=TRUE
; IF (A_ComputerName<>"4-ASUS-GL522VW")
	; WANT_GO=TRUE
TempPID=
GOSUB GOSUB_R
; GOSUB CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK_MIDNIGHT_2
; GOSUB PROCESS_PRIORITY_CLASS_MODIFY_2
; GOSUB PROCESS_PRIORITY_CLASS_MODIFY
; GOSUB PROCESS_PRIORITY_CLASS_MODIFY_REAL

; SETTIMER PROCESS_PRIORITY_CLASS_MODIFY_2,100
; SETTIMER CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK_MIDNIGHT_2,1000

	
RETURN
 




; This can be used to identify the HWND and make a context sensitive hotkey a window for which shares ahk_exe and Window title with multiple other windows,
; and is killed/remade during script run time. For this to work, the window needs to have a unique set of controls.
;
; Possible extention: Consider WinText and do Exclude if possible.
; 
; Start script.
; Setup

;; WinList is a listbox control with delim set to newline

; If InStr(TempPID, "Uploader for Dropbox, Drive - Google Chrome")>0 
; {
	; MSGBOX %LIST%
	; MSGBOX %PID%


	; Process, Priority, %PROCESS_PID%, L

; }
; Return

/* PROCESS_PRIORITY_CLASS_MODIFY_2:

Process, Exist, chrome.exe
NewPID = %ErrorLevel%  ; Save the value immediately since ErrorLevel is often changed.
	
if NewPID
{ ; process exists!
    WinGetClass, ClassID, ahk_pid %NewPID%   ; ClassID will be read here for the process
    WinGetTitle, Title, ahk_pid %NewPID% ; Title will contain the processe's first window's title
	IfWinExist ahk_class %ClassID% ; this will find the first window by the ClassID
    {
        WinGet, WinID, ID ; this will get the ID of the window into WinID variable
        ; WinActivate ; this will bring this window to front (not necessary for example)  
        ; ListVars ; this will display your variables
        ; Pause
    }
    IfWinExist %Title% ; this will find the first window with the window title
    {
        WinGet, WinID, ID
        ; WinActivate ; this will bring this window to front (not necessary for example)  
        ; ListVars
        ; Pause
		; msgbox %Title%

		

    }
}
RETURN
 */
 
GOSUB_R:
this_title=
this_class=
this_id=
PID_8=

hwndDesktop := DllCall("User32.dll\GetDesktopWindow", "UPtr")
; WinGet, id, list,,, Program Manager 
WinGet, id, list,,, ahk_id %hwndDesktop%
Loop, %id% { this_id := id%A_Index% 
; WinActivate, ahk_id %this_id%
WinGetClass, this_class, ahk_id %this_id% 
WinGetTitle, this_title, ahk_id %this_id%
WinGet, id, list,,, Program Manager
Loop, %id%
{
    this_id := id%A_Index%
    ; WinActivate, ahk_id %this_id%
    
	WinGet, PID_8, PID, ahk_id %this_id%
	WinGetClass, this_class, ahk_id %this_id%
    WinGetTitle, this_title, ahk_id %this_id%
	WinGet, PATH_FULL, ProcessPath, ahk_id %this_id%
    ; MsgBox, 4, , Visiting All Windows`n%a_index% of %id%`nahk_id %this_id%`nahk_class %this_class%`n%this_title%`n`nContinue?
    ;IfMsgBox, NO, break
	
	IF INSTR(this_title,"Uploader for Dropbox, Drive - Google Chrome")>0
	{
		Process, Priority, %PID_8%, n
		; msgbox "set low"
	}
	
	
	String2 = %PATH_FULL%
	StringUpper, String2, String2
	PATH_FULL = %String2%
	IF INSTR(this_title,"Uploader for Dropbox, Drive - Google Chrome")=0
	IF INSTR(PATH_FULL,"CHROME.EXE")>0
	{
		; MSGBOX %this_title%
		Process, Priority, %PID_8%, R
		; msgbox "set REAL" %PID_8%
	}

	
} 
 
return
 
 
CLOSE_ALL_VB__AHK_CLASS_WNDCLASS_DESKED_GSK_MIDNIGHT_2:

	HDesktop := DllCall("User32.dll\GetDesktopWindow", "UPtr")	
	
	WinGet, List, List, ahk_class Chrome_WidgetWin_1
	; WinGet, List, List, "ahk_id" %HDesktop%
	; WinGet, List, List, "ahk_exe " chrome.exe
	; WinGet, List, List, "ahk_exe " i)\\chrome\.exe$



	PATH_ID_BUILD=
	Loop %List%  
	{
			
		
		WinGet, PID_8, PID, % "ahk_id " List%A_Index% 
		
		WinGetTitle, Title, % "ahk_id " List%A_Index% 
		ifWinExist "ahk_id " List%A_Index%
			msgbox %Title%
		msgbox List%A_Index%
		WinGetClass, class_Title, % "ahk_id " List%A_Index%
		WinGet, PATH_FULL, ProcessPath, % "ahk_id " List%A_Index% 
		WinGet, PATH_EXE, ProcessName, % "ahk_id " List%A_Index% 
		StringUpper PATH_EXE, PATH_EXE
		StringUpper PATH_FULL, PATH_FULL
				
			IF INSTR(class_Title,"Uploader for Dropbox, Drive - Google Chrome")>0
			{
				Process, Priority, %PID_8%, n
				; msgbox "set low"
			}
			
			IF INSTR(Title,"Uploader for Dropbox, Drive - Google Chrome")>0
			{
				Process, Priority, %PID_8%, n
				; msgbox "set low"
			}
			
			IF INSTR(Title,"Uploader for Dropbox, Drive - Google Chrome")=0
			IF INSTR(PATH_FULL,"VB_KEEP_RUNNER.exe")=0
			IF INSTR(PATH_FULL,"C:\PROGRAM FILES (X86)\GOOGLE\CHROME\")>0
			{
			
				IF INSTR(PATH_ID_BUILD,PATH_EXE)=0
				{
					PATH_ID_BUILD=%PATH_ID_BUILD%%PATH_EXE%`n
					; msgbox %PATH_ID_BUILD% %PID_8%
					PROCESS, EXIST, %PID_8%
					IF ERRORLEVEL
					{
					
						Process, Priority, %PID_8%, R
						; msgbox "set real"	
						
					
						; TOOLTIP % PATH_ID_BUILD
						; SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
						; WINCLOSE ahk_exe %PATH_EXE%
						; SLEEP 200
						; PROCESS, EXIST, %PATH_EXE%
						; IF ERRORLEVEL
						; {
						;	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
							; WINCLOSE ahk_exe %PATH_EXE%
						; }
						; SLEEP 200
						; PROCESS, EXIST, %PATH_EXE%
						; IF ERRORLEVEL
						; {
						; 	SOUNDPLAY, %a_scriptDir%\Autokey -- 10-READ MOUSE CURSOR ICON\start.wav
							; Process, Close, %PATH_EXE%
						; }
					}
				}

			}
	}	
	; TOOLTIP
	; SLEEP 500
	; ---------------------------------------------------------------
	; AFTER ALL GONE
	; RE_RUNNER VB_KEEP_RUNNER
	; ---------------------------------------------------------------
	; GOSUB SUB_RESTORE_VB_KEEP_RUNNER
	; FN_VAR_1 := "D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.exe"
	; IfExist, %FN_VAR_1%
		; Run, %FN_VAR_1%
	; ---------------------------------------------------------------
	; ---------------------------------------------------------------
RETURN


; -------------------------------------------------------------------
; CODE IS RUN TIMER FROM HERE
; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2.ahk
; RUN_CHKDSK_FOR_MEDIA_CAR_V_DRIVE:
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; WRITE CODER TIME
; -------------------------------------------------------------------
; Sat 11-Jan-2020 12:18:21
; Sat 11-Jan-2020 15:30:00 -- 3 HOUR 15 MINUTE
; -------------------------------------------------------------------
PROCESS_PRIORITY_CLASS_MODIFY:
	; PROCESS_TITLE:="Uploader for Dropbox, Drive - Google Chrome"
	; Controls=
	; OutputVar1=
	; OutputVar2=
	; MSG=
	; hwndDesktop=
	; DetectHiddenWindows, on
	; DetectHiddenText, On
	; SetTitleMatchMode, 3
 	
	; hwndDesktop := DllCall("User32.dll\GetDesktopWindow", "UPtr")
	; CHROME:="chrome.exe"
	; WinGet, ControlList, ControlList, ahk_exe %CHROME%
	; HDesktop := DllCall("User32.dll\GetDesktopWindow", "UPtr")
	; WinGet ControlList, ProcessName, ahk_id %HDeskTop%
	; Loop, Parse, ControlList, `n
	; {
		; ClassNN := A_LoopField
		; ControlGetText, OutputVar, %ClassNN%, ahk_id %hwndDesktop%
		; WinGet ControlList, ProcessName, ahk_id %hwndDesktop%
		
		; Loop, Parse, A_LoopField, %A_Tab%
		; {
			; RowNumber := A_Index
			; MSG=%MSG% RowNumber #%RowNumber% Col #%A_Index% is #%A_LoopField%
			; MSGBOX %MSG%
			
			; If InStr(MSG, PROCESS_TITLE)
			; {
				; Process, Priority, %PROCESS_PID%, L

			; }
		; }
	; }


	
RETURN


PROCESS_PRIORITY_CLASS_MODIFY_REAL:

	; PROCESS_TITLE:="Uploader for Dropbox, Drive - Google Chrome"
	; PROCESS_PID=
	
	; ControlGet, Items, List,, ComboBox1, ahk_class Chrome_WidgetWin_1
	; Loop %List%
	; { 
		; GL:=List%A_Index% 
		; WinGet, PROCESS_PID, PID, AHK_ID %GL%
		; WinGetTitle, THIS_TITLE, AHK_ID %GL%
		; If InStr(THIS_TITLE, PROCESS_TITLE)
		; {
		; Process, Priority, %PROCESS_PID%, L
		; MSGBOX %THIS_TITLE%
		; }
	; }
	
RETURN


