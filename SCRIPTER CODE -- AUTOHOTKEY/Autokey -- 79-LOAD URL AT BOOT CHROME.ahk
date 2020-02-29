;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 79-LOAD URL AT BOOT CHROME.ahk
;# __ 
;# __ Autokey -- 79-LOAD URL AT BOOT CHROME.ahk
;# __ 
;# __ BY Matthew Lancaster 
;# __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 01
; -------------------------------------------------------------------
; WANT TO SEPARATE THIS ONE SO ABLE RUN FROM MY CODE MENU OPTION
; -------------------------------------------------------------------
; FROM __ Fri 07-Jun-2019 07:16:33
; TO   __ Fri 07-Jun-2019 08:51:00
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
SoundBeep , 1500 , 400
; SetStoreCapslockMode, off


GOSUB CHROME_RUN_AND_MIN

; -------------------------------------------------------------------
EXITAPP
; -------------------------------------------------------------------

; -------------------------------------------------------------------
RETURN
; -------------------------------------------------------------------


CHROME_RUN_AND_MIN:
			
	SetTitleMatchMode 2  ; PARTIAL PATH

	AUTO_RELOAD_FACEBOOK_VAR=0
	AUTO_RELOAD_FIREFOX_VAR_2=0
	AUTO_RELOAD_RAIN_ALARM_VAR=0
	
	Element=matt.lan@btinternet.com - BT Yahoo Mail - Mozilla Firefox
	WinGetTITLE, TITLE_VAR, %Element% ahk_class MozillaWindowClass
	IF INSTR(TITLE_VAR,Element)
		AUTO_RELOAD_FIREFOX_VAR_2=1

	Element:=Your notifications - Google Chrome
	WinGetTITLE, TITLE_VAR, %Element% ahk_class Chrome_WidgetWin_1
	IF INSTR(TITLE_VAR,Element)
		AUTO_RELOAD_FACEBOOK_VAR=1
		
	Element:=Rain Alarm - Google Chrome
	WinGetTITLE, TITLE_VAR, %Element% ahk_class Chrome_WidgetWin_1
	IF INSTR(TITLE_VAR,Element)
		AUTO_RELOAD_RAIN_ALARM_VAR=1
	
	; ---------------------------------------------------------------
	; USE COLON WITH VARIABLE -- BUT NOT IF GOT @ SIGN
	; EXAMPLE ABOVE NOT FIND WITHOUT
	; AND OKAY HERE
	; ---------------------------------------------------------------
	
		
	SET_DONE=FALSE
	IF AUTO_RELOAD_FACEBOOK_VAR=0
	{
		; ahk_class rctrl_renwnd32
		FN_VAR_EXE:="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
		AttributeString := FileExist(FN_VAR_EXE)
		IF !AttributeString
			FN_VAR_EXE:="C:\Program Files\Google\Chrome\Application\chrome.exe"
		AttributeString := FileExist(FN_VAR_EXE)
		IF !AttributeString
		{
			AUTO_RELOAD_FACEBOOK_VAR=1
			MSGBOX CHROME.EXE NOT EXIST TO FIND
		}	
		
		IF (A_ComputerName = "1-ASUS-X5DIJ") 
		CHROME_PAGE=https://www.facebook.com/notifications
		IF (A_ComputerName = "2-ASUS-EEE") 
		CHROME_PAGE=https://www.rain-alarm.com/?from=chrome2
		; IF (A_ComputerName = "3-LINDA-PC") 
		; CHROME_PAGE=https://www.facebook.com/notifications
		IF (A_ComputerName = "4-ASUS-GL522VW") 
		CHROME_PAGE=https://www.facebook.com/notifications
		IF (A_ComputerName = "5-ASUS-P2520LA") 
		CHROME_PAGE=https://www.rain-alarm.com/?from=chrome2
		IF (A_ComputerName = "7-ASUS-GL522VW") 
		CHROME_PAGE=https://www.facebook.com/notifications
		IF (A_ComputerName = "8-MSI-GP62M-7RD") 
		CHROME_PAGE=https://www.facebook.com/notifications
		
		IF AUTO_RELOAD_FACEBOOK_VAR=0
		IF CHROME_PAGE
			FN_VAR_EXE=%FN_VAR_EXE%" "%CHROME_PAGE%
		Run, "%FN_VAR_EXE%" ; ,,MIN --- SET MIN AT LOAD RUN DOES NOT WORKK
		SoundBeep , 2500 , 100
		SET_DONE=TRUE
	}
	

	SET_DONE_2=FALSE
	IF AUTO_RELOAD_FIREFOX_VAR_2=0
	{
		; ahk_class rctrl_renwnd32
		FN_VAR_EXE:="C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
		AttributeString := FileExist(FN_VAR_EXE)
		IF !AttributeString
			FN_VAR_EXE:="C:\Program Files\Mozilla Firefox\firefox.exe"
		AttributeString := FileExist(FN_VAR_EXE)
		IF !AttributeString
		{
			AUTO_RELOAD_FIREFOX_VAR_2=1
			MSGBOX FIREFOX.EXE NOT EXIST TO FIND
		}	
		
		; IF (A_ComputerName = "1-ASUS-X5DIJ") 
		; FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
		; IF (A_ComputerName = "2-ASUS-EEE") 
		; FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
		; IF (A_ComputerName = "3-LINDA-PC") 
		; FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
		IF (A_ComputerName = "4-ASUS-GL522VW") 
		FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
		IF (A_ComputerName = "5-ASUS-P2520LA") 
		FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
		IF (A_ComputerName = "7-ASUS-GL522VW") 
		FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
		IF (A_ComputerName = "8-MSI-GP62M-7RD") 
		FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
		
		IF AUTO_RELOAD_FIREFOX_VAR_2=0
		IF FIREFOX_PAGE
		FN_VAR_EXE=%FN_VAR_EXE%" "%FIREFOX_PAGE%
		Run, "%FN_VAR_EXE%" ; ,,MIN --- SET MIN AT LOAD RUN DOES NOT WORKK
		SoundBeep , 2500 , 100
		SET_DONE_2=TRUE
	}
	
	SET_DONE=TRUE
	IF SET_DONE=TRUE 
	{
		STYLE_CHROME=-2
		WinWait, ahk_class Chrome_WidgetWin_1
		EXIT_LOOP=10 ; ---- DO A FEW MIGHT AS WELL
		LOOP
		{
			; WinMinimize ahk_class Chrome_WidgetWin_1
			WinMAXIMIZE ahk_class Chrome_WidgetWin_1
			WinGet STYLE_CHROME, MinMax, ahk_class Chrome_WidgetWin_1
			SoundBeep , 2500 , 100
			;1 maximized 0 normal -1 minimized
			If STYLE_CHROME=1
			{
				EXIT_LOOP-=1
				IF EXIT_LOOP<0
					BREAK
			}
			SLEEP 50
		}
	}

	IF SET_DONE=TRUE
	{
		STYLE_FIREFOX=-2
		WinWait, ahk_class MozillaWindowClass
		EXIT_LOOP=10 ; ---- DO A FEW MIGHT AS WELL
		LOOP
		{
			WINMINIMIZE ahk_class MozillaWindowClass
			WinGet STYLE_FIREFOX, MinMax, ahk_class MozillaWindowClass
			SoundBeep , 2500 , 100
			;1 maximized 0 normal -1 minimized
			If STYLE_FIREFOX=-1
			{
				EXIT_LOOP-=1
				IF EXIT_LOOP<0
					BREAK
			}
			SLEEP 50
		}
	}
	
	
RETURN
