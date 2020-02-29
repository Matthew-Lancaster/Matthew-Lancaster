;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 79-LOAD URL AT BOOT CHROME.ahk
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 79-BROWSER LOAD URL AT AUTO BOOT CHROME.ahk
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 79-BROWSER LOAD URL BOOT CHROME.ahk
;# __ 
;# __ Autokey -- 79-LOAD URL AT BOOT CHROME.ahk                -- BEFORE
;# __ Autokey -- 79-BROWSER LOAD URL AT AUTO BOOT CHROME.ahk   -- BEFORE
;# __ Autokey -- 79-BROWSER LOAD URL BOOT CHROME.ahk           -- HERE
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

; -------------------------------------------------------------------
; -------------------------------------------------------------------
; SESSION 02
; -------------------------------------------------------------------
; RENAME AND COMPLETE WORK
; -------------------------------------------------------------------
; FROM __ Mon 17-Feb-2020 03:14:57
; TO   __ Mon 17-Feb-2020 04:27:46 -- 1 HOUR 13 MINUTE
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

	AUTO_RELOAD_FACEBOOK_VAR=1
	AUTO_RELOAD_FIREFOX_VAR=1
	AUTO_RELOAD_RAIN_ALARM_VAR=1
	
	; -- DOUBLE FOR FIREFOX 01 OF 02
	Element:="matt.lan@btinternet.com - BT Yahoo Mail - Mozilla Firefox"
	WinGetTITLE, TITLE_VAR, %Element% ahk_class MozillaWindowClass
	IF INSTR(TITLE_VAR,Element)
		AUTO_RELOAD_FIREFOX_VAR=
	
	; -- DOUBLE FOR FIREFOX 02 OF 02 
	; -- CHARACTER NOT PULLER CORRECT -- Yahoo – login - Mozilla Firefox
	Element:="login - Mozilla Firefox"
	WinGetTITLE, TITLE_VAR, %Element% ahk_class MozillaWindowClass
	IF INSTR(TITLE_VAR,Element)
	IF AUTO_RELOAD_FIREFOX_VAR
		AUTO_RELOAD_FIREFOX_VAR=

	Element:="Your notifications - Google Chrome"
	WinGetTITLE, TITLE_VAR, %Element% ahk_class Chrome_WidgetWin_1
	IF INSTR(TITLE_VAR,Element)
		AUTO_RELOAD_FACEBOOK_VAR=
		
	Element:="Rain Alarm - Google Chrome"
	WinGetTITLE, TITLE_VAR, %Element% ahk_class Chrome_WidgetWin_1
	IF INSTR(TITLE_VAR,Element)
		AUTO_RELOAD_RAIN_ALARM_VAR=

	FIREFOX_EXE:="C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
	AttributeString := FileExist(FIREFOX_EXE)
	IF !AttributeString
		FIREFOX_EXE:="C:\Program Files\Mozilla Firefox\firefox.exe"
	AttributeString := FileExist(FIREFOX_EXE)
	IF !AttributeString
		FIREFOX_EXE=

	CHROME_EXE:="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
	AttributeString := FileExist(CHROME_EXE)
	IF !AttributeString
		CHROME_EXE:="C:\Program Files\Google\Chrome\Application\chrome.exe"
	AttributeString := FileExist(CHROME_EXE)
	IF !AttributeString
		CHROME_EXE=

	; SET DO AR
	IF (A_ComputerName = "1-ASUS-X5DIJ") 
		CHROME_PAGE=https://www.facebook.com/notifications
	IF (A_ComputerName = "2-ASUS-EEE") 
		CHROME_PAGE=https://www.rain-alarm.com/?from=chrome2
	IF (A_ComputerName = "4-ASUS-GL522VW") 
	{
		CHROME_PAGE=https://www.facebook.com/notifications
		FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
	}
	IF (A_ComputerName = "5-ASUS-P2520LA") 
		CHROME_PAGE=https://www.rain-alarm.com/?from=chrome2
	IF (A_ComputerName = "7-ASUS-GL522VW") 
		CHROME_PAGE=https://www.facebook.com/notifications
	IF (A_ComputerName = "8-MSI-GP62M-7RD") 
	{
		CHROME_PAGE=https://www.facebook.com/notifications
		FIREFOX_PAGE=https://mail.yahoo.com/d/folders/1
	}

	; REMOVE IF NOT REQUIRE
	IF (A_ComputerName = "1-ASUS-X5DIJ" and !AUTO_RELOAD_FACEBOOK_VAR) 
		CHROME_PAGE=
	IF (A_ComputerName = "2-ASUS-EEE" and !AUTO_RELOAD_RAIN_ALARM_VAR) 
		CHROME_PAGE=
	IF (A_ComputerName = "4-ASUS-GL522VW" and !AUTO_RELOAD_FACEBOOK_VAR) 
		CHROME_PAGE=
	IF (A_ComputerName = "4-ASUS-GL522VW" and !AUTO_RELOAD_FIREFOX_VAR)
		FIREFOX_PAGE=
	IF (A_ComputerName = "5-ASUS-P2520LA" and !AUTO_RELOAD_RAIN_ALARM_VAR) 
		CHROME_PAGE=
	IF (A_ComputerName = "7-ASUS-GL522VW" and !AUTO_RELOAD_FACEBOOK_VAR) 
		CHROME_PAGE=
	IF (A_ComputerName = "8-MSI-GP62M-7RD" and !AUTO_RELOAD_FACEBOOK_VAR) 
		CHROME_PAGE=
	IF (A_ComputerName = "8-MSI-GP62M-7RD" and !AUTO_RELOAD_FIREFOX_VAR) 
		FIREFOX_PAGE=
	
	IF !CHROME_EXE
	{
		MSGBOX CHROME.EXE NOT EXIST TO FIND
		CHROME_PAGE=
	}
		
	IF !FIREFOX_EXE
	{
		MSGBOX FIREFOX.EXE NOT EXIST TO FIND
		FIREFOX_PAGE=
	}

		
	SET_DONE_FIREFOX=
	SET_DONE_CHROME=
		
	; ---------------------------------------------------------------
	; USE COLON WITH VARIABLE -- BUT NOT IF GOT @ SIGN
	; EXAMPLE ABOVE NOT FIND WITHOUT
	; AND OKAY HERE
	; ---------------------------------------------------------------
	IF CHROME_PAGE
	{
		LAUNCH_EXE_GO=%CHROME_EXE%" "%CHROME_PAGE%
		Run, "%LAUNCH_EXE_GO%" ; ,,MIN --- SET MIN AT LOAD RUN DOES NOT WORKK
		SoundBeep , 2500 , 100
		SET_DONE_CHROME=TRUE
	}
	
	IF FIREFOX_PAGE
	{
		LAUNCH_EXE_GO=%FIREFOX_EXE%" "%FIREFOX_PAGE%
		Run, "%LAUNCH_EXE_GO%" ; ,,MIN --- SET MIN AT LOAD RUN DOES NOT WORKK
		SoundBeep , 2500 , 100
		SET_DONE_FIREFOX=TRUE
	}
	
	IF 	SET_DONE_CHROME=TRUE 
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

	IF SET_DONE_FIREFOX=TRUE
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
