
CHROME_RUN_AND_MIN:
			
	IF (A_ComputerName = "7-ASUS-GL522VW") 
		RETURN
	; IF (A_ComputerName = "4-ASUS-GL522VW") 
		; RETURN
		
	; Process, Exist, chrome.exe
	; If ErrorLevel=0  ; errorlevel will = 0 if process doesn't exist
	
	AUTO_RELOAD_FACEBOOK_VAR=0
	Element=Your Notifications - Google Chrome
	WinGetTITLE, TITLE_VAR, %Element% ahk_class Chrome_WidgetWin_1
	IF INSTR(TITLE_VAR,Element)
		AUTO_RELOAD_FACEBOOK_VAR=1
	Element=Rain Alarm - Google Chrome
	WinGetTITLE, TITLE_VAR, %Element% ahk_class Chrome_WidgetWin_1
	IF INSTR(TITLE_VAR,Element)
		AUTO_RELOAD_FACEBOOK_VAR=1
		
	; MSGBOX % AUTO_RELOAD_FACEBOOK_VAR " -- " Element

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
			MSGBOX CHROME.EXE NOT EXIST TO FIND
			RETURN
		}	
		
		; IF (A_ComputerName = "2-ASUS-EEE") 
		IF (A_ComputerName = "1-ASUS-X5DIJ") 
		CHROME_PAGE=https://www.rain-alarm.com/?from=chrome2
		IF (A_ComputerName = "2-ASUS-EEE") 
		CHROME_PAGE=https://www.facebook.com/notifications
		IF (A_ComputerName = "3-LINDA-PC") 
		CHROME_PAGE=https://www.facebook.com/notifications
		IF (A_ComputerName = "4-ASUS-GL522VW") 
		CHROME_PAGE=https://www.facebook.com/notifications
		IF (A_ComputerName = "4-LINDA-PC") 
		CHROME_PAGE=https://www.facebook.com/notifications
		IF (A_ComputerName = "5-ASUS-P2520LA") 
		CHROME_PAGE=https://www.rain-alarm.com/?from=chrome2
		IF (A_ComputerName = "8-MSI-GP62M-7RD") 
		CHROME_PAGE=https://www.facebook.com/notifications
		
		FN_VAR_EXE=%FN_VAR_EXE%" "%CHROME_PAGE%
		Run, "%FN_VAR_EXE%" ; ,,MIN --- SET MIN AT LOAD RUN DOES NOT WORKK
		SoundBeep , 2500 , 100
		SET_DONE=TRUE
	}
	
	
	SET_DONE=TRUE
	IF SET_DONE=TRUE 
	{
		style_CHROME=-2
		WinWait, ahk_class Chrome_WidgetWin_1
		EXIT_LOOP=10 ; ---- DO A FEW MIGHT AS WELL
		LOOP
		{
			; WinMinimize ahk_class Chrome_WidgetWin_1
			WinMAXIMIZE ahk_class Chrome_WidgetWin_1
			WinGet style_CHROME, MinMax, ahk_class Chrome_WidgetWin_1
			SoundBeep , 2500 , 100
			;1 maximized 0 normal -1 minimized
			If style_CHROME=1
			{
				EXIT_LOOP-=1
				IF EXIT_LOOP<0
					BREAK
			}
			SLEEP 50
		}
	}
RETURN
