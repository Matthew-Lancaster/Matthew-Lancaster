#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;

;---------------------------------------------------------
;WANT SELECT ALL LINE AND PASTE ONTO IT
;WANT COPY ON IT OWN
;WANT HOLD CTRL UNTIL ASK IT STOP FOR LINK URL IN WEB PAGE
;---------------------------------------------------------
;
;--------------------
#SingleInstance force
;--------------------
;--------------------
^F4::
		x = 0
		SoundBeep
	   
		;WinGetTitle, active_title, A
		
		;MSGBOX %O_active_title%
		;MSGBOX %active_title%
		;---------------------------------------------------------------------------------------------------------------------
		;1ST ANSWER FOUND BUT NOT GOOD ENOUGH BY WANT NUMERIC HWND INSTEAD
		;---------------------------------------------------------------------------------------------------------------------
		;GOOGLE SEARCH -- BODY SEARCH
		;--------------------------------------------------
		;FINAL GOOGLE SERCH -- AUTOHOTKEYS o_active_title = active_title
		;----
		;<SOLVED> IfWinActive, ahk_class, Title (exact name) Co - Ask for Help - AutoHotkey Community
		;https://autohotkey.com/board/topic/74614-solved-ifwinactive-ahk-class-title-exact-name-co/
		;----
		
		WinGetTitle, active_title, A
		IF active_title <> O_active_title
		{
			SoundBeep
			break
		}
		O_active_title = %active_title%
		
		;------------------------------------------
		;CHROME ACTIVE FIND
		;------------------------------------------
		;NEXT ANSWER IN THIS STAGE OF 2ND BUT NOT TOTALLY
		;IS WANTING NUMERIC OF HWND CHANGE IDEA
		;------------------------------------------
		IfWinNotActive ahk_class Chrome_WidgetWin_1
		{
			;SoundBeep
			break
		}
		;--------------------------------------------------------------------
		;ALL CONDITION LOOK THEN OKAY TO PROCEED 
		;--------------------------------------------------------------------
	    Loop {
				x += 1
				;Send,{PgDn} 
				Send,{Down}

				sleep 40

				if GetKeyState("RButton")
				{
					SoundBeep
					break
				}
				if (x > 4000000)
				{
					SoundBeep
					break
				}
	   }
;}



;-------------------------------------------------------------------------------------------   
;UN USEFUL CODE BLOCK ABOVE
;WinGet nChromeWindows, Count, ahk_class Chrome_WidgetWin_1
;If(nChromeWindows > 0)
;-------------------------------------------------------------------------------------------   
;		if WinNOTActive(ahk_exe chrome.exe)
;		{
;			SoundBeep
;			break
;		}
;-------------------------------------------------------------------------------------------   
   
   
;-------------------------------------------------------------------------
;REPLACE F10 TO DO CONTROL PRINT SCREEN
;FOR CLIPBOARD SCREEN SHOT -- CODE PIC___ WON'T HOT KEY F0 ON WIN 7 AND UP
;-------------------------------------------------------------------------
;F10::
;SEND ^{PrintScreen}

Return

;-------------------------------------------------------------
;MAKE WIN + Y WORK TO DO A SCROLL LOCK FOR WINDOWSE STOP CLEAR
;-------------------------------------------------------------
;NOT WORKING TO DO A WIN + KEY TO RELACE
;LIKE AUTOKEY USE WIN + A -- TO STOP A SPY
;-------------------------------------------------------------
;^::

;#Y::^L

;#Y::^ScrollLock

;#Y::
;SEND ^L
;SEND ScrollLock
;-------------------------------------------------------------

;---------------------------------------------------
;CTRL F FIND
;---------------------------------------------------
;F6::
;   Send,{ctrl down}f{ctrl up}{ENTER}
;   Send,{ctrl down}f{ctrl up}
;---------------------------------------------------
;CTRL F FIND AND ENTER ON HIGH-LIGHTED SELECTED TEXT
;---------------------------------------------------
;F6::
;   Send,{ctrl down}f{ctrl up}{ENTER}