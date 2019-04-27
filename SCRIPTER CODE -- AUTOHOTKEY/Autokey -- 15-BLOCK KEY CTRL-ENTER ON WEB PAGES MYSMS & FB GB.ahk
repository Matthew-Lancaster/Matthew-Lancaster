;# ------------------------------------------------------------------
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES MYSMS & FB GB.ahk
;# __ 
;# __ Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES MYSMS & FB GB.ahk
;# __ 
;# __ BY Matthew Lancaster __ Matt.Lan@Btinternet.com __ 
;# 
;# Wed 18 October 2017 02:05:00 Work Time Beginner
;# Sat 21 April   2018 04:00:00 2nd Wave of Work Better Coder Emerged
;# Tue 09 October 2018 05:25:00 3nd Wave of Code Facebook Grin Book Name Scriptor & Array 
;#                              Work File Rather Than Hardcode-ed Info for Sharing On-line Privacy
;# __ 
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; Published on Google Blogger
; -------------------------------------------------------------------
; Search Description
; -------------------------------------------------------------------
; Code to Change Shift-Enter Control Enter and Enter for Use when 
; Sending Messenger and Not Wanting to Send on The Enter Key or By 
; Accident Control Enter For Use with MYSMS and GrinBook FB And Any 
; Other May Find Later
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location Internet
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set
;--------------------------------------------------------------------
; ----
; Matthew-Lancaster/Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES MYSMS & FB GB.ahk at master · Matthew-Lancaster/Matthew-Lancaster
; https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2015-BLOCK%20KEY%20CTRL-ENTER%20ON%20WEB%20PAGES%20MYSMS%20%26%20FB%20GB.ahk
; ----
;# ------------------------------------------------------------------

; -------------------------------------------------------------------
; 002 SESSION _ Get Lift Off for Hotkey-er Coder __ On AutoHotKeys 
;               Multiple Use of One or Two Hot Key Intercept Hooker
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; Here is the Code in Autohotkeys to Change Shift-Enter Control Enter 
; and Enter for Use when Sending Messenger and Not Wanting to Send on 
; entering or By Accident Control Enter
; For Use with MYSMS and GrinBook FB And Any Other May Find Later
; My New Improved Style of Coding Read a Bit More in Rem Line
; Took 4 Hours Work Additional on Code Before the Last Attempt Early 
; Hour Until Dawn
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; New Code Improver New Style Of Coder Improved Working
; Multi-Keys Set Multi Operations On One Key
; Little Tricky 
; To Make Code Less Complex Having To Use So Many Variables
; I Could Have Used A Timer If A Key Was Not Used And Want 
; To Place Back As Original When Not User
; But For Better Speed Do While Intercept Key As Quickly As Possible
; The Tricky Bit Was If A Key Was Intercept And Decided Not 
; Any Action Required To Change It
; The Key Has To Be Placed Back In The Buffer As Originally Was
; And That Plays A Knock-On Effect Later On In The Code 
; That Has To Be Taken Care Of
; The Idea Is I Want A Series Of Swapping Control And Shift-Return
; To Enter Key And Vice Versa Depending On If In The MYSMS Program That I Use 
; A Lot And Also Facebook 
; Facebook Is Only Detected Because Of My Name As The Title Of The Window 
; And Also The Class Name Not Much More Can Do On That One
; New Update Write Code Time
; Now Facebook Is Going To Have A Problem Because Of The Enter Key
; Messenger Post If Different To A Messenger Comment
; While I Got Messenger Comment Sorted
; I Don't Know The Detection For Messenger Post
; Oh Well
; Will Have To Go Back A Step And Have A Call-Up Routine 
; Which Is Hard To Use It Kick Into Place
; Unless A Hot-Key Is Done
; Quite Possible
; Yes Able To Get Around That With Work Of Shift-Enter Rather Than
;  Control Enter
; You Will Hear A Beep For The Enter Key In All Your Programs As Key Intercepted
; Not Just Isolated To GrinBook FB And MYSMS With Little Extra Code 
; You Could Disable Beeper Unless In FB And MYSMS
; Might Be Helpful To Remove The Beeper There As It Less Quicker 
; A Lot The Repeat Return Key Compared To Other Keys In Editing 
; Besides MYSMS And Grin Book
; Yes That Was The Answer But You Can Also Change The Sound 
; Speeded Up
; Well I Finished That Off With User-Friendly Variable Setting Choice 
; -------------------------------------------------------------------
; This Type Of Code Is Not Clearly Laid Out For Beginner Novice 
; AutoHotKeys __ Professional Expert Programmer
; -------------------------------------------------------------------
; WORK TIME
; Sat 21-Apr-2018 01:33:27 BEGINNER 
; Sat 21-Apr-2018 04:00:00 TOTAL 2 HOUR 27 MINUTE
; Sat 21-Apr-2018 05:38:00 TOTAL 4 HOUR 05 MINUTE + PUBLISH TIME QUICKER
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 003 SESSION
; -------------------------------------------------------------------
; WRITE CODE TO INCLUDE MORE NAME FOR FACEBOOK WORKAROUND ABOUT USE 
; CONTROL ENTER TO STOP SEND ON ENTER KEY PREMATURELY
;
; AND ADDITIONAL WORK OF MAKE IT PRIVACY WHEN SHARE CODE ON-LINE 
; IN SEPARATE TEXT FILE FOR DATA STORE-ER
;
; WORK READ FROM FILE INTO AN ARRAY FILTER OUT REM LINE FOR EXTRA INFO 
; WANTED AND NONE BLANK INFO 
; EXTRA INVISIBLE BLANK LINES BY ACCIDENT
;
; PRETTY EASY WHEN LOOKER FEW EXTRA LINE OF CODE - MAKE BETTER WORKING
; COULDN'T FIND HOW TO READ ONE ITEM OF NUMERIC FROM ARRAY ONLY TO READ ALL IN LOOP
;
; ADDITION THE NUMERIC IS THE FRIEND COUNT 
; WHEN YOU JUST PASTE ALL YOU FRIEND CONTACT INFO IN HERE
; TRIM OFF FEW TO MAKE TIDIER
; KEEP ADDITIONAL CONTACT INFO MIGHT WANTER
; ADDITIONALLY NUMERIC COUNT ON ONE LINE ONLY ARE EXCLUDED LIKE THE 
; SAME REM ; LINE ARE ALSO FEW FRIEND CONTACT WILL ONLY SHOW MUTUAL 
; CONTACT COUNTER AND THE FEW SHOW NOTHING
; SO KEEP THAT INFO AN FILTER EXCLUDE LINES WITH FILTER __ mutual friends
;
; SOME NUMERIC BY FACEBOOK HAVE THOUSAND SEPARATOR ,
; -------------------------------------------------------------------
; Add On __ 
; Count The Total Network of Friend Just For Additional Extra Programmer Fun
; Only an Extra Hour to Add of Code in Early Hour
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; TO     Tue 09-Oct-2018 01:25:00 __ DEFINITELY HARD WORK AS USUALLY MAKING AN ARRAY 
; FROM   Tue 09-Oct-2018 05:25:00 __ 4 HOUR
; -------------------------------------------------------------------

; GLOBAL SETTINGS ===================================================
#Warn
#NoEnv
#SingleInstance Force
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

; Create the popup menu by adding some items to it.
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Terminate Script, MenuHandler  ; Creates a new menu item.
Menu, Tray, Add, Terminate All AutoHotKey.exe, MenuHandler  ; Creates a new menu item.

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; SCRIPT ============================================================

DetectHiddenWindows, on
SetStoreCapslockMode, off
Set_Key_FB_Control=False 
Dont_Send_2_Enter=False
Sound_Speed=20
TEXT_STRING_VAR=	
	
; User Setting as Required
; If Set False You Get beep for Return Enter Key for Everything
; -------------------------------------------------------------------
Mute_Beep_In_Other_Program_Beside_GrinBook_an_Mysms=true
Mute_Beep_In_Other_Program_Beside_GrinBook_an_Mysms=false

current := Object()


; -------------------------------------------------------------------
; GOSUB ROUTINE_FINDER_ADD_TOTAL_COUNT_PEOPLE_IN_FACEBOOK_FRIEND_OF_FRIEND_NETOWRK_COUNT:
; OUTPUT FILE
; SourceFile := % A_ScriptDir "\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES NAME SCRIPT.txt
; -------------------------------------------------------------------

;--------------------------------------------------------------------
; CAREFUL ABOUT YOUR EDITOR OF FACEBOOK GRINBOOK FB FRIEND MEMBER SET 
; YOU MIGHT WANT TO RUN IT THROUGH DOUBLE CHECKER BEFORE GO
; WHILE EDITING
;--------------------------------------------------------------------
; For i in current
;	MSGBOX % current[i]



SoundBeep , 2000 , 200
SoundBeep , 2500 , 100

SetTimer TIMER_PREVIOUS_INSTANCE,1
; Timer_Previous_Instance Is A New Code I Set Up To Detect If A 
; The Previous Instance Of The Script Is A Runner
; Some Slow Computer Require That As A Limitation Of AHK
; Especially When Launch All You AHK Script At Once At 
; Speed When Boot Computer And Then Again _ Seems To Do The 
; Job Working Well

SetTimer TIMER_TEST,off

return

TIMER_TEST:    
    WinGet, HWND, ID, A  ; Get Active Window
    WinGet, OutputVar, ControlList, ahk_id %HWND%
    Tooltip, % OutputVar ; List All Controls of Active Window
return

enter::
{
    Set_Key=False
    Set_Key_FB_Control=False
    Dont_Send_2_Enter=False
    SetTitleMatchMode 2  ; Avoid specify full path
    
	; ---------------------------------------------------------------
	; Some of the Person I Write to Including Myself on Facebook FB Grin Book 
	; Which Stopper the Entering Of Message on Return Enter Key
	; And Let Continue Without Message Being Sent Premature
	; Previously You Have to Do By Hand Control and Enter to Enter the HTML Alternative Method Editor 
	; That Would Stop That Thing
	; This Does it Auto 
	; And It Has Extra When Press Enter the Word Replacement to Fan Page Or People 
	; Is Prevented By Adding and Extra Space at Enter Key
	;
	; I Call It Simpleton with One Little Message
	; Already got a Restriction if It Too Long
	; People are Hampered by there Little Mobile Keyboard
	; Pity I Have to Type The Person Name in Because of This One
	;
	; Maybe I Keep them on File and Software them In
	; Should Be Easy Enough Reading in File and Acting on Them
	; Kept Separately
	; A Typical One Liner Form of a Name Would Be Myself 
	; Matthew Lancaster -
	; With a Dash at End is Allowed
	; It is a Search in String So Doesn't Matter About the Proceed-er Numeric In Brackets 
	; For Notifications Missed But After Can Be Anything Else
	;
	; All Done New Improved Code Heps Privacy When Sharer 
	; of Scripts On-line
	;
	; CHUCK IN ALL FACEBOOK MEMBERS NOW
	; 
	; ---------------------------------------------------------------
	
	SET_GO_ENTER_FACEBOOK=FALSE
	; ---------------------------------------------------------------
	; REMM-ED OUT OLD STYLE OF CODE MAYBE USEFUL IN OTHER WORKER _ HARDCODE-ED
	; if WinActive("Matthew Lancaster -") 
		; SET_GO_ENTER_FACEBOOK=TRUE
	; if WinActive("Will Dee -") 
		; SET_GO_ENTER_FACEBOOK=TRUE
	; if WinActive("Steve Owen -") 
		; SET_GO_ENTER_FACEBOOK=TRUE
	; if WinActive("Jayla Holmes -") 
		; SET_GO_ENTER_FACEBOOK=TRUE
	; ---------------------------------------------------------------

	; ---------------------------------------------------------------
	For i in current
		IF WinActive(current[i])
			SET_GO_ENTER_FACEBOOK=TRUE
	; ---------------------------------------------------------------
		
	IF SET_GO_ENTER_FACEBOOK=TRUE
		IF !WinActive("ahk_class Chrome_WidgetWin_1")
			SET_GO_ENTER_FACEBOOK=FALSE
	
	IF SET_GO_ENTER_FACEBOOK = TRUE
	{
		SendInput {space} 
		; Add a Space when Enter Pressed in FB page So It It Doesn't 
		; Ask to Highlight any Fan Pages
		;------------------------------------------------------------
		Sleep 100 
		;------------------------------------------------------------
		; MINIMAL SLEEP I COULD FIND TO NOT CLASH THE KEYS OVERLAPPING
		; DEPEND ON YOUR PROCESSOR OR SOMETHING 100 MILLISECOND
		;------------------------------------------------------------
		; SetCapsLockState , Off
		; FOUND NEW SYSTEM WITH    SetStoreCapslockMode, off
		; THIS WILL PREVENT ERROR OF CAPS LOCK STATUS 
		; SHOWER FROM LOGITECH
		;------------------------------------------------------------
		; THE LOGITECH REMOTE KEYBOARD HAS A FLASHER BOX ON ABOUT CAPS LOCK STATE LEAVE IT WITH CAPS LOCK OFF 
		; AND THESE SendInput KEY DON'T GET IN THE WAY
		;------------------------------------------------------------
		SendInput {shift down}{enter}{shift up}
		SoundBeep , 2500 , 100
		Set_Key=True
		Set_Key_FB_Control=True
    }
    if Set_Key=False 
    {
		Dont_Send_2_Enter=True
        SendInput {enter}
        Set_Go=True
        if SET_GO_ENTER_FACEBOOK=TRUE
            Set_Go=False
        if Mute_Beep_In_Other_Program_Beside_GrinBook_an_Mysms=False
            Set_Go=True
        if Set_Go=True
            SoundBeep , 2500 , Sound_Speed
    }
}

; CTRL enter unable to be swapped back to Enter so use SendInput
;--------------------------------------------------------------
^enter::
{
    Set_Key=False
    SetTitleMatchMode 1  ; Specify the full path
    if Set_Key_FB_Control=False 
    {
        if (WinActive("mysms") and WinActive("ahk_class Chrome_WidgetWin_1"))
        {
            if Dont_Send_2_Enter=False
            {
                SendInput {ENTER}
                SoundBeep , 2500 , 100
                Set_Key=True
            }
        }
        if Set_Key=False 
        {
            if Dont_Send_2_Enter=False
            {
                SendInput {ctrl down}{enter}{ctrl up}
                Set_Go=True
                if (WinActive("mysms") and WinActive("ahk_class Chrome_WidgetWin_1"))
                    Set_Go=False
                if Mute_Beep_In_Other_Program_Beside_GrinBook_an_Mysms=False
                    Set_Go=True
                if Set_Go=True
                SoundBeep , 2500 , Sound_Speed
            }
        }
    }
    Set_Key_FB_Control=False 
    Dont_Send_2_Enter=False
}
Return

+enter::
{
    Set_Key=False
    if Set_Key_FB_Control=False 
    {
        if (WinActive("mysms") and WinActive("ahk_class Chrome_WidgetWin_1"))
        {
            {
                SendInput {ENTER}
                ; SetCapsLockState , Off
                SoundBeep , 2500 , 100
                SET_KEY=True
            }
        }
        if Set_Key=False 
        {
            if Dont_Send_2_Enter=False
                {
                SendInput {shift down}{enter}{shift up}
                Set_Go=True
                if (WinActive("mysms") and WinActive("ahk_class Chrome_WidgetWin_1"))
                    Set_Go=False
                if Mute_Beep_In_Other_Program_Beside_GrinBook_an_Mysms=False
                    Set_Go=True
                if Set_Go=True
                SoundBeep , 2500 , Sound_Speed
            }
        }
        Set_Key_FB_Control=False 
        Dont_Send_2_Enter=False
    }
}
Return



ROUTINE_FINDER_ADD_TOTAL_COUNT_PEOPLE_IN_FACEBOOK_FRIEND_OF_FRIEND_NETOWRK_COUNT:

; SOURCE CREDIT
; -------------------------------------------------------------------
; mixing pseudo arrays (array%n%) with real arrays (array[n], array.n)
; This is a real array example:
; -------------------------------------------------------------------
; Reading file and storing value in array problem - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/92137-reading-file-and-storing-value-in-array-problem/
; -------------------------------------------------------------------
; BLANK LINES        NOT READ IN
; LINES THAT BEGIN ; NOT READ IN
; -------------------------------------------------------------------
SourceFile := % A_ScriptDir "\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES NAME SCRIPT.txt"
IfExist, %SourceFile%
{
	Loop, read, % A_ScriptDir "\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES NAME SCRIPT.txt"
	{
		; ---------------------------------------------------------------
		; SOME NUMERIC BY FACEBOOK HAVE THOUSAND SEPARATOR
		; ---------------------------------------------------------------
		TEST_STRING := A_LoopReadLine
		TEST_STRING := StrReplace(TEST_STRING, ",", "")
		SET_GO=TRUE
		IF SubStr(A_LoopReadLine, 1, 1)=";"
			SET_GO=FALSE
		IF INSTR(A_LoopReadLine,"mutual friends")>0
			SET_GO=FALSE
		IF A_LoopReadLine is number
			SET_GO=FALSE
		IF TEST_STRING is number
			SET_GO=FALSE
		IF !A_LoopReadLine
			SET_GO=FALSE
		IF SET_GO=TRUE
		{
			TEXT_STRING_VAR = %A_LoopReadLine%%A_Space%-
			current[A_Index] := TEXT_STRING_VAR
			
			; MSGBOX % current[A_Index]
			; EXITAPP
			
		}
	}
}

; ------------------------------------------------------------------------
; Count The Total Network of Friend Just For Additional Extra Programmer Fun
; ------------------------------------------------------------------------
COUNTER_VALUE=0
SourceFile := % A_ScriptDir "\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES NAME SCRIPT.txt"
IfExist, %SourceFile%
{
	DestFile = % A_ScriptDir "\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES NUMERIC COUNT TOTAL.txt"
	IfExist, %DestFile%
	{
		FileDelete, %DestFile%
	}

	Loop, read, %SourceFile%, %DestFile%
	{
		TEST_STRING := A_LoopReadLine
		TEST_STRING := StrReplace(TEST_STRING, ",", "")
		TEST_STRING := StrReplace(TEST_STRING, "mutual friends", "")
		TEST_STRING := Trim(TEST_STRING)
		SET_GO=FALSE
		IF TEST_STRING is number
			SET_GO=TRUE

		If SET_GO=TRUE
			{
			FileAppend, %TEST_STRING%`n
			COUNTER_VALUE+=%TEST_STRING%
			}
	}
	TEXT_STRING_VAR:="="
	; StringReplace, TEXT_STRING_VAR, TEXT_STRING_VAR,",,All
	
	FileAppend, =`n, %DestFile%
	; FileAppend, %TEXT_STRING_VAR%`n, %DestFile%
	FileAppend, %COUNTER_VALUE%`n, %DestFile%
}

RETURN

MenuHandler:
	; MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
	if A_ThisMenuItem=Terminate Script
		Process, Close,% DllCall("GetCurrentProcessId")
	
	if A_ThisMenuItem=Terminate All AutoHotKey.exe
	{
		; Run, "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS" /F /IM AutoHotKey.exe /T , , Max
		DetectHiddenWindows, On 
		WinGet, List, List, ahk_class AutoHotkey 
		Loop %List% 
		  { 
			WinGet, PID, PID, % "ahk_id " List%A_Index% 
			If ( PID <> DllCall("GetCurrentProcessId") ) 
				 ; PostMessage,0x111,65405,0,, % "ahk_id " List%A_Index% 
				 Process, Close, List%A_Index% 
		  }
		Process, Close,% DllCall("GetCurrentProcessId")
		
		;  ----------------------------------------------------------
		; PROBLEM HERE IF PROGRAM THAT CALL THE BATCH FILE IS KILL SO IS THEN BATCH FILE
		; AND WE GET OVER THAT BY GO EXTRA VIA VBSCRIPT ANOTHER FILE
		; COULD OF RUN A  LOOP AND KILL BUT TRY NOT LOSE OWN ONE FIRST
		; [ Saturday 14:55:00 Pm_02 March 2019 ]
		;  ----------------------------------------------------------

		;  ----------------------------------------------------------
		; OTHER OPTION SET PROCESS KILLER
		;  ----------------------------------------------------------
		; Run, BAT_03_PROCESS_KILLER.BAT /F /IM AutoHotKey.exe /T , , Max
		; Run, %ComSpec% /k ""BAT_03_PROCESS_KILLER.BAT" "/F" "/IM" "AutoHotKey.exe" "/T"" , , Max
		; Process, Close, AutoHotKey.exe
		;  ----------------------------------------------------------
	
		; AUTO GENERATED FILE BY HERE VISUAL BASIC ORIGINAL LONG BEFORE AUTOHOTKEY WANT
		; D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB KEEP RUNNER.exe
		; D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe
		; -------------------------------------------------------------------
		; AND USED BY HERE
		; LOT OF AUTOHOTKEYS TRAY MENU ITEM
		; -------------------------------------------------------------------
		; [ Saturday 14:52:10 Pm_02 March 2019 ]
		; -------------------------------------------------------------------
		; EDITOR COPY PASTE FROM VBS 39-KILL PROCESS.VBS
		; THIS FILE BECAME USE BY
		; LOT OF AUTOHOTKEYS TRAY MENU ITEM
		; AND THEY USE IT HERE THIS ONE
		; C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT
		; ORIGINAL AT HERE LOCATION 
		; C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
		; AND MOVED HERE MAYBE 
		; -------------------------------------------------------------------
		; MOST LIKELY TRY AND KEEP IN SYNC LATER
		; EXCEPT THE AUTO GENERATOR
		; -------------------------------------------------------------------
}
return


TIMER_PREVIOUS_INSTANCE:
SETTIMER TIMER_PREVIOUS_INSTANCE,10000
if ScriptInstanceExist()
{
    Exitapp
}
return

ScriptInstanceExist() {
    static title := " - AutoHotkey v" A_AhkVersion
    dhw := A_DetectHiddenWindows
    DetectHiddenWindows, On
    WinGet, match, List, % A_ScriptFullPath . title
    DetectHiddenWindows, % dhw
    return (match > 1)
    }
Return


;# ------------------------------------------------------------------
EOF:                           ; on exit
ExitApp     
;# ------------------------------------------------------------------

;# ------------------------------------------------------------------
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
;# ------------------------------------------------------------------
; exit the app

