;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 21-AutoRun.ahk
;# __ 
;# __ Autokey -- 21-AutoRun.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Wed 18:22:50 Pm_20 Dec 2017 ]
;# END         TIME [ Wed 19:42:10 Pm_20 Dec 2017 ]
;# LAST EDITOR TIME [ Tue 14:55:00 Pm_01 May 2018 ]
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; 001 ---------------------------------------------------------------
; AUTOEXEC BOOT SCRIPT
; WORK SESSION AT END AN MID WAY
; TAKE THE REGISTRY BOOTS AND PUT HERE
; AND DO A DELETE REGISTRY KEY SO HERE TAKES CONTROL
; PRETTY EASY
; THERE IS A LOT MORE REGISTRY BOOTER TO DO IF WANT
; USING AUTORUNS
; -------------------------------------------------------------------
; FROM TIME __ Fri 27-Apr-2018 22:54:28 
; TO   TIME __ Sat 28-Apr-2018 00:02:00 1 HOUR+
; -------------------------------------------------------------------

; 002 ---------------------------------------------------------------
; LETHAL SKILL BIT OF CODE TO DETECT IF A PROCESS EXIST AND IF IT BY 
; CURRENT USERNAME
; USEFUL FOR MULTI USER COMPUTER SYSTEM AND WANT YOU FAVORITE PROGRAMS 
; ON THEM ALL
; LEARN FUNCTION TO SUBROUTINE AND GLOBAL VARIABLE DECLARE
; LONG TIME LOT OF PAGES OPEN ON INTERNET
; WAS DIFFICULT BECAUSE SOMEBODY GAVE CODE INA GUI ROUTINE WHERE YOU 
; PRESS BUTTON TO MAKE GO
; AND TWO LINES ON ONE COMMAND SEPARATED BY COMMA
; WAS 3 NICE CODE THIS ONE THE BEST AND A FEW LESS QUICK CODE
; WHICH I MIGHT OF BEEN USING ONE IF PASS THE VARIABLE TO ONE WOULD OF 
; WORKED
; 51 URL'S OPEN WORKER ON THIS ONE AND OTHERS BEHIND BACK
; I COULD OF DONE THE FUNCTION ROUTINE BETTER WHEN LEARN GLOBAL 
; VARIABLE SETTING _ BIT TIRED AFTER THAT ONE
; -------------------------------------------------------------------
; FROM TIME __ Sun 29-Apr-2018 06:10:13
; TO   TIME __ Sun 29-Apr-2018 14:50:00 8 HOUR 40 & 51 URL'S OPEN
; -------------------------------------------------------------------

; 003 ---------------------------------------------------------------
; LOST SOME TEXT SO COMBINE IN ONE
; THE NEW CODE TO DETECT PROCESS NAME AND BY USERNAME
; FAILED ON 32 BIT COMPUTER OR WIN XP AND ALSO WIN STARTER WIN 7
; I MAGANED TO SET VARIABLE TO SOMETHING TO ALLOW CONTINUE
; AS THE PROCESS LIST REQUEST CAME BACK WITH NOTHING
; NOW I HAVE TO FINISHA THE FUNCTION TO AST AS NORMAL WITHOUT EXTRA CHECKING
; IF THE PROCESS REQUEST WITH NAME RETURNS NOTHING
; PROBLEM FOR WINDOWS 7 STARTER AS THE REQUEST CRASHES OUT THE SCRIPT
; ALSO WITH THE ERROR IN PLACE
; ALL PROGRAM GET REQUEST TO RUN AGAIN
; AND SOME HIDEN TASKBAR ICON PROGRAMS ALLOW TO BE SHOWN ON RUN AGAIN
; SO HAVE TO MAKE SURE MINIMIZED AFTER RUN 
; WHICH LEADS ONTO NEXT HERE
;--------------------------------------------------------------------
; SOME MORE TO DO
; DFX HAD SAME PROBLEM ABOVE CALLS ITSELF SHOWING ON LAUNCH EXE
; BUT AHK FOUND THE WRONG HWND HANDLE I WANTED SO PROBLEM REQUEST 
; TO MAKE HIDDEN WINDOW
; HELPED OUT BY MY VISUAL BASIC CODE
; I FOUND I HAD TO REQUEST THE PARENT HWND WINDOW
; AND THEN WORKED
; ALSO EXTRA TOPPING ON TOP I FOUND HOW TO DETECT IF A WINDOW IS HIDDEN
; SO THEN QUICKER WITHOUT DELAY IN A LOOP UNTIL REQUEST TO MAKE HIDDEN
; -------------------------------------------------------------------
; FROM TIME __ Mon 30-Apr-2018 19:37:18
; AND  TIME __ Mon 30-Apr-2018 21:00:00 1 HOUR 20 MINUTE 
; TO   TIME __ Mon 30-Apr-2018 23:00:00 2 HOUR 20 MINUTE _22_ URL'S OPENED

; FROM TIME __ Tue 01-May-2018 07:18:15 WORK FOR XP WORKING FINE TUNING
; TO   TIME __ Tue 01-May-2018 11:21:54 4 HOUR 4 MINUTE 
; -------------------------------------------------------------------

; 004 ---------------------------------------------------------------
; ADD REGISTRY SEARCH AND DELETE FOR CHROME AUTO STARTING
; -------------------------------------------------------------------
; FROM TIME __ Tue 01-May-2018 11:53:18
; TO   TIME __ Tue 01-May-2018 12:38:00 40 MINUTE _1_ URL OPENED
; -------------------------------------------------------------------

; 005 ---------------------------------------------------------------
; MADE THE CODE WORK FOR XP BACK AGAIN THROUGH THE PROCEDURE FUNCTION
; PROBLEM WITH GET PROCESS RUNNER
; -------------------------------------------------------------------
; FROM TIME __ Tue 01-May-2018 14:38:50
; TO   TIME __ Tue 01-May-2018 14:55:00 17 MINUTE _ NONE URL OPENED
; -------------------------------------------------------------------

; 00* 
; BOOT UP'S RUNNING SMOOTHER
; -------------------------------------------------------------------
; Sun 20-May-2018 22:35:20
; Sun 21-May-2018 02:14:20 4 HOUR
; -------------------------------------------------------------------

; 007 ---------------------------------------------------------------
; MAJOR CHANGE USE COMMAND LINE PARAMETER SAVE DUPLICATE CODE
; -------------------------------------------------------------------
; FROM TIME __ Wed 30-May-2018 07:47:13
; TO   TIME __ Wed 30-May-2018 08:10:00
; -------------------------------------------------------------------

; 008 ---------------------------------------------------------------
; WORK BLUETOOTHLOGGER NIRSOFT WORK TO FIX CAN EASY CRASH 
; WINDOW IF NOT FULLY STARTED YET AND IF TRY AND MINIMIZE TWICE
; -------------------------------------------------------------------
; FROM TIME __ Tue 19-Jun-2018 22:30:41
; TO   TIME __ Wed 20-Jun-2018 00:35:00 __ 2 HOUR
; -------------------------------------------------------------------

; 009 ---------------------------------------------------------------
; PROCESSEXIST FUNCTION WAS WRONG REQUIRED %% AROUND VARIABLE
; WHEN CHECKER FOR XP COMPUTER AND WIN 7 STARTER
; -------------------------------------------------------------------
; FROM TIME __ Wed 20-Jun-2018 14:27:06
; TO   TIME __ Wed 20-Jun-2018 14:38:00 __ 
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
; Location OnLine
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set Google Drive
; Possible Censorship of Code Detected By Google as Malicious Happen Here
; unlike DropBox that has All Available
; https://drive.google.com/open?id=0BwoB_cPOibCPTnRZZVFuRFpHOTg
;--------------------------------------------------------------------
; Link to Folder of all My Scriptor Project Set DropBox
; https://www.dropbox.com/sh/ntghoncyb8py1tf/AACWYrfkVn9PlqpYzNNSMcpMa?dl=0
;--------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
; https://www.dropbox.com/s/w22p90h81uwni71/Autokey%20--%2021-AUTORUN.ahk?dl=0
;# ------------------------------------------------------------------


; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force

DetectHiddenWindows, on
SetStoreCapslockMode, off

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

CODE_RUN_FOR_BRUTE_BOOT_DOWN_AHK=FALSE

global CLOSE_SOME_LEFT_OVER_WINDOWS_VAR
global FLAG_GET_PROCESS_MATCH=0
global ProcessSearch:=""
global HWND=0
global SET_GO
global OSVER_N_VAR

CLOSE_SOME_LEFT_OVER_WINDOWS_VAR=FALSE

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
OSVER_N_VAR:=a_osversion
IF INSTR(a_osversion,".")>0
	OSVER_N_VAR:=substr(a_osversion, 1, INSTR(a_osversion,".")-1)
IF OSVER_N_VAR=WIN_XP
	OSVER_N_VAR=5
IF OSVER_N_VAR=WIN_7
	OSVER_N_VAR=6

GOSUB TEST_STARTER_RUN_IN


Param_in=
Loop, %0% {               ;for each command line parameter
    If (%A_Index%)        ;check if known command line parameter exists
        Param_in=TRUE
}

If Param_in
{
	GOSUB MAIN_ROUTINE_2
	EXITAPP
}
ELSE
{
	GOSUB MAIN_ROUTINE_2
	GOSUB MAIN_ROUTINE
}

; -------------------------------------------------------------------
;GuiEscape:
;GuiClose:

SOUNDBEEP 1000,200
SOUNDBEEP 1500,200
SOUNDBEEP 1000,200
; -------------------------------------------------------------------

Exitapp

RETURN


;RegWrite, REG_SZ, HKEY_LOCAL_MACHINE\SOFTWARE\TestKey, MyValueName, Test Value
;RegWrite, REG_BINARY, HKEY_CURRENT_USER\Software\TEST_APP, TEST_NAME, 01A9FF77
;RegWrite, REG_MULTI_SZ, HKEY_CURRENT_USER\Software\TEST_APP, TEST_NAME, Line1`nLine2


GUI_GO:
; GUI ===============================================================

global WTS_EX := "SID|PID|Process Name|User SID|Name|Domain|Number of Threads|Handle Count|Pagefile Usage|Peak Pagefile Usage|Working Set Size|Peak Working Set Size|User Time|Kernel Time"

Gui, Margin, 5, 5
Gui, Add, Edit, xm ym w200 hWndhSearch vprcsearch
DllCall("user32.dll\SendMessage", "Ptr", hSearch, "UInt", 0x1501, "Ptr", 1, "Str", "Prozess Name here", "Ptr")
Gui, Add, Button, x+5 yp-1 w95 gSTART, % "Start"
Gui, Add, ListView, xm y+5 w1200 h600, % WTS_EX
LV_ModifyCol(1, 50), LV_ModifyCol(2, 50), LV_ModifyCol(3, 140), LV_ModifyCol(4, 60)
Gui, Add, Edit, xm y+5 w60 0x800 vprccount
Gui, Show, AutoSize
return


; SCRIPT ============================================================

START:
{ 
	Gui, Submit, NoHide
    LV_Delete()
	WTSEnumerateProcessesEx()
	prcsearch:="Notepad++.exe"
	prcsearch=%ProcessSearch%
    loop % WTS.MaxIndex()
    {
        if (InStr(WTS[A_Index, "PN"], prcsearch))
            {
			USER_2:=WTS[A_Index, "Name"]
			IF A_USERNAME=%USER_2%
			;MSGBOX, % A_USERNAME
			
			LV_Add("", WTS[A_Index, "SID"]
                     , WTS[A_Index, "PID"]
                     , WTS[A_Index, "PN"]
                     , WTS[A_Index, "USID"]
                     , WTS[A_Index, "Name"]
                     , WTS[A_Index, "Domain"]
                     , WTS[A_Index, "NOT"]
                     , WTS[A_Index, "HC"]
                     , Num_Comma(Round(WTS[A_Index, "PU"] / 1024)) " K"
                     , Num_Comma(Round(WTS[A_Index, "PPU"] / 1024)) " K"
                     , Num_Comma(Round(WTS[A_Index, "WSS"] / 1024)) " K"
                     , Num_Comma(Round(WTS[A_Index, "PWSS"] / 1024)) " K"
                     , GetDurationFormat(WTS[A_Index, "UT"])
                     , GetDurationFormat(WTS[A_Index, "KT"]))
			}		 
    }
    GuiControl,, prccount, % LV_GetCount()
}
return


START_GET_PROCESS:
{ 
	WTS := []
	WTSEnumerateProcessesEx()
	IF WTS.MaxIndex()>0
	{	
	loop % WTS.MaxIndex()
    {
        if (InStr(WTS[A_Index, "PN"], ProcessSearch))
            {
			USER_3:=WTS[A_Index, "Name"]
			;MSGBOX, % USER_3 
			IF A_USERNAME=%USER_3%
				{
				FLAG_GET_PROCESS_MATCH=1
				RETURN
				}
			}		 
    }
	}
FLAG_GET_PROCESS_MATCH=0
}
return


; FUNCTIONS =========================================================

WTSEnumerateProcessesEx()
{
    local PI := pPI := nTTL := 0, LIST := use := ""
    hModule := DllCall("kernel32.dll\LoadLibrary", "Str", "wtsapi32.dll", "Ptr")
    if !(DllCall("wtsapi32.dll\WTSEnumerateProcessesEx", "Ptr", 0, "UInt*", 1, "UInt", 0xFFFFFFFE, "Ptr*", pPI, "UInt*", nTTL))
        return "", DllCall("kernel32.dll\SetLastError", "UInt", -1)

    PI := pPI, WTS := []
    loop % (nTTL)
    {
        WTS[A_Index, "SID"]     := NumGet(PI + 0, "UInt")      ; SessionId             ( The Remote Desktop Services session identifier for the session associated with the process. )
        WTS[A_Index, "PID"]     := NumGet(PI + 4, "UInt")      ; ProcessId             ( The process identifier that uniquely identifies the process on the RD Session Host server. )
        WTS[A_Index, "PN"]      := StrGet(NumGet(PI + 8))      ; pProcessName          ( A pointer to a null-terminated string that contains the name of the executable file associated with the process. )
        WTS[A_Index, "USID"]    := NumGet(PI + 16)             ; pUserSid              ( A pointer to the user security identifiers (SIDs) in the primary access token of the process. )
        VarSetCapacity(Name, 512, 0), VarSetCapacity(Domain, 512, 0)
        DllCall("advapi32.dll\LookupAccountSid", "Ptr", 0, "Ptr", NumGet(PI + 16), "Str", Name, "UIntP", 256, "Str", Domain, "UIntP", 256, "UIntP", use)
        WTS[A_Index, "Name"]    := Name
        WTS[A_Index, "Domain"]  := Domain
        WTS[A_Index, "NOT"]     := NumGet(PI + 24, "UInt")     ; NumberOfThreads       ( The number of threads in the process. )
        WTS[A_Index, "HC"]      := NumGet(PI + 28, "UInt")     ; HandleCount           ( The number of handles in the process. )
        WTS[A_Index, "PU"]      := NumGet(PI + 32, "UInt")     ; PagefileUsage         ( The page file usage of the process, in bytes. )
        WTS[A_Index, "PPU"]     := NumGet(PI + 36, "UInt")     ; PeakPagefileUsage     ( The peak page file usage of the process, in bytes. )
        WTS[A_Index, "WSS"]     := NumGet(PI + 40, "UInt")     ; WorkingSetSize        ( The working set size of the process, in bytes. )
        WTS[A_Index, "PWSS"]    := NumGet(PI + 44, "UInt")     ; PeakWorkingSetSize    ( The peak working set size of the process, in bytes. )
        WTS[A_Index, "UT"]      := NumGet(PI + 48, "Int64")    ; UserTime              ( The amount of time, in milliseconds, the process has been running in user mode. )
        WTS[A_Index, "KT"]      := NumGet(PI + 56, "Int64")    ; KernelTime            ( The amount of time, in milliseconds, the process has been running in kernel mode. )
        PI += 64
    }

    DllCall("wtsapi32.dll\WTSFreeMemoryEx", "Ptr", 1, "Ptr", pPI, "UInt", nTTL)

    if (hModule)
        DllCall("kernel32.dll\FreeLibrary", "Ptr", hModule)

    return WTS, DllCall("kernel32.dll\SetLastError", "UInt", nTTL)
}




GetDurationFormat(VarIn, Format := "hh:mm:ss.fff")
{
    VarSetCapacity(VarOut, 128, 0), VarIn := VarIn
    if !(DllCall("kernel32.dll\GetDurationFormat", "UInt", 0x400, "UInt", 0, "Ptr", 0, "Int64", VarIn, "WStr", Format, "WStr", VarOut, "Int", 2048))
        return DllCall("kernel32.dll\GetLastError")
    return VarOut
}

Num_Comma(num)
{
    n := DllCall("kernel32.dll\GetNumberFormat", "UInt", 0x0400, "UInt", 0, "Str", num, "Ptr", 0, "Ptr", 0, "Int", 0)
    VarSetCapacity(s, n * (A_IsUnicode ? 2 : 1), 0)
    DllCall("kernel32.dll\GetNumberFormat", "UInt", 0x0400, "UInt", 0, "Str", num, "Ptr", 0, "Str", s, "Int", n)
    return SubStr(s, 1, StrLen(s) - 3)
}

; ----------------------------

ProcessExist(processName,Test_A_UserName) 
{
	SET_GO=FALSE
	IF (A_ComputerName = "1-ASUS-X5DIJ") 
		SET_GO=TRUE
	IF (A_ComputerName = "2-ASUS-EEE") 
		SET_GO=TRUE
	IF (A_ComputerName = "3-LINDA-PC") 
		SET_GO=TRUE
	
	;If OSVER_N_VAR<10
	;	SET_GO=TRUE
	
	IF SET_GO=TRUE
	{
		Process, Exist, %processName%
		RETURN %ErrorLevel%
		; If Not ErrorLevel ; errorlevel will = 0 if process doesn't exist
	}

	ProcessSearch:="Notepad++.exe"
	ProcessSearch=%processName%
	GOSUB START_GET_PROCESS
	IF FLAG_GET_PROCESS_MATCH=1 
		RETURN 1
	RETURN 0
}

MINIMIZE_ALL_CMD_AT_BOOT:

	; MINIMIZE ALL THE CMD AT BOOT
	WinGet, id, list,ahk_class ConsoleWindowClass
	Loop, %id%
	{
		;WinGetTitle, title, ahk_id %table%
		;some_variable%A_Index%=%title%

		table := id%A_Index%
		WinMinimize  ahk_id %table%
	} 
RETURN	

MINIMIZE_ALL__EXPLORER_AT_BOOT:

WinGet, id, list,ahk_class CabinetWClass
Loop, %id%
{
	;WinGetTitle, title, ahk_id %table%
	;some_variable%A_Index%=%title%

	table := id%A_Index%
	WinGetTitle, title, ahk_id %table%
	WinMinimize  ahk_id %table%
	IF (A_ComputerName="2-ASUS-EEE")
	IF title=D:\#
		WinClose,  ahk_id %table% ; Got an Annoyer Explorer Window Every Boot
	IF title=D:\0 CLOUD\ASUS WEBSTORAGE _5GB
		WinClose,  ahk_id %table% ; Got an Annoyer Explorer Window Every Run Twice Booter
	
} 
RETURN

MINIMIZE_ALL_CHROME_AT_BOOT:
; MINIMIZE ALL THE CHROME AT BOOT
WinGet, id, list,ahk_class Chrome_WidgetWin_1
Loop, %id%
{

	table := id%A_Index%
	WinMinimize  ahk_id %table%
} 
RETURN





; -------------------------------------------------------------------
MAIN_ROUTINE_2:

GOSUB MINIMIZE_ALL__EXPLORER_AT_BOOT
GOSUB MINIMIZE_ALL_CMD_AT_BOOT
GOSUB MINIMIZE_ALL_CHROME_AT_BOOT

;--------------------------------------------------------------------
; SLEEP FOR 60 SECOND IF LESS QUICK COMPUTER & BOOT TICKER SHOWS LOWER
;--------------------------------------------------------------------

SET_GO_1=TRUE
SET_GO_2=20
IF (A_ComputerName = "1-ASUS-X5DIJ") 
	SET_GO_2=60
IF (A_ComputerName = "2-ASUS-EEE") 
	SET_GO_2=80
IF (A_ComputerName = "3-LINDA-PC") 
	SET_GO_2=60
IF (A_ComputerName = "5-ASUS-P2520LA") 
	SET_GO_2=60

COUNT_TICK_TIME=% 1000*60*8
; IF OUR SET TIME IS LESS THEN TICK TIME LET IT GO
; ONLY IF BOOT IN EARLY
;-------------------------------------------------
IF COUNT_TICK_TIME<%A_TICKCOUNT%
	SET_GO_1=FALSE
	
IF SET_GO_1=TRUE
{
	I_COUNT=0
	LOOP 
	{
		I_COUNT+=1
		SLEEP 1000
		IF I_COUNT>%SET_GO_2%
			BREAK
	}
}

SKIP_CODE=FALSE
IF (A_ComputerName = "3-LINDA-PC") 
	SKIP_CODE=TRUE

IF SKIP_CODE=FALSE
{
If ProcessExist("ClipBoard Logger.exe", A_UserName)=0
	{
		FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe"
		IfExist, %FN_VAR%
		{
				SoundBeep , 2000 , 100
				Run, "%FN_VAR%"
		}
	}
}

SKIP_CODE=FALSE
IF (A_ComputerName = "1-ASUS-X5DIJ") 
	SKIP_CODE=TRUE
IF (A_ComputerName = "2-ASUS-EEE") 
	SKIP_CODE=TRUE
IF (A_ComputerName = "3-LINDA-PC") 
	SKIP_CODE=TRUE

IF SKIP_CODE=FALSE
{

	If ProcessExist("EliteSpy.exe", A_UserName)=0
		{
			FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2000 , 100
				Run, "%FN_VAR%"
			}
		}

	SET_GO=TRUE
	IF (A_ComputerName = "2-ASUS-EEE") 
		SET_GO=FALSE
	IF (A_ComputerName = "1-ASUS-X5DIJ") 
		SET_GO=FALSE
	IF (A_ComputerName = "3-LINDA-PC") 
		SET_GO=FALSE
	IF (A_ComputerName = "5-ASUS-P2520LA") 
		SET_GO=FALSE

	IF SET_GO=TRUE
	{
		If ProcessExist("Tidal.exe", A_UserName)=0
			{
				FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe"
				IfExist, %FN_VAR%
				{
					SoundBeep , 2000 , 100
					Run, "%FN_VAR%"
				}
			}
	}

}

SET_GO=TRUE
IF (A_ComputerName = "1-ASUS-X5DIJ") 
	SET_GO=FALSE
IF (A_ComputerName = "2-ASUS-EEE") 
	SET_GO=FALSE
IF (A_ComputerName = "3-LINDA-PC") 
	SET_GO=FALSE
If (OSVER_N_VAR<10)
	SET_GO=FALSE
	
IF SET_GO=TRUE
	{
	If ProcessExist("CPU % INDIVIDUAL PROCESS.exe", A_UserName)=0
		{
			FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\CPU % OF A PROGRAM\CPU % INDIVIDUAL PROCESS.exe"
			IfExist, %FN_VAR%
			{
				;SoundBeep , 2500 , 100
				;Run, "%FN_VAR%" , , MIN
			}
		}
	}
		
If ProcessExist("VB KEEP RUNNER.exe", A_UserName)=0
{
	FN_VAR:="D:\VB6\VB-NT\00_Best_VB_01\VB KEEP RUNNER\VB KEEP RUNNER.exe"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%"
		; SLEEP 8000
		; -----------------------------------------------------------------
		; GIVE TIME TO RUN FOR XP THE TWO BLUEOOTH APP ON TASK BAR TOGETHER
		; DOING A SEARCH ON XP TERMINATES THE APP
		; -----------------------------------------------------------------
	}
}


;--------------------------------------------------------------------
;AUTOHOTKEYS
;--------------------------------------------------------------------
	
FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"	
IfExist, %FN_VAR%
{
	SoundBeep , 2500 , 100
	Run, "%FN_VAR%"
}

FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 04-SET CAP NUM & SCROLL TO LIKING - ONCE.ahk"	
IfExist, %FN_VAR%
	IF !WinExist(FN_VAR) 
	{
		SoundBeep , 2000 , 100
		Run, "%FN_VAR%"
	}


RETURN
; -------------------------------------------------------------------


; -------------------------------------------------------------------
MAIN_ROUTINE:

SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.


; --------------------------------------------
; Main Boot up File Source Written to Registry
; IT SETS THE UAC TO ADMIN IF POSSIBLE
; --------------------------------------------


RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, STARTUP_COMMON_04_ALL, "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 00 ELEVATED PRIV ADMIN _ START_UP.BAT"


SET_GO_1=1


; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
If (OSVER_N_VAR>5 
	and SET_GO_1=1)
	{
	Process, Exist, AmazonDrive.exe
	If Not ErrorLevel ; errorlevel will = 0 if process doesn't exist
		{
		FN_VAR=C:\Users\%A_UserName%\AppData\Local\Amazon Drive\AmazonDrive.exe
		;IfExist, %FN_VAR%
			;Run, "%FN_VAR%" , , HIDE
		}
	}
	
	; PROBLEM HERE SOLVED IF REGNAME HAS SPACES INIT DO A LOOP AND FIND
	; -----------------------------------------------------------------
	Loop, Reg, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\, VR
	{
		if A_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		
		IfInString, A_LoopRegName, Amazon Drive
		{
			SoundBeep , 2500 , 100
			RegDelete, %A_LoopRegKey%\%A_LoopRegSubKey%, %A_LoopRegName%
			;MsgBox, 4, , %A_LoopRegKey%\%A_LoopRegSubKey%\%A_LoopRegName%
			;MsgBox, 4, , %A_LoopRegKey%\%A_LoopRegSubKey%
			;MsgBox, 4, , %A_LoopRegName% = %value% (%A_LoopRegType%)`n`nContinue?
			;IfMsgBox, NO, break
		}
	}

SET_GO_1=0
IF (A_ComputerName="5-ASUS-P2520LA" and A_UserName="MATT 01")
	SET_GO_1=0
IF (A_ComputerName="4-ASUS-GL522VW" and A_UserName="MATT 01")
	SET_GO_1=0
IF (A_ComputerName="7-ASUS-GL522VW" and A_UserName="MATT 04")
	SET_GO_1=1
IF (A_ComputerName="8-MSI-GP62M-7RD" and A_UserName="MATT 01")
	SET_GO_1=0

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
If (OSVER_N_VAR>5 
	and SET_GO_1=1)
	{
	Process, Exist, BTCloud.exe
	If Not ErrorLevel ; errorlevel will = 0 if process doesn't exist
		{
			FN_VAR:="C:\Program Files\BT Cloud\BT Cloud\BTCloud.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				Run, "%FN_VAR%"
			}
		}
	}
RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, SynchronossPC


; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
SET_GO_1=0
If (OSVER_N_VAR>5 
	and SET_GO_1=1)
	{
	If ProcessExist("VideoDownloaderUltimate.exe", A_UserName)=0
		{
			FN_VAR:="C:\ProgramData\VideoDownloaderUltimateWinApp\VideoDownloaderUltimate.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				Run, "%FN_VAR%" /repair
			}
		}
	}
RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, VideoDownloaderUltimate

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
If (OSVER_N_VAR>5)
{
Process, Exist, lastapp_x64.exe
If Not ErrorLevel
	{
		FN_VAR:="C:\Program Files (x86)\LastPass\lastapp_x64.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%"
		}	
	}	
}


SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

	
; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
If (OSVER_N_VAR>5)
{
If ProcessExist("LogiOptions.exe", A_UserName)=0
	{
		FN_VAR:="C:\Program Files\Logitech\LogiOptions\LogiOptions.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" /noui
			SLEEP 1000
			CLOSE_SOME_LEFT_OVER_WINDOWS_VAR=TRUE
			; GOSUB CLOSE_SOME_LEFT_OVER_WINDOWS
		}
	}
}

RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, LogiOptions

If ProcessExist("ProcessLasso.exe", A_UserName)=0
	{
	FN_VAR:="C:\Program Files\Process Lasso\ProcessLasso.exe"
	IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" /tray , , HIDE
			
			;WinWait, ahk_class Class_PLMain, , 80
			;WinGet, HWND, ID, ahk_class Class_PLMain
			;IS_WINDOW_HIDDEN_AND_HIDE(HWND, 50)
		}
	}

If ProcessExist("picpick.exe", A_UserName)=0
	{
		FN_VAR:="C:\PStart\Progs\#_PortableApps\PortableApps\PicPickPortable\App\picpick\picpick.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" /startup , , HIDE
		}
	}

If ProcessExist("RoboTaskBarIcon.exe", A_UserName)=0
{
	FN_VAR:="C:\Program Files (x86)\Siber Systems\AI RoboForm\RoboTaskBarIcon.exe"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%"
	}

	FN_VAR:="C:\Program Files\Siber Systems\AI RoboForm\RoboTaskBarIcon.exe"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%"
	}

	DetectHiddenWindows, OFF
}	

RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, RoboForm

;ExitApp


DetectHiddenWindows, On

If ProcessExist("wweb32.exe", A_UserName)=0
	{
		FN_VAR:="C:\Program Files\WordWeb\wweb32.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" -startup , , HIDE
			WinGet, HWND, ID, WordWeb ahk_class TTheDi
			IS_WINDOW_HIDDEN_AND_HIDE(HWND, 50)
			WinGet, HWND, ID, ahk_class Wordweb Tray Icon
			IS_WINDOW_HIDDEN_AND_HIDE(HWND, 50)
		}
	}

RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, WordWeb
	
SKIP_CODE=FALSE
IF (A_ComputerName = "1-ASUS-X5DIJ") 
	SKIP_CODE=TRUE
IF (A_ComputerName = "2-ASUS-EEE") 
	SKIP_CODE=TRUE
IF (A_ComputerName = "3-LINDA-PC") 
	SKIP_CODE=TRUE

IF SKIP_CODE=FALSE
{
	PID=1
	Process, Exist, pushbullet_client.exe
	If Not ErrorLevel PID=0
	Process, Exist, pushbullet.exe
	If Not ErrorLevel PID=0
	If PID=0
		{
			FN_VAR:="C:\Program Files (x86)\Pushbullet\pushbullet.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				Run, "%FN_VAR%" -show false
			}
		}

	RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, Pushbullet

		

	SET_GO_1=0
	IF (A_ComputerName="1-ASUS-X5DIJ")
		SET_GO_1=1
	IF (A_ComputerName="2-ASUS-EEE")
		SET_GO_1=1
	IF (A_ComputerName="3-LINDA-PC")
		SET_GO_1=1
	IF (A_ComputerName="5-ASUS-P2520LA" and A_UserName="MATT 01")
		SET_GO_1=0
	IF (A_ComputerName="4-ASUS-GL522VW" and A_UserName="MATT 01")
		SET_GO_1=1
	IF (A_ComputerName="7-ASUS-GL522VW" and A_UserName="MATT 04")
		SET_GO_1=1
	IF (A_ComputerName="8-MSI-GP62M-7RD" and A_UserName="MATT 01")
		SET_GO_1=1

	; WIN_XP 5 WIN_7 6 WIN_10 10  
	; --------------------------
	If (OSVER_N_VAR>5 
		and SET_GO_1=1)
		{
		Process, Exist, NokiaSuite.exe
		If Not ErrorLevel
			{
				FN_VAR:="C:\Program Files (x86)\Nokia\Nokia Suite\NokiaSuite.exe"
				IfExist, %FN_VAR%
				{
					SoundBeep , 2500 , 100
					Run, "%FN_VAR%" -tray
				}
			}
		}
	RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, NokiaSuite.exe

	; WIN_XP 5 WIN_7 6 WIN_10 10  
	; --------------------------
	If (OSVER_N_VAR>5 
		and SET_GO_1=1)
		{
		Process, Exist, BoxSync.exe
		If Not ErrorLevel
			{
				FN_VAR:="C:\Program Files\Box\Box Sync\BoxSync.exe"
				IfExist, %FN_VAR%
				{
					SoundBeep , 2500 , 100
					Run, "%FN_VAR%" -m
					SLEEP 1000
					CLOSE_SOME_LEFT_OVER_WINDOWS_VAR=TRUE
					; GOSUB CLOSE_SOME_LEFT_OVER_WINDOWS
				}
			}
		}

	RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, BoxSync
		
	
	; WIN_XP 5 WIN_7 6 WIN_10 10  
	; --------------------------
	If (OSVER_N_VAR>5 
		and SET_GO_1=1)
		{
		ProgramFilesX86 := A_ProgramFiles . (A_PtrSize=8 ? " (x86)" : "")

		; 32 bit Installed
		Loop Files, %ProgramFilesX86%\ASUS\WebStorage\ASUSWSLoader.exe, R
		{
			FILENAME_ASUSWSLoader = %A_LoopFileFullPath%
		}
		Loop Files, %ProgramFilesX86%\ASUS\WebStorage\AsusWSPanel.exe, R
		{
			FILENAME_AsusWSPanel = %A_LoopFileFullPath%
		}

		Process, Exist, ASUSWSLoader.exe
		If Not ErrorLevel
			{
				FN_VAR=%FILENAME_ASUSWSLoader%
				IfExist, %FN_VAR%
				{
					SoundBeep , 2500 , 100
					Run, "%FN_VAR%"
				}
			}

		Process, Exist, AsusWSPanel.exe
		If Not ErrorLevel
			{
				FN_VAR=%FILENAME_AsusWSPanel%
				IfExist, %FN_VAR%
				{
					SoundBeep , 2500 , 100
					Run, "%FN_VAR%" /S , , HIDE
				}
			}
	}

	RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, WebStorage
	RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, WebStorage
		
		
	IF SET_GO_1=1
	{
	Process, Exist, GoodSync-v10.exe
	If Not ErrorLevel
		{
			FN_VAR:="C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				; Run, "%FN_VAR%" , , MIN ; -- __ -- __ /min
				; STARTING UP MIN HAS WIN 10 PROBLEM LIKE BLUETOOTH LOGGER ONE WAS NOT SHOW FROM TAB UP
				Run, "%FN_VAR%" 
			}
		}
	}
	
	Process, Exist, gs-server.exe
	If Not ErrorLevel
	{
		FN_VAR:="C:\Program Files\Siber Systems\GoodSync\gs-server.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			; Run, "%FN_VAR%"  /service, , MIN
			Run, "%FN_VAR%"  /service
		}
	}
		
	IF SET_GO_1=1
	{
	Process, Exist, googledrivesync.exe
	If Not ErrorLevel
		{
			FN_VAR:="C:\Program Files (x86)\Google\Drive\googledrivesync.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				Run, "%FN_VAR%" /autostart
			}
		}
	}
	RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, GoogleDriveSync
	
	; WIN_XP 5 WIN_7 6 WIN_10 10  
	; --------------------------
	If (OSVER_N_VAR>7 
		and SET_GO_1=1)
		{
		Process, Exist, OneDrive.exe
		If Not ErrorLevel
			{
				FN_VAR:="C:\Users\%A_UserName%\AppData\Local\Microsoft\OneDrive\OneDrive.exe"
				IfExist, %FN_VAR%
				{
					SoundBeep , 2500 , 100
					Run, "%FN_VAR%" /background
				}
			}
		}
	RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, OneDrive

	IF SET_GO_1=1
	{
	Process, Exist, Dropbox.exe
	If Not ErrorLevel
		{
			FN_VAR:="C:\Program Files (x86)\Dropbox\Client\Dropbox.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				Run, "%FN_VAR%" /systemstartup , , HIDE
			}
		}
	}
	RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, Dropbox

	
}

;-------------------------------------------
; REQUIRE LITTLE DELAY ON RUN HUBIC
; GETS SOME CLASH RUN FROM REGISTRY ALSO RUN
;-------------------------------------------
SET_GO_1=0
IF (A_ComputerName="5-ASUS-P2520LA" and A_UserName="MATT 01")
	SET_GO_1=1
IF (A_ComputerName="4-ASUS-GL522VW" and A_UserName="MATT 01")
	SET_GO_1=1
IF (A_ComputerName="7-ASUS-GL522VW" and A_UserName="MATT 04")
	SET_GO_1=1
IF (A_ComputerName="8-MSI-GP62M-7RD" and A_UserName="MATT 01")
	SET_GO_1=1

; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
If (OSVER_N_VAR>5 
	and SET_GO_1=1)
	{
		Process, Exist, hubiC.exe
		If Not ErrorLevel ; errorlevel will = 0 if process doesn't exist
		{
			IfWinNotExist , hubiC
			{
				FN_VAR:="C:\Program Files\OVH\hubiC\hubiC.exe"
				IfExist, %FN_VAR%
				{
					SoundBeep , 2500 , 100
					Run, "%FN_VAR%"
				}
			}
		}
	}

RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, hubiC


Process, Exist, ViceVersa.exe
;If Not ErrorLevel
If TRUE=FALSE
{
	VV_VAR:="C:\Program Files\ViceVersa Pro\ViceVersa.exe"
	IfExist, %VV_VAR%
	{
		;VF_VAR:="C:\SCRIPTER\SCRIPTER -- ViceVersa PRO\VICE V -- VB-VB-NT--  EXE -- AT BOOT.fsf"
		;IfExist, %VF_VAR%
		; Run, "%VV_VAR%" "%VF_VAR%" /dialogautoexec /autoclose , , MIN
		
		VF_VAR:="C:\SCRIPTER\SCRIPTER -- ViceVersa PRO\VV C DRIVE ROOT\VV C DRIVE __  SYSINTERNALS TO PROGRAM FILES.fsf"
		IfExist, %VF_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%VV_VAR%" "%VF_VAR%" /dialogautoexec /autoclose , , MIN
		}
		
		VF_VAR:="C:\SCRIPTER\SCRIPTER -- ViceVersa PRO\VV C DRIVE ROOT\VV C DRIVE __  PROGRAM FILES MOVER ROOT.fsf"
		IfExist, %VF_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%VV_VAR%" "%VF_VAR%" /dialogautoexec /autoclose , , MIN
		}

		VF_VAR:="C:\SCRIPTER\SCRIPTER -- ViceVersa PRO\VV C DRIVE ROOT\VV C DRIVE __  PROGRAM FILES WINRAR.fsf"
		IfExist, %VF_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%VV_VAR%" "%VF_VAR%" /dialogautoexec /autoclose , , MIN
		}
	}
}
	
;FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 10-VICEVERSA _ SHELL FOLDERING__.VBS"
;IfExist, %FN_VAR%
;	Run, "%FN_VAR%" , , MIN

	

FN_VAR:="C:\Program Files\WinRAR\WINRAR REGKEY.BAT"
IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%" , , HIDE
	}
	

If ProcessExist("PStart.exe", A_UserName)=0
	{
		FN_VAR:="C:\PStart\PStart.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" , , HIDE
			WinWait, ahk_class TMainForm.UnicodeClass, , 80
			WinGet, HWND, ID, ahk_class TMainForm.UnicodeClass
			IS_WINDOW_HIDDEN_AND_HIDE(HWND, 50)
		}
	}


SetTitleMatchMode 2  ; Avoids the need to specify the full path of the file below.

IfWinNotExist BAT 01 BOOT KILLER
	{
		FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01 BOOT KILLER.BAT"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" , , MIN
			GOSUB MINIMIZE_ALL_CMD_AT_BOOT
		}
	}

	
IfWinNotExist BAT 01 REGISTRY AT BOOTER
	{
		FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01 REGISTRY AT BOOTER.BAT"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" , , MIN
			GOSUB MINIMIZE_ALL_CMD_AT_BOOT
		}
	}
	
; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
If (OSVER_N_VAR=10)
{
	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 23-ASSOCIATION EXTENSIONS WIN_10 _ ORGINAL.BAT"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%" "/QUITE" , , MIN
		GOSUB MINIMIZE_ALL_CMD_AT_BOOT
	}

	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 24-ASSOCIATION EXTENSIONS WIN_10 _ MY CUSTOM.BAT"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%" "/QUITE" , , MIN
		GOSUB MINIMIZE_ALL_CMD_AT_BOOT
	}

	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 27-ASSOCIATION EXTENSIONS WIN_10 _ STOP_RESET MY_APPS.BAT"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%" "/QUITE" , , MIN
		GOSUB MINIMIZE_ALL_CMD_AT_BOOT
	}
}

	
;--------------------------------------------------------------------

IfWinNotExist SendSMTP_REBOOT_BATCH
	{
		FN_VAR:="C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP_REBOOT_BATCH.BAT"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" , , MIN
		}
	}

;--------------------------------------------------------------------
; THIS COMMAND PROCESS TERMINATES THE APP IN WINDOWS XP	
; TRY ANOTHER METHOD
;--------------------------------------------------------------------
;DetectHiddenWindows, ON
;WinWait, VB KEEP RUNNER ahk_class ThunderRT6FormDC, , 30
;--------------------------------------------------------------------

; TRY ANOTHER METHOD DIDN'T WORK EITHER
;--------------------------------------------------------------------
;LCOUNT=20
;LOOP
;{
;	IFWINEXIST, VB KEEP RUNNER ahk_class ThunderRT6FormDC
;		BREAK
;	LCOUNT-=1
;	IF LCOUNT<0
;		BREAK
;	SLEEP 1000
;}

; BOOT UPS RUNNING SMOOTHER
; 2018-05-21 02:12

	
SET_GO=FALSE
IF (A_ComputerName = "1-ASUS-X5DIJ") 
	SET_GO=TRUE
IF (A_ComputerName = "2-ASUS-EEE") 
	SET_GO=TRUE
IF (A_ComputerName = "4-ASUS-GL522VW") 
	SET_GO=TRUE
IF (A_ComputerName = "7-ASUS-GL522VW") 
	SET_GO=TRUE
IF (A_ComputerName = "8-MSI-GP62M-7RD") 
	SET_GO=TRUE

DetectHiddenWindows, off	
IF SET_GO=TRUE
{
	Process, Exist, bluetoothlogview.exe
	If Not ErrorLevel
	{
		FN_VAR:="C:\PStart\Progs\0_Nirsoft_Package\NirSoft\bluetoothlogview.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%"
			
			; If OSVER_N_VAR>5
			; {
				; Run, "%FN_VAR%"
			; }
			; ELSE
			; {
				; Run, "%FN_VAR%", , MIN
			; }			
		}
	}
	SLEEP 2000
	
	Process, Exist, bluetoothview.exe
	If Not ErrorLevel
	{
		FN_VAR:="C:\PStart\Progs\0_Nirsoft_Package\NirSoft\bluetoothview.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%"

			; If OSVER_N_VAR>5
			; {
				; Run, "%FN_VAR%"
			; }
			; ELSE
			; {
				; Run, "%FN_VAR%", , MIN
			; }
		}
	}
	SLEEP 2000
	
	;GOSUB MINIMIZE_ALL_BLUETOOTH

}

DetectHiddenWindows, ON

SET_GO=FALSE
IF (A_ComputerName = "8-MSI-GP62M-7RD") 
	SET_GO=TRUE

SET_GO=TRUE

IF SET_GO=TRUE
{
	FILE_PATH:="E:\01 Start Menu\#_8-MSI-GP62M-7RD\Programs\Startup\Killer Control Center.lnk"
	if FileExist(FILE_PATH) 
		FileDelete, %FILE_PATH%
		
	If ProcessExist("KillerControlCenter.exe", A_UserName)=0
	{	
		FN_VAR:="C:\Program Files\Killer Networking\Killer Control Center\KillerControlCenter.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" -minimized , , HIDE
		}

		IfWinExist Killer Control Center	
		{
			WinGet, HWND, ID, Killer Control Center	
			IS_WINDOW_HIDDEN_AND_HIDE(HWND, 100)
		}

	}

	
}


SET_GO=FALSE
IF (A_ComputerName = "5-ASUS-P2520LA") 
	SET_GO=TRUE
IF (A_ComputerName = "4-ASUS-GL522VW") 
	SET_GO=TRUE
IF (A_ComputerName = "7-ASUS-GL522VW") 
	SET_GO=TRUE
IF (A_ComputerName = "8-MSI-GP62M-7RD") 
	SET_GO=TRUE

DetectHiddenWindows, off
IF SET_GO=TRUE
{
	Process, Exist, QfinderPro.exe
	If Not ErrorLevel
		{
			FN_VAR:="C:\Program Files (x86)\QNAP\Qfinder\QfinderPro.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				Run, "%FN_VAR%" /min /auto	, , HIDE
				WinWait, QNAP Qfinder Pro, , 240
				WINHIDE
			}
		}
}

DetectHiddenWindows, on



RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, QfinderPro

	
SET_GO=FALSE
IF (A_ComputerName = "4-ASUS-GL522VW") 
	SET_GO=TRUE
IF (A_ComputerName = "7-ASUS-GL522VW") 
	SET_GO=TRUE
IF (A_ComputerName = "8-MSI-GP62M-7RD") 
	SET_GO=TRUE

IF SET_GO=TRUE
{
	Process, Exist, NetworkDriveAgent.exe
	If Not ErrorLevel
		{
			FN_VAR:="C:\Program Files (x86)\QNAP\myQNAPcloud Connect\NetworkDriveAgent.exe"
			IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				Run, "%FN_VAR%" /min , , HIDE
			}
		}
}
	
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, NetworkDriveAgent

	
SET_GO=FALSE
IF (A_ComputerName = "4-ASUS-GL522VW") 
	SET_GO=FALSE
IF (A_ComputerName = "7-ASUS-GL522VW") 
	SET_GO=TRUE

IF SET_GO=TRUE
{
	If ProcessExist("dfx.exe", A_UserName)=0
		{
		FN_VAR:="C:\Program Files (x86)\DFX\dfx.exe"
		IfExist, %FN_VAR%
			{
				SoundBeep , 2500 , 100
				Run, "%FN_VAR%" -startup , , HIDE
				WinWait, AHK_CLASS #32770, , 120
				LOOP, 1000
				{
					SLEEP 10
					WinGet, HWND, ID, AHK_EXE dfx.exe AHK_CLASS #32770
					WinGetClass, This_Class, ahk_id %HWND%
					WinGet, path, ProcessName, ahk_id %HWND%
					IF (This_Class="#32770" 
					and PATH="dfx.exe")
					{
						IS_WINDOW_HIDDEN_AND_HIDE(HWND, 10)
						aParent:=DllCall( "GetParent", UInt, HWND) + 0
						HWND=%aParent%
						IS_WINDOW_HIDDEN_AND_HIDE(HWND, 50)
						
					}
					IF HWND>0 
						BREAK
				}
				
				
		}
	}
}
	
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, FxSound Enhancer



; WIN_XP 5 WIN_7 6 WIN_10 10  
; --------------------------
If (OSVER_N_VAR>5)
{
	If ProcessExist("SetPoint.exe", A_UserName)=0
	{
		FN_VAR:="C:\Program Files\Logitech\SetPointP\SetPoint.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" /launchGaming
		}
	}
	
	WinWait, SetPoint Settings ahk_class KEMUIMainWndClass, , 10
	SLEEP 400
	WINCLOSE

}




RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, EvtMgr6


; WIN_XP 5 WIN_7 6 WIN_10 10  
If (OSVER_N_VAR=5)
{
		FN_VAR:="C:\PStart\# NOT INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\Bginfo.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" /timer:0 , , HIDE
		}
}

;--------------------------------------------------------------------
; STOP THE CHROME AUTO LAUNCH MAINLY FOR XP BUT ALSO REDUCE 
; MESS IN AUTO START __ MESSY LITTLE ITEM
; StubPath = "C:\Program Files (x86)\Google\Chrome\Application\66.0.3359.139\Installer\chrmstp.exe" --configure-user-settings --verbose-logging --system-level (REG_SZ)
;--------------------------------------------------------------------
; WIN_XP 5 WIN_7 6 WIN 10 10  
; If (OSVER_N_VAR=5)
;--------------------------------------------------------------------

	Loop, Reg, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Active Setup\Installed Components\, VR
	{
		if A_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		
		IfInString, value, Google\Chrome\Application
		IfInString, value, \Installer\chrmstp.exe
		{
		SoundBeep , 2500 , 100
		RegDelete, %A_LoopRegKey%\%A_LoopRegSubKey%, %A_LoopRegName%
		;MsgBox, 4, , %A_LoopRegKey%\%A_LoopRegSubKey%
		;MsgBox, 4, , %A_LoopRegName% = %value% (%A_LoopRegType%)`n`nContinue?
		;IfMsgBox, NO, break
		}
	}
;--------------------------------------------------------------------



;--------------------------------------------------------------------
	;C:\Program Files (x86)\Google\Chrome\Application\chrome.exe  --flag-switches-begin --flag-switches-end --restore-last-session

	RegDelete, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Runonce, Application Restart #0

	Loop, Reg, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Runonce\, VR
	{
		if A_LoopRegType = key
			value =
		else
		{
			RegRead, value
			if ErrorLevel
				value = *error*
		}
		
		IfInString, value, Google\Chrome\Application
		IfInString, value, chrome.exe
		{
		SoundBeep , 2500 , 100
		RegDelete, %A_LoopRegKey%\%A_LoopRegSubKey%, %A_LoopRegName%
		;MsgBox, 4, , %A_LoopRegKey%\%A_LoopRegSubKey%
		;MsgBox, 4, , %A_LoopRegName% = %value% (%A_LoopRegType%)`n`nContinue?
		;IfMsgBox, NO, break
		}
	}
;--------------------------------------------------------------------


;--------------------------------------------------------------------

If ProcessExist("lastapp_x64.exe", A_UserName)=0
{
	FN_VAR:="C:\Program Files (x86)\LastPass\lastapp_x64.exe"
	IfExist, %FN_VAR%
	{
		; SoundBeep , 2500 , 100
		; Run, "%FN_VAR%"
	}
}
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, LastApp

Process, Exist, FileZilla Server.exe
If Not ErrorLevel
{
	SoundBeep , 2500 , 100
	RunWait,sc start "FileZilla Server" 
	CODE_RUN_FOR_BRUTE_BOOT_DOWN_AHK=TRUE
	GOSUB BRUTE_BOOT_DOWN_AHK_SUB

}
Process, Exist, FileZilla Server Interface.exe
If Not ErrorLevel
{
	FN_VAR:="C:\Program Files (x86)\FileZilla Server\FileZilla Server Interface.exe"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%" , , HIDE
		CODE_RUN_FOR_BRUTE_BOOT_DOWN_AHK=TRUE
		GOSUB BRUTE_BOOT_DOWN_AHK_SUB
	}
}
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, FileZilla Server Interface

Process, Exist, NoIPDUCService4
If Not ErrorLevel
{
	SoundBeep , 2500 , 100
	RunWait,sc start "NoIPDUCService4" 
}
Process, Exist, DUC40.exe
If Not ErrorLevel
{
	FN_VAR:="C:\Program Files (x86)\No-IP\DUC40.exe"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, "%FN_VAR%" , , HIDE
	}
}


DetectHiddenWindows, OFF
SET_GO=FALSE

If OSVER_N_VAR=10
	SET_GO=TRUE

;IF (A_ComputerName = "4-ASUS-GL522VW") 
;	SET_GO=FALSE
	
IF SET_GO=TRUE 
{	
	Process, Exist, SystemExplorer.exe
	If Not ErrorLevel
		{	FN_VAR:="C:\PStart\Progs\#_PortableApps\PortableApps\SystemExplorerPortable\App\SystemExplorer\SystemExplorer.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" /TRAY
			;WinWait, System Explorer, , 40
			;SLEEP 800
			;WINCLOSE
			;WinMinimize
			;WinHIDE, System Explorer
			SLEEP 4000
			
			CODE_RUN_FOR_BRUTE_BOOT_DOWN_AHK=TRUE
			GOSUB BRUTE_BOOT_DOWN_AHK_SUB
		}
	}
}
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run, SystemExplorerAutoStart


DetectHiddenWindows, OFF
SET_GO=FALSE
If OSVER_N_VAR>6
SET_GO=TRUE 

IF SET_GO=TRUE 
{	
	Process, Exist, procexp64.exe
	If Not ErrorLevel
		{	;FN_VAR:="C:\PStart\Progs\#_PortableApps\PortableApps\ProcessExplorerPortable\ProcessExplorerPortable.exe"
		FN_VAR:="C:\PStart\# NOT INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\procexp.exe"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" /t
			
			;WinWait, System Explorer, , 40
			;SLEEP 800
			;WINCLOSE
			;WinMinimize
			;WinHIDE, System Explorer
		}
	}
}

; WIN XP 2-ASUS-EEE
; "C:\Program Files\Common Files\Adobe\ARM\1.0\AdobeARM.exe"
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, Adobe ARM
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled, Adobe ARM
;
; DISABLED AUTORUNS 2-ASUS-EEE
; HotKeysCmds
; C:\WINDOWS\system32\hkcmd.exe
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, HotKeysCmds
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled, HotKeysCmds

; DISABLED AUTORUNS 2-ASUS-EEE
; NeroFilterCheck
; C:\WINDOWS\system32\NeroCheck.exe
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, NeroFilterCheck
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled, NeroFilterCheck

; DISABLED AUTORUNS 2-ASUS-EEE
; PATHPILOT
; C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.exe
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, PATHPILOT
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled, PATHPILOT

; DISABLED AUTORUNS 2-ASUS-EEE
; Persistence
; C:\WINDOWS\system32\igfxpers.exe
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, Persistence
RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled, Persistence

; MSMSGS AUTORUNS 2-ASUS-EEE
; "C:\Program Files\Messenger\msmsgs.exe" /background
RegDelete, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run, MSMSGS
RegDelete, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled, MSMSGS

; WINDOWS 10
FILE_PATH:="C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\NeroDesktopSwitcher.scf"
if FileExist(FILE_PATH) 
	{
		SoundBeep , 2500 , 100
		FileSetAttrib, -RH, %FILE_PATH% 
		FileDelete, %FILE_PATH%
	}

; WINDOWS 10
FILE_PATH:="E:\01 Start Menu\#_7-ASUS-GL522VW\Programs\Startup\NeroDesktopSwitcher.scf"
if FileExist(FILE_PATH)
	{
		SoundBeep , 2500 , 100
		FileSetAttrib, -RH, %FILE_PATH% 
		FileDelete, %FILE_PATH%
	}

	
DetectHiddenWindows, ON


; ----------------------------------------
; ONLY IN THE FIRST FIVE MINUTES OF TICKER
; ----------------------------------------
COUNT_TICK_TIME=% 1000*60*5
;IF A_TICKCOUNT< %COUNT_TICK_TIME%
{
	SET_GO=TRUE
	IF (A_ComputerName = "3-LINDA-PC") 
		SET_GO=FALSE

	If OSVER_N_VAR<6
		SET_GO=FALSE
	
	IF SET_GO=TRUE 
	{	
		FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 12-PinItem BATCH.VBS"
		IfExist, %FN_VAR%
		{
			SoundBeep , 2500 , 100
			Run, "%FN_VAR%" /Q , , MIN
		}
	}	
}	

GOSUB POWERSHELL

GOSUB MINIMIZE_ALL__EXPLORER_AT_BOOT
GOSUB MINIMIZE_ALL_CMD_AT_BOOT
GOSUB MINIMIZE_ALL_CHROME_AT_BOOT
GOSUB MINIMIZE_ALL_BLUETOOTH

	
;----------------------------------------------------------------
; NORTON WANTS TO RUN AT END IT'S ENGINE HASN'T STARTED UP PROPER	
; CALL TOO QUICK AND BE HASN'T RUN
;----------------------------------------------------------------
FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 18-NORTON CONTROL BOOTER.ahk"
IfExist, %FN_VAR%
{
	SoundBeep , 2500 , 100
	Run, "%FN_VAR%"
}


;--------------------------------------------------------------------
; ONLY SOMETIMES ASUS WEBSTORAGE DOESN'T OPEN QUIETLY
; AND WITH A DELAY
; AND LOAD THE EXPLORER FOLDER
; SAVE TO END 
; KEEPS CODE RUNNING A BIT
; 4 MINUTES WAIT
; OPPS GOT NEW SETTING FOR STARTUP NOW
;--------------------------------------------------------------------
;WinWait, D:\0 CLOUD\ASUS WEBSTORAGE _5GB ahk_class CabinetWClass, , 300
;WINCLOSE


; JUST IN CASE OPEN WITH HAS BEEN SET WRONG PREVENTING WINMERGE FROM RUNNING PROPERLY WITH NOT AT ALL
; SOMETIME EDIT WITH NOTEPAD GET PUT IN
RegDelete, HKEY_CLASSES_ROOT\WinMerge.Project.File\shell\edit

; 3-LINDA-PC
; C:\Program Files\Winamp\winamp.exe

RegDelete, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run, OODITRAY.EXE
; C:\Program Files\Laplink\DiskImage\ooditray.exe

;HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run
;SecurityHealth
;REG_EXPAND_SZ
;%ProgramFiles%\Windows Defender\MSASCuiL.exe

GOSUB BRUTE_BOOT_DOWN_AHK_SUB

GOSUB CLOSE_SOME_LEFT_OVER_WINDOWS

GOSUB GRAMMARLY_CREATE_SHORTCUT

RETURN

;--------------------------------------------------------------------
TEST_STARTER_RUN_IN:


; ExitApp

RETURN


GRAMMARLY_CREATE_SHORTCUT:

; --------------------------------------------------------------------------------------------
; "C:\Users\MATT 04\AppData\Local\GrammarlyForWindows\app-1.5.32\GrammarlyForWindows.exe"
; --------------------------------------------------------------------------------------------
; GRAMMARLY NOW HAS OWN LAUNCHER FOR EASIER LOAD WITHOUT VERSION PATH DROP DOWN ONE PATH LOWER
; --------------------------------------------------------------------------------------------

SET_GO=TRUE
	
IF SET_GO=TRUE
{
	
	FileGetShortcut, %A_Desktop%\Grammarly.lnk, OutTarget, OutDir, OutArgs, OutDesc, OutIcon, OutIconNum, OutRunState

	; MsgBox %OutTarget%`n%OutDir%`n%OutArgs%`n%OutDesc%`n%OutIcon%`n%OutIconNum%`n%OutRunState%
	
	SET_GO=FALSE
	; IF FILE CONTAIN A NETWORK PATH BY MISTAKE AND FAULT -----------
	IF INSTR(OutTarget,"\\") > 0 
		SET_GO=TRUE
	
	; IF FILE TARGET DOS NOT EXIST ----------------------------------
	IF !FileExist(OutTarget)
		SET_GO=TRUE
	
	; IF FILE SHORTCUT DOS NOT EXIST --------------------------------
	if FileExist("%A_Desktop%\Grammarly.lnk")=false
		SET_GO=TRUE
	
	; IF FILE TARGET MATCHES USERNAME--------------------------------
	TEST_PATH_1 = C:\Users\%A_UserName%\AppData\Local
	IF INSTR(OutTarget,TEST_PATH_1) = 0 
		SET_GO=TRUE
	
	ProgramFilesX86 = C:\Users\%A_UserName%\AppData\Local\GrammarlyForWindows\GrammarlyForWindows.exe

	; GRAMMARLY Install
	FILENAME = ""
	Loop Files, %ProgramFilesX86%, R  ; Recurse into subfolders.
	{
		FILENAME = %A_LoopFileFullPath%
		SHORTCUT_FILENAME = %A_LoopFileFullPath%
	}

	IF FILENAME="" 
	{
		;SoundBeep , 3500 , 100
		;SoundBeep , 2500 , 100
		;MSGBOX, "AutoHotkeys __ Norton Security Was Not Found - ENDER"
		;SoundBeep , 3500 , 100
		;SoundBeep , 2500 , 100
		;Exitapp
	}
	
	IF FILENAME<>"" 
		IF SET_GO=TRUE
		{
			SoundBeep , 2000 , 100
			FileCreateShortcut, %SHORTCUT_FILENAME%, %A_Desktop%\Grammarly.lnk, , "%A_ScriptFullPath%"
		}
}
RETURN




MINIMIZE_ALL_BLUETOOTH:

	SetTitleMatchMode 3
	DetectHiddenWindows, OFF

	SET_GO=FALSE
	If OSVER_N_VAR>5
		SET_GO=TRUE
	
	SET_GO=TRUE

	IF SET_GO=FALSE
		RETURN
	
	IF WinExist("BluetoothView") 
	{
		SLEEP 1000
		WinGet, HWND, ID, BluetoothView
		IS_WINDOW_MINIMIZED_THEN_MINIMIZE(HWND,800)
	}
	
	IF WinExist("ahk_class BluetoothLogView")
	{
		SLEEP 1000
		WinGet, HWND, ID, ahk_class BluetoothLogView
		IS_WINDOW_MINIMIZED_THEN_MINIMIZE(HWND,800)
	}
	
RETURN

BRUTE_BOOT_DOWN_AHK_SUB:
	IF CODE_RUN_FOR_BRUTE_BOOT_DOWN_AHK=TRUE
	{
		FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 32-BRUTE BOOT DOWN.ahk"
		IfExist, %FN_VAR%
			IF !WinExist(FN_VAR) 
			{
				SoundBeep , 2000 , 100
				Run, "%FN_VAR%"
			}
	}
RETURN

CLOSE_SOME_LEFT_OVER_WINDOWS:
	SetTitleMatchMode 3
	DetectHiddenWindows, OFF
	
	COUNT_TICK_TIME=% 1000*60*8
	; IF OUR SET TIME IS LESS THEN TICK TIME LET IT GO
	; ONLY IF BOOT IN EARLY
	;-------------------------------------------------
	LOOP
	{
		SET_GO=FALSE
		If OSVER_N_VAR>5
			SET_GO=TRUE
		IF COUNT_TICK_TIME>%A_TICKCOUNT%
			SET_GO=TRUE
		IF CLOSE_SOME_LEFT_OVER_WINDOWS_VAR=TRUE
			SET_GO=TRUE

		BOXSYNC:="C:\Program Files\Box\Box Sync\BoxSyncMonitor.exe ahk_class"
		LOGIOPTIONS:="C:\ProgramData\Logishrd\LogiOptions\Software\Current\laclient\laclient.exe ahk_class ConsoleWindowClass"
		
		IF SET_GO=TRUE
		{
			IFWinEXIST, %BOXSYNC% ConsoleWindowClass
			{
				WINCLOSE
				SoundBeep , 2500 , 100
			}
			
			IFWinEXIST, %LOGIOPTIONS%
			{
				WINCLOSE
				SoundBeep , 2500 , 100
			}
			
			;IFWinEXIST, C:\ProgramData\Logishrd\LogiOptions\Software\Current\laclient\laclient.exe ahk_class ConsoleWindowClass
			;{
			;	WinGet, HWND, ID, C:\ProgramData\Logishrd\LogiOptions\Software\Current\laclient\laclient.exe ahk_class ConsoleWindowClass
			;	WinGet, PID_VAR, PID, ahk_id %HWND%
			;	Process, Close, %PID_VAR%

			;	SoundBeep , 2500 , 100
			;}
			
		}
		SET_GO=FALSE
		IFWinEXIST, %BOXSYNC%
		SET_GO=TRUE
		IFWinEXIST, %LOGIOPTIONS%
		SET_GO=TRUE
		IF SET_GO=FALSE 
			BREAK
		IF COUNT_TICK_TIME>%A_TICKCOUNT%
			BREAK
	}
RETURN


POWERSHELL:

	IF (A_ComputerName = "1-ASUS-X5DIJ") 
		RETURN
	IF (A_ComputerName = "2-ASUS-EEE") 
		RETURN

	;-----------------------------------------------
	; PROBLEM WITH RUN THIS AT THE END OF THE SCRIPT
	; PREFER AT BEGINNING OR GET LOCKUP SCREEN
	; REQUIRE SOME SORT OF SLEEP AFTER MAYBE
	; IT'S OKAY JUST DELAY TO GET RUNNER SOMETIMES
	;-----------------------------------------------
	DetectHiddenWindows, OFF
	
	FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 01 - WINRM QUICKCONFIG.PS1"
	IfExist, %FN_VAR%
	{
		SoundBeep , 2500 , 100
		Run, C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -File "%FN_VAR%" ,MIN
	}
	WinWait, WindowsPowerShell ahk_class ConsoleWindowClass, , 20
		WinMinimize
	;MSGBOX HERE
	
	
	;powershell -noexit -executionpolicy bypass -File "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 01 - WINRM QUICKCONFIG.PS1"

RETURN

;--------------------------------------------------------------------
; FUNCTION SET 
;--------------------------------------------------------------------


IS_WINDOW_MINIMIZED_THEN_MINIMIZE(HWND, TIME_LENGHT)
{
	IF (!HWND)
		RETURN

	WinGet, Style, Style, ahk_id %HWND%
	; 0x20000000 = WS_MINIMIZE
	; MSGBOX % (Style & 0x20000000)
	IF (Style & 0x20000000)=0
	{
		WINMINIMIZE ahk_id %HWND%
		DONE_BLUETOOTH_MINIMIZE=TRUE
	}
}
RETURN

IS_WINDOW_HIDDEN_AND_HIDE(HWND, TIME_LENGHT)
{
WINGETCLASS,OUTPUTVAR, ahk_id %HWND%
;tooltip, %OUTPUTVAR% __ GO 1
SLEEP 100
IF (!HWND)
	RETURN
; FEW QUICK COUNT TO CHECK WAIT TO SHOW
XY=0
LOOP, %TIME_LENGHT%
{
	XY+=1
	WINHIDE ahk_id %HWND%
	WinGet, Style, Style, ahk_id %HWND%
	; 0x10000000 is WS_VISIBLE  Style & 0x10000000 = 0 IS HIDDEN > 0 NOT HIDDEN
	; TOOLTIP % XY
	IF (Style & 0x10000000)=0
		{
		;MSGBOX STYLE 0
		SLEEP 10
		}
	IF (Style & 0x10000000)>0
		{
		;MSGBOX BREAK
		BREAK
		}
}
;tooltip, %OUTPUTVAR% __ GO 2
; AND THEN DO THE HIDE
LOOP {
	WINHIDE ahk_id %HWND%
	WinGet, Style, Style, ahk_id %HWND%
	; 0x10000000 is WS_VISIBLE  Style & 0x10000000 = 0 IS HIDDEN > 0 NOT HIDDEN
	IF (Style & 0x10000000)>0
	{
		SLEEP 10
		WINHIDE ahk_id %HWND%
		WINCLOSE ahk_id %HWND%
	}
	WINGET, HWND, ID, ahk_id %HWND%
} UNTIL ((Style & 0x10000000)=0 or HWND=0)
}
RETURN

;--------------------------------------------------------------------
;--------------------------------------------------------------------


;# ------------------------------------------------------------------
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

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))


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
;# -----------------------------------------------------------------0
; exit the app



; -------------------------------------------------------------------
; JUNK BOX SNIPPETS USED IN CODE PROGRAMMER
; -------------------------------------------------------------------
; CONVERT HEX TO DECIMAL
; -------------------------------------------------------------------
;SetFormat, integer, decimal
;result := HWND+0 
;MSGBOX, %This_Class% -- %PATH% -- %result% -- %aParent%
; -------------------------------------------------------------------

; -------------------------------------------------------------------
;SetTitleMatchMode 3
;SetTitleMatchMode, FAST
; -------------------------------------------------------------------
