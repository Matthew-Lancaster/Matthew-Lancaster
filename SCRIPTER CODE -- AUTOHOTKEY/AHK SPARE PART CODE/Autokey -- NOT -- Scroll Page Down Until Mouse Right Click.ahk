;Autokey Project to Use Enter Key on (Save As) Window Dialogue of Chrome.exe & Exit On Chrome Gone
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;
;-----------------------------------------------------------------------------------------------------------------------------------------
; Code to Press Enter Key on Save As Window in Chrome.exe & Exit When Chrome.exe Gone
;-----------------------------------------------------------------------------------------------------------------------------------------
; Thu 25 Aug 2016 01:47:51 -- Length Time 1 hr 45 Min -- Finish Code
; Thu 25 Aug 2016 02:35:00 -- Length Time 2 hr 30 Min -- Project End Time -- Fine Tune & Documentation
; Thu 25 Aug 2016 03:00:00 -- Length Time 3 hr 00 Min -- Fine Tune Documentation And Publish Online

; A Second Session Bash in the Morning
; Thu 25-Aug-2016 06:16:41 --
; Thu 25 Aug 2016 07:00:00 -- Length Time 0 hr 40 Min -- Tune Documentation And Publish Another ReFresh

;-----------------------------------------------------------------------------------------------------------------------------------------
; 2nd Project in a Line of AutoKey -- First On Blogger Here
; This is All Copy Paste in and Go Project for AutoKey Working
; Building Block  Head Start
;-----------------------------------------------------------------------------------------------------------------------------------------
;
;--------------------
#SingleInstance force
;--------------------
;Here --- SingleInstance force -- Reduce Question Ask When Reload From Edit and More
;--------------------

;----------------------------------------------------------------------
;----------------------------------------------------------------------
;Here --- a code block to use around --- Persistent
#Persistent

; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))

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
;----------------------------------------------------------------------
;----------------------------------------------------------------------

;----------------------------------------------------------------------
;NOTE After a Little Work Test Method and Find Error Fix Through Help File
;Because Window 10 Won't Except Keyboard Send From The AutoKey Unless 
;It Set to Load As Administrator 
;By Set Properties on This Program in 64 Bit or (x86) 32 Bit -- Of the Main Program Folder
;C:\Program Files\AutoHotkey\AutoHotkey.exe
;Yes Fixed the Error Variable Increment -- Seem Like a C++ Method Style -- TEST_VAR += 1

;This Program Has Problem in Win 10 The Dialog to Save Do Stay So Move On After Save to Next Icon to Do Save JPG Screenshot
;----
;Save as Shortcut - Chrome Web Store
;https://chrome.google.com/webstore/detail/save-as-shortcut/flehofiklehmnnolpjcamplcnmhgcbkk?utm_source=chrome-app-launcher-info-dialog
;----

;A Tip to Learn With This Program is 
;When Ask to Save Select JPG From PNG and It will Remember
;----
;Capture Webpage Screenshot Entirely. FireShot - Chrome Web Store
;https://chrome.google.com/webstore/detail/capture-webpage-screensho/mcbpblocgmgfnpjjppndjkmgjaogfceg?utm_source=chrome-app-launcher-info-dialog
;----


;----------------------------------------------------------------------
;Work Logg -- Start time not that clear
;Yes Found Time Begin Phew Charging Someone for This Automation -- RSI Injury

;Thu 25-Aug-2016 11:20:00 
;Thu 25-Aug-2016 14:59:29
;Thu 25-Aug-2016 16:31:00 11 TO 4pm 5 Half Hour Hurt Ouch
;Long Time to Sort a Admin Out Improve Code at Same TIme
;Next New Time of CPU User -- getting 0.03%
;----------------------------------------------------------------------

;----------------------------------------------------------------------
;My code begin
;----------------------------------------------------------------------
While EverLoop=EverLoop
	{
    if WinActive("Save As" ahk_class #32770)
        {
        if WinActive(ahk_exe chrome.exe)
            {
            if WinActive("Save As" ahk_class #32770)
				{
				;--------------------------------------
				; For Some Reason It Maybe I Think Best to Find This Window
				; Because to Use WinActive With Look for Process Maybe Not Quite Right -- Wanting to Focus Put on the Window HWND Handle Than Process ID PID
				; And the Step Method 1.. 2.. 3.. is Speed Save As at 1st is Quicker and Poll a Process Will Use More CPU
				; process Tick's Ticker
				; Even Tough It is Done Every 500 Mill Sec Later -- More Code With Counter Can Make That Less
				;--------------------------------------
				; Modified - Is Not Much Difference to Check Below the Window Title and Class Name
				; in Term of Lead On to and The Process ID and Process Usage CPU Ticker Hungry Check to Do Both
				; The Two << Title and Class >> Name Both Same In CPU Ticker Check Term
				; Two Not same as Process Hungry CPU Process ID Check
				; WinActive("Save As" ahk_class #32770)

                WinActivate  ; Activate the window it found.
				;SLEEP 100
				;Send {Enter}
                ;Send {Enter Down}
				;Sleep 200
                ;Send {Enter Up}
				;Sendinput !+S
				Sendinput {Enter}

				TEST_VAR += 1 ; in help file under set timer learn to increment the variable
				IF TEST_VAR < 2
					{
					;SLEEP 100
					;MsgBox, First Run Working Okay Press Send {Enter} Done
					}
                }
			}
		}

	Var_Count += 1
	
	;if var_count < 3 
	;	msgbox, test
		
	if Var_Count > 500
		{

		;500  Milli Sec with 500 -- is 250 Second without Chrome Exist and then Exit
		;1000 Milli Sec with 500 -- is 500 Second without Chrome Exist and then Exit		
		;-------------------------------------------------------------------------------------
		;-------------------
		Var_Count = 0
		;Up and Down Numeric 0
		;-------------------
		;Less Frequent Check on a Process ID a CPU Ticker Usage is High -- Enumerate the Processes
		;Double Check and Continuous Run Through With Var_Count Loop Odd Check and was Run at CPU 0.04%
		;A Continuous KeyBoard Check Run CPU Ticker At 0.01< and Less and None CPU Usage -- Process Explorer Result
		;---------------------------------------------------------------------------------------
		;Result State is Half CPU Ticker Usage at 0.02%
		;Ticker Meaning How Many Instruction Flop to Perform the Operation Command CPU Flop Thinker the Flip Flop Logic IC Integrated Circuit Terminology
		;---------------------------------------------------------------------------------------

		IfWinNotExist, ahk_exe chrome.exe
			{
				Exitapp
			}
		}

	;------------------------------------------------------------
	;--------------
	;Sleep, 1000 ; tend to use more CPU Usage with Smaller Sleep 
	; -----------  at 1000 Get 0.02 and 0.03% on The Asus Gl522VW Intel 7
	; -----------  Slight Less Than 0.03 and 0.04 at 500 millisecond
	;Sleep 500
	Sleep 100
	;--------------
	;It Don't Make Any Difference to Sleep 500 or 1000 Millisecond to the CPU 0.02% Usage But Take It Out and CPU Ticker 10%
	;Get a Quicker Response to Press Save As Button on 500 Than 1000 Sleeper
	;------------------------------------------------------------
	;EVER LOOP END
	}
;----------------------------------------------------------------------
;----------------------------------------------------------------------
;My Code Block Done
;----------------------------------------------------------------------
;----------------------------------------------------------------------

;----------------------------------------------------------------------
;----------------------------------------------------------------------
;Info Note
;----------------------------------------------------------------------
;Project Creation, Intention and Purpose for Use On-Line
;----------------------------------------------------------------------
;Some Website Specially Filter Out Question that Been Ask Before to Avoid Repeat
;While Other Coder Person Keep Them Code to Their Own
;Line Above As Is the Philosophy of Work Employment Not give to Giving Your Trade Secret Away
;Which I feel is Such a Conflict
;Other Person Own Web Have It Written Newbie For Dummy's Which Seem to Be a Repeat Issue as a New Operating Sys and Product arrive Every year and Almost Every Year
;Sometime Avoiding Make Clear And the Actual Complete Code -- Got to Read the Dribble
;Other Person As in PWGen Code Author Generate the Code Free and source Code Free
;----------------------------------------------------------------------
;Project Creation, Intention and Purpose
;----------------------------------------------------------------------
;----
;To Save Item Detailed Technical Info on Page Of thing Buy on eBay and Other Online
;Especially With Intention of Sell on Again Later After Buy Less Expensive at First
;And Or Sell on Again Later an Item Not Good Enough for a Job
;----
;Also to Strengthen My Library of Code for this AutoKey Purpose
;As I Immediately Saw as Much Better Alternative to Knock up a Code in Visual Basic
;----
;Also the Recent Change By Google to Make Everything that is a Download a Risk Question Including Own Product Chrome
;And to Switch Download to Back on When Before was Auto
;And Talk is that is Easier than a Dot You Want Keep Button By Instead Use Enter Key On Save
;But Now Download Ask Option is Not always the Download Folder
;
;Also Was Talk When Use Select Download Method Rather Than Keep Button
;It Begin Download Soon As Before Give Request to Save
;Humm Thinker AutoKey Answer
;----

;----
;And Here Following Something I Mean to Accomplish a Long Time
;----
;Amazon.com the US -- Recently Removed all My Buy from There -- But Not The .co.uk
;Which Keep Forever
;----
;Ebay Every End Of Year a Whole Year in Past Is Wiped Out

;Use of these 4 Link in Turn to Save Page of the Chrome Extension
;1.. Copy All Urls - Chrome Web Store
;2.. Save as Shortcut - Chrome Web Store
;3.. Capture Webpage Screenshot Entirely. FireShot - Chrome Web Store
;4.. SingleFile - Chrome Web Store -- to Save an HTML Version
;----------------------------------------------------------------------
;----------------------------------------------------------------------

;----------------------------------------------------------------------
;Ref: Source
;----------------------------------------------------------------------
;----------------------------------------------------------------------
;Info Techno --
;Coder To the Next Level Part Two -- Detect Window Form -- To Learn and Coder With Efficiently
;Coded Style --
;----------------------------------------------------------------------
;Efficiently and Worked All Redundant Code That Possible to Remove Compacting
;That Make Logic Proof Reading Eye Sight Better
;----------------------------------------------------------------------
;Efficiently and Noted Numeric 0 Use of Processor CPU time Even With Poll the
;Process Chrome.exe Every 500 Mill Sec
;----------------------------------------------------------------------
; Length Time 1 hr 45 Min %+- Min
; Length Time 2 hr 30 Min %+- Min
;----------------------------------------------------------------------
;Thu 25-Aug-2016 00:04:42
;Thu 25 Aug 2016 01:47:51 -- Finish Code
;Thu 25 Aug 2016 02:35:00 -- Fine Tune Documentation --- Length Time 2 hr 30 Min %+-
;Thu 25 Aug 2016 03:00:00 -- Fine Tune Documentation And Publish On-Line --- Length Time 3 hr

; A Second Session Bash in the Morning
; Thu 25-Aug-2016 06:16:41 --
; Thu 25 Aug 2016 07:00:00 -- Length Time 0 hr 40 Min -- Tune Documentation And Publish Another ReFresh

;----------------------------------------------------------------------
;----------------------------------------------------------------------

;----------------------------------------------------------------------------------------------------------------
;Ref: - Credit to Source Helper Ref:

;----------------------------------------------------------------------------------------------------------------
;AutoKey Source to Learn
;----------------------------------------------------------------------------------------------------------------
;Begin Here Google Search
;----------------------------------------------------------------------------------------------------------------
;1..
;How to save files using Enter key? - Ask for Help - AutoHotkey Community
;https://autohotkey.com/board/topic/124006-how-to-save-files-using-enter-key/

;----------------------------------------------------------------------------------------------------------------
;AutoKey Info Help Web
;----------------------------------------------------------------------------------------------------------------
;1..
;OnExit
;https://autohotkey.com/docs/commands/OnExit.htm
;--------
;2..
;While-loop
;https://autohotkey.com/docs/commands/While.htm
;--------
;3..
;Else
;https://autohotkey.com/docs/commands/Else.htm
;--------
;4..
;OnExit
;https://autohotkey.com/docs/commands/OnExit.htm
;--------
;5..
;Functions
;https://autohotkey.com/docs/Functions.htm
;----------------------------------------------------------------------------------------------------------------


;----------------------------------------------------------------------------------------------------------------
;Ref
;Chrome Extension
;----------------------------------------------------------------------------------------------------------------

;1..
;--------
;This One not Much To do With Here But Still Used
;--------
;Copy All Urls - Chrome Web Store
;https://chrome.google.com/webstore/detail/copy-all-urls/djdmadneanknadilpjiknlnanaolmbfk?utm_source=chrome-app-launcher-info-dialog
;--------

;2..
;--------
;Yes Save to URL to Download Folder
;--------
;Save as Shortcut - Chrome Web Store
;https://chrome.google.com/webstore/detail/save-as-shortcut/flehofiklehmnnolpjcamplcnmhgcbkk?utm_source=chrome-app-launcher-info-dialog
;--------

;3..
;--------
;This One Save Auto to Download Folder Without Our Coding Requirement
;--------
;Capture Webpage Screenshot Entirely. FireShot - Chrome Web Store
;https://chrome.google.com/webstore/detail/capture-webpage-screensho/mcbpblocgmgfnpjjppndjkmgjaogfceg?utm_source=chrome-app-launcher-info-dialog
;--------

;4..
;--------
;As Well as Save Screen Shot About Do a Web Page HTML Save also
;Be Good If Done Without an Extra Ask to Download Save As
;But Setup to Store a Page on a Web On-line Auto -- PageArchiver
;--------
;SingleFile - Chrome Web Store
;https://chrome.google.com/webstore/detail/singlefile/mpiodijhokgodhhofbcjdecpffjipkle?utm_source=chrome-app-launcher-info-dialog
;----------------------------------------------------------------------------------------------------------------