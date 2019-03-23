#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;
;---------------------------------------------------------------------------------------------
;---------------------------------------------------------------------------------------------
; RESULT OF USER CODE 
; in Brief, Quick Scroll Down A Web Page F4 Key And Mouse Button Stop Start
;---------------------------------------------------------------------------------------------
;---------------------------------------------------------------------------------------------
;--1.. Press Control F4 To Begin
;---    It Will Scroll Down The Page Of Example Space Book Grin Book Land  FB
;--     Next In Line Of Job 
;--2.. It Will Stop When The Form Window Has Switch To Another By Detect The Text Title Has Changed
;--3.. In The Begin Was Set Not To Scroll More Than 4million RSI Repeat Injury
;--     If It Happen Without Sleeper Delay It Would Dish Out 4 Million Keys Very Quick
;--     Info -- It Will Use Arrow Key Down Or You Can Change That Hard Coder Remark And 
;--     Have It Page Down
;--     Not Much Speed Difference Bit Jerky When Page Down
;--4.. Next Most Complicated -- Is When It Happen To Be Scroll Down
;----- Is Option To Stop And Start With The Same Mouse 
;----- Right Click Button When Want Stop Anything Interesting
;----- And Toggle That Yes -- As You Learn It Is Initiated By The Control F4 Key
;----- Careful Not To Press ALT-F4 Or It Will Close A Window
;----- Also The Last Bit To Learn And Wonderful Hard 
;----- Was When Use The Mouse Right Button Stop Start Toggle And The Button 
;----- Is Disabled For The Action So That It Does Not Bright The Context Menu Up
;----- I Thing Also Maybe Well Originally Or Maybe Or-Whore It Only Work Happen when for Chrome
;----- Unless You Want To Change
;---------------------------------------------------------------------------------------------
; The Beginning Of My AutoHotKeys Learn
; and Careful To Keep The Source Code Exampling Found For My Learning Curve
;---------------------------------------------------------------------------------------------
; AutoHotkey
; https://autohotkey.com/
;---------------------------------------------------------------------------------------------

;---------------------------------------------------------------------------------------------
; TRY THIS TOOL LINE FOR WHEN STUCK DODGY CODE
; OPEN A COMMAND LINE WINDOW
;---------------------------------------------------------------------------------------------
; TASKKILL /IM AUTOHOTKEY.EXE
;---------------------------------------------------------------------------------------------
; KILLER ALL INSTANCES
;---------------------------------------------------------------------------------------------
;---------------------------------------------------------------------------------------------

;---------------------------------------------------------------------------------------------
;---------------------------------------------------------------------------------------------
; WORK FEW HOUR UNTIL RESULT 
;---------------------------------------------------------------------------------------------
; INITIAL ENTRY DATE FIRST DAY MISSING AT THE MOMENT
;---------------------------------------------------------------------------------------------
; Sat 29 October 2016 20:03:41---------- End
; A Run Here Between This Time In One Day
; Sat 29 October 2016 22:20:11---------- End 2nd
; Sat 29 October 2016 22:42:29---------- Finial Editing
; Sat 29 October 2016 23:15:00---------- Last Debug And Tidy -- 
;---------------------------------------------              Return Ending Inside The End Bracket
; Sat 29 October 2016 23:31:00---------- Done Finial Mouse Button Block 
;                                                       ---------- Hotkey Had To Be Moved Next To End Of Code
;---------------------------------------------------------------------------------------------
; Sun 30 October 2016 09:30:00---------- Task Canvas -- Quick Show I Been An Hour First Thing 
; Sun 30 October 2016 10:21:17---------- End Time -- Final Product
; Sun 30 October 2016 10:36:48---------- End Time -- Final Product Text Editor
; Sun 30 October 2016 11:12:31---------- End Time -- Final Product Text Editor
; Sun 30 October 2016 11:32:13---------- Notice One Minor Bugg Document With-in
; -------------------------------------------------------- Work Around Answer -- Do It Later ESCAPE Key Repeater
; Sun 30 October 2016 11:43:09---------- Final 

;---------------------------------------------------------------------------------------------
; ONE MORE
; LEFT CLICK TO STOP
; ------------------
; SUCCESS MINOR TYPO 
; FOUND PROBLEM CAPS LOCK IF YOU GOT IT ON IT MAKE A PROBLEM OF FLICK OFF ON BUT IF NOT OKAY
; WORKAROUND WOULD BE TURN IT OFF AT BEGIN
; OR LOOK ANOTHER WAY
; Mon 07 November 2016 02:05:00---------- SEARCH TIME
; Mon 07 November 2016 02:15:00----------
; REMINDER THE WAY MICROSOFT ARE WORKER WIN 10 -- LEFT CLICK HITT TO OFF SOMETHING
; FROM WORK TODAY NOTICE
;---------------------------------------------------------------------------------------------


; 11:54:07 - Day Light! 51.994%, Opposite = 48%
; 11:54:15 - Sunday 30th of October
; 11:54:22 - Sun Rise . 6 48 and 1 Seconds
; 11:54:26 - Sun Noon . 11 42 and 46 Seconds
; 11:54:29 - Sun Set . 4 36 and 44 Seconds
; 11:54:34 - Day Light! 52.071%, Opposite = 48%

;---------------------------------------------------------------------------------------------
; Sunday Bed Time 1 AM Up at 8:50 Work Begin 9:30
; Sunday Bed Was After Friday and Saturday All-Nighter Before and Productive Every --2 Hour-- Coder
;---------------------------------------------------------------------------------------------
;--------------------------------
; 3 Fault Fix Sunday
;--------------------------------
; 1.. A VAR Need Clear Empty To Reuse Again ---- O_Active_Title =
; 2..Mouse Wheel Idea I Had Was Not Going To Work  As User Need That When Stop
; 3..Messy But Okay -- As This Is The Route Method Most AutoHotKeys User Have
; --- The Code Has To Have Press ESC On The Context Menu On RButton When Click On
; --- The Code Sometime Code Does Happen Very Quick Not To Even See -- and Again The Code Sometime Code Does Do Little Flash
; 4.. After Them 3 -- And Number 3 -- This Programmer AutoHotKeys Is Not What I Expect As A Key Hooker ; Bastard 
; As I'm Used To In Visual Basic Key Hooking -- Example -- When I Want A Key And Use It I Can Decide That ; It Is Thrown Away After Use And Not Allow Further Use -- Maybe Or Maybe Not Option 
; All In -- It Is A Good Key Hooker As Visual Basic Can't Detect A Lot Like Print Screen And Joy Pad In The ; Normal Key Hooker and Mouse -- Has JoyPad Else Where But Not Mouse Anything YET
;----
; Another Thinker Note When Use a HotKey The Hooker Don't Pass it Back Further Use
; Example Here The CTRL Or ALT F4 Would Close Window
; But How do You Do It in a Loop -- Same Problem Programmer Are Send Escape on Context Menu
; Not Any Looper Example Demo -- Search on Google

;---------------------------------------------------------------------------------------------
; 4 Thing Left
;--------------------------------
; 1.. A Numeric HWND Handle to Detect Window Change Better than Text Method
; 2.. And Thinker TOT -- It Was -- KeyHooker Block a Key After USE
; 3.. There is Not a Stop Key Totally as Yet Only to Swap Window To Stop or Toggle Mouse
; 4.. Just When I Thought It Was Over Context Menu Repeat Escape Is Learn Minor Problem
;       Document Further in Code Area -- To Fix Later
;       Release With a Mini Minor Bugg -- Don't Say I Don't Give You Anything  -- At Least It Is Usable
;
;       Install AutoHotKey Paste the Code in NotePad Extension AHK and RUNNER AHK Compiler & Enjoy Yourself
;       Lovely Scroll Downer Scripter
;       Tested Thoroughly If Anything Not Working Remark Statement Syntax Error
;       or Any Not Load After EDIT Final And Working
;       Bouncing Ball Scroll Up and Down
;       Technology Pooper Space Cadet Especially On Grin Book FB
;---------------------------------------------------------------------------------------------
; Couldn't Resist RePlay Retry Repeat Return Revise ReLooker ReCoder Pick It Up Again 
; Sunday To Get It Done -- Enjoy Like It's a Pleasure
; Productivity Answer
;------------------------------------------------------------------------------------------------------------------
;------------------------------------------------------------------------------------------------------------------
; Next Project Working
;----------------------------------------------------------------------------------
; 1.. Killer of Process in Time Delay After It Has Launched -- to Use a Bit -- But And Not Stay
;----------------------------------------------------------------------------------
; 2.. a Process Changer -- Save Setting up of a Pro Version
;----------------------------------------------------------------------------------
; 3.. Another Project Already Done and Last To talk About
;----------------------------------------------------------------------------------
; Example for at Boot Up
; This Project Hard To Display than Example Because It Require 
; Big Chunk Code to Exit after One Run
; Here the Example
; Secretary Word Processor Want the Caps Lock Gone But I'm Suffer Opposite
;----------------------------------------------------------------------------------
; The 1 Page Looked up Sunday Today Left Open From Day Before
; Press Escape on Context Menu
;----------------------------------------------------------------------------------
; Just Disable right-click menu after finish my hotkey. plz... - Ask for Help - AutoHotkey Community
; https://autohotkey.com/board/topic/65705-just-disable-right-click-menu-after-finish-my-hotkey-plz/
;----------------------------------------------------------------------------------

;----------------------------------------------------------------------------------
; SetCapsLockState , On
; SetNumLockState , Off
; SetScrollLockState , Off
;----------------------------------------------------------------------------------
;----------------------------------------------------------------------------------

;---------------------------------------------------------------------------------------------
; ** \/\/\/ -- Code Begin Here After Header Remark Block -- \/\/\/ **
;---------------------------------------------------------------------------------------------
; Use this When Launch the Code From Click HITT Button Explorer I Will Reload Itself
; Also Use the Double Click on the Green ICON and it Happen Give Route Code a Trace Path of Source Code Actions
; to Debug It
;----------------------------------------------------------------------------------
#SingleInstance force
;----------------------------------------------------------------------------------

;----------------------------------------------------------------------------------
; Initialize Beep If Wanter
;----------------------------------------------------------------------------------
SoundBeep , 1800 , 400
;----------------------------------------------------------------------------------

^F4::
 {
 x = 0
 O_Active_title =
 SoundBeep , 1800 , 400

 ;---------------------------------------------------------------------------------------------
 ; Nightmare Zombie Land Working Out Double Hot Key OFF And ON
 ; Not Any Asleep Pattern Saturday And Friday Day
 ;---------------------------------------------------------------------------------------------
 ; Do From Keyboard F4 And Then Control Off On Same With Mouse Right Button
 ; With Flag Hooking -- F4_Key = 1
 ;---------------------------------------------------------------------------------------------
 
 F4_KEY_TOGGLE = True
    
 ;---------------------------------------------------------------------------------------------------------------------
 ; 1st Answer Found But Not Good Enough By Want Numeric Hwnd Instead
 ;---------------------------------------------------------------------------------------------------------------------
 ; Google Search -- Body Search
 ;---------------------------------------------------------------------------------------------------------------------
 ; Final Google Search -- Autohotkeys O_Active_Title = Active_Title
 ;---------------------------------------------------------------------------------------------------------------------
 ; <SOLVED> IfWinActive, ahk_class, Title (exact name) Co - Ask for Help - AutoHotkey Community
 ; https://autohotkey.com/board/topic/74614-solved-ifwinactive-ahk-class-title-exact-name-co/
 ;---------------------------------------------------------------------------------------------------------------------
 
 ;------------------------------------------
 ; Chrome With Active Find
 ;------------------------------------------
 ; Next Answer In This Stage Of 2nd But Not Totally
 ; Idea Is Wanting Numeric Of Hwnd Change
 ;------------------------------------------

 ;---------------------------------------------------------------------------------------------------------------
 ; ALL CONDITION LOOK THEN OKAY TO PROCEED WITH LOOP ROUTINE
 ;---------------------------------------------------------------------------------------------------------------

 Loop 
 {
 WinGetTitle, Active_Title, A
 
 ;-------------------------------------------------------------------------------------------   
 ;HOW TO SET AND READ A VARIABLE CLEAR VOID EMPTY
 ;----
 ;Check Variable isn't blank - Ask for Help - AutoHotkey Community
 ;https://autohotkey.com/board/topic/70709-check-variable-isnt-blank/
 ;-------------------------------------------------------------------------------------------   
 if !O_Active_Title 
 {
 O_Active_title := Active_Title
 }
 
 ;---------------------------------------------------------------------------------------------------------------------
 ;----
 ;Compare two variables as strings - Ask for Help - AutoHotkey Community
 ;https://autohotkey.com/board/topic/28987-compare-two-variables-as-strings/
 ;---------------------------------------------------------------------------------------------------------------------
 IF ( "Z" Active_Title = "Z" O_Active_Title)
 {
 }
 Else
 {
 SoundBeep , 1800 , 400
 Break
 }
 O_Active_title := Active_Title


 ;---------------------------------------------------------------------------------------------------------------------
 ; Mon 07 November 2016 02:05:00----------
 ; LEFT CLICK STOP
 ;---------------------------------------------------------------------------------------------------------------------
 IF GetKeyState("LButton", "P")
 {
 O_Active_title := 
 SoundBeep
 Break
 }

 
 x += 1
 
 IF (F4_KEY_TOGGLE = "True")
 {
 ;Send,{PgDn} 
 ;-------------------- Remark Out or Use One or the Other
 Send,{Down}
 }
 
 Sleep 15
 ;---------------------------------------------------------------------------------------------------------------------------------
 ; Key Send By The Page Load Or Down Arrow And Wait A Bit Slower Down
 ;---------------------------------------------------------------------------------------------------------------------------------

 IfWinNotActive ahk_class Chrome_WidgetWin_1
 {
 SoundBeep
 Break
 ;--------------------------------------------------
 ;BREAK ARE FOR ONLY IN LOOP
 ;--------------------------------------------------
 }
 
 ;----------------------
 ;P FOR PRESS 
 ;----------------------
 if GetKeyState("RButton", "P")
 {
 SoundBeep , 1800 , 400
 
 ;-------------------------------------------------------------------------------
 ;---- Notice The Line Here -- Look Wonderful Semi-Difficult To Toggle 
 ;---- A Variable OFF And ON
 ;---- Might Be A Better Way If It Was Numeric -- People Want It in TEXT
 ;-------------------------------------------------------------------------------

 F4_KEY_TOGGLE := F4_KEY_TOGGLE = "true" ? "False" : "True"
 
 ;-------------------------------------------------------------------------------
 ;How work toggle (booleans) vars? - Ask for Help - AutoHotkey Community
 ;https://autohotkey.com/board/topic/67783-how-work-toggle-booleans-vars/
 ;MyVar := Myvar = "true" ? "False" : "True"
 ;-----------------------------------------------
 ;DEBUG TESTING AND REDUNDANT
 ;Msgbox %F4_KEY_TOGGLE%
 ;-----------------------------------------------
 }
 
 ;------------------------------------------------------------------------------------------------------------------
 ; Here Is the Same Check Of RButton -- But Hash In Front Mean Not Sure
 ; The Order Here Is After Or RButton Wait Mean First Button Suffering
 ;------------------------------------------------------------------------------------------------------------------
 ; Big Problem Here Source Tracer Show Escape Is Being Press All The Time
 ; Does Make Some Task Bar Item Flash With It
 ; But Acceptable Does Exit at Change Window
 ; Remove Hash # and Context Menu Won't Work Barges It Way In
 ;------------------------------------------------------------------------------------------------------------------
 
 #If GetKeyState("RButton", "p")
 {
 KeyWait, Rbutton ;As soon as RButton is released...
 SendInput {Esc}  ;... kill the context menu
 }


 
 ; 4 Million Key Down and Then Halt at Some Millisecond Interval
 if (x > 4000000)
 {
 SoundBeep , 1800 , 400
 Break
 }
 }
 } 

Return

;-------------------------------------------------------------------------------------------   
; Ender Code
;-------------------------------------------------------------------------------------------   
;-------------------------------------------------------------------------------------------   
;-------------------------------------------------------------------------------------------




; INTRO -- ON-LINE MEDIA
; --------------------------------------
; Hi Room
; Here Is the Second 2nd Version And Working Project 
; After That Code All Morning Sunday 2 Hour Worth
; I Able Have a Walk Around The Room Now
; All Is Explained
; Also On Blogger 28 Hitt Counter in a Few Moment
; Especially as I Put on Google+ Plus Also
; ----
; Roids Polaroids Mach 2: AutoHotKeys Project Before Mach 2 Page Scroller Down
; http://roidsrim-minimal.blogspot.co.uk/2016/10/autohotkeys-project-before-mach-2-page.html
; --------
; Hi Room Here Is the Second 2nd Version And... - Matthew Lancaster
; https://www.facebook.com/matthew.lancaster.4/posts/10206629081999481
; ----

; EDIT -- Grammar Check 100% Minor Change
; Sun 30 October 2016 12:17:05----------
; 12:17:19 - Sun Rise . 6 48 and 1 Seconds
; 12:17:21 - Day Light! 55.941%, Opposite = 44%

; Good Afternoon Sunday
; Space Book Grin Book Land FB

; EDIT -- 21 Minute In 
; 11:54:07 - Day Light! 51.994%, Opposite = 48%
; 11:54:15 - Sunday 30th of October
; 11:54:22 - Sun Rise . 6 48 and 1 Seconds
; 11:54:26 - Sun Noon . 11 42 and 46 Seconds
; 11:54:29 - Sun Set . 4 36 and 44 Seconds
; 11:54:34 - Day Light! 52.071%, Opposite = 48%

; EDIT -- Grammar Check 100% Minor Change
; Sun 30 October 2016 12:17:05----------
; 12:17:19 - Sun Rise . 6 48 and 1 Seconds
; 12:17:21 - Day Light! 55.941%, Opposite = 44%

; EDIT -- and 1 Hour 10 Minute in Left Intention Word-in Away
; Sun 30 October 2016 13:06:50----------
; 13:00:04 - Sunday 30th of October
; 13:07:12 - Sun Rise . 6 48 and 1 Seconds
; 13:07:16 - Day Light! 64.420%, Opposite = 36%
; 13:07:22 - Sun Noon . 11 42 and 46 Seconds
; 13:07:27 - Sun Set . 4 36 and 44 Seconds

; ---- Left Intention Word-in Away -- Hard Coder Remark
; EDIT -- Text 11:55a TO 01:35p -- 1 Hr 40 Minute in
; Sun 30 October 2016 13:39:45----------
; 13:40:00 - TIME #01:40
; 13:40:06 - Day Light! 69.997%, Opposite = 30%
; 13:40:19 - 01:40 And 19 Seconds
; 13:40:26 - Day Light! 70.054%, Opposite = 30%

; 13:49:46 - Day Light! 71.639%, Opposite = 28%
; 13:49:51 - Day Light! 71.653%, Opposite = 28%

; ---------------------------------------
; CODE BEGIN COPY PASTER
; ---------------------------------------
; INTRO END ON-LINE MEDIA
; ---------------------------------------



;-------------------------------------------------------------------------------
;IDEA STRAIN NOTE TO SELF AND ACCOMPLISHED
;-------------------------------------------------------------------------------
;WANT TO REPEATER RIGHT CLICK INSTEAD F4 TO BEGIN UNLESS ANOTHER LIKE LButton
;if GetKeyState("LButton", "P")
;{
;	SoundBeep
;	break
;	F4_KEY = 1
;}

;if (GetKeyState("RButton") & F4_KEY = 1)  or (GetKeyState("F4")  and GetKeyState("ctrl"))


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


;-------------------------------------------------------
;UN-USED CODE 
;-------------------------------------------------------
;----
;How to temporarly disable mouse buttons - AutoHotkey Community
;https://autohotkey.com/boards/viewtopic.php?t=21358
;----
				;;;#if brm
				;;;$RButton::
				;;;MButton::
				;;;return
;-------------------------------------------------------
   
   
;-------------------------------------------------------------------------
;REPLACE F10 TO DO CONTROL PRINT SCREEN
;FOR CLIPBOARD SCREEN SHOT -- CODE PIC___ WON'T HOT KEY F0 ON WIN 7 AND UP
;-------------------------------------------------------------------------
;10::
;SEND ^{PrintScreen}

;Return

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

