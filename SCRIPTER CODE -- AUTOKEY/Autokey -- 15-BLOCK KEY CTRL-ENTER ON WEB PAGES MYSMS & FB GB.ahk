;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES MYSMS & FB GB.ahk
;# __ 
;# __ Autokey -- 15-BLOCK KEY CTRL-ENTER ON WEB PAGES MYSMS & FB GB.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# Wed 18 October 2017 02:05:00 Work Time Beginner
;# Sat 21 April   2018 04:00:00 2nd Wave of Work Better Coder Emerged
;# __ 


; Published on Blogger
; -------------------------------------------------------------------
; Search Description
; -------------------------------------------------------------------
; Code to Change Shift-Enter Control Enter and Enter for Use when Sending Messenger and Not Wanting to Send on The Enter Key or By Accident Control Enter For Use with MYSMS and GrinBook FB And Any Other May Find Later
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
; https://www.dropbox.com/s/xdjfywh4u5vos0y/Autokey%20--%2015-BLOCK%20KEY%20CTRL-ENTER%20ON%20WEB%20PAGES%20MYSMS%20%26%20FB%20GB.ahk?dl=0
;# ------------------------------------------------------------------



; -------------------------------------------------------------------
; Here is the Code in Autohotkeys to Change Shift-Enter Control Enter and Enter for Use when Sending Messenger and Not Wanting to Send on entering or By Accident Control Enter
; For Use with MYSMS and GrinBook FB And Any Other May Find Later
; My New Improved Style of Coding Read a Bit More in Rem Line
; Took 4 Hours Work Additional on Code Before the Last Attempt Early Hour Until Dawn
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
; To Enter Key And Vice Versa Depending On If In The Mysms Program That I Use 
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
; Not Just Isolated To GrinBook FB And Mysms With Little Extra Code 
; You Could Disable Beeper Unless In FB And Mysms
; Might Be Helpful To Remove The Beeper There As It Less Quicker 
; A Lot The Repeat Return Key Compared To Other Keys In Editing 
; Besides Mysms And Grin Book
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

; GLOBAL SETTINGS ===================================================
#Warn
#NoEnv
#SingleInstance Force
#Persistent
;IT USER ExitFunc TO EXIT FROM #Persistent
;--------------------

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; SCRIPT ============================================================

DetectHiddenWindows, on
SetStoreCapslockMode, off
Set_Key_FB_Control=False 
Dont_Send_2_Enter=False
Sound_Speed=20

;User Setting as Required
Mute_Beep_In_Other_Program_Beside_GrinBook_an_Mysms=true
Mute_Beep_In_Other_Program_Beside_GrinBook_an_Mysms=false

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
    if (WinActive("Matthew Lancaster -") and WinActive("ahk_class Chrome_WidgetWin_1"))
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
        if (WinActive("Matthew Lancaster -") and WinActive("ahk_class Chrome_WidgetWin_1"))
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
;# ------------------------------------------------------------------
; exit the app

