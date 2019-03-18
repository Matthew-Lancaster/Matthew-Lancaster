;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 37-Block_All_Keys.ahk
;# __ 
;# __ Autokey -- 37-Block_All_Keys.ahk
;# __ 
;# Copied BY Matthew __ Matt.Lan@Btinternet.com
;# __ 
;  =============================================================

;---------------------------------------------------------------
; DESCRIPTION
;---------------------------------------------------------------
; Block All the Keys for Keyboard Cleaning
;---------------------------------------------------------------

;# ------------------------------------------------------------------
; Location Internet
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
; 
;# ------------------------------------------------------------------

; 001
; --------------------------------------------------------------
; FROM  Wed 18-Jul-2018 17:36:24
; TO    Wed 18-Jul-2018 17:48:00 __ Copied from On-Line Searcher
; --------------------------------------------------------------

; --------------------------------------------------------------
; Credit Source
; --------------------------------------------------------------
;Disable all keyboard buttons - AutoHotkey Community
;https://autohotkey.com/boards/viewtopic.php?t=33925
; --------------------------------------------------------------


hk(keyboard:=0, mouse:=0, message:="", timeout:=3, displayonce:=0) { 
   static AllKeys, z, d
   z:=message
   d:=displayonce
      For k,v in AllKeys {
           Hotkey, *%v%, Block_Input, off         ; initialisation
      }
   if !AllKeys {
      s := "||NumpadEnter|Home|End|PgUp|PgDn|Left|Right|Up|Down|Del|Ins|"
      Loop, 254
         k := GetKeyName(Format("VK{:0X}", A_Index))
       , s .= InStr(s, "|" k "|") ? "" : k "|"
      For k,v in {Control:"Ctrl",Escape:"Esc"}
         AllKeys := StrReplace(s, k, v)
      AllKeys := StrSplit(Trim(AllKeys, "|"), "|")
   }
   ;------------------
   if (mouse!=2)  ; if mouse=1 disable right and left mouse buttons  if mouse=0 don't disable mouse buttons
    {
        For k,v in AllKeys {
           IsMouseButton := Instr(v, "Wheel") || Instr(v, "Button")
           Hotkey, *%v%, Block_Input, % (keyboard && !IsMouseButton) || (mouse && IsMouseButton) ? "On" : "Off"
        }
    }
   if (mouse=2)   ;disable right mouse button (but not left mouse)
    {                
     ExcludeKeys:="LButton"
      For k,v in AllKeys {
           IsMouseButton := Instr(v, "Wheel") || Instr(v, "Button")
           if v not in %ExcludeKeys%
           Hotkey, *%v%, Block_Input, % (keyboard && !IsMouseButton) || (mouse && IsMouseButton) ? "On" : "Off"
        }
    }
   if d
    {
   if (z != "") {
      Progress, B1 ZH0, %z%
      SetTimer, TimeoutTimer, % -timeout*1000
   }
   else
      Progress, Off
     }
   Block_Input:
   if (d!=1)
    {
   if (z != "") {
      Progress, B1 ZH0, %z%
      SetTimer, TimeoutTimer, % -timeout*1000
   }
   else
      Progress, Off
     }
   Return
   TimeoutTimer:
   Progress, Off
   Return
}

!F1::hk(1,1,"Keyboard keys and mouse buttons disabled!`nPress Alt+F2 to enable")   ; Disable all keyboard keys and mouse buttons
!F2::hk(0,0,"Keyboard keys and mouse buttons restored!")         ; Enable all keyboard keys and mouse buttons
!F3::hk(1,0,"Keyboard keys disabled!`nPress Alt+F2 to enable")   ; Disable all keyboard keys (but not mouse buttons)
!F4::hk(0,1,"Mouse buttons disabled!`nPress Alt+F2 to enable")   ; Disable all mouse buttons (but not keyboard keys)
!F5::hk(1,2,"Keyboard keys and Right Mouse button disabled!`nPress Alt+F2 to enable")   ; Disable all keyboard keys and right mouse button only
!F6::hk(0,2,"Right Mouse button disabled!`nPress Alt+F2 to enable",3,1)   ; Disable right mouse button only and show message only once