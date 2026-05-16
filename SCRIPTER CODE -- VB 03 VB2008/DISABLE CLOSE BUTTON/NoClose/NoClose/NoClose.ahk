
;NoClose.ahk
; Disable the Close button (X) of selected windows
; To run, save to BlockInput.ahk and install AutoHotkey from www.autohotkey.com
;Skrommel @2006

#SingleInstance,Force
SetTitleMatchMode,2

applicationname=NoClose

ids=
oldids=

Gosub,INIREAD
Gosub,TRAYMENU
Gosub,STARTUP
OnExit,EXIT
Hotkey,%add%,ADD
Hotkey,%swap%,SWAP

Loop
{
  Sleep,500
  allids=
  activeids=
  WinGet,id_,List,,,Program Manager
  Loop,%id_%
  {
    Sleep,0
    id:=id_%A_Index%
    allids=%allids%%id%`,
    IfInString,ids,%id%`,
      activeids=%activeids%%id%`,
    If autodisable=0
      Continue
    IfInString,oldids,%id%`,
      Continue
    WinGetTitle,title,ahk_id %id%
    WinGetClass,class,ahk_id %id%
    rule=%title% ahk_class %class%|||
    IfInString,rules,%rule%
    {
      DISABLE(id)
      activeids=%activeids%%id%`,
    }
  }
  oldids:=allids
  ids:=activeids
}
Return


STARTUP:
allids=
WinGet,id_,List,,,Program Manager
Loop,%id_%
{
  id:=id_%A_Index%
  allids=%allids%%id%`,
  If disableonstartup=0
    Continue
  WinGetTitle,title,ahk_id %id%
  WinGetClass,class,ahk_id %id%
  rule=%title% ahk_class %class%|||
  IfInString,rules,%rule%
  {
    DISABLE(id)
    ids=%ids%%id%`,
  }
}
oldids:=allids
Return


EXIT:
If enableonexit=0
  ExitApp
WinGet,id_,List,,,Program Manager
Loop,%id_%
{
  id:=id_%A_Index%
  IfInString,ids,%id%`,
  {
    ENABLE(id)
    StringReplace,ids,ids,%id%`,,
  }
}
ExitApp


ADD:
WinGet,id,ID,A
WinGetTitle,title,ahk_id %id%
WinGetClass,class,ahk_id %id%
rule=%title% ahk_class %class%|||
IfInString,rules,%rule%
  Return
Else
{
  DISABLE(id)
  rules=%rules%%rule%
  ids=%ids%%id%`,
  IniWrite,%rules%,%applicationname%.ini,Settings,rules
}
Return


SWAP:
WinGet,id,ID,A
WinGetTitle,title,ahk_id %id%
WinGetClass,class,ahk_id %id%
IfInString,ids,%id%`,
{
  ENABLE(id)
  StringReplace,ids,ids,%id%`,,
  Return
}
DISABLE(id)
ids=%ids%%id%`,
Return


DISABLE(id) ;By RealityRipple at http://www.xtremevbtalk.com/archive/index.php/t-258725.html
{
  menu:=DllCall("user32\GetSystemMenu","UInt",id,"UInt",0)
  DllCall("user32\DeleteMenu","UInt",menu,"UInt",0xF060,"UInt",0x0)
  WinGetPos,x,y,w,h,ahk_id %id%
  WinMove,ahk_id %id%,,%x%,%y%,%w%,% h-1
  WinMove,ahk_id %id%,,%x%,%y%,%w%,% h+1
}


ENABLE(id) ;By Mosaic1 at http://www.xtremevbtalk.com/archive/index.php/t-258725.html
{
  menu:=DllCall("user32\GetSystemMenu","UInt",id,"UInt",1)
  DllCall("user32\DrawMenuBar","UInt",id)
}


TRAYMENU:
Menu,Tray,NoStandard 
Menu,Tray,DeleteAll 
Menu,Tray,Add,%applicationname%,ABOUT
Menu,Tray,Add,
Menu,Tray,Add,&Settings...,SETTINGS
Menu,Tray,Add,&About...,ABOUT
Menu,Tray,Add,E&xit,EXIT
Menu,Tray,Default,%applicationname%
Menu,Tray,Tip,%applicationname%
Return


INIREAD:
IfNotExist,%applicationname%.ini
{
  disableonstartup=1
  autodisable=1
  enableonexit=1
  swap=^1
  add=^2
  rules=
  Gosub,INIWRITE
  Gosub,ABOUT
}
IniRead,disableonstartup,%applicationname%.ini,Settings,disableonstartup
IniRead,autodisable,%applicationname%.ini,Settings,autodisable
IniRead,enableonexit,%applicationname%.ini,Settings,enableonexit
IniRead,swap,%applicationname%.ini,Settings,swap
IniRead,add,%applicationname%.ini,Settings,add
IniRead,rules,%applicationname%.ini,Settings,rules
Return


INIWRITE:
IniWrite,%disableonstartup%,%applicationname%.ini,Settings,disableonstartup
IniWrite,%autodisable%,%applicationname%.ini,Settings,autodisable
IniWrite,%enableonexit%,%applicationname%.ini,Settings,enableonexit
IniWrite,%swap%,%applicationname%.ini,Settings,swap
IniWrite,%add%,%applicationname%.ini,Settings,add
IniWrite,%rules%,%applicationname%.ini,Settings,rules
Return


SETTINGS:
HotKey,%swap%,Off
HotKey,%add%,Off
Gui,Destroy
Gui,Add,Tab,W340 H330 xm,Options|Rules
Gui,Tab,1
Gui,Add,GroupBox,xm+10 ym+40 w320 h70,&Hotkey to Enable/Disable the active windows' close button
Gui,Add,Hotkey,xp+10 yp+20 w300 vsswap
StringReplace,current,swap,+,Shift +%A_Space%
StringReplace,current,current,^,Ctrl +%A_Space%
StringReplace,current,current,!,Alt +%A_Space%
Gui,Add,Text,xm+20 y+5,Current hotkey: %current%

Gui,Add,GroupBox,xm+10 y+30 w320 h70,Hotkey to &Add a new rule
Gui,Add,Hotkey,xm+20 yp+20 w300 vsadd
StringReplace,current,add,+,Shift +%A_Space%
StringReplace,current,current,^,Ctrl +%A_Space%
StringReplace,current,current,!,Alt +%A_Space%
Gui,Add,Text,xm+20 y+5,Current hotkey: %current%

Gui,Add,GroupBox,xm+10 y+30 w320 h80,Automatic rule execution
Gui,Add,CheckBox,xm+20 yp+20 Checked%disableonstartup% vsdisableonstartup,Disable close buttons on NoClose &Startup
Gui,Add,CheckBox,xm+20 y+5 Checked%autodisable% vsautodisable,Disable close buttons on &Window Creation
Gui,Add,CheckBox,xm+20 y+5 Checked%enableonexit% vsenableonexit,Enable close buttons on NoClose &Exit

Gui,Tab,2
StringReplace,rules,rules,|||,`n,All
Gui,Add,GroupBox,w320 h280 xm+10 y+10,&Windows Titles and Classes
Gui,Add,Edit,xm+20 yp+20 w300 h180 Multi -Wrap vsrules,%rules%
Gui,Add,Text,xm+20 y+5,Syntax: <Part of a Window Title> <ahk_class Class Name>
Gui,Add,Text,xm+20 y+5,Example: Calculator ahk_class SciCalc   
Gui,Add,Text,xm+30 y+5,will disable all Calculator close buttons.
Gui,Add,Text,xm+20 y+5,Either part is optional.

Gui,Tab
Gui,Add,Button,xm+10 y+30 w75 GSETTINGSOK,&OK
Gui,Add,Button,x+5 w75 GSETTINGSCANCEL,&Cancel
Gui,Show,,%applicationname% Settings
Return

SETTINGSOK:
Gui,Submit
If sswap<>
{
  swap:=sswap
  HotKey,%swap%,SWAP
}
HotKey,%swap%,On
If sadd<>
{
  add:=sadd
  HotKey,%add%,ADD
}
HotKey,%add%,On
If sdelay<>
  delay:=sdelay
StringReplace,rules,srules,`n,|||,All
rules=%rules%|||
Loop
{
  StringReplace,rules,rules,||||||,|||,All
  StringGetPos,pos,rules,||||||
  If pos<0
    Break
}
StringLeft,start,rules,3
If start=|||
  StringTrimLeft,rules,rules,3
disableonstartup:=sdisableonstartup
autodisable:=sautodisable
enableonexit:=senableonexit
Gosub,INIWRITE
Return

SETTINGSCANCEL:
HotKey,%swap%,SWAP
HotKey,%swap%,On
HotKey,%add%,ADD
HotKey,%add%,On
Gui,Destroy
Return


ABOUT:
Gui,99:Destroy
Gui,99:Margin,20,20
Gui,99:Add,Picture,xm Icon1,%applicationname%.exe
Gui,99:Font,Bold
Gui,99:Add,Text,x+10 yp+10,%applicationname% v1.1
Gui,99:Font
Gui,99:Add,Text,y+10,Disable the Close button (X) of selected windows.
Gui,99:Add,Text,y+10,- Press Ctrl+1 to Enable or Disable a close button.
Gui,99:Add,Text,y+5 ,- Press Ctrl+2 to Add a rule.
Gui,99:Add,Text,y+10,- To change the settings, choose Settings in the tray menu.

Gui,99:Add,Picture,xm y+20 Icon5,%applicationname%.exe
Gui,99:Font,Bold
Gui,99:Add,Text,x+10 yp+10,1 Hour Software by Skrommel
Gui,99:Font
Gui,99:Add,Text,y+10,For more tools, information and donations, please visit 
Gui,99:Font,CBlue Underline
Gui,99:Add,Text,y+5 G1HOURSOFTWARE,www.1HourSoftware.com
Gui,99:Font

Gui,99:Add,Picture,xm y+20 Icon7,%applicationname%.exe
Gui,99:Font,Bold
Gui,99:Add,Text,x+10 yp+10,DonationCoder
Gui,99:Font
Gui,99:Add,Text,y+10,Please support the contributors at
Gui,99:Font,CBlue Underline
Gui,99:Add,Text,y+5 GDONATIONCODER,www.DonationCoder.com
Gui,99:Font

Gui,99:Add,Picture,xm y+20 Icon6,%applicationname%.exe
Gui,99:Font,Bold
Gui,99:Add,Text,x+10 yp+10,AutoHotkey
Gui,99:Font
Gui,99:Add,Text,y+10,This tool was made using the powerful
Gui,99:Font,CBlue Underline
Gui,99:Add,Text,y+5 GAUTOHOTKEY,www.AutoHotkey.com
Gui,99:Font

Gui,99:Show,,%applicationname% About
hCurs:=DllCall("LoadCursor","UInt",NULL,"Int",32649,"UInt") ;IDC_HAND
OnMessage(0x200,"WM_MOUSEMOVE") 
Return

1HOURSOFTWARE:
  Run,http://www.1hoursoftware.com,,UseErrorLevel
Return

DONATIONCODER:
  Run,http://www.donationcoder.com,,UseErrorLevel
Return

AUTOHOTKEY:
  Run,http://www.autohotkey.com,,UseErrorLevel
Return

99GuiClose:
  Gui,99:Destroy
  OnMessage(0x200,"")
  DllCall("DestroyCursor","Uint",hCur)
Return

WM_MOUSEMOVE(wParam,lParam)
{
  Global hCurs
  MouseGetPos,,,,ctrl
  If ctrl in Static10,Static14,Static18
    DllCall("SetCursor","UInt",hCurs)
  Return
}
Return
