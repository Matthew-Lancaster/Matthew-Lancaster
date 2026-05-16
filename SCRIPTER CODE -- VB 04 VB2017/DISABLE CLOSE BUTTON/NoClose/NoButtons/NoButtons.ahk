;NoButtons.ahk
; Hold the mouse still for one second to show a window with clicking actions.
;Skrommel @2006

#SingleInstance,Force
CoordMode,Mouse,Screen
SetWinDelay,0

applicationname=NoButtons

Gosub,INIREAD
Gosub,BUILDGUI

showgui=0
hidden=0
action=LEFT
leftdown=0
rightdown=0
counter=0
repeats=0

Loop
{
  Sleep,50
  x1=%x2%
  y1=%y2%
  MouseGetPos,x2,y2,id,ctrl

  If (x2=x1 And y2=y1)
    counter+=1
  Else
    counter=0
  
  If hidden=1
  {
    If (x2<>x1 Or y2<>y1)
      Gosub,GUISHOW
    Continue
  }

  If counter>20
  {
    counter=0
    repeats+=1
    If repeats>0
      Gosub,GUISHOW
    Else
      Gosub,%action%
  }

  If showgui=1
  If id<>%guiid%
  {
    counter=0
    repeats=0
    Gosub,GUIHIDE
  }
}
Return


GUISHOW:
If showgui=1
{
  x:=Ceil((x2-guix)/buttonw)
  y:=Ceil((y2-guiy)/buttonh)
  IniRead,action,%applicationname%.ini,%y%-%x%,action
  IniRead,parameters,%applicationname%.ini,%y%-%x%,parameters
  IniRead,hide,%applicationname%.ini,%y%-%x%,hide
  If action<>Error
  If action<>
    Gosub,%action%
  counter=0
  repeats=0
  Return
}
hidden=0
counter=0
repeats=0
showgui=1
MouseGetPos,x3,y3,id,ctrl
guix:=x3-(defaultw-defaultx)/2
guiy:=y3-(defaulth-defaulty)/2
If guix<0
  guix=0
If guiy<0
  guiy=0
If (guix+guiw>A_ScreenWidth)
  guix:=A_ScreenWidth-guiw
If (guiy+guih>A_ScreenHeight)
  guiy:=A_ScreenHeight-guih
WinSet,TopMost,,ahk_id %guiid%
WinSet,AlwaysOnTop,,ahk_id %guiid%
WinMove,ahk_id %guiid%,,%guix%,%guiy%
MouseMove,% guix+defaultx+buttonw/2,% guiy+defaulty+buttonh/2,0
id=%guiid%
Return


GUIHIDE:
showgui=0
WinMove,ahk_id %guiid%,,% -buttonw*buttonsw,% -buttonh*buttonsh
Return


REST:
Return

HIDE:
hidden=1
Gosub,GUIHIDE
MouseMove,%x3%,%y3%,0
MouseGetPos,x2,y2
Return

SETTINGS:
Gosub,GUIHIDE
Run,%applicationname%.ini
Return

LEFT:
Gosub,GUIHIDE
MouseClick,Left,%x3%,%y3%,1,0
leftdown=0
Return

RIGHT:
Gosub,GUIHIDE
MouseClick,Right,%x3%,%y3%,1,0
rightdown=0
Return

DOUBLELEFT:
Gosub,GUIHIDE
MouseClick,Left,%x3%,%y3%,2,0
leftdown=0
Return

DOUBLERIGHT:
Gosub,GUIHIDE
MouseClick,Right,%x3%,%y3%,2,0
rightdown=0
Return

DRAGLEFT:
Gosub,GUIHIDE
If leftdown=0
{
  MouseClick,Left,%x3%,%y3%,1,0,D
  leftdown=1
}
Else
{
  MouseClick,Left,%x3%,%y3%,1,0,U
  leftdown=0
}
Return

DRAGRIGHT:
Gosub,GUIHIDE
If rightdown=0
{
  MouseClick,Right,%x3%,%y3%,1,0,D
  rightdown=1
}
Else
{
  MouseClick,Right,%x3%,%y3%,1,0,U
  rightdown=0
}
Return

SHIFTLEFT:
Gosub,GUIHIDE
Send,{Shift Down}
MouseClick,Left,%x3%,%y3%,1,0
Send,{Shift Up}
leftdown=0
Return

SHIFTRIGHT:
Gosub,GUIHIDE
Send,{Shift Down}
MouseClick,Right,%x3%,%y3%,1,0
Send,{Shift Up}
rightdown=0
Return

SEND:
Send,%parameters%
If hide=1
  Gosub,GUIHIDE
Return


RUN:
Run,%parameters% ;,,UseErrorLevel
If hide=1
  Gosub,GUIHIDE
Return


INIREAD:
IfNotExist,%applicationname%.ini
{
  IniWrite,5,%applicationname%.ini,Settings,buttonsw
  IniWrite,3,%applicationname%.ini,Settings,buttonsh
  IniWrite,50,%applicationname%.ini,Settings,buttonw
  IniWrite,50,%applicationname%.ini,Settings,buttonh
  IniWrite,Left,%applicationname%.ini,Settings,default

  IniWrite,Left,%applicationname%.ini,1-1,name
  IniWrite,LEFT,%applicationname%.ini,1-1,action
  IniWrite,Double Left,%applicationname%.ini,1-2,name
  IniWrite,DOUBLELEFT,%applicationname%.ini,1-2,action
  IniWrite,Right,%applicationname%.ini,1-3,name
  IniWrite,RIGHT,%applicationname%.ini,1-3,action
  IniWrite,Toggle Left,%applicationname%.ini,2-1,name
  IniWrite,DRAGLEFT,%applicationname%.ini,2-1,action
  IniWrite,+,%applicationname%.ini,2-2,name
  IniWrite,REST,%applicationname%.ini,2-2,action
  IniWrite,Toggle Right,%applicationname%.ini,2-3,name
  IniWrite,DRAGRIGHT,%applicationname%.ini,2-3,action
  IniWrite,Shift Left,%applicationname%.ini,3-1,name
  IniWrite,SHIFTLEFT,%applicationname%.ini,3-1,action
  IniWrite,Settings,%applicationname%.ini,3-2,name
  IniWrite,SETTINGS,%applicationname%.ini,3-2,action
  IniWrite,Hide,%applicationname%.ini,3-3,name
  IniWrite,HIDE,%applicationname%.ini,3-3,action

  IniWrite,Copy,%applicationname%.ini,1-4,name
  IniWrite,SEND,%applicationname%.ini,1-4,action
  IniWrite,^c,%applicationname%.ini,1-4,parameters
  IniWrite,1,%applicationname%.ini,1-4,hide
  
  IniWrite,Cut,%applicationname%.ini,2-4,name
  IniWrite,SEND,%applicationname%.ini,2-4,action
  IniWrite,^x,%applicationname%.ini,2-4,parameters
  IniWrite,1,%applicationname%.ini,2-4,hide

  IniWrite,Paste,%applicationname%.ini,3-4,name
  IniWrite,SEND,%applicationname%.ini,3-4,action
  IniWrite,^v,%applicationname%.ini,3-4,parameters
  IniWrite,1,%applicationname%.ini,3-4,hide

  IniWrite,Keys,%applicationname%.ini,1-5,name
  IniWrite,RUN,%applicationname%.ini,1-5,action
  IniWrite,%WinDir%\System32\osk.exe,%applicationname%.ini,1-5,parameters
  IniWrite,1,%applicationname%.ini,1-5,hide
  
  IniWrite,Notepad,%applicationname%.ini,2-5,name
  IniWrite,RUN,%applicationname%.ini,2-5,action
  IniWrite,Notepad.exe,%applicationname%.ini,2-5,parameters
  IniWrite,1,%applicationname%.ini,2-5,hide

  IniWrite,Calc,%applicationname%.ini,3-5,name
  IniWrite,RUN,%applicationname%.ini,3-5,action
  IniWrite,Calc.exe,%applicationname%.ini,3-5,parameters
  IniWrite,1,%applicationname%.ini,3-5,hide
}

IniRead,buttonsw,%applicationname%.ini,Settings,buttonsw
IniRead,buttonsh,%applicationname%.ini,Settings,buttonsh
IniRead,buttonw,%applicationname%.ini,Settings,buttonw
IniRead,buttonh,%applicationname%.ini,Settings,buttonh
IniRead,default,%applicationname%.ini,Settings,default
Return


BUILDGUI:
Gui,+AlwaysOnTop +ToolWindow +Border -Caption
Gui,Margin,0,0
Loop,%buttonsh%
{
  y:=A_Index
  Loop,%buttonsw%
  {
    x:=A_Index
    IniRead,name,%applicationname%.ini,%y%-%x%,name
    If name=Error
      name=
    Else
    {
      IniRead,action,%applicationname%.ini,%y%-%x%,action
      If action=Error
        action=
    }
    buttonx:=x*buttonw-buttonw
    buttony:=y*buttonh-buttonh
    Gui,Add,Button,x%buttonx% y%buttony% w%buttonw% h%buttonh% G%action%,%name%
  }
}
Gui,Show,x0 y0 NoActivate,NoButtonsGUI
WinGet,guiid,ID,NoButtonsGUI
;WinSet,Transparent,150,ahk_id %guiid%
WinGetPos,guix,guiy,guiw,guih,ahk_id %guiid%
ControlGetPos,defaultx,defaulty,defaultw,defaulth,%default%,ahk_id %guiid%
Return