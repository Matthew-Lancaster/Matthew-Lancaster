/*
 * Copyright (c) 2004-2005 by Enovatic-Solutions. All rights reserved.
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
 * ----------------------------------------------------------------------
 * If you have any suggestions of new features or further questions 
 * feel free to contact the author.
 *
 * Company:  Enovatic-Solutions (IT Service Provider)
 * Author:   Oliver Pfeiffer, Bremen (GERMANY)
 * Homepage: http://www.enovatic.org/
 * Email:    contact@enovatic.org
 */



; NiftyWindows Version 0.9.3.1
; http://www.enovatic.org/products/niftywindows/

; AutoHotkey Version 1.0.36.01
; http://www.autohotkey.com/



#SingleInstance force
#HotkeyInterval 1000
#MaxHotkeysPerInterval 100
#NoTrayIcon



; [SYS] autostart section

SplitPath, A_ScriptFullPath, SYS_ScriptNameExt, SYS_ScriptDir, SYS_ScriptExt, SYS_ScriptNameNoExt, SYS_ScriptDrive
SYS_ScriptVersion = 0.9.3.1
SYS_ScriptBuild = 20050702195845
SYS_ScriptInfo = %SYS_ScriptNameNoExt% %SYS_ScriptVersion%

Process, Priority, , HIGH
SetBatchLines, -1
; TODO : a nulled key delay may produce problems for WinAmp control
SetKeyDelay, 0, 0
SetMouseDelay, 0
SetDefaultMouseSpeed, 0
SetWinDelay, 0
SetControlDelay, 0

Gosub, SYS_ParseCommandLine
Gosub, CFG_LoadSettings
Gosub, CFG_ApplySettings

MIR_MirandaFullPath = %ProgramFiles%\Miranda\Miranda32.exe
SplitPath, MIR_MirandaFullPath, , MIR_MirandaDir

if ( !A_IsCompiled )
	SetTimer, REL_ScriptReload, 1000

OnExit, SYS_ExitHandler

Gosub, TRY_TrayInit
Gosub, SYS_ContextCheck
Gosub, UPD_AutoCheckForUpdate

Return



; [SYS] parses command line parameters

SYS_ParseCommandLine:
	Loop %0%
		If ( (%A_Index% = "/x") or (%A_Index% = "/exit") )
			ExitApp
Return



; [SYS] exit handler

SYS_ExitHandler:
	Gosub, AOT_ExitHandler
	Gosub, ROL_ExitHandler
	Gosub, TRA_ExitHandler
	Gosub, CFG_SaveSettings
ExitApp



; [SYS] context check

SYS_ContextCheck:
	Gosub, SYS_TrayTipBalloonCheck
	If ( !SYS_TrayTipBalloon )
	{
		Gosub, SUS_SuspendSaveState
		Suspend, On
		MsgBox, 4148, Balloon Handler - %SYS_ScriptInfo%, The balloon messages are disabled on your system. These visual messages`nabove the system tray are often used by tools as additional information four`nyour interest.`n`nNiftyWindows uses balloon messages to show you some important operating`ndetails. If you leave the messages disabled NiftyWindows will show some plain`nmessages as tooltips instead (in front of the system tray).`n`nDo you want to enable balloon messages now (highly recommended)?
		Gosub, SUS_SuspendRestoreState
		IfMsgBox, Yes
		{
			SYS_TrayTipBalloon = 1
			RegWrite, REG_DWORD, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced, EnableBalloonTips, %SYS_TrayTipBalloon%
			RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced, EnableBalloonTips, %SYS_TrayTipBalloon%
			SendMessage, 0x001A, , , , ahk_id 0xFFFF ; 0x001A is WM_SETTINGCHANGE ; 0xFFFF is HWND_BROADCAST
			Sleep, 500 ; lets the other windows relax
		}
	}
	
	IfNotExist, %A_ScriptDir%\readme.txt
	{
		TRY_TrayEvent := "Help"
		Gosub, TRY_TrayEvent
		Suspend, On
		Sleep, 10000
		ExitApp, 1
	}

	IfNotExist, %A_ScriptDir%\license.txt
	{
		TRY_TrayEvent := "View License"
		Gosub, TRY_TrayEvent
		Suspend, On
		Sleep, 10000
		ExitApp, 1
	}
	
	TRY_TrayEvent := "About"
	Gosub, TRY_TrayEvent
Return



; [SYS] handles tooltips

SYS_ToolTipShow:
	If ( SYS_ToolTipText )
	{
		If ( !SYS_ToolTipSeconds )
			SYS_ToolTipSeconds = 2
		SYS_ToolTipMillis := SYS_ToolTipSeconds * 1000
		CoordMode, Mouse, Screen
		CoordMode, ToolTip, Screen
		If ( !SYS_ToolTipX or !SYS_ToolTipY )
		{
			MouseGetPos, SYS_ToolTipX, SYS_ToolTipY
			SYS_ToolTipX += 16
			SYS_ToolTipY += 24
		}
		ToolTip, %SYS_ToolTipText%, %SYS_ToolTipX%, %SYS_ToolTipY%
		SetTimer, SYS_ToolTipHandler, %SYS_ToolTipMillis%
	}
	SYS_ToolTipText =
	SYS_ToolTipSeconds =
	SYS_ToolTipX =
	SYS_ToolTipY =
Return

SYS_ToolTipFeedbackShow:
	If ( SYS_ToolTipFeedback )
		Gosub, SYS_ToolTipShow
	SYS_ToolTipText =
	SYS_ToolTipSeconds =
	SYS_ToolTipX =
	SYS_ToolTipY =
Return

SYS_ToolTipHandler:
	SetTimer, SYS_ToolTipHandler, Off
	ToolTip
Return



; [SYS] handles balloon messages

SYS_TrayTipShow:
	If ( SYS_TrayTipText )
	{
		If ( !SYS_TrayTipTitle )
			SYS_TrayTipTitle = %SYS_ScriptInfo%
		If ( !SYS_TrayTipSeconds )
			SYS_TrayTipSeconds = 10
		If ( !SYS_TrayTipOptions )
			SYS_TrayTipOptions = 17
		SYS_TrayTipMillis := SYS_TrayTipSeconds * 1000
		Gosub, SYS_TrayTipBalloonCheck
		If ( SYS_TrayTipBalloon and !A_IconHidden )
		{
			TrayTip, %SYS_TrayTipTitle%, %SYS_TrayTipText%, %SYS_TrayTipSeconds%, %SYS_TrayTipOptions%
			SetTimer, SYS_TrayTipHandler, %SYS_TrayTipMillis%
		}
		Else
		{
			TrayTip
			SYS_ToolTipText = %SYS_TrayTipTitle%:`n`n%SYS_TrayTipText%
			SYS_ToolTipSeconds = %SYS_TrayTipSeconds%
			SysGet, SYS_TrayTipDisplay, Monitor
			SYS_ToolTipX = %SYS_TrayTipDisplayRight%
			SYS_ToolTipY = %SYS_TrayTipDisplayBottom%
			Gosub, SYS_ToolTipShow
		}
	}
	SYS_TrayTipTitle =
	SYS_TrayTipText =
	SYS_TrayTipSeconds =
	SYS_TrayTipOptions =
Return

SYS_TrayTipHandler:
	SetTimer, SYS_TrayTipHandler, Off
	TrayTip
Return

SYS_TrayTipBalloonCheck:
	RegRead, SYS_TrayTipBalloonCU, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced, EnableBalloonTips
	SYS_TrayTipBalloonCU := ErrorLevel or SYS_TrayTipBalloonCU
	RegRead, SYS_TrayTipBalloonLM, HKEY_LOCAL_MACHINE, Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced, EnableBalloonTips
	SYS_TrayTipBalloonLM := ErrorLevel or SYS_TrayTipBalloonLM
	SYS_TrayTipBalloon := SYS_TrayTipBalloonCU and SYS_TrayTipBalloonLM
Return



; [SUS] provides suspend services

#Esc::
SUS_SuspendToggle:
	Suspend, Permit
	If ( !A_IsSuspended )
	{
		Suspend, On
		SYS_TrayTipText = NiftyWindows is suspended now.`nPress WIN+ESC to resume it again.
		SYS_TrayTipOptions = 2
	}
	Else
	{
		Suspend, Off
		SYS_TrayTipText = NiftyWindows is resumed now.`nPress WIN+ESC to suspend it again.
	}
	Gosub, SYS_TrayTipShow
	Gosub, TRY_TrayUpdate
Return

SUS_SuspendSaveState:
	SUS_Suspended := A_IsSuspended
Return

SUS_SuspendRestoreState:
	If ( SUS_Suspended )
		Suspend, On
	Else
		Suspend, Off
Return

SUS_SuspendHandler:
	IfWinActive, A
	{
		WinGet, SUS_WinID, ID
		If ( !SUS_WinID )
			Return
		WinGet, SUS_WinMinMax, MinMax, ahk_id %SUS_WinID%
		WinGetPos, SUS_WinX, SUS_WinY, SUS_WinW, SUS_WinH, ahk_id %SUS_WinID%
		
		If ( (SUS_WinMinMax = 0) and (SUS_WinX = 0) and (SUS_WinY = 0) and (SUS_WinW = A_ScreenWidth) and (SUS_WinH = A_ScreenHeight) )
		{
			WinGetClass, SUS_WinClass, ahk_id %SUS_WinID%
			WinGet, SUS_ProcessName, ProcessName, ahk_id %SUS_WinID%
			SplitPath, SUS_ProcessName, , , SUS_ProcessExt
			If ( (SUS_WinClass != "Progman") and (SUS_ProcessExt != "scr") and !SUS_FullScreenSuspend )
			{
				SUS_FullScreenSuspend = 1
				SUS_FullScreenSuspendState := A_IsSuspended
				If ( !A_IsSuspended )
				{
					Suspend, On
					SYS_TrayTipText = A full screen window was activated.`nNiftyWindows is suspended now.`nPress WIN+ESC to resume it again.
					SYS_TrayTipOptions = 2
					Gosub, SYS_TrayTipShow
					Gosub, TRY_TrayUpdate
				}
			}
		}
		Else
		{
			If ( SUS_FullScreenSuspend )
			{
				SUS_FullScreenSuspend = 0
				If ( A_IsSuspended and !SUS_FullScreenSuspendState )
				{
					Suspend, Off
					SYS_TrayTipText = A full screen window was deactivated.`nNiftyWindows is resumed now.`nPress WIN+ESC to suspend it again.
					Gosub, SYS_TrayTipShow
					Gosub, TRY_TrayUpdate
				}
			}
		}
	}
Return



; [SYS] provides reversion of all visual effects

/**
 * This powerful hotkey removes all visual effects (like on exit) that have 
 * been made before by NiftyWindows. You can use this action as a fall-back 
 * solution to quickly revert any always-on-top, rolled windows and 
 * transparency features you've set before.
 */

^#BS::
^!BS::
SYS_RevertVisualEffects:
	Gosub, AOT_SetAllOff
	Gosub, ROL_RollDownAll
	Gosub, TRA_TransparencyAllOff
	SYS_TrayTipText = All visual effects (AOT, Roll, Transparency) were reverted.
	Gosub, SYS_TrayTipShow
Return



; [NWD] nifty window dragging

/**
 * This is the most powerful feature of NiftyWindows. The area of every window 
 * is tiled in a virtual 9-cell grid with three columns and rows. The center 
 * cell is the largest one and you can grab and move a window around by clicking 
 * and holding it with the right mouse button. The other eight corner cells are 
 * used to resize a resizable window in the same manner.
 */

$RButton::
$+RButton::
$+!RButton::
$+^RButton::
$+#RButton::
$+!^RButton::
$+!#RButton::
$+^#RButton::
$+!^#RButton::
$!RButton::
$!^RButton::
$!#RButton::
$!^#RButton::
$^RButton::
$^#RButton::
$#RButton::
	NWD_ResizeGrids = 5
	CoordMode, Mouse, Screen
	MouseGetPos, NWD_MouseStartX, NWD_MouseStartY, NWD_WinID
	If ( !NWD_WinID )
		Return
	WinGetPos, NWD_WinStartX, NWD_WinStartY, NWD_WinStartW, NWD_WinStartH, ahk_id %NWD_WinID%
	WinGet, NWD_WinMinMax, MinMax, ahk_id %NWD_WinID%
	WinGet, NWD_WinStyle, Style, ahk_id %NWD_WinID%
	WinGetClass, NWD_WinClass, ahk_id %NWD_WinID%
	GetKeyState, NWD_CtrlState, Ctrl, P
	
	; the and'ed condition checks for popup window:
	; (WS_POPUP) and !(WS_DLGFRAME | WS_SYSMENU | WS_THICKFRAME)
	If ( (NWD_WinClass = "Progman") or ((NWD_CtrlState = "U") and (((NWD_WinStyle & 0x80000000) and !(NWD_WinStyle & 0x4C0000)) or (NWD_WinClass = "ExploreWClass") or (NWD_WinClass = "CabinetWClass") or (NWD_WinClass = "IEFrame") or (NWD_WinClass = "MozillaWindowClass") or (NWD_WinClass = "OpWindow") or (NWD_WinClass = "ATL:ExplorerFrame") or (NWD_WinClass = "ATL:ScrapFrame"))) )
	{
		NWD_ImmediateDownRequest = 1
		NWD_ImmediateDown = 0
		NWD_PermitClick = 1
	}
	Else
	{
		NWD_ImmediateDownRequest = 0
		NWD_ImmediateDown = 0
		NWD_PermitClick = 1
	}
	
	NWD_Dragging := (NWD_WinClass != "Progman") and ((NWD_CtrlState = "D") or ((NWD_WinMinMax != 1) and !NWD_ImmediateDownRequest))

	; checks wheter the window has a sizing border (WS_THICKFRAME)
	If ( (NWD_CtrlState = "D") or (NWD_WinStyle & 0x40000) )
	{
		If ( (NWD_MouseStartX >= NWD_WinStartX + NWD_WinStartW / NWD_ResizeGrids) and (NWD_MouseStartX <= NWD_WinStartX + (NWD_ResizeGrids - 1) * NWD_WinStartW / NWD_ResizeGrids) )
			NWD_ResizeX = 0
		Else
			If ( NWD_MouseStartX > NWD_WinStartX + NWD_WinStartW / 2 )
				NWD_ResizeX := 1
			Else
				NWD_ResizeX := -1

		If ( (NWD_MouseStartY >= NWD_WinStartY + NWD_WinStartH / NWD_ResizeGrids) and (NWD_MouseStartY <= NWD_WinStartY + (NWD_ResizeGrids - 1) * NWD_WinStartH / NWD_ResizeGrids) )
			NWD_ResizeY = 0
		Else
			If ( NWD_MouseStartY > NWD_WinStartY + NWD_WinStartH / 2 )
				NWD_ResizeY := 1
			Else
				NWD_ResizeY := -1
	}
	Else
	{
		NWD_ResizeX = 0
		NWD_ResizeY = 0
	}
	
	If ( NWD_WinStartW and NWD_WinStartH )
		NWD_WinStartAR := NWD_WinStartW / NWD_WinStartH
	Else
		NWD_WinStartAR = 0
	
	; TODO : this is a workaround (checks for popup window) for the activation 
	; bug of AutoHotkey -> can be removed as soon as the known bug is fixed
	If ( !((NWD_WinStyle & 0x80000000) and !(NWD_WinStyle & 0x4C0000)) )
		IfWinNotActive, ahk_id %NWD_WinID%
			WinActivate, ahk_id %NWD_WinID%
	
	; TODO : the hotkeys must be enabled in the 2nd block because the 1st block 
	; activates them only for the first call (historical problem of AutoHotkey)
	Hotkey, Shift, NWD_IgnoreKeyHandler
	Hotkey, Ctrl, NWD_IgnoreKeyHandler
	Hotkey, Alt, NWD_IgnoreKeyHandler
	Hotkey, LWin, NWD_IgnoreKeyHandler
	Hotkey, RWin, NWD_IgnoreKeyHandler
	Hotkey, Shift, On
	Hotkey, Ctrl, On
	Hotkey, Alt, On
	Hotkey, LWin, On
	Hotkey, RWin, On
	SetTimer, NWD_IgnoreKeyHandler, 100
	SetTimer, NWD_WindowHandler, 10
Return

NWD_SetDraggingOff:
	NWD_Dragging = 0
Return

NWD_SetClickOff:
	NWD_PermitClick = 0
	NWD_ImmediateDownRequest = 0
Return

NWD_SetAllOff:
	Gosub, NWD_SetDraggingOff
	Gosub, NWD_SetClickOff
Return

NWD_IgnoreKeyHandler:
	GetKeyState, NWD_RButtonState, RButton, P
	GetKeyState, NWD_ShiftState, Shift, P
	GetKeyState, NWD_CtrlState, Ctrl, P
	GetKeyState, NWD_AltState, Alt, P
	; TODO : unlike the other modifiers, Win does not exist 
	; as a virtual key (but Ctrl, Alt and Shift do)
	GetKeyState, NWD_LWinState, LWin, P
	GetKeyState, NWD_RWinState, RWin, P
	If ( (NWD_LWinState = "D") or (NWD_RWinState = "D") )
		NWD_WinState = D
	Else
		NWD_WinState = U
	
	If ( (NWD_RButtonState = "U") and (NWD_ShiftState = "U") and (NWD_CtrlState = "U") and (NWD_AltState = "U") and (NWD_WinState = "U") )
	{
		SetTimer, NWD_IgnoreKeyHandler, Off
		Hotkey, Shift, Off
		Hotkey, Ctrl, Off
		Hotkey, Alt, Off
		Hotkey, LWin, Off
		Hotkey, RWin, Off
	}
Return

NWD_WindowHandler:
	SetWinDelay, -1
	CoordMode, Mouse, Screen
	MouseGetPos, NWD_MouseX, NWD_MouseY
	WinGetPos, NWD_WinX, NWD_WinY, NWD_WinW, NWD_WinH, ahk_id %NWD_WinID%
	GetKeyState, NWD_RButtonState, RButton, P
	GetKeyState, NWD_ShiftState, Shift, P
	GetKeyState, NWD_AltState, Alt, P
	; TODO : unlike the other modifiers, Win does not exist 
	; as a virtual key (but Ctrl, Alt and Shift do)
	GetKeyState, NWD_LWinState, LWin, P
	GetKeyState, NWD_RWinState, RWin, P
	If ( (NWD_LWinState = "D") or (NWD_RWinState = "D") )
		NWD_WinState = D
	Else
		NWD_WinState = U
	
	If ( NWD_RButtonState = "U" )
	{
		SetTimer, NWD_WindowHandler, Off
		
		If ( NWD_ImmediateDown )
			MouseClick, RIGHT, %NWD_MouseX%, %NWD_MouseY%, , , U
		Else
			If ( NWD_PermitClick and (!NWD_Dragging or ((NWD_MouseStartX = NWD_MouseX) and (NWD_MouseStartY = NWD_MouseY))) )
			{
				MouseClick, RIGHT, %NWD_MouseStartX%, %NWD_MouseStartY%, , , D
				MouseClick, RIGHT, %NWD_MouseX%, %NWD_MouseY%, , , U
			}

		Gosub, NWD_SetAllOff
		NWD_ImmediateDown = 0
	}
	Else
	{
		NWD_MouseDeltaX := NWD_MouseX - NWD_MouseStartX
		NWD_MouseDeltaY := NWD_MouseY - NWD_MouseStartY

		If ( NWD_MouseDeltaX or NWD_MouseDeltaY )
		{
			If ( NWD_ImmediateDownRequest and !NWD_ImmediateDown )
			{
				MouseClick, RIGHT, %NWD_MouseStartX%, %NWD_MouseStartY%, , , D
				MouseMove, %NWD_MouseX%, %NWD_MouseY%
				NWD_ImmediateDown = 1
				NWD_PermitClick = 0
			}

			If ( NWD_Dragging )
			{
				If ( !NWD_ResizeX and !NWD_ResizeY )
				{
					NWD_WinNewX := NWD_WinStartX + NWD_MouseDeltaX
					NWD_WinNewY := NWD_WinStartY + NWD_MouseDeltaY
					NWD_WinNewW := NWD_WinStartW
					NWD_WinNewH := NWD_WinStartH
				}
				Else
				{
					NWD_WinDeltaW = 0
					NWD_WinDeltaH = 0
					If ( NWD_ResizeX )
						NWD_WinDeltaW := NWD_ResizeX * NWD_MouseDeltaX
					If ( NWD_ResizeY )
						NWD_WinDeltaH := NWD_ResizeY * NWD_MouseDeltaY
					If ( NWD_WinState = "D" )
					{
						If ( NWD_ResizeX )
							NWD_WinDeltaW *= 2
						If ( NWD_ResizeY )
							NWD_WinDeltaH *= 2
					}
					NWD_WinNewW := NWD_WinStartW + NWD_WinDeltaW
					NWD_WinNewH := NWD_WinStartH + NWD_WinDeltaH
					If ( NWD_WinNewW < 0 )
						If ( NWD_WinState = "D" )
							NWD_WinNewW *= -1
						Else
							NWD_WinNewW := 0
					If ( NWD_WinNewH < 0 )
						If ( NWD_WinState = "D" )
							NWD_WinNewH *= -1
						Else
							NWD_WinNewH := 0
					If ( (NWD_AltState = "D") and NWD_WinStartAR )
					{
						NWD_WinNewARW := NWD_WinNewH * NWD_WinStartAR
						NWD_WinNewARH := NWD_WinNewW / NWD_WinStartAR
						If ( NWD_WinNewW < NWD_WinNewARW )
							NWD_WinNewW := NWD_WinNewARW
						If ( NWD_WinNewH < NWD_WinNewARH )
							NWD_WinNewH := NWD_WinNewARH
					}
					NWD_WinDeltaX = 0
					NWD_WinDeltaY = 0
					If ( NWD_WinState = "D" )
					{
						NWD_WinDeltaX := NWD_WinStartW / 2 - NWD_WinNewW / 2
						NWD_WinDeltaY := NWD_WinStartH / 2 - NWD_WinNewH / 2
					}
					Else
					{
						If ( NWD_ResizeX = -1 )
							NWD_WinDeltaX := NWD_WinStartW - NWD_WinNewW
						If ( NWD_ResizeY = -1 )
							NWD_WinDeltaY := NWD_WinStartH - NWD_WinNewH
					}
					NWD_WinNewX := NWD_WinStartX + NWD_WinDeltaX
					NWD_WinNewY := NWD_WinStartY + NWD_WinDeltaY
				}
				
				If ( NWD_ShiftState = "D" )
					NWD_WinNewRound = -1
				Else
					NWD_WinNewRound = 0
				
				Transform, NWD_WinNewX, Round, %NWD_WinNewX%, %NWD_WinNewRound%
				Transform, NWD_WinNewY, Round, %NWD_WinNewY%, %NWD_WinNewRound%
				Transform, NWD_WinNewW, Round, %NWD_WinNewW%, %NWD_WinNewRound%
				Transform, NWD_WinNewH, Round, %NWD_WinNewH%, %NWD_WinNewRound%
				
				If ( (NWD_WinNewX != NWD_WinX) or (NWD_WinNewY != NWD_WinY) or (NWD_WinNewW != NWD_WinW) or (NWD_WinNewH != NWD_WinH) )
				{
					WinMove, ahk_id %NWD_WinID%, , %NWD_WinNewX%, %NWD_WinNewY%, %NWD_WinNewW%, %NWD_WinNewH%
					
					If ( SYS_ToolTipFeedback )
					{
						WinGetPos, NWD_ToolTipWinX, NWD_ToolTipWinY, NWD_ToolTipWinW, NWD_ToolTipWinH, ahk_id %NWD_WinID%
						SYS_ToolTipText = Window Drag: (X:%NWD_ToolTipWinX%, Y:%NWD_ToolTipWinY%, W:%NWD_ToolTipWinW%, H:%NWD_ToolTipWinH%)
						Gosub, SYS_ToolTipFeedbackShow
					}
				}
			}
		}
	}
Return



; [MIW {NWD}] minimize/roll on right + left mouse button

/**
 * Minimizes the selected window (if minimizable) to the task bar. If you press 
 * the left button over the titlebar the selected window will be rolled up 
 * instead of being minimized. You have to apply this action again to roll the 
 * window back down.
 */

$LButton::
$^LButton::
	GetKeyState, MIW_RButtonState, RButton, P
	If ( (MIW_RButtonState = "D") and (!NWD_ImmediateDown) and (NWD_WinClass != "Progman") )
	{
		GetKeyState, MIW_CtrlState, Ctrl, P
		WinGet, MIW_WinStyle, Style, ahk_id %NWD_WinID%
		SysGet, MIW_CaptionHeight, 4 ; SM_CYCAPTION
		SysGet, MIW_BorderHeight, 7 ; SM_CXDLGFRAME
		MouseGetPos, , MIW_MouseY

		If ( MIW_MouseY <= MIW_CaptionHeight + MIW_BorderHeight )
		{
			; checks wheter the window has a sizing border (WS_THICKFRAME)
			If ( (MIW_CtrlState = "D") or (MIW_WinStyle & 0x40000) )
			{
				Gosub, NWD_SetAllOff
				ROL_WinID = %NWD_WinID%
				Gosub, ROL_RollToggle
			}
		}
		Else
		{
			; the second condition checks for minimizable window:
			; (WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX)
			If ( (MIW_CtrlState = "D") or (MIW_WinStyle & 0xCA0000 = 0xCA0000) )
			{
				Gosub, NWD_SetAllOff
				WinMinimize, ahk_id %NWD_WinID%
				SYS_ToolTipText = Window Minimize
				Gosub, SYS_ToolTipFeedbackShow
			}
		}
	}
	Else
	{
		; this feature should be implemented by using a timer because 
		; AutoHotkeys threading blocks the first thread if another 
		; one is started (until the 2nd is stopped)
		
		Thread, priority, 1
		MouseClick, LEFT, , , , , D
		KeyWait, LButton
		MouseClick, LEFT, , , , , U
	}
Return



; [CLW {NWD}] close/send bottom on right + middle mouse button || double click on middle mouse button

/**
 * Closes the selected window (if closeable) as if you click the close button 
 * in the titlebar. If you press the middle button over the titlebar the 
 * selected window will be sent to the bottom of the window stack instead of 
 * being closed.
 */

$MButton::
$^MButton::
	GetKeyState, CLW_RButtonState, RButton, P
	If ( (CLW_RButtonState = "D") and (!NWD_ImmediateDown) and (NWD_WinClass != "Progman") )
	{
		GetKeyState, CLW_CtrlState, Ctrl, P
		WinGet, CLW_WinStyle, Style, ahk_id %NWD_WinID%
		SysGet, CLW_CaptionHeight, 4 ; SM_CYCAPTION
		SysGet, CLW_BorderHeight, 7 ; SM_CXDLGFRAME
		MouseGetPos, , CLW_MouseY

		If ( CLW_MouseY <= CLW_CaptionHeight + CLW_BorderHeight )
		{
			Gosub, NWD_SetAllOff
			Send, !{Esc}
			SYS_ToolTipText = Window Bottom
			Gosub, SYS_ToolTipFeedbackShow
		}
		Else
		{
			; the second condition checks for closeable window:
			; (WS_CAPTION | WS_SYSMENU)
			If ( (CLW_CtrlState = "D") or (CLW_WinStyle & 0xC80000 = 0xC80000) )
			{
				Gosub, NWD_SetAllOff
				WinClose, ahk_id %NWD_WinID%
				SYS_ToolTipText = Window Close
				Gosub, SYS_ToolTipFeedbackShow
			}
		}
	}
	Else
	{
		; TODO : workaround for "MouseClick, LEFT, , , 2" due to inactive titlebar problem
		
		Thread, Priority, 1
		CoordMode, Mouse, Screen
		MouseGetPos, CLW_MouseX, CLW_MouseY
		MouseClick, LEFT, %CLW_MouseX%, %CLW_MouseY%
		Sleep, 10
		MouseGetPos, CLW_MouseNewX, CLW_MouseNewY
		MouseClick, LEFT, %CLW_MouseX%, %CLW_MouseY%
		MouseMove, %CLW_MouseNewX%, %CLW_MouseNewY%
	}
Return



; [TSM {NWD}] toggles windows start menu || moves window to previous display || maximize to multiple windows on the left

/**
 * This additional button is used to toggle the windows start menu.
 */

$XButton1::
$^XButton1::
	If ( NWD_ImmediateDown )
		Return
		
	GetKeyState, TSM_RButtonState, RButton, P
	If ( TSM_RButtonState = "U" )
	{
		Send, {LWin}
	}
	Else
		IfWinActive, A
		{
			WinGet, TSM_WinID, ID
			If ( !TSM_WinID )
				Return
			WinGetClass, TSM_WinClass, ahk_id %TSM_WinID%

			If ( TSM_WinClass != "Progman" )
			{
				Gosub, NWD_SetAllOff
				GetKeyState, TSM_CtrlState, Ctrl, P
				If ( TSM_CtrlState = "U" )
				{
					Send, ^<
					SYS_ToolTipText = Window Move: LEFT
					Gosub, SYS_ToolTipFeedbackShow
				}
				; Else
				; TODO : maximize to multiple displays on the left (planned feature)
			}
		}
Return



; [MAW {NWD}] toggles window maximizing || moves window to next display || maximize to multiple windows on the right

/**
 * This additional button is used to toggle the maximize state of the active 
 * window (if maximizable).
 */

$XButton2::
$^XButton2::
	If ( NWD_ImmediateDown )
		Return
	
	IfWinActive, A
	{
		WinGet, MAW_WinID, ID
		If ( !MAW_WinID )
			Return
		WinGetClass, MAW_WinClass, ahk_id %MAW_WinID%
		If ( MAW_WinClass = "Progman" )
			Return
		
		GetKeyState, MAW_RButtonState, RButton, P
		If ( MAW_RButtonState = "U" )
		{
			GetKeyState, MAW_CtrlState, Ctrl, P
			WinGet, MAW_WinStyle, Style
			
			; the second condition checks for maximizable window:
			; (WS_CAPTION | WS_SYSMENU | WS_MAXIMIZEBOX)
			If ( (MAW_CtrlState = "D") or (MAW_WinStyle & 0xC90000 = 0xC90000) )
			{
				WinGet, MAW_MinMax, MinMax
				If ( MAW_MinMax = 0 )
				{
					WinMaximize
					SYS_ToolTipText = Window Maximize
					Gosub, SYS_ToolTipFeedbackShow
				}
				Else
					If ( MAW_MinMax = 1 )
					{
						WinRestore
						SYS_ToolTipText = Window Restore
						Gosub, SYS_ToolTipFeedbackShow
					}
			}
		}
		Else
		{
			Gosub, NWD_SetAllOff
			GetKeyState, MAW_CtrlState, Ctrl, P
			If ( MAW_CtrlState = "U" )
			{
				Send, ^>
				SYS_ToolTipText = Window Move: RIGHT
				Gosub, SYS_ToolTipFeedbackShow
			}
			; Else
			; TODO : maximize to multiple displays on the right (planned feature)
		}
	}
Return



; [TSW {NWD}] provides alt-tab-menu to the right mouse button + mouse wheel

/**
 * Provides a quick task switcher (alt-tab-menu) controlled by the mouse wheel.
 */

WheelDown::
	GetKeyState, TSW_RButtonState, RButton, P
	If ( (TSW_RButtonState = "D") and (!NWD_ImmediateDown) )
	{
		; TODO : this is a workaround because the original tabmenu 
		; code of AutoHotkey is buggy on some systems
		GetKeyState, TSW_LAltState, LAlt
		If ( TSW_LAltState = "U" )
		{
			Gosub, NWD_SetAllOff
			Send, {LAlt down}{Tab}
			SetTimer, TSW_WheelHandler, 1
		}
		Else
			Send, {Tab}
	}
	Else
		Send, {WheelDown}
Return

WheelUp::
	GetKeyState, TSW_RButtonState, RButton, P
	If ( (TSW_RButtonState = "D") and (!NWD_ImmediateDown) )
	{
		; TODO : this is a workaround because the original tabmenu 
		; code of AutoHotkey is buggy on some systems
		GetKeyState, TSW_LAltState, LAlt
		If ( TSW_LAltState = "U" )
		{
			Gosub, NWD_SetAllOff
			Send, {LAlt down}+{Tab}
			SetTimer, TSW_WheelHandler, 1
		}
		Else
			Send, +{Tab}
	}
	Else
		Send, {WheelUp}
Return

TSW_WheelHandler:
	GetKeyState, TSW_RButtonState, RButton, P
	If ( TSW_RButtonState = "U" )
	{
		SetTimer, TSW_WheelHandler, Off
		GetKeyState, TSW_LAltState, LAlt
		If ( TSW_LAltState = "D" )
			Send, {LAlt up}
	}
Return



; [AOT] toggles always on top

/**
 * Toggles the always-on-top attribute of the selected/active window.
 */

#SC029::
#LButton::
AOT_SetToggle:
	Gosub, AOT_CheckWinIDs
	SetWinDelay, -1
	
	IfInString, A_ThisHotkey, LButton
	{
		MouseGetPos, , , AOT_WinID
		If ( !AOT_WinID )
			Return
		IfWinNotActive, ahk_id %AOT_WinID%
			WinActivate, ahk_id %AOT_WinID%
	}
	
	IfWinActive, A
	{
		WinGet, AOT_WinID, ID
		If ( !AOT_WinID )
			Return
		WinGetClass, AOT_WinClass, ahk_id %AOT_WinID%
		If ( AOT_WinClass = "Progman" )
			Return
			
		WinGet, AOT_ExStyle, ExStyle, ahk_id %AOT_WinID%
		If ( AOT_ExStyle & 0x8 ) ; 0x8 is WS_EX_TOPMOST
		{
			SYS_ToolTipText = Always on Top: OFF
			Gosub, AOT_SetOff
		}
		Else
		{
			SYS_ToolTipText = Always on Top: ON
			Gosub, AOT_SetOn
		}
		Gosub, SYS_ToolTipFeedbackShow
	}
Return

AOT_SetOn:
	Gosub, AOT_CheckWinIDs
	SetWinDelay, -1
	IfWinNotExist, ahk_id %AOT_WinID%
		Return
	IfNotInString, AOT_WinIDs, |%AOT_WinID%
		AOT_WinIDs = %AOT_WinIDs%|%AOT_WinID%
	WinSet, AlwaysOnTop, On, ahk_id %AOT_WinID%
Return

AOT_SetOff:
	Gosub, AOT_CheckWinIDs
	SetWinDelay, -1
	IfWinNotExist, ahk_id %AOT_WinID%
		Return
	StringReplace, AOT_WinIDs, AOT_WinIDs, |%A_LoopField%, , All
	WinSet, AlwaysOnTop, Off, ahk_id %AOT_WinID%
Return

AOT_SetAllOff:
	Gosub, AOT_CheckWinIDs
	Loop, Parse, AOT_WinIDs, |
		If ( A_LoopField )
		{
			AOT_WinID = %A_LoopField%
			Gosub, AOT_SetOff
		}
Return

#^SC029::
	Gosub, AOT_SetAllOff
	SYS_ToolTipText = Always on Top: ALL OFF
	Gosub, SYS_ToolTipFeedbackShow
Return

AOT_CheckWinIDs:
	DetectHiddenWindows, On
	Loop, Parse, AOT_WinIDs, |
		If ( A_LoopField )
			IfWinNotExist, ahk_id %A_LoopField%
				StringReplace, AOT_WinIDs, AOT_WinIDs, |%A_LoopField%, , All
Return

AOT_ExitHandler:
	Gosub, AOT_SetAllOff
Return



; [ROL] rolls up/down a window to/from its title bar

ROL_RollToggle:
	Gosub, ROL_CheckWinIDs
	SetWinDelay, -1
	IfWinNotExist, ahk_id %ROL_WinID%
		Return
	WinGetClass, ROL_WinClass, ahk_id %ROL_WinID%
	If ( ROL_WinClass = "Progman" )
		Return
	
	IfNotInString, ROL_WinIDs, |%ROL_WinID%
	{
		SYS_ToolTipText = Window Roll: UP
		Gosub, ROL_RollUp
	}
	Else
	{
		WinGetPos, , , , ROL_WinHeight, ahk_id %ROL_WinID%
		If ( ROL_WinHeight = ROL_WinRolledHeight%ROL_WinID% )
		{
			SYS_ToolTipText = Window Roll: DOWN
			Gosub, ROL_RollDown
		}
		Else
		{
			SYS_ToolTipText = Window Roll: UP
			Gosub, ROL_RollUp
		}
	}
	Gosub, SYS_ToolTipFeedbackShow
Return

ROL_RollUp:
	Gosub, ROL_CheckWinIDs
	SetWinDelay, -1
	IfWinNotExist, ahk_id %ROL_WinID%
		Return
	WinGetClass, ROL_WinClass, ahk_id %ROL_WinID%
	If ( ROL_WinClass = "Progman" )
		Return
	
	WinGetPos, , , , ROL_WinHeight, ahk_id %ROL_WinID%
	IfInString, ROL_WinIDs, |%ROL_WinID%
		If ( ROL_WinHeight = ROL_WinRolledHeight%ROL_WinID% ) 
			Return
	SysGet, ROL_CaptionHeight, 4 ; SM_CYCAPTION
	SysGet, ROL_BorderHeight, 7 ; SM_CXDLGFRAME
	If ( ROL_WinHeight > (ROL_CaptionHeight + ROL_BorderHeight) )
	{
		IfNotInString, ROL_WinIDs, |%ROL_WinID%
			ROL_WinIDs = %ROL_WinIDs%|%ROL_WinID%
		ROL_WinOriginalHeight%ROL_WinID% := ROL_WinHeight
		WinMove, ahk_id %ROL_WinID%, , , , , (ROL_CaptionHeight + ROL_BorderHeight)
		WinGetPos, , , , ROL_WinRolledHeight%ROL_WinID%, ahk_id %ROL_WinID%
	}
Return

ROL_RollDown:
	Gosub, ROL_CheckWinIDs
	SetWinDelay, -1
	If ( !ROL_WinID )
		Return
	IfNotInString, ROL_WinIDs, |%ROL_WinID%
		Return
	WinGetPos, , , , ROL_WinHeight, ahk_id %ROL_WinID%
	If( ROL_WinHeight = ROL_WinRolledHeight%ROL_WinID% )
		WinMove, ahk_id %ROL_WinID%, , , , , ROL_WinOriginalHeight%ROL_WinID%
	StringReplace, ROL_WinIDs, ROL_WinIDs, |%ROL_WinID%, , All
	ROL_WinOriginalHeight%ROL_WinID% =
	ROL_WinRolledHeight%ROL_WinID% =
Return

ROL_RollDownAll:
	Gosub, ROL_CheckWinIDs
	Loop, Parse, ROL_WinIDs, |
		If ( A_LoopField )
		{
			ROL_WinID = %A_LoopField%
			Gosub, ROL_RollDown
		}
Return

#^r::
	Gosub, ROL_RollDownAll
	SYS_ToolTipText = Window Roll: ALL DOWN
	Gosub, SYS_ToolTipFeedbackShow
Return

ROL_CheckWinIDs:
	DetectHiddenWindows, On
	Loop, Parse, ROL_WinIDs, |
		If ( A_LoopField )
			IfWinNotExist, ahk_id %A_LoopField%
			{
				StringReplace, ROL_WinIDs, ROL_WinIDs, |%A_LoopField%, , All
				ROL_WinOriginalHeight%A_LoopField% =
				ROL_WinRolledHeight%A_LoopField% =
			}
Return

ROL_ExitHandler:
	Gosub, ROL_RollDownAll
Return



; [TRA] provides window transparency

/**
 * Adjusts the transparency of the active window in ten percent steps 
 * (opaque = 100%) which allows the contents of the windows behind it to shine 
 * through. If the window is completely transparent (0%) the window is still 
 * there and clickable. If you loose a transparent window it will be extremly 
 * complicated to find it again because it's invisible (see the first hotkey 
 * in this list for emergency help in such situations). 
 */

#WheelUp::
#+WheelUp::
#WheelDown::
#+WheelDown::
	Gosub, TRA_CheckWinIDs
	SetWinDelay, -1
	IfWinActive, A
	{
		WinGet, TRA_WinID, ID
		If ( !TRA_WinID )
			Return
		WinGetClass, TRA_WinClass, ahk_id %TRA_WinID%
		If ( TRA_WinClass = "Progman" )
			Return
		
		IfNotInString, TRA_WinIDs, |%TRA_WinID%
			TRA_WinIDs = %TRA_WinIDs%|%TRA_WinID%
		TRA_WinAlpha := TRA_WinAlpha%TRA_WinID%
		TRA_PixelColor := TRA_PixelColor%TRA_WinID%
		
		IfInString, A_ThisHotkey, +
			TRA_WinAlphaStep := 255 * 0.01 ; 1 percent steps
		Else
			TRA_WinAlphaStep := 255 * 0.1 ; 10 percent steps

		If ( TRA_WinAlpha = "" )
			TRA_WinAlpha = 255

		IfInString, A_ThisHotkey, WheelDown
			TRA_WinAlpha -= TRA_WinAlphaStep
		Else
			TRA_WinAlpha += TRA_WinAlphaStep

		If ( TRA_WinAlpha > 255 )
			TRA_WinAlpha = 255
		Else
			If ( TRA_WinAlpha < 0 )
				TRA_WinAlpha = 0

		If ( !TRA_PixelColor and (TRA_WinAlpha = 255) )
		{
			Gosub, TRA_TransparencyOff
			SYS_ToolTipText = Transparency: OFF
		}
		Else
		{
			TRA_WinAlpha%TRA_WinID% = %TRA_WinAlpha%

			If ( TRA_PixelColor )
				WinSet, TransColor, %TRA_PixelColor% %TRA_WinAlpha%, ahk_id %TRA_WinID%
			Else
				WinSet, Transparent, %TRA_WinAlpha%, ahk_id %TRA_WinID%

			TRA_ToolTipAlpha := TRA_WinAlpha * 100 / 255
			Transform, TRA_ToolTipAlpha, Round, %TRA_ToolTipAlpha%
			SYS_ToolTipText = Transparency: %TRA_ToolTipAlpha% `%
		}
		Gosub, SYS_ToolTipFeedbackShow
	}
Return

#^LButton::
#^MButton::
	Gosub, TRA_CheckWinIDs
	SetWinDelay, -1
	CoordMode, Mouse, Screen
	CoordMode, Pixel, Screen
	MouseGetPos, TRA_MouseX, TRA_MouseY, TRA_WinID
	If ( !TRA_WinID )
		Return
	WinGetClass, TRA_WinClass, ahk_id %TRA_WinID%
	If ( TRA_WinClass = "Progman" )
		Return
	
	IfWinNotActive, ahk_id %TRA_WinID%
		WinActivate, ahk_id %TRA_WinID%
	IfNotInString, TRA_WinIDs, |%TRA_WinID%
		TRA_WinIDs = %TRA_WinIDs%|%TRA_WinID%
	
	IfInString, A_ThisHotkey, MButton
	{
		AOT_WinID = %TRA_WinID%
		Gosub, AOT_SetOn
		TRA_WinAlpha%TRA_WinID% := 25 * 255 / 100
	}
	
	TRA_WinAlpha := TRA_WinAlpha%TRA_WinID%
	
	; TODO : the transparency must be set off first, 
	; this may be a bug of AutoHotkey
	WinSet, TransColor, OFF, ahk_id %TRA_WinID%
	PixelGetColor, TRA_PixelColor, %TRA_MouseX%, %TRA_MouseY%, RGB
	WinSet, TransColor, %TRA_PixelColor% %TRA_WinAlpha%, ahk_id %TRA_WinID%
	TRA_PixelColor%TRA_WinID% := TRA_PixelColor

	IfInString, A_ThisHotkey, MButton
		SYS_ToolTipText = Transparency: 25 `% + %TRA_PixelColor% color (RGB) + Always on Top
	Else
		SYS_ToolTipText = Transparency: %TRA_PixelColor% color (RGB)
	Gosub, SYS_ToolTipFeedbackShow
Return

#MButton::
	Gosub, TRA_CheckWinIDs
	SetWinDelay, -1
	MouseGetPos, , , TRA_WinID
	If ( !TRA_WinID )
		Return
	IfWinNotActive, ahk_id %TRA_WinID%
		WinActivate, ahk_id %TRA_WinID%
	IfNotInString, TRA_WinIDs, |%TRA_WinID%
		Return
	Gosub, TRA_TransparencyOff

	SYS_ToolTipText = Transparency: OFF
	Gosub, SYS_ToolTipFeedbackShow
Return

TRA_TransparencyOff:
	Gosub, TRA_CheckWinIDs
	SetWinDelay, -1
	If ( !TRA_WinID )
		Return
	IfNotInString, TRA_WinIDs, |%TRA_WinID%
		Return
	StringReplace, TRA_WinIDs, TRA_WinIDs, |%TRA_WinID%, , All
	TRA_WinAlpha%TRA_WinID% =
	TRA_PixelColor%TRA_WinID% =
	; TODO : must be set to 255 first to avoid the black-colored-window problem
	WinSet, Transparent, 255, ahk_id %TRA_WinID%
	WinSet, TransColor, OFF, ahk_id %TRA_WinID%
	WinSet, Transparent, OFF, ahk_id %TRA_WinID%
	WinSet, Redraw, , ahk_id %TRA_WinID%
Return

TRA_TransparencyAllOff:
	Gosub, TRA_CheckWinIDs
	Loop, Parse, TRA_WinIDs, |
		If ( A_LoopField )
		{
			TRA_WinID = %A_LoopField%
			Gosub, TRA_TransparencyOff
		}
Return

#^t::
	Gosub, TRA_TransparencyAllOff
	SYS_ToolTipText = Transparency: ALL OFF
	Gosub, SYS_ToolTipFeedbackShow
Return

TRA_CheckWinIDs:
	DetectHiddenWindows, On
	Loop, Parse, TRA_WinIDs, |
		If ( A_LoopField )
			IfWinNotExist, ahk_id %A_LoopField%
			{
				StringReplace, TRA_WinIDs, TRA_WinIDs, |%A_LoopField%, , All
				TRA_WinAlpha%A_LoopField% =
				TRA_PixelColor%A_LoopField% =
			}
Return

TRA_ExitHandler:
	Gosub, TRA_TransparencyAllOff
Return



; [EJC] opens/closes a drive

/**
 * Opens or closes an installed CD/DVD-ROM reader/writer drive tray. The drives 
 * are assigned to their hotkeys by the certain drive number in your system. The 
 * hotkeys are used in the sequence of the key placement on your physical 
 * keyboard from left to right (1 refers to the first and 0 to the tenth drive). 
 * So you are limited to a total number of 10 drives controlled by NiftyWindows.
 */

#MaxThreadsBuffer On
$#0::
$#1::
$#2::
$#3::
$#4::
$#5::
$#6::
$#7::
$#8::
$#9::
	DriveGet, EJC_DriveList, List, CDROM
	StringLen, EJC_DriveListLength, EJC_DriveList
	StringRight, EJC_PressedDrive, A_ThisHotkey, 1
	
	If ( !EJC_PressedDrive )
		EJC_PressedDrive := 10
	
	If ( EJC_PressedDrive <= EJC_DriveListLength )
	{
		StringMid, EJC_DriveLabel, EJC_DriveList, EJC_PressedDrive, 1
		SYS_ToolTipText = Drive Eject: %EJC_DriveLabel%
		Gosub, SYS_ToolTipFeedbackShow
		Gosub, SUS_SuspendSaveState
		Suspend, On
		Drive, Eject, %EJC_DriveLabel%:
		If ( A_TimeSinceThisHotkey < 250 )
			Drive, Eject, %EJC_DriveLabel%:, 1
		Gosub, SUS_SuspendRestoreState
	}
	Else
	{
		Transform, EJC_PressedDrive, Mod, %EJC_PressedDrive%, 10
		Send, #%EJC_PressedDrive%
	}
Return
#MaxThreadsBuffer Off



; [MUT] toggles the audio mute

/**
 * Toggles the muteness of an installed audio card.
 */

Pause::
	SoundSet, +1, , Mute
	SoundGet, MUT_MuteState, , MUTE
	SYS_ToolTipText = Audio Mute: %MUT_MuteState%
	Gosub, SYS_ToolTipFeedbackShow
Return



; [SCR] starts the user defined screensaver

/**
 * Starts the user defined screensaver (password protection aware). 
 */

#s up::
^#s up::
	RegRead, SCR_Saver, HKEY_CURRENT_USER, Control Panel\Desktop, SCRNSAVE.EXE
	If ( !ErrorLevel and SCR_Saver )
	{
		SendMessage, 0x112, 0xF140, 0, , Program Manager ; 0x112 is WM_SYSCOMMAND ; 0xF140 is SC_SCREENSAVE
		
		If ( A_ThisHotkey != "^#s up" )
			Return
		
		SplitPath, SCR_Saver, SCR_SaverFileName
		Process, Wait, %SCR_SaverFileName%, 5
		If ( ErrorLevel )
		{
			Gosub, SUS_SuspendSaveState
			Suspend, On
			Sleep, 5000
			Gosub, SUS_SuspendRestoreState
			Process, Exist, %SCR_SaverFileName%
			If ( ErrorLevel )
				SendMessage, 0x112, 0xF170, 2, , Program Manager ; 0x112 is WM_SYSCOMMAND ; 0xF170 is SC_MONITORPOWER ; (2 = off, 1 = standby, -1 = on)
		}
	}
	Else
	{
		SYS_TrayTipText = No screensaver specified in display settings (control panel).
		SYS_TrayTipOptions = 2
		Gosub, SYS_TrayTipShow
	}
Return



; [SIZ {NWD}] provides several size adjustments to windows

/**
 * Adjusts the transparency of the active window in ten percent steps 
 * (opaque = 100%) which allows the contents of the windows behind it to shine 
 * through. If the window is completely transparent (0%) the window is still 
 * there and clickable. If you loose a transparent window it will be extremly 
 * complicated to find it again because it's invisible (see the first hotkey in 
 * this list for emergency help in such situations). 
 */

!WheelUp::
!+WheelUp::
!^WheelUp::
!#WheelUp::
!+^WheelUp::
!+#WheelUp::
!^#WheelUp::
!+^#WheelUp::
!WheelDown::
!+WheelDown::
!^WheelDown::
!#WheelDown::
!+^WheelDown::
!+#WheelDown::
!^#WheelDown::
!+^#WheelDown::
	; TODO : the following code block is a workaround to handle 
	; virtual ALT calls in WheelDown/Up functions
	GetKeyState, SIZ_AltState, Alt, P
	If ( SIZ_AltState = "U" )
	{
		IfInString, A_ThisHotkey, WheelDown
			Gosub, WheelDown
		Else
			Gosub, WheelUp
		Return
	}

	If ( NWD_Dragging or NWD_ImmediateDown )
		Return
	
	SetWinDelay, -1
	CoordMode, Mouse, Screen
	IfWinActive, A
	{
		WinGet, SIZ_WinID, ID
		If ( !SIZ_WinID )
			Return
		WinGetClass, SIZ_WinClass, ahk_id %SIZ_WinID%
		If ( SIZ_WinClass = "Progman" )
			Return
		
		GetKeyState, SIZ_CtrlState, Ctrl, P
		WinGet, SIZ_WinMinMax, MinMax, ahk_id %SIZ_WinID%
		WinGet, SIZ_WinStyle, Style, ahk_id %SIZ_WinID%

		; checks wheter the window isn't maximized and has a sizing border (WS_THICKFRAME)
		If ( (SIZ_CtrlState = "D") or ((SIZ_WinMinMax != 1) and (SIZ_WinStyle & 0x40000)) )
		{
			WinGetPos, SIZ_WinX, SIZ_WinY, SIZ_WinW, SIZ_WinH, ahk_id %SIZ_WinID%
			
			If ( SIZ_WinW and SIZ_WinH )
			{
				SIZ_AspectRatio := SIZ_WinW / SIZ_WinH

				IfInString, A_ThisHotkey, WheelDown
					SIZ_Direction = 1
				Else
					SIZ_Direction = -1
				
				IfInString, A_ThisHotkey, +
					SIZ_Factor = 0.01
				Else
					SIZ_Factor = 0.1
				
				SIZ_WinNewW := SIZ_WinW + SIZ_Direction * SIZ_WinW * SIZ_Factor
				SIZ_WinNewH := SIZ_WinH + SIZ_Direction * SIZ_WinH * SIZ_Factor
				
				IfInString, A_ThisHotkey, #
				{
					SIZ_WinNewX := SIZ_WinX + (SIZ_WinW - SIZ_WinNewW) / 2
					SIZ_WinNewY := SIZ_WinY + (SIZ_WinH - SIZ_WinNewH) / 2
				}
				Else
				{
					SIZ_WinNewX := SIZ_WinX
					SIZ_WinNewY := SIZ_WinY
				}
				
				If ( SIZ_WinNewW > A_ScreenWidth )
				{
					SIZ_WinNewW := A_ScreenWidth
					SIZ_WinNewH := SIZ_WinNewW / SIZ_AspectRatio
				}
				If ( SIZ_WinNewH > A_ScreenHeight )
				{
					SIZ_WinNewH := A_ScreenHeight
					SIZ_WinNewW := SIZ_WinNewH * SIZ_AspectRatio
				}
				
				Transform, SIZ_WinNewX, Round, %SIZ_WinNewX%
				Transform, SIZ_WinNewY, Round, %SIZ_WinNewY%
				Transform, SIZ_WinNewW, Round, %SIZ_WinNewW%
				Transform, SIZ_WinNewH, Round, %SIZ_WinNewH%
				
				WinMove, ahk_id %SIZ_WinID%, , SIZ_WinNewX, SIZ_WinNewY, SIZ_WinNewW, SIZ_WinNewH
				
				If ( SYS_ToolTipFeedback )
				{
					WinGetPos, SIZ_ToolTipWinX, SIZ_ToolTipWinY, SIZ_ToolTipWinW, SIZ_ToolTipWinH, ahk_id %SIZ_WinID%
					SYS_ToolTipText = Window Size: (X:%SIZ_ToolTipWinX%, Y:%SIZ_ToolTipWinY%, W:%SIZ_ToolTipWinW%, H:%SIZ_ToolTipWinH%)
					Gosub, SYS_ToolTipFeedbackShow
				}
			}
		}
	}
Return

!NumpadAdd::
!^NumpadAdd::
!#NumpadAdd::
!^#NumpadAdd::
!NumpadSub::
!^NumpadSub::
!#NumpadSub::
!^#NumpadSub::
	If ( NWD_Dragging or NWD_ImmediateDown )
		Return

	SetWinDelay, -1
	CoordMode, Mouse, Screen
	IfWinActive, A
	{
		WinGet, SIZ_WinID, ID
		If ( !SIZ_WinID )
			Return
		WinGetClass, SIZ_WinClass, ahk_id %SIZ_WinID%
		If ( SIZ_WinClass = "Progman" )
			Return
		
		GetKeyState, SIZ_CtrlState, Ctrl, P
		WinGet, SIZ_WinMinMax, MinMax, ahk_id %SIZ_WinID%
		WinGet, SIZ_WinStyle, Style, ahk_id %SIZ_WinID%

		; checks wheter the window isn't maximized and has a sizing border (WS_THICKFRAME)
		If ( (SIZ_CtrlState = "D") or ((SIZ_WinMinMax != 1) and (SIZ_WinStyle & 0x40000)) )
		{
			WinGetPos, SIZ_WinX, SIZ_WinY, SIZ_WinW, SIZ_WinH, ahk_id %SIZ_WinID%
			
			IfInString, A_ThisHotkey, NumpadAdd
				If ( SIZ_WinW < 160 )
					SIZ_WinNewW = 160
				Else
					If ( SIZ_WinW < 320 )
						SIZ_WinNewW = 320
					Else
						If ( SIZ_WinW < 640 )
							SIZ_WinNewW = 640
						Else
							If ( SIZ_WinW < 800 )
								SIZ_WinNewW = 800
							Else
								If ( SIZ_WinW < 1024 )
									SIZ_WinNewW = 1024
								Else
									If ( SIZ_WinW < 1152 )
										SIZ_WinNewW = 1152
									Else
										If ( SIZ_WinW < 1280 )
											SIZ_WinNewW = 1280
										Else
											If ( SIZ_WinW < 1400 )
												SIZ_WinNewW = 1400
											Else
												If ( SIZ_WinW < 1600 )
													SIZ_WinNewW = 1600
												Else
													SIZ_WinNewW = 1920
			Else
				If ( SIZ_WinW <= 320 )
					SIZ_WinNewW = 160
				Else
					If ( SIZ_WinW <= 640 )
						SIZ_WinNewW = 320
					Else
						If ( SIZ_WinW <= 800 )
							SIZ_WinNewW = 640
						Else
							If ( SIZ_WinW <= 1024 )
								SIZ_WinNewW = 800
							Else
								If ( SIZ_WinW <= 1152 )
									SIZ_WinNewW = 1024
								Else
									If ( SIZ_WinW <= 1280 )
										SIZ_WinNewW = 1152
									Else
										If ( SIZ_WinW <= 1400 )
											SIZ_WinNewW = 1280
										Else
											If ( SIZ_WinW <= 1600 )
												SIZ_WinNewW = 1400
											Else
												If ( SIZ_WinW <= 1920 )
													SIZ_WinNewW = 1600
												Else
													SIZ_WinNewW = 1920
			
			If ( SIZ_WinNewW > A_ScreenWidth )
				SIZ_WinNewW := A_ScreenWidth
			SIZ_WinNewH := 3 * SIZ_WinNewW / 4
			If ( SIZ_WinNewW = 1280 )
				SIZ_WinNewH := 1024
			
			IfInString, A_ThisHotkey, #
			{
				SIZ_WinNewX := SIZ_WinX + (SIZ_WinW - SIZ_WinNewW) / 2
				SIZ_WinNewY := SIZ_WinY + (SIZ_WinH - SIZ_WinNewH) / 2
			}
			Else
			{
				SIZ_WinNewX := SIZ_WinX
				SIZ_WinNewY := SIZ_WinY
			}
			
			Transform, SIZ_WinNewX, Round, %SIZ_WinNewX%
			Transform, SIZ_WinNewY, Round, %SIZ_WinNewY%
			Transform, SIZ_WinNewW, Round, %SIZ_WinNewW%
			Transform, SIZ_WinNewH, Round, %SIZ_WinNewH%
			
			WinMove, ahk_id %SIZ_WinID%, , SIZ_WinNewX, SIZ_WinNewY, SIZ_WinNewW, SIZ_WinNewH
			
			If ( SYS_ToolTipFeedback )
			{
				WinGetPos, SIZ_ToolTipWinX, SIZ_ToolTipWinY, SIZ_ToolTipWinW, SIZ_ToolTipWinH, ahk_id %SIZ_WinID%
				SYS_ToolTipText = Window Size: (X:%SIZ_ToolTipWinX%, Y:%SIZ_ToolTipWinY%, W:%SIZ_ToolTipWinW%, H:%SIZ_ToolTipWinH%)
				Gosub, SYS_ToolTipFeedbackShow
			}
		}
	}
Return



; [XWN] provides X Window like focus switching (focus follows mouse)

/**
 * Provided a 'X Window' like focus switching by mouse cursor movement. After 
 * activation of this feature (by using the responsible entry in the tray icon 
 * menu) the focus will follow the mouse cursor with a delayed focus change 
 * (after movement end) of 500 milliseconds (half a second). This feature is 
 * disabled per default to avoid any confusion due to the new user-interface-flow.
 */

XWN_FocusHandler:
	CoordMode, Mouse, Screen
	MouseGetPos, XWN_MouseX, XWN_MouseY, XWN_WinID
	If ( !XWN_WinID )
		Return
	
	If ( (XWN_MouseX != XWN_MouseOldX) or (XWN_MouseY != XWN_MouseOldY) )
	{
		IfWinNotActive, ahk_id %XWN_WinID%
			XWN_FocusRequest = 1
		Else
			XWN_FocusRequest = 0
		
		XWN_MouseOldX := XWN_MouseX
		XWN_MouseOldY := XWN_MouseY
		XWN_MouseMovedTickCount := A_TickCount
	}
	Else
		If ( XWN_FocusRequest and (A_TickCount - XWN_MouseMovedTickCount > 500) )
		{
			WinGetClass, XWN_WinClass, ahk_id %XWN_WinID%
			If ( XWN_WinClass = "Progman" )
				Return
			
			; checks wheter the selected window is a popup menu
			; (WS_POPUP) and !(WS_DLGFRAME | WS_SYSMENU | WS_THICKFRAME)
			WinGet, XWN_WinStyle, Style, ahk_id %XWN_WinID%
			If ( (XWN_WinStyle & 0x80000000) and !(XWN_WinStyle & 0x4C0000) )
				Return
			
			IfWinNotActive, ahk_id %XWN_WinID%
				WinActivate, ahk_id %XWN_WinID%
				
			XWN_FocusRequest = 0
		}
Return



; [GRP] groups windows for quick task switching

/**
 * Activates the next window in a process window group that was defined 
 * gradually before with the given CTRL modifier. This feature causes the first 
 * window of the responsible group to be activated. Using it a second time will 
 * activate the next window in the series and so on. By using process window 
 * groups you can organize and access your process windows in semantic groups 
 * quickly. 
 */

^#F1::
^#F2::
^#F3::
^#F4::
^#F5::
^#F6::
^#F7::
^#F8::
^#F9::
^#F10::
^#F11::
^#F12::
^#F13::
^#F14::
^#F15::
^#F16::
^#F17::
^#F18::
^#F19::
^#F20::
^#F21::
^#F22::
^#F23::
^#F24::
	IfWinActive, A
	{
		WinGet, GRP_WinID, ID
		If ( !GRP_WinID )
			Return
		WinGetClass, GRP_WinClass, ahk_id %GRP_WinID%
		If ( GRP_WinClass = "Progman" )
			Return
		WinGet, GRP_WinPID, PID
		If ( !GRP_WinPID )
			Return
			
		StringMid, GRP_GroupNumber, A_ThisHotkey, 3, 3
		GroupAdd, Group%GRP_GroupNumber%, ahk_PID %GRP_WinPID%
		
		SYS_ToolTipText = Active window was added to group %GRP_GroupNumber%.
		Gosub, SYS_ToolTipFeedbackShow
	}
Return

#F1::
#F2::
#F3::
#F4::
#F5::
#F6::
#F7::
#F8::
#F9::
#F10::
#F11::
#F12::
#F13::
#F14::
#F15::
#F16::
#F17::
#F18::
#F19::
#F20::
#F21::
#F22::
#F23::
#F24::
	StringMid, GRP_GroupNumber, A_ThisHotkey, 2, 3
	GroupActivate, Group%GRP_GroupNumber%
	
	SYS_ToolTipText = Activated next window in group %GRP_GroupNumber%.
	Gosub, SYS_ToolTipFeedbackShow
Return

!#F1::
!#F2::
!#F3::
!#F4::
!#F5::
!#F6::
!#F7::
!#F8::
!#F9::
!#F10::
!#F11::
!#F12::
!#F13::
!#F14::
!#F15::
!#F16::
!#F17::
!#F18::
!#F19::
!#F20::
!#F21::
!#F22::
!#F23::
!#F24::
	StringMid, GRP_GroupNumber, A_ThisHotkey, 3, 3
	GroupClose, Group%GRP_GroupNumber%, A
	
	SYS_ToolTipText = Closed all windows in group %GRP_GroupNumber%.
	Gosub, SYS_ToolTipFeedbackShow
Return



; [MIR] toggles the visibility of miranda buddy list

/**
 * Toggles the visibility of the Miranda buddy list (if installed). Currently 
 * Miranda does not provide a hotkey to activate the buddy list if the window 
 * is still visible. Instead the opened (but not activated) buddy list will be 
 * minimized. This is not expected so this NiftyWindows feature provides the 
 * needed service asked by so many people.
 */

^+b::
	IfExist, %MIR_MirandaFullPath%
	{
		SetTitleMatchMode, 3
		DetectHiddenWindows, On
		MIR_MirandaStart = 0
		IfWinNotExist, Miranda IM
		{
			Run, %MIR_MirandaFullPath%, %MIR_MirandaDir%
			WinWait, Miranda IM
			MIR_MirandaStart=1
			Sleep, 500
		}
		DetectHiddenWindows, Off
		IfWinActive, Miranda IM
		{
			If ( !MIR_MirandaStart )
				WinHide
		}
		Else
			IfWinExist, Miranda IM
				WinActivate
			Else
			{
				DetectHiddenWindows, On
				IfWinExist, Miranda IM
				{
					WinShow
					WinActivate
				}
			}
	}
	Else
		Send, ^+b
Return



; [MIR] toggles the visibility of last used miranda message container

/**
 * Toggles the visibility of the last used Miranda message container 
 * (if installed). Currently Miranda does not provide a hotkey to activate the 
 * last used message container if there is no unread message waiting for your 
 * attention. So this hotkey will make a container visible (if it is minimized) 
 * and activate it. If there is no existing message container, this hotkey will 
 * do nothing. 
 */

~^+u::
	IfExist, %MIR_MirandaFullPath%
	{
		Sleep, 500	
		SetTitleMatchMode, 3
		IfWinExist, ahk_class #32770
		{
			WinGetTitle, MIR_Title
			IfNotInString MIR_Title, Mail
				IfWinNotActive
					WinActivate
		}
	}
	Else
		Send, ^+u
Return



; [TRY] handles the tray icon/menu

TRY_TrayInit:
	Menu, TRAY, NoStandard
	Menu, TRAY, Tip, %SYS_ScriptInfo%

	If ( !A_IsCompiled )
	{
		Menu, AutoHotkey, Standard
		Menu, TRAY, Add, AutoHotkey, :AutoHotkey
		Menu, TRAY, Add
	}

	Menu, TRAY, Add, Help, TRY_TrayEvent
	Menu, TRAY, Default, Help
	Menu, TRAY, Add
	Menu, TRAY, Add, About, TRY_TrayEvent
	Menu, TRAY, Add
	Menu, TRAY, Add, Mail Author, TRY_TrayEvent
	Menu, TRAY, Add, View License, TRY_TrayEvent
	Menu, TRAY, Add, Visit Website, TRY_TrayEvent
	Menu, TRAY, Add, Check For Update, TRY_TrayEvent
	Menu, TRAY, Add

	Menu, MouseHooks, Add, Left Mouse Button, TRY_TrayEvent
	Menu, MouseHooks, Add, Middle Mouse Button, TRY_TrayEvent
	Menu, MouseHooks, Add, Right Mouse Button, TRY_TrayEvent
	Menu, MouseHooks, Add, Fourth Mouse Button, TRY_TrayEvent
	Menu, MouseHooks, Add, Fifth Mouse Button, TRY_TrayEvent
	Menu, TRAY, Add, Mouse Hooks, :MouseHooks

	Menu, TRAY, Add, ToolTip Feedback, TRY_TrayEvent
	Menu, TRAY, Add, Auto Suspend, TRY_TrayEvent
	Menu, TRAY, Add, Focus Follows Mouse, TRY_TrayEvent
	Menu, TRAY, Add, Suspend All Hooks, TRY_TrayEvent
	Menu, TRAY, Add, Revert Visual Effects, TRY_TrayEvent
	Menu, TRAY, Add, Hide Tray Icon, TRY_TrayEvent
	Menu, TRAY, Add
	Menu, TRAY, Add, Exit, TRY_TrayEvent
	
	Gosub, TRY_TrayUpdate

	If ( A_IconHidden )
		Menu, TRAY, Icon
Return

TRY_TrayUpdate:
	If ( CFG_LeftMouseButtonHook )
		Menu, MouseHooks, Check, Left Mouse Button
	Else
		Menu, MouseHooks, UnCheck, Left Mouse Button
	If ( CFG_MiddleMouseButtonHook )
		Menu, MouseHooks, Check, Middle Mouse Button
	Else
		Menu, MouseHooks, UnCheck, Middle Mouse Button
	If ( CFG_RightMouseButtonHook )
		Menu, MouseHooks, Check, Right Mouse Button
	Else
		Menu, MouseHooks, UnCheck, Right Mouse Button
	If ( CFG_FourthMouseButtonHook )
		Menu, MouseHooks, Check, Fourth Mouse Button
	Else
		Menu, MouseHooks, UnCheck, Fourth Mouse Button
	If ( CFG_FifthMouseButtonHook )
		Menu, MouseHooks, Check, Fifth Mouse Button
	Else
		Menu, MouseHooks, UnCheck, Fifth Mouse Button
	If ( SYS_ToolTipFeedback )
		Menu, TRAY, Check, ToolTip Feedback
	Else
		Menu, TRAY, UnCheck, ToolTip Feedback
	If ( SUS_AutoSuspend )
		Menu, TRAY, Check, Auto Suspend
	Else
		Menu, TRAY, UnCheck, Auto Suspend
	If ( XWN_FocusFollowsMouse )
		Menu, TRAY, Check, Focus Follows Mouse
	Else
		Menu, TRAY, UnCheck, Focus Follows Mouse
	If ( A_IsSuspended )
		Menu, TRAY, Check, Suspend All Hooks
	Else
		Menu, TRAY, UnCheck, Suspend All Hooks
Return

TRY_TrayEvent:
	If ( !TRY_TrayEvent )
		TRY_TrayEvent = %A_ThisMenuItem%
	
	If ( TRY_TrayEvent = "Help" )
		IfExist, %A_ScriptDir%\readme.txt
			Run, "%A_ScriptDir%\readme.txt"
		Else
		{
			SYS_TrayTipText = File couldn't be accessed:`n%A_ScriptDir%\readme.txt
			SYS_TrayTipOptions = 3
			Gosub, SYS_TrayTipShow
		}

	If ( TRY_TrayEvent = "About" )
	{
		SYS_TrayTipText = Copyright (c) 2004-2005 by Enovatic-Solutions.`nAll rights reserved. Use is subject to license terms.`n`nCompany:`tEnovatic-Solutions (IT Service Provider)`nAuthor:`t`tOliver Pfeiffer`, Bremen (GERMANY)`nEmail:`t`tniftywindows@enovatic.org
		Gosub, SYS_TrayTipShow
	}

	If ( TRY_TrayEvent = "Mail Author" )
		Run, mailto:niftywindows@enovatic.org?subject=%SYS_ScriptInfo% (build %SYS_ScriptBuild%)

	If ( TRY_TrayEvent = "View License" )
		IfExist, %A_ScriptDir%\license.txt
			Run, "%A_ScriptDir%\license.txt"
		Else
		{
			SYS_TrayTipText = File couldn't be accessed:`n%A_ScriptDir%\license.txt
			SYS_TrayTipOptions = 3
			Gosub, SYS_TrayTipShow
		}

	If ( TRY_TrayEvent = "Visit Website" )
		Run, http://www.enovatic.org/products/niftywindows/

	If ( TRY_TrayEvent = "Check For Update" )
		Gosub, UPD_CheckForUpdate

	If ( TRY_TrayEvent = "ToolTip Feedback" )
		SYS_ToolTipFeedback := !SYS_ToolTipFeedback

	If ( TRY_TrayEvent = "Auto Suspend" )
	{
		SUS_AutoSuspend := !SUS_AutoSuspend
		Gosub, CFG_ApplySettings
	}

	If ( TRY_TrayEvent = "Focus Follows Mouse" )
	{
		XWN_FocusFollowsMouse := !XWN_FocusFollowsMouse
		Gosub, CFG_ApplySettings
	}

	If ( TRY_TrayEvent = "Suspend All Hooks" )
		Gosub, SUS_SuspendToggle
	
	If ( TRY_TrayEvent = "Revert Visual Effects" )
		Gosub, SYS_RevertVisualEffects

	If ( TRY_TrayEvent = "Hide Tray Icon" )
	{
		SYS_TrayTipText = Tray icon will be hidden now.`nPress WIN+X to show it again.
		SYS_TrayTipOptions = 2
		SYS_TrayTipSeconds = 5
		Gosub, SYS_TrayTipShow
		SetTimer, TRY_TrayHide, 5000
	}

	If ( TRY_TrayEvent = "Exit" )
		ExitApp

	If ( TRY_TrayEvent = "Left Mouse Button" )
	{
		CFG_LeftMouseButtonHook := !CFG_LeftMouseButtonHook
		Gosub, CFG_ApplySettings
	}
	
	If ( TRY_TrayEvent = "Middle Mouse Button" )
	{
		CFG_MiddleMouseButtonHook := !CFG_MiddleMouseButtonHook
		Gosub, CFG_ApplySettings
	}
	
	If ( TRY_TrayEvent = "Right Mouse Button" )
	{
		CFG_RightMouseButtonHook := !CFG_RightMouseButtonHook
		Gosub, CFG_ApplySettings
	}
	
	If ( TRY_TrayEvent = "Fourth Mouse Button" )
	{
		CFG_FourthMouseButtonHook := !CFG_FourthMouseButtonHook
		Gosub, CFG_ApplySettings
	}
	
	If ( TRY_TrayEvent = "Fifth Mouse Button" )
	{
		CFG_FifthMouseButtonHook := !CFG_FifthMouseButtonHook
		Gosub, CFG_ApplySettings
	}

	Gosub, TRY_TrayUpdate
	TRY_TrayEvent =
Return

TRY_TrayHide:
	SetTimer, TRY_TrayHide, Off
	Menu, TRAY, NoIcon
Return



; [EDT] edits this script in notepad

^#!F9::
	If ( A_IsCompiled )
		Return
	
	Gosub, SUS_SuspendSaveState
	Suspend, On
	MsgBox, 4129, Edit Handler - %SYS_ScriptInfo%, You pressed the hotkey for editing this script:`n`n%A_ScriptFullPath%`n`nDo you really want to edit?
	Gosub, SUS_SuspendRestoreState
	IfMsgBox, OK
		Run, notepad.exe %A_ScriptFullPath%
Return



; [REL] reloads this script on change

REL_ScriptReload:
	If ( A_IsCompiled )
		Return

	FileGetAttrib, REL_Attribs, %A_ScriptFullPath%
	IfInString, REL_Attribs, A
	{
		FileSetAttrib, -A, %A_ScriptFullPath%
		If ( REL_InitDone )
		{
			Gosub, SUS_SuspendSaveState
			Suspend, On
			MsgBox, 4145, Update Handler - %SYS_ScriptInfo%, The following script has changed:`n`n%A_ScriptFullPath%`n`nReload and activate this script?
			Gosub, SUS_SuspendRestoreState
			IfMsgBox, OK
				Reload
		}
	}
	REL_InitDone = 1
Return



; [EXT] exits this script

#x::
	If ( A_IconHidden )
	{
		Menu, TRAY, Icon
		SYS_TrayTipText = Tray icon is shown now.`nPress WIN+X again to exit NiftyWindows.
		SYS_TrayTipSeconds = 5
		Gosub, SYS_TrayTipShow
		Return
	}

	If ( A_IsCompiled )
	{
		SYS_TrayTipText = NiftyWindows will exit now.`nYou can find it here (to start it again):`n%A_ScriptFullPath%
		SYS_TrayTipOptions = 2
		SYS_TrayTipSeconds = 5
		Gosub, SYS_TrayTipShow
		Suspend, On
		Sleep, 5000
		ExitApp
	}

	Gosub, SUS_SuspendSaveState
	Suspend, On
	MsgBox, 4145, Exit Handler - %SYS_ScriptInfo%, You pressed the hotkey for exiting this script:`n`n%A_ScriptFullPath%`n`nDo you really want to exit?
	Gosub, SUS_SuspendRestoreState
	IfMsgBox, OK
		ExitApp
Return



; [CFG] handles the persistent configuration

CFG_LoadSettings:
	CFG_IniFile = %A_ScriptDir%\%SYS_ScriptNameNoExt%.ini
	IniRead, SUS_AutoSuspend, %CFG_IniFile%, Main, AutoSuspend, 1
	IniRead, XWN_FocusFollowsMouse, %CFG_IniFile%, WindowHandling, FocusFollowsMouse, 0
	IniRead, SYS_ToolTipFeedback, %CFG_IniFile%, Visual, ToolTipFeedback, 1
	IniRead, UPD_LastUpdateCheck, %CFG_IniFile%, UpdateCheck, LastUpdateCheck, %A_MM%
	IniRead, CFG_LeftMouseButtonHook, %CFG_IniFile%, MouseHooks, LeftMouseButton, 1
	IniRead, CFG_MiddleMouseButtonHook, %CFG_IniFile%, MouseHooks, MiddleMouseButton, 1
	IniRead, CFG_RightMouseButtonHook, %CFG_IniFile%, MouseHooks, RightMouseButton, 1
	IniRead, CFG_FourthMouseButtonHook, %CFG_IniFile%, MouseHooks, FourthMouseButton, 1
	IniRead, CFG_FifthMouseButtonHook, %CFG_IniFile%, MouseHooks, FifthMouseButton, 1
Return

CFG_SaveSettings:
	CFG_IniFile = %A_ScriptDir%\%SYS_ScriptNameNoExt%.ini
	IniWrite, %SUS_AutoSuspend%, %CFG_IniFile%, Main, AutoSuspend
	IniWrite, %XWN_FocusFollowsMouse%, %CFG_IniFile%, WindowHandling, FocusFollowsMouse
	IniWrite, %SYS_ToolTipFeedback%, %CFG_IniFile%, Visual, ToolTipFeedback
	IniWrite, %UPD_LastUpdateCheck%, %CFG_IniFile%, UpdateCheck, LastUpdateCheck
	IniWrite, %CFG_LeftMouseButtonHook%, %CFG_IniFile%, MouseHooks, LeftMouseButton
	IniWrite, %CFG_MiddleMouseButtonHook%, %CFG_IniFile%, MouseHooks, MiddleMouseButton
	IniWrite, %CFG_RightMouseButtonHook%, %CFG_IniFile%, MouseHooks, RightMouseButton
	IniWrite, %CFG_FourthMouseButtonHook%, %CFG_IniFile%, MouseHooks, FourthMouseButton
	IniWrite, %CFG_FifthMouseButtonHook%, %CFG_IniFile%, MouseHooks, FifthMouseButton
Return

CFG_ApplySettings:
	If ( SUS_AutoSuspend )
		SetTimer, SUS_SuspendHandler, 1000
	Else
		SetTimer, SUS_SuspendHandler, Off
		
	If ( XWN_FocusFollowsMouse )
		SetTimer, XWN_FocusHandler, 100
	Else
		SetTimer, XWN_FocusHandler, Off
		
	If ( CFG_LeftMouseButtonHook )
		CFG_LeftMouseButtonHookStr = On
	Else
		CFG_LeftMouseButtonHookStr = Off

	If ( CFG_MiddleMouseButtonHook )
		CFG_MiddleMouseButtonHookStr = On
	Else
		CFG_MiddleMouseButtonHookStr = Off

	If ( CFG_RightMouseButtonHook )
		CFG_RightMouseButtonHookStr = On
	Else
		CFG_RightMouseButtonHookStr = Off

	If ( CFG_FourthMouseButtonHook )
		CFG_FourthMouseButtonHookStr = On
	Else
		CFG_FourthMouseButtonHookStr = Off

	If ( CFG_FifthMouseButtonHook )
		CFG_FifthMouseButtonHookStr = On
	Else
		CFG_FifthMouseButtonHookStr = Off
	
	Hotkey, $LButton, %CFG_LeftMouseButtonHookStr%
	Hotkey, $^LButton, %CFG_LeftMouseButtonHookStr%
	Hotkey, #LButton, %CFG_LeftMouseButtonHookStr%
	Hotkey, #^LButton, %CFG_LeftMouseButtonHookStr%
	
	Hotkey, #MButton, %CFG_MiddleMouseButtonHookStr%
	Hotkey, #^MButton, %CFG_MiddleMouseButtonHookStr%
	Hotkey, $MButton, %CFG_MiddleMouseButtonHookStr%
	Hotkey, $^MButton, %CFG_MiddleMouseButtonHookStr%
	
	Hotkey, $RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $+RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $+!RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $+^RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $+#RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $+!^RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $+!#RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $+^#RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $+!^#RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $!RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $!^RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $!#RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $!^#RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $^RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $^#RButton, %CFG_RightMouseButtonHookStr%
	Hotkey, $#RButton, %CFG_RightMouseButtonHookStr%
	
	Hotkey, $XButton1, %CFG_FourthMouseButtonHookStr%
	Hotkey, $^XButton1, %CFG_FourthMouseButtonHookStr%
	
	Hotkey, $XButton2, %CFG_FifthMouseButtonHookStr%
	Hotkey, $^XButton2, %CFG_FifthMouseButtonHookStr%
Return



; [UPD] checks for a new build

UPD_CheckForUpdate:
	UPD_CheckSuccess =
	Random, UPD_Random
	If ( TEMP )
		UPD_BuildFile = %TEMP%\%SYS_ScriptNameNoExt%.%UPD_Random%.tmp
	Else
		UPD_BuildFile = %SYS_ScriptDir%\%SYS_ScriptNameNoExt%.%UPD_Random%.tmp
	Gosub, SUS_SuspendSaveState
	Suspend, On
	URLDownloadToFile, http://www.enovatic.org/products/niftywindows/files/build.txt?random=%UPD_Random%, %UPD_BuildFile%
	Gosub, SUS_SuspendRestoreState
	If ( !ErrorLevel )
	{
		FileReadLine, UPD_Build, %UPD_BuildFile%, 1
		If ( !ErrorLevel )
			If UPD_Build is digit
			{
				UPD_CheckSuccess = 1
				UPD_LastUpdateCheck = %A_MM%
				If ( UPD_Build > SYS_ScriptBuild )
				{
					SYS_TrayTipText = There is a new version available. Please check website.
					SYS_TrayTipOptions = 1
					Run, http://www.enovatic.org/products/niftywindows/
				}
				Else
					SYS_TrayTipText = There is no new version available.
			}
			Else
				SYS_TrayTipText = wrong build pattern in downloaded build file
		Else
			SYS_TrayTipText = downloaded build file couldn't be read
	}
	Else
		SYS_TrayTipText = build file couldn't be downloaded
	FileDelete, %UPD_BuildFile%
	If ( !UPD_CheckSuccess )
	{
		SYS_TrayTipText = Check for update failed:`n%SYS_TrayTipText%
		SYS_TrayTipOptions = 3
	}
	Gosub, SYS_TrayTipShow
Return

UPD_AutoCheckForUpdate:
	If ( UPD_LastUpdateCheck != A_MM )
	{
		Gosub, SUS_SuspendSaveState
		Suspend, On
		MsgBox, 4132, Update Handler - %SYS_ScriptInfo%, You haven't checked for updates for a long period of time (at least one month).`n`nDo you want NiftyWindows to check for a new version now (highly recommended)?
		Gosub, SUS_SuspendRestoreState
		IfMsgBox, Yes
			Gosub, UPD_CheckForUpdate
		Else
			UPD_LastUpdateCheck = %A_MM%
	}
Return

^#!b::
	If ( !A_IsCompiled )
	{
		UPD_VersionFile = %SYS_ScriptDir%\version.txt
		IfExist, %UPD_VersionFile%
		{
			FileDelete, %UPD_VersionFile%
			If ( ErrorLevel )
				Return
		}
		FileAppend, %SYS_ScriptVersion%, %UPD_VersionFile%
		If ( ErrorLevel )
			Return
			
		UPD_BuildFile = %SYS_ScriptDir%\build.txt
		IfExist, %UPD_BuildFile%
		{
			FileDelete, %UPD_BuildFile%
			If ( ErrorLevel )
				Return
		}
		FileAppend, %A_NowUTC%, %UPD_BuildFile%
		If ( ErrorLevel )
			Return
		
		SYS_TrayTipText = Version and build files were written successfully:`n%UPD_VersionFile%`n%UPD_BuildFile%
		SYS_TrayTipOptions = 2
		SYS_TrayTipSeconds = 5
		Gosub, SYS_TrayTipShow
	}
Return
