; ======================================================================================================================
; AHK 1.1.04 +
; ======================================================================================================================
; Function:          Support for SysLink controls - http://msdn.microsoft.com/en-us/library/bb760704%28VS.85%29.aspx
;                    They may be useful in some cases, but they offer very few options for customizing.
; AHK version:       1.1.04.00 (U32)
; Language:          English
; Tested on:         Win XPSP3, Win VistaSP2 (32 Bit)
; Version:           0.0.02.012/2011-09-16/just me
; How to use:        To create a new Syslink control use MySysLink := New SysLink_Control()
;                    passing up to four parameters:
;						Name - Name of the control
;                       Text   - Text to show in the control                              (String)
;                       ----------- Optional ---------------------------------------------------------
;                       Options   - Options for positioning and sizing (X, Y, W, H)          (String)
;                                   as used in "Gui, Add" command.
;                                   Defaults: X - automatically positioned by GUI
;                                             Y - automatically positioned by GUI
;                                             W - 200
;                                             H - 0 (the height will be adjusted to fit the text)
;                                   The options "BackgroundTrans" and "Border" are supported also.
;                       Function  - Function to be called on WM_NOTIFY messages              (String)
;                                   Default: "_SysLink_Notifications"
;                    On success   - New returns an object with six public keys:
;                       HWND      - HWND of the control                                      (Pointer)
;                       X         - X-position                                               (Integer)
;                       Y         - Y-position                                               (Integer)
;                       W         - Width of the control                                     (Integer)
;                       H         - Height of the control                                    (Integer)
;                       ID        - Id of the control (>= 0x4000)                            (Integer)
;                    On failure   - New returns False, ErrorLevel contains additional informations.
;
;                    To destroy the control afterwards simply assign an empty string or zero (e.g. MySysLink := "").
;
; Remarks:           MSDN: "The SysLink control supports the anchor tag(<a>) along with the attributes HREF and ID.
;                    An HREF can be any protocol, such as http, ftp, and mailto. An ID is an optional name,
;                    unique within a SysLink control, and it is associated with an individual link.
;                    Links are also assigned a zero-based index according to their position within the string."
;
;                    The scripts provides the function _SysLink_Notifications() as default handler for
;                    NM_CLICK and NM_RETURN notifications (the only notifications associated with the control).
;                    This function only will run clicked URLs. If you want to do another action or to respond also
;                    to clicks on IDs, you have to pass the name of a function accepting exactly four parameters
;                    (Hwnd, Msg, wParam, lParam) to be called instead.
;
;                    This class is using Class_Notification.ahk from http://www.autohotkey.com/forum/topic75591.html
;                    because I think it makes things easier. If you want to handle other notifications in your script,
;                    you should use it too.
; ======================================================================================================================
; This software is provided 'as-is', without any express or implied warranty.
; In no event will the authors be held liable for any damages arising from the use of this software.
; ======================================================================================================================

/*
Class: CSysLinkControl
A class for displaying hyperlinks in text.

Property: Text
Simply use <A HREF="URL">URL Text</A> notation inside the text to show a link. This control supports multiple links and regular text as well.

Event: Click(URL)
Called when a link is clicked. If this is not handled the control will try to open the URL.

*/
Class CSysLinkControl extends CControl {
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; PRIVATE Properties and Methods ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   ; ===================================================================================================================
   ; CONSTRUCTOR     __New()
   ; ===================================================================================================================
   __New(Name, Options, Text, GUINum)
	{
		Static INIT := False
		Static HDEFFONT := 0
		Static DEFAULT_GUI_FONT := 17
		Static ICC_LINK_CLASS := 0x00008000
		Static IDC_SYSLINK := 0x4000
		Static WC_LINK    := "SysLink"
		Static LWS_TRANSPARENT := 0x0001
		Static WM_USER := 0x400
		Static LM_GETIDEALHEIGHT := (WM_USER + 0x301)
		Static WM_SETFONT := 0x0030
		Static WM_GETFONT := 0x0031
		Static NM_FIRST   := 0
		Static NM_CLICK   := (NM_FIRST-2)
		Static NM_RETURN  := (NM_FIRST-4)
		Static WS_TABSTOP := 0x00010000
		Static WS_BORDER  := 0x00800000
		Static WS_VISIBLE := 0x10000000
		Static WS_CHILD   := 0x40000000
		Static DefaultW := 200
		Static DefaultH := 0
		base.__New(Name, Options, Text, GUINum)
		this.Type := "SysLink"
		HGUI := CGUI.GUIList[GUINum].hwnd
		If !(INIT) {
		 VarSetCapacity(ICCX, 8, 0)
		 NumPut(8, ICCX, 0, "UInt")
		 NumPut(ICC_LINK_CLASS, ICCX, 4, "UInt")
		 If !DllCall("Comctl32.dll\InitCommonControlsEx", "Ptr", &ICCX)
			Return False
		 HDEFFONT := DllCall("Gdi32.dll\GetStockObject", "Int", DEFAULT_GUI_FONT)
		 INIT := True
		}
		If !DllCall("User32.dll\IsWindow", "UPtr", HGUI) {
		 ErrorLevel := "Invalid parameter HGUI: " HGUI
		 Return False
		}
		W := DefaultW
		H := DefaultH
		Opts := {X: "", Y: "", W: "", H: ""}
		Options := Trim(RegExReplace(Options, "\s+", " "))
		SL_STYLES := WS_TABSTOP | WS_VISIBLE | WS_CHILD
		Loop, Parse, Options, %A_Space%
		{
		 If (A_LoopField = "BackgroundTrans") {
			SL_STYLES |= LWS_TRANSPARENT
			Continue
		 }
		 If (A_LoopField = "Border") {
			SL_STYLES |= WS_BORDER
			Continue
		 }
		 O := SubStr(A_LoopField, 1, 1)
		 If InStr("XYWH", O)
			Opts[O] := A_LoopField
		}
		If (Opts.W <> "")
		 W := Opts.W
		If (Opts.H <> "")
		 H := Opts.H
		Options := Opts.X " " Opts.Y " " Opts.W " " Opts.H
		Gui, %HGUI%:Add, Text, %Options% hwndHDUMMY, X
		GuiControlGet, CP, Pos, %HDUMMY%
		HFONT := DllCall("SendMessage", "Ptr", HDUMMY, "Int", WM_GETFONT, "Ptr", 0, "Ptr", 0, "UPtr")
		If !(HFONT)
		 HFONT := HDEFFONT
		X := CPX
		Y := CPY
		If (W = Opts.W)
		 W := CPW
		If (H = Opts.H)
		 H := CPH
		HSL := DllCall("CreateWindowExW", "UInt", 0, "WStr", WC_LINK, "WStr", Text, "UInt", SL_STYLES
					, "Int", X, "Int", Y, "Int", W, "Int", H
					, "Ptr", HGUI, "Ptr", IDC_SYSLINK, "Ptr", 0, "Ptr", 0, "Ptr")
		If ((ErrorLevel) || !(HSL)) {
		 ErrorMsg := "Couldn't create SysLink control!`nErrorLevel: " ErrorLevel " - HWND: " HSL
		 ErrorLevel := ErrorMsg
		 Return False
		}
		DllCall("SendMessage", "Ptr", HSL, "Int", WM_SETFONT, "Ptr", HFONT, "Ptr", False)
		If (H = DefaultH) {
		 H := DllCall("SendMessage", "Ptr", HSL, "Int", LM_GETIDEALHEIGHT, "Ptr", W, "Ptr", 0)
		 DllCall("MoveWindow", "Ptr", HSL, "Int", X, "Int", Y, "Int", W, "Int", H, "Int", False)
		 If (H <> CPH)
			Gui, %HGUI%:Add, Text, x%X% y%Y% w%W% h%H% +BackgroundTrans
		}
		This.HWND := HSL
		;~ This.X := X
		;~ This.Y := Y
		;~ This.W := W
		;~ This.H := H
		This.ID := IDC_SYSLINK
		IDC_SYSLINK += 1
		return HSL
   }
   ; ===================================================================================================================
   ; DESTRUCTOR      __Delete()
   ; ===================================================================================================================
   __Delete() {
      Static NM_FIRST  := 0
      Static NM_CLICK  := (NM_FIRST-2)
      Static NM_RETURN := (NM_FIRST-4)
      If (This.HWND) {
         DllCall("DestroyWindow", "Ptr", This.HWND)
      }
   }
   OnWM_NOTIFY(wParam, lParam) {
		; lParam is a pointer to a NMLINK structure containing NMHDR and LITEM structures
		; -> http://msdn.microsoft.com/en-us/library/bb760714%28VS.85%29.aspx
		Static NM_FIRST  := 0
		Static NM_CLICK  := (NM_FIRST-2)
		Static NM_RETURN := (NM_FIRST-4)
		Static MAX_LINKID_TEXT := 48
		Static L_MAX_URL_LENGTH := (2048 + 32 + 3) ; 3 : sizeof("://"))
		Static OffHwnd   := 0
		Static OffIDFrom     := OffHwnd + A_PtrSize
		Static OffCode   := OffIDFrom + A_PtrSize
		Static OffLITEM  := OffCode + 4
		Static OffMask   := OffLITEM
		Static OffItem   := OffMask + 4
		Static OffState  := OffItem + 4
		Static OffStmask := OffState + 4
		Static OffID     := OffStmask + 4
		Static OffUrl    := OffID + (MAX_LINKID_TEXT * 2)
		Code := NumGet(lParam + 0, OffCode, "INT")
		if(Code = NM_CLICK || Code = NM_RETURN)
		{
			Url := StrGet(lParam + OffUrl, L_MAX_URL_LENGTH, "UTF-16")
			Url := Trim(Url)
			if(!this.CallEvent("Click", Url).Handled)
				if(Url)
					Run, *Open %Url%
		}
	}
}