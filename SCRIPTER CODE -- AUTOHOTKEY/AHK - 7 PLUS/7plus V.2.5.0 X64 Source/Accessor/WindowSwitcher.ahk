Accessor_WindowSwitcher_Init(ByRef WindowSwitcher, PluginSettings)
{
	WindowSwitcher.Settings.Keyword := "switch"
	WindowSwitcher.DefaultKeyword := "switch"
	WindowSwitcher.KeywordOnly := false
	WindowSwitcher.MinChars := 0
	WindowSwitcher.OKName := "Activate"
	WindowSwitcher.Settings.FuzzySearch := PluginSettings.HasKey("FuzzySearch") ? PluginSettings.FuzzySearch : 0
	WindowSwitcher.Settings.IgnoreFileExtensions := PluginSettings.HasKey("IgnoreFileExtensions") ? PluginSettings.IgnoreFileExtensions : 1
	WindowSwitcher.Description := "Activate windows by typing a part of their title or their executable filename. `nThis also shows CPU usage, shows/sets Always on Top state and `nallows to close and kill processes."
	WindowSwitcher.HasSettings := True
}
Accessor_WindowSwitcher_ShowSettings(WindowSwitcher, PluginSettings, PluginGUI)
{
	AddControl(PluginSettings, PluginGUI, "Edit", "Keyword", "", "", "Keyword:")
	AddControl(PluginSettings, PluginGUI, "Edit", "BasePriority", "", "", "Base Priority:")
	AddControl(PluginSettings, PluginGUI, "Checkbox", "FuzzySearch", "Use fuzzy search (slower)", "", "")
	AddControl(PluginSettings, PluginGUI, "Checkbox", "IgnoreFileExtensions", "Ignore .exe extension in program paths", "", "")	
}
Accessor_WindowSwitcher_GetDisplayStrings(WindowSwitcher, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
	Path := AccessorListEntry.ExeName
	Detail1 := "CPU: " AccessorListEntry.CPU "%"
	Detail2 := AccessorListEntry.OnTop
}
Accessor_WindowSwitcher_IsInSinglePluginContext(WindowSwitcher, Filter, LastFilter)
{
	return false
}
Accessor_WindowSwitcher_OnAccessorOpen(WindowSwitcher, Accessor)
{
	WindowSwitcher.List := GetWindowInfo()
	SetTimer, UpdateTimes, 1000	
	WindowSwitcher.Priority := WindowSwitcher.Settings.BasePriority
}
Accessor_WindowSwitcher_OnAccessorClose(WindowSwitcher, Accessor)
{
	;This is apparently not desired for icons obtained by WM_GETICON or GetClassLong since they are shared? See http://msdn.microsoft.com/en-us/library/windows/desktop/ms648063(v=vs.85).aspx
	;~ if(IsObject(WindowSwitcher.List))
		;~ for index, ListEntry in WindowSwitcher.List
			;~ if(ListEntry.Icon != Accessor.GenericIcons.Application)
				;~ DestroyIcon(ListEntry.Icon)
	SetTimer, UpdateTimes, Off
}
Accessor_WindowSwitcher_FillAccessorList(WindowSwitcher, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet)
{
	strippedFilter := WindowSwitcher.Settings.IgnoreFileExtensions ? strTrimRight(Filter, ".exe") : Filter
	FuzzyList := Array()
	Loop % WindowSwitcher.List.MaxIndex()
	{
		x := 0
		window := WindowSwitcher.List[A_Index]
		ExeName := WindowSwitcher.Settings.IgnoreFileExtensions ? strTrimRight(window.ExeName,".exe") : window.ExeName
		if(x := (Filter = "" || InStr(window.Title,Filter) || InStr(ExeName,strippedFilter)) || (WindowSwitcher.Settings.FuzzySearch && FuzzySearch(ExeName,strippedFilter) < 0.4))
		{
			if((IconIndex := ImageList_ReplaceIcon(Accessor.ImageListID, -1, window.Icon)) != -1)
				IconCount++
			else
				IconIndex := 0
			if(x)
				Accessor.List.Insert(Object("Icon", IconIndex + 1, "Title", window.Title, "Path", window.Path, "ExeName", window.ExeName, "CPU", window.CPU, "OnTop", window.OnTop, "Type", "WindowSwitcher", "PID", window.PID, "hwnd", window.hwnd))			
			else
				FuzzyList.Insert(Object("Icon", IconIndex + 1, "Title", window.Title, "Path", window.Path, "ExeName", window.ExeName, "CPU", window.CPU, "OnTop", window.OnTop, "Type", "WindowSwitcher", "PID", window.PID, "hwnd", window.hwnd))			
		}
	}
	Accessor.List.extend(FuzzyList)
}
Accessor_WindowSwitcher_PerformAction(WindowSwitcher, Accessor, AccessorListEntry)
{
	WinActivate % "ahk_id " AccessorListEntry.hwnd
}
Accessor_WindowSwitcher_ListViewEvents(WindowSwitcher, AccessorListEntry)
{
}
Accessor_WindowSwitcher_EditEvents(WindowSwitcher, AccessorListEntry, Filter, LastFilter)
{
	return true
}
Accessor_WindowSwitcher_SetupContextMenu(WindowSwitcher, AccessorListEntry)
{
	Menu, AccessorMenu, add, Activate window,AccessorOK
	Menu, AccessorMenu, Default,Activate window			
	Menu, AccessorMenu, add, End process,WindowSwitcherEndProcess
	Menu, AccessorMenu, add, Close window,WindowSwitcherCloseWindow
	Menu, AccessorMenu, add, Open executable path in explorer,AccessorOpenExplorer
	Menu, AccessorMenu, add, Open executable path in CMD,AccessorOpenCMD
	Menu, AccessorMenu, add, Copy executable path (CTRL+C), AccessorCopyPath
	
	hwnd := window.hwnd
	WinGet, es, ExStyle, ahk_id %hwnd%
	Menu, AccessorMenu, add, Always on top,WindowSwitcherAlwaysOnTop
	if(es & 0x8 > 0)
		Menu, AccessorMenu, Check, Always on top
}
Accessor_WindowSwitcher_OnExit(WindowSwitcher)
{
}
WindowSwitcherEndProcess:
WindowSwitcherEndProcess()
return
WindowSwitcherCloseWindow:
WindowSwitcherCloseWindow()
return
WindowSwitcherAlwaysOnTop:
WindowSwitcherAlwaysOnTop()
return
WindowSwitcherEndProcess()
{
	global Accessor, AccessorListView
	GUINum := Accessor.GUINum
	Gui, %GUINum%: Default
	Gui, ListView, AccessorListView
	selected := LV_GetNext()
	if(!selected)
		return
	LV_GetText(id,selected,2)
	hwnd := Accessor.List[id].hwnd
	WinKill ahk_id %hwnd%
	AccessorClose()
}
WindowSwitcherCloseWindow()
{
	global Accessor, AccessorListView,AccessorPlugins
	GUINum := Accessor.GUINum
	Gui, %GUINum%: Default
	Gui, ListView, AccessorListView
	selected := LV_GetNext()
	if(!selected)
		return
	LV_GetText(id,selected,2)
	hwnd := Accessor.List[id].hwnd
	PostMessage, 0x112, 0xF060,,, ahk_id %hwnd%
	AccessorClose()
}
WindowSwitcherAlwaysOnTop()
{
	global Accessor, AccessorListView
	GUINum := Accessor.GUINum
	Gui, %GUINum%: Default
	Gui, ListView, AccessorListView
	selected := LV_GetNext()
	if(!selected)
		return
	LV_GetText(id,selected,2)
	hwnd := Accessor.List[id].hwnd
	WinSet, AlwaysOnTop, Toggle, ahk_id %hwnd%
	if(Accessor.List[id].OnTop)
		Accessor.List[id].OnTop := ""
	else
		Accessor.List[id].OnTop := "OnTop"
	LV_Modify(selected, "Col6", Accessor.List[id].OnTop)
}
GetWindowInfo()
{
	global WindowList
	WS_EX_CONTROLPARENT :=0x10000
	WS_EX_DLGMODALFRAME :=0x1
	WS_CLIPCHILDREN :=0x2000000
	WS_EX_APPWINDOW :=0x40000
	WS_EX_TOOLWINDOW :=0x80
	WS_DISABLED :=0x8000000
	WS_VSCROLL :=0x200000
	WS_POPUP :=0x80000000
	
	windows := Array()
	DetectHiddenWindows, Off
	WinGet, Window_List, List ; Gather a list of running programs
	hInstance := GetModuleHandle(0)
	Loop, %Window_List%
	{
		wid := Window_List%A_Index%
		WinGetTitle, wid_Title, ahk_id %wid%
		WinGet, Style, Style, ahk_id %wid%

		If ((Style & WS_DISABLED) or ! (wid_Title)) ; skip unimportant windows ; ! wid_Title or 
			Continue

		WinGet, es, ExStyle, ahk_id %wid%
		Parent := Parent := GetParent(wid)
		WinGet, Style_parent, Style, ahk_id %Parent%
		Owner := Owner := GetWindow(wid, 4) ; GW_OWNER = 4
		WinGet, Style_Owner, Style, ahk_id %Owner%

		If (((es & WS_EX_TOOLWINDOW)  and !(Parent)) ; filters out program manager, etc
		or ( !(es & WS_EX_APPWINDOW)
		and (((Parent) and ((Style_parent & WS_DISABLED) =0)) ; These 2 lines filter out windows that have a parent or owner window that is NOT disabled -
		or ((Owner) and ((Style_Owner & WS_DISABLED) =0))))) ; NOTE - some windows result in blank value so must test for zero instead of using NOT operator!
			continue
		if(wid_Title = "Accessor")
			continue
		if(WindowList.HasKey(wid))
			Exe_Name := WindowList[wid].Executable
		else
			WinGet, Exe_Name, ProcessName, ahk_id %wid%
		WinGet, PID, PID, ahk_id %wid%
		FullPath := GetModuleFileNameEx(PID)
		WinGetClass, Win_Class, ahk_id %wid%
		hw_popup := hw_popup := DllCall("GetLastActivePopup", "Ptr", wid)

		Dialog := 0 ; init/reset
		If (Parent and ! Style_parent)
			CPA_file_name := GetCPA_file_name( wid ) ; check if it's a control panel window
		Else
			CPA_file_name =
		If (CPA_file_name or (Win_Class ="#32770") or ((style & WS_POPUP) and (es & WS_EX_DLGMODALFRAME)))
			Dialog =1 ; found a Dialog window
		If (CPA_file_name)
			hIcon := ExtractIcon(hInstance, CPA_file_name, 1)
		Else
			hIcon := GetWindowIcon(wid, 1) ; (window id, whether to get large icons)
		
		windows.Insert(Object("hwnd",wid,"Title", wid_Title, "Class", Win_Class, "Style", Style, "ExStyle", es, "ExeName", Exe_Name, "Path", FullPath, "OnTop", (es&0x8 > 0 ? "On top" : ""), "PID", PID, "Type", "Window", "Icon", hIcon))
	}
	return windows
}
GetWindowIcon(wid, LargeIcons) ; (window id, whether to get large icons)
{
	Local NR_temp, h_icon
	; check status of window - if window is responding or "Not Responding"
	NR_temp =0 ; init
	h_icon =
	Responding := DllCall("SendMessageTimeout", "Ptr", wid, "UInt", 0x0, "Int", 0, "Int", 0, "UInt", 0x2, "UInt", 150, "UInt *", NR_temp) ; 150 = timeout in millisecs
	If (Responding)
	{
		; WM_GETICON values -    ICON_SMALL =0,   ICON_BIG =1,   ICON_SMALL2 =2
		If LargeIcons =1
		{
			SendMessage, 0x7F, 1, 0,, ahk_id %wid%
			h_icon := ErrorLevel
		}
		If ( ! h_icon )
		{
			SendMessage, 0x7F, 2, 0,, ahk_id %wid%
			h_icon := ErrorLevel
			If ( ! h_icon )
			{
				SendMessage, 0x7F, 0, 0,, ahk_id %wid%
				h_icon := ErrorLevel
				If ( ! h_icon )
				{
					If LargeIcons =1
						h_icon := DllCall( "GetClassLong", "Ptr", wid, "int", -14 ) ; GCL_HICON is -14
					If ( ! h_icon )
					{
						h_icon := DllCall( "GetClassLong", "Ptr", wid, "int", -34 ) ; GCL_HICONSM is -34
						If ( ! h_icon )
							h_icon := DllCall( "LoadIcon", "Ptr", 0, "uint", 32512 ) ; IDI_APPLICATION is 32512
					}
				}
			}
		}
	}
	If ! ( h_icon = "" or h_icon = "FAIL") ; Add the HICON directly to the icon list
		return h_icon
	Else	; use a generic icon
		return Accessor.GenericIcons.Application
}
GetCPA_file_name( p_hw_target ) ; retrives Control Panel applet icon
{
   WinGet, pid_target, PID, ahk_id %p_hw_target%
   hp_target := DllCall( "OpenProcess", "uint", 0x18, "int", false, "uint", pid_target, "Ptr")
   hm_kernel32 := GetModuleHandle("kernel32.dll")
   pGetCommandLine := DllCall( "GetProcAddress", "Ptr", hm_kernel32, "str", A_IsUnicode ? "GetCommandLineW"  : "GetCommandLineA")
   buffer_size := 6
   VarSetCapacity( buffer, buffer_size )
   DllCall( "ReadProcessMemory", "Ptr", hp_target, "uint", pGetCommandLine, "uint", &buffer, "uint", buffer_size, "uint", 0 )
   loop, 4
      ppCommandLine += ( ( *( &buffer+A_Index ) ) << ( 8*( A_Index-1 ) ) )
   buffer_size := 4
   VarSetCapacity( buffer, buffer_size, 0 )
   DllCall( "ReadProcessMemory", "Ptr", hp_target, "uint", ppCommandLine, "uint", &buffer, "uint", buffer_size, "uint", 0 )
   loop, 4
      pCommandLine += ( ( *( &buffer+A_Index-1 ) ) << ( 8*( A_Index-1 ) ) )
   buffer_size := 260
   VarSetCapacity( buffer, buffer_size, 1 )
   DllCall( "ReadProcessMemory", "Ptr", hp_target, "uint", pCommandLine, "uint", &buffer, "uint", buffer_size, "uint", 0 )
   DllCall( "CloseHandle", "Ptr", hp_target )
   IfInString, buffer, desk.cpl ; exception to usual string format
     return, "C:\WINDOWS\system32\desk.cpl"

   ix_b := InStr( buffer, "Control_RunDLL" )+16
   ix_e := InStr( buffer, ".cpl", false, ix_b )+3
   StringMid, CPA_file_name, buffer, ix_b, ix_e-ix_b+1
   if ( ix_e )
      return, CPA_file_name
   else
      return, false
}

GetSystemTimes()    ;Total CPU Load
{
   Static oldIdleTime, oldKrnlTime, oldUserTime
   Static newIdleTime, newKrnlTime, newUserTime

   oldIdleTime := newIdleTime
   oldKrnlTime := newKrnlTime
   oldUserTime := newUserTime

   DllCall("GetSystemTimes", "int64P", newIdleTime, "int64P", newKrnlTime, "int64P", newUserTime)
   Return (1 - (newIdleTime-oldIdleTime)/(newKrnlTime-oldKrnlTime + newUserTime-oldUserTime)) * 100
}
UpdateTimes:
UpdateCPUTimes()
return

UpdateCPUTimes()
{
	global AccessorListView, Accessor, AccessorPlugins
	GUINum := Accessor.GUINum
	if(!GuiNum)
		return
	Gui, %GUINum%: Default
	GUI, ListView, AccessorListView
	Loop % LV_GetCount()
	{
		LV_GetText(index,A_Index,2)
		if(Accessor.List[index].Type = "WindowSwitcher")
		{
			Accessor.List[index].oldKrnlTime := Accessor.List[index].newKrnlTime
			Accessor.List[index].oldUserTime := Accessor.List[index].newUserTime

			hProc := DllCall("OpenProcess", "Uint", 0x400, "int", 0, "Uint", Accessor.List[index].PID, "Ptr")
			DllCall("GetProcessTimes", "Ptr", hProc, "int64P", CreationTime, "int64P", ExitTime, "int64P", newKrnlTime, "int64P", newUserTime, "Ptr")
			DllCall("CloseHandle", "Ptr", hProc)
			Accessor.List[index].newKrnlTime := newKrnlTime
			Accessor.List[index].newUserTime := newUserTime
			Accessor.List[index].CPU := Round(min(max((Accessor.List[index].newKrnlTime-Accessor.List[index].oldKrnlTime + Accessor.List[index].newUserTime-Accessor.List[index].oldUserTime)/10000000 * 100,0),100), 2)   ; 1sec: 10**7
			WindowSwitcher := AccessorPlugins[AccessorPlugins.FindKeyWithValue("Type", "WindowSwitcher")]
			outputdebug % WindowSwitcher.List[WindowSwitcher.List.FindKeyWithValue("hwnd", Accessor.List[index].hwnd)].CPU
			WindowSwitcher.List[WindowSwitcher.List.FindKeyWithValue("hwnd", Accessor.List[index].hwnd)].CPU := Accessor.List[index].CPU
			; Accessor.List[index].CPU := GetProcessTimes(Accessor.List[index].PID)
			LV_Modify(A_Index,"Col5","CPU: " Accessor.List[index].CPU "%")
		}
	}
	if(usage := Round(GetSystemTimes(),2))
	SB_SetText(" CPU Usage: " usage "%",1,2)
	return
}