Accessor_NotepadPlusPlus_Init(ByRef NotepadPlusPlus, PluginSettings)
{	
	path := GetNotepadPlusPlusPath()
	if(path)
		NotepadPlusPlus.Icon := ExtractIcon(path "\Notepad++.exe", 1, 64)
	NotepadPlusPlus.Settings.Keyword := "np"
	NotepadPlusPlus.DefaultKeyword := "np"
	NotepadPlusPlus.KeywordOnly := false
	NotepadPlusPlus.MinChars := 0 ;This is actually 2, but not when notepad is active
	NotepadPlusPlus.MRUList := Array()
	NotepadPlusPlus.OKName := "Open Tab"
	NotepadPlusPlus.Settings.FuzzySearch := PluginSettings.HasKey("FuzzySearch") ? PluginSettings.FuzzySearch : 0
	NotepadPlusPlus.Description := "Activate a specific Notepad++ tab by typing a part of its name. This plugin restores the text `nwhich was previously entered when the current tab was last active. `nThis way you can quicly switch between the most used tabs."
	NotepadPlusPlus.HasSettings := True
}
Accessor_NotepadPlusPlus_ShowSettings(NotepadPlusPlus, PluginSettings, PluginGUI)
{
	AddControl(PluginSettings, PluginGUI, "Edit", "Keyword", "", "", "Keyword:")
	AddControl(PluginSettings, PluginGUI, "Edit", "BasePriority", "", "", "Base Priority:")
	AddControl(PluginSettings, PluginGUI, "Checkbox", "FuzzySearch", "Use fuzzy search (slower)", "", "")
}
Accessor_NotepadPlusPlus_IsInSinglePluginContext(NotepadPlusPlus, Filter, LastFilter)
{
}
Accessor_NotepadPlusPlus_GetDisplayStrings(NotepadPlusPlus, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
}
Accessor_NotepadPlusPlus_OnAccessorOpen(NotepadPlusPlus, Accessor)
{
	global AccessorEdit, AccessorListView
	NotepadPlusPlus.List1 := GetListOfOpenNotepadPlusPlusTabs()
	NotepadPlusPlus.List2 := GetListOfOpenNotepadPlusPlusTabs(2)
	NotepadPlusPlus.Priority := NotepadPlusPlus.Settings.BasePriority
	if(!NotePadPlusPlus.Icon)
	{
		path := GetNotepadPlusPlusPath()
		NotepadPlusPlus.Icon := ExtractIcon(path "\Notepad++.exe", 1, 64)
	}
	;if Notepad++ is open and there is an entry with the last usd command for the current tab, put it in edit box
	if(WinExist("ahk_class Notepad++") = Accessor.PreviousWindow)
	{
		NotepadPlusPlus.Priority := 10000
		if(index := NotepadPlusPlus.MRUList.FindKeyWithValue("Path", Path := GetNotepadPlusPlusActiveTab()))
		{
			GUINum := Accessor.GUINum
			if(!GUINum)
				return
			Gui, %GUINum%: Default
			NotepadPlusPlus.Waiting := true
			GuiControl,, AccessorEdit, % NotepadPlusPlus.MRUList[index].Command
			hwndEdit := Accessor.hwndEdit
			SendMessage, 0xC1, -1,,, AHK_ID %hwndEdit%  ; EM_LINEINDEX (Gets index number of line)
			CaretTo := ErrorLevel
			SendMessage, 0xB1, 0, CaretTo,, AHK_ID %hwndEdit% ;EM_SETSEL
			
			while(NotepadPlusPlus.Waiting || !Accessor.ListViewReady)
				Sleep 10
			Gui, ListView, AccessorListView
			Entry := NotepadPlusPlus.MRUList[index].Entry
			Loop % LV_GetCount()
			{
				LV_GetText(Path,A_Index,4)
				LV_GetText(Type,A_Index,5)
				if(Path = Entry && Type = "NP++")
				{
					LV_Modify(A_Index, "Select")
					break
				}
			}
		}
	}
}
Accessor_NotepadPlusPlus_OnAccessorClose(NotepadPlusPlus, Accessor)
{
}
Accessor_NotepadPlusPlus_OnExit(NotepadPlusPlus)
{
	DestroyIcon(NotepadPlusPlus.Icon)
}
Accessor_NotepadPlusPlus_FillAccessorList(NotepadPlusPlus, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet)
{
	if(NotepadPlusPlus.Waiting)
	{
		NotepadPlusPlus.Waiting := false
	}
	if(!WinActive("ahk_class Notepad++") && strLen(Filter) < 2 && !KeywordSet)
		return
	if(!NotepadPlusPlus.List1.MaxIndex() && !NotepadPlusPlus.List2.MaxIndex())
		return
	outputdebug count %IconCount%
	ImageList_ReplaceIcon(Accessor.ImageListID, -1, NotepadPlusPlus.Icon)
	IconCount++
	if(!Filter && KeywordSet)
	{
		Loop % NotepadPlusPlus.List1.MaxIndex()
		{
			Path := NotepadPlusPlus.List1[A_Index]
			SplitPath, Path, Name
			Accessor.List.Insert(Object("Title",Name,"Path",Path, "Type","NotepadPlusPlus", "Detail1", "NP++", "Icon", IconCount))
		}
		Loop % NotepadPlusPlus.List2.MaxIndex()
		{
			Path := NotepadPlusPlus.List2[A_Index]
			SplitPath, Path, Name
			Accessor.List.Insert(Object("Title",Name,"Path",Path, "Type","NotepadPlusPlus", "Detail1", "NP++", "Icon", IconCount))
		}
		return
	}
	InStrList := Array()
	FuzzyList := Array()
	Loop % NotepadPlusPlus.List1.MaxIndex()
	{
		Path := NotepadPlusPlus.List1[A_Index]
		SplitPath, Path, Name
		pos := InStr(Name, Filter)
		if(pos = 1)
			Accessor.List.Insert(Object("Title",Name,"Path",Path, "Type","NotepadPlusPlus", "Detail1", "NP++", "Icon", IconCount))
		else if(pos > 1)
			InStrList.Insert(Object("Title",Name,"Path",Path, "Type","NotepadPlusPlus", "Detail1", "NP++", "Icon", IconCount))
		else if(NotepadPlusPlus.Settings.FuzzySearch && FuzzySearch(Name,Filter) < 0.3)
			FuzzyList.Insert(Object("Title",Name,"Path",Path, "Type","NotepadPlusPlus", "Detail1", "NP++", "Icon", IconCount))
	}
	Loop % NotepadPlusPlus.List2.MaxIndex()
	{
		Path := NotepadPlusPlus.List2[A_Index]
		SplitPath, Path, Name		
		pos := InStr(Name, Filter)
		if(pos = 1)
			Accessor.List.Insert(Object("Title",Name,"Path",Path, "Type","NotepadPlusPlus", "Detail1", "NP++", "Icon", IconCount))
		else if(pos > 1)
			InStrList.Insert(Object("Title",Name,"Path",Path, "Type","NotepadPlusPlus", "Detail1", "NP++", "Icon", IconCount))
		else if(NotepadPlusPlus.Settings.FuzzySearch && FuzzySearch(Name,Filter) < 0.3)
			FuzzyList.Insert(Object("Title",Name,"Path",Path, "Type","NotepadPlusPlus", "Detail1", "NP++", "Icon", IconCount))
	}
	Accessor.List.Extend(InStrList)
	Accessor.List.Extend(FuzzyList)
	return
}
Accessor_NotepadPlusPlus_PerformAction(NotepadPlusPlus, Accessor, AccessorListEntry)
{
	global AccessorEdit
	if(WinExist("ahk_class Notepad++") = Accessor.PreviousWindow)
	{
		GUINum := Accessor.GUINum
		if(GUINum)
		{
			Gui, %GUINum%: Default
			GuiControlGet, Filter, , AccessorEdit
			if(!(index := NotepadPlusPlus.MRUList.FindKeyWithValue("Path", ActiveTab := GetNotepadPlusPlusActiveTab())))
				NotepadPlusPlus.MRUList.Insert(Object("Path", ActiveTab, "Command", Filter, "Entry", AccessorListEntry.Path))
			else
			{
				NotepadPlusPlus.MRUList[index].Command := Filter
				NotepadPlusPlus.MRUList[index].Entry := AccessorListEntry.Path
			}
		}
	}
	if(strEndsWith(AccessorListEntry.Path, "[2]"))
		ActivateNotepadPlusPlusTab(NotepadPlusPlus.List2.indexOf(AccessorListEntry.Path), 2)
	else
		ActivateNotepadPlusPlusTab(NotepadPlusPlus.List1.indexOf(AccessorListEntry.Path), 1)
	
	return
}
Accessor_NotepadPlusPlus_ListViewEvents(NotepadPlusPlus, AccessorListEntry)
{
}
Accessor_NotepadPlusPlus_EditEvents(NotepadPlusPlus, AccessorListEntry, Filter, LastFilter)
{
	return false
}
Accessor_NotepadPlusPlus_SetupContextMenu(NotepadPlusPlus, AccessorListEntry)
{
	Menu, AccessorMenu, add, Open Tab,AccessorOK
	Menu, AccessorMenu, Default,Open Tab
	Menu, AccessorMenu, add, Open File,NotepadPlusPlusOpenFile
	Menu, AccessorMenu, add, Open file path in explorer,AccessorOpenExplorer
	Menu, AccessorMenu, add, Open file path in CMD,AccessorOpenCMD
	Menu, AccessorMenu, add, Copy file path (CTRL+C),AccessorCopyPath
	Menu, AccessorMenu, add, Explorer context menu, AccessorExplorerContextMenu
}
NotepadPlusPlusOpenFile:
NotepadPlusPlusOpenFile()
return

NotepadPlusPlusOpenFile()
{
	global Accessor, AccessorListView
	GUINum := Accessor.GUINum
	Gui, %GUINum%: Default
	Gui, ListView, AccessorListView
	selected := LV_GetNext()
	if(!selected)
		return
	LV_GetText(id,selected,2)
	Run(Accessor.List[id].Path)
}
ActivateNotepadPlusPlusTab(Index, View = 1)
{
	hwnd := WinExist("ahk_class Notepad++")
	if(!hwnd)
		return
	Index--
	WinActivate ahk_id %hwnd%
	view--
	SendMessage, 0x804, %view%, %index%, ,ahk_id %hwnd% ;NPPM_ACTIVATEDOC
}
GetListOfOpenNotepadPlusPlusTabs(WhichView = 1)
{
	IsUnicode := IsNotepadPlusPlusUnicode()
	hwnd := WinExist("ahk_class Notepad++")
	if(!hwnd) ;Notepad++ not running, empty list
		return Array()
	if(WhichView = 1)
		SendMessage, 0x7EF, 0, 1, ,ahk_id %hwnd% ;NPPM_GETNBOPENFILES
	else
		SendMessage, 0x7EF, 0, 2, ,ahk_id %hwnd% ;NPPM_GETNBOPENFILES
	count := Errorlevel ;Count of tabs
	list := Array()
	Loop % count
	{
		pos := A_Index - 1
		if(WhichView = 1)
			SendMessage, 0x823, %pos%, 0, ,ahk_id %hwnd% ;NPPM_GETBUFFERIDFROMPOS
		else
			SendMessage, 0x823, %pos%, 1, ,ahk_id %hwnd% ;NPPM_GETBUFFERIDFROMPOS
		BufferID := Errorlevel
		SendMessage, 0x822, %BufferID%, 0, ,ahk_id %hwnd% ;NPPM_GETFULLPATHFROMBUFFERID
		MsgReply := ErrorLevel > 0x7FFFFFFF ? -(~ErrorLevel) - 1 : ErrorLevel
		if(MsgReply = -1) ;MsgReply = -1 -> Some error occured
			return Array()
		size := IsUnicode ? MsgReply * 2 + 2 : MsgReply + 1
		RemoteBuf_Open(H, hwnd, size)
		RemoteAddress := RemoteBuf_Get(H) ;Get address of buffer
		VarSetCapacity(localBuffer, size, 0) ;Create local buffer of same size
		SendMessage, 0x822, %BufferID%, %RemoteAddress%, ,ahk_id %hwnd% ;NPPM_GETFULLPATHFROMBUFFERID
		RemoteBuf_Read(H, TabName, size, 0) ;Try to read the first string
		if(A_IsUnicode && !IsUnicode)
			Ansi2Unicode(TabName, ConvertedTabName)
		else if(!A_IsUnicode && IsUnicode)
			Unicode2Ansi(TabName, ConvertedTabName)
		if(ConvertedTabName)
			TabName := ConvertedTabName
		if(WhichView != 1)
			TabName .= " [2]"
		list.Insert(TabName)
		RemoteBuf_Close(H) ;Close/free remote buffer
	}
	return list
}
IsNotepadPlusPlusUnicode()
{
	hwnd := WinExist("ahk_class Notepad++")
	if(!hwnd)
		return 0
	MAXPATH := 260
	size := Maxpath * 2 + 2
	RemoteBuf_Open(H, hwnd, size)
	RemoteAddress := RemoteBuf_Get(H) ;Get address of buffer
	VarSetCapacity(localBuffer, size, 0) ;Create local buffer of same size
	SendMessage, 0xFBF, MAXPATH, RemoteAddress, ,ahk_id %hwnd% ;NPPM_GETNPPDIRECTORY
	RemoteBuf_Read(H, path1, size, 0) ;Try to read the first string
	RemoteBuf_Close(H) ;Close/free remote buffer
	if(NumGet(path1,0, "Char") != 0 && NumGet(path1,1, "Char") = 0)
		return 1
	return 0
}
GetNotepadPlusPlusPath()
{
	hwnd := WinExist("ahk_class Notepad++")
	if(!hwnd)
		return ""
	MAXPATH := 260
	size := Maxpath * 2 + 2
	RemoteBuf_Open(H, hwnd, size)
	RemoteAddress := RemoteBuf_Get(H) ;Get address of buffer
	VarSetCapacity(localBuffer, size, 0) ;Create local buffer of same size
	SendMessage, 0xFBF, MAXPATH, RemoteAddress, ,ahk_id %hwnd% ;NPPM_GETNPPDIRECTORY
	RemoteBuf_Read(H, Path, size, 0) ;Try to read the first string
	RemoteBuf_Close(H) ;Close/free remote buffer
	IsUnicode := IsNotepadPlusPlusUnicode()
	if(A_IsUnicode && !IsUnicode)
		Ansi2Unicode(Path, Path)
	return Path
}
;May only be used while Accessor window is visible
GetNotepadPlusPlusActiveTab()
{
	global AccessorPlugins
	NotepadPlusPlus := AccessorPlugins.GetItemWithValue("Type", "NotepadPlusPlus")
	hwnd := WinExist("ahk_class Notepad++")
	if(!hwnd)
		return ""
	RemoteBuf_Open(H, hwnd, 4)
	RemoteAddress := RemoteBuf_Get(H) ;Get address of buffer
	SendMessage, 0x7EC, 0, RemoteAddress, ,ahk_id %hwnd% ;NPPM_GETCURRENTSCINTILLA
	RemoteBuf_Read(H, View, 4, 0) ;Try to read the first string
	RemoteBuf_Close(H) ;Close/free remote buffer
	View := NumGet(View, 0, "UInt")
	if(!View) ;Main View
	{
		SendMessage, 0x7FF, 0, 0, ,ahk_id %hwnd% ;NPPM_GETCURRENTDOCINDEX		
		index := Errorlevel + 1
		outputdebug % "index " index " tab " NotepadPlusPlus.List1[index]
		return NotepadPlusPlus.List1[index]
	}
	else ;Secondary View
	{
		SendMessage, 0x7FF, 0, 1, ,ahk_id %hwnd% ;NPPM_GETCURRENTDOCINDEX		
		index := Errorlevel + 1
		outputdebug % "index " index " tab " NotepadPlusPlus.List2[index]
		return NotepadPlusPlus.List2[index]
	}
}