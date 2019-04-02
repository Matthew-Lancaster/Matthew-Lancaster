Class CAccessorPlugin
{
	__New(PluginSettings)
	{
	}
	ShowSettings(PluginSettings, PluginGUI)
	{
	}
	IsInSinglePluginContext(Filter, LastFilter)
	{
	}
	GetDisplayStrings(AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
	{
	}
	OnAccessorOpen(Accessor)
	{
	}
	OnAccessorClose(Accessor)
	{
	}
	OnExit()
	{
	}
	FillAccessorList(Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet)
	{
	}
	PerformAction(Accessor, AccessorListEntry)
	{
	}
	ListViewEvents(AccessorListEntry)
	{
	}
	EditEvents(AccessorListEntry, Filter, LastFilter)
	{
	}
	OnKeyDown()
	{
	}
	SetupContextMenu(Calc, AccessorListEntry)
	{
	}
}
#include %A_ScriptDir%\Accessor\AccessorEventPlugin.ahk
#include %A_ScriptDir%\Accessor\Calc.ahk
#include %A_ScriptDir%\Accessor\FastFolders.ahk
#include %A_ScriptDir%\Accessor\FileSystem.ahk
#include %A_ScriptDir%\Accessor\Google.ahk
#include %A_ScriptDir%\Accessor\Notepad++.ahk
#include %A_ScriptDir%\Accessor\Notes.ahk
#include %A_ScriptDir%\Accessor\ProgramLauncher.ahk
#include %A_ScriptDir%\Accessor\SciTE4AutoHotkey.ahk
#include %A_ScriptDir%\Accessor\WindowSwitcher.ahk
#include %A_ScriptDir%\Accessor\Uninstall.ahk
#include %A_ScriptDir%\Accessor\URL.ahk
#include %A_ScriptDir%\Accessor\Weather.ahk
#include %A_ScriptDir%\Accessor\Run.ahk
Accessor_Init()
{
	global AccessorPlugins, Accessor
	AccessorPluginsList := "WindowSwitcher,FileSystem,EventPlugin,Google,Calc,ProgramLauncher,NotepadPlusPlus,SciTE4AutoHotkey,Notes,FastFolders,Uninstall,URL,Weather,Run" ;The order here partly determines the order in the window, so choose carefully
	AccessorPlugins := Array()
	Accessor := Object("Base", Object("OnExit", "Accessor_OnExit", "History", Array(), "EditQueue", Array()))
	FileRead, xml, % Settings.ConfigPath "\Accessor.xml"
	if(!xml)
	{
		xml = 
		( LTrim
			<Keywords>
			<Keyword>
			<Command>http://dict.leo.org/ende?search=${1}</Command>
			<Key>leo</Key>
			</Keyword><Keyword>
			<Command>http://google.com/search?q=${1}</Command>
			<Key>google</Key>
			</Keyword><Keyword>
			<Command>http://en.wikipedia.org/wiki/Special:Search?search=${1}</Command>
			<Key>w</Key>
			</Keyword><Keyword>
			<Command>http://maps.google.com/maps?q=${1}</Command>
			<Key>gm</Key>
			</Keyword><Keyword>
			<Command>http://www.amazon.com/s?url=search-alias`%3Daps&field-keywords=${1}</Command>
			<Key>a</Key>
			</Keyword><Keyword>
			<Command>http://www.bing.com/search?q=${1}</Command>
			<Key>bing</Key>
			</Keyword><Keyword>
			<Command>http://www.youtube.com/results?search_query=${1}</Command>
			<Key>y</Key>
			</Keyword><Keyword>
			<Command>http://www.imdb.com/find?q=${1}</Command>
			<Key>i</Key>
			</Keyword><Keyword>
			<Command>http://www.wolframalpha.com/input/?i=${1}</Command>
			<Key>wa</Key>
			</Keyword></Keywords>
		)
	}
	XMLObject := XML_Read(xml)
	outputdebug Accessor Plugins: %AccessorPluginsList%
	Loop, Parse, AccessorPluginsList, `,,%A_Space%
	{		
		tmpobject := RichObject()
		tmpobject.Type := A_LoopField
		tmpobject.Init := "Accessor_" A_LoopField "_Init"
		tmpobject.ReadXML := "Accessor_" A_LoopField "_ReadXML"
		tmpobject.IsInSinglePluginContext := "Accessor_" A_LoopField "_IsInSinglePluginContext"
		tmpobject.FillAccessorList := "Accessor_" A_LoopField "_FillAccessorList"
		tmpobject.OnAccessorOpen := "Accessor_" A_LoopField "_OnAccessorOpen"
		tmpobject.OnAccessorClose := "Accessor_" A_LoopField "_OnAccessorClose"
		tmpobject.GetDisplayStrings := "Accessor_" A_LoopField "_GetDisplayStrings"
		tmpobject.ListViewEvents := "Accessor_" A_LoopField "_ListViewEvents"
		tmpobject.EditEvents := "Accessor_" A_LoopField "_EditEvents"
		tmpobject.PerformAction := "Accessor_" A_LoopField "_PerformAction"
		tmpobject.SetupContextMenu := "Accessor_" A_LoopField "_SetupContextMenu"
		;tmpobject.OnKeyDown := "Accessor_" A_LoopField "_OnKeyDown"
		tmpobject.OnCopy := "Accessor_" A_LoopField "_OnCopy"
		tmpobject.OnExit := "Accessor_" A_LoopField "_OnExit"
		tmpobject.ShowSettings := "Accessor_" A_LoopField "_ShowSettings"
		tmpobject.SaveSettings := "Accessor_" A_LoopField "_SaveSettings"
		tmpobject.GUISubmit := "Accessor_" A_LoopField "_GUISubmit"
		tmpobject.Priority := 0
		tmpobject.HandlesEnter := 0
		tmpobject.OKName := "OK"
		tmpobject.Settings := RichObject()
		tmpobject.Settings.BasePriority := IsObject(XMLObject[A_LoopField]) ? XMLObject[A_LoopField].Settings.BasePriority : 0
		tmpobject.Enabled := IsObject(XMLObject[A_LoopField]) ? XMLObject[A_LoopField].Enabled : 1
		Accessor_%A_LoopField%_Init(tmpobject, XMLObject[A_LoopField].Settings)
		if(IsObject(XMLObject[A_LoopField]))
		{
			if(XMLObject[A_LoopField].Settings.Keyword) ;Set keyword automatically for each plugin to save some lines
				tmpobject.Settings.Keyword := XMLObject[A_LoopField].Settings.Keyword
			if(XMLObject[A_LoopField].Settings.BasePriority) ;Set base priority automatically for each plugin to save some lines
				tmpobject.Settings.BasePriority := XMLObject[A_LoopField].Settings.BasePriority	
		}
		AccessorPlugins.Insert(Object("Base",tmpobject))
	}
		
	Accessor.Keywords := Array()
	if(!IsObject(XMLObject.Keywords))
		XMLObject.Keywords := Object()
	if(!IsObject(XMLObject.Keywords.Keyword) || !XMLObject.Keywords.Keyword.MaxIndex())
		XMLObject.Keywords.Keyword := IsObject(XMLObject.Keywords.Keyword) ? Array(XMLObject.Keywords.Keyword) : Array()
	for index, Keyword in XMLObject.Keywords.Keyword
		Accessor.Keywords.Insert(Object("Key", Keyword.Key, "Command", Keyword.Command))
	Accessor.GenericIcons := Object()
	Accessor.GenericIcons.Application := ExtractIcon("shell32.dll", 3, 64)
	Accessor.GenericIcons.File := ExtractIcon("shell32.dll", 1, 64)
	Accessor.GenericIcons.Folder := ExtractIcon("shell32.dll", 4, 64)
	FileAppend, test, %A_Temp%\7plus\test.htm
	Accessor.GenericIcons.URL := ExtractAssociatedIcon(0, A_Temp "\7plus\test.htm", iIndex)
	FileDelete, %A_Temp%\7plus\test.htm
}
Accessor_OnExit(Accessor)
{
	global AccessorPlugins
	if(Accessor.GUINum)
		AccessorClose()
	
	FileDelete, % Settings.ConfigPath "\Accessor.xml"
	XMLObject := Object()
	for index, AccessorPlugin in AccessorPlugins
	{
		XMLObject[AccessorPlugin.Type] := Object("Enabled", AccessorPlugin.Enabled, "Keyword", AccessorPlugin.Settings.Keyword, "Settings", AccessorPlugin.Settings)
		AccessorPlugin.OnExit()
	}
	XMLObject.Keywords := Object("Keyword",Array())
	for index, Keyword in Accessor.Keywords
		XMLObject.Keywords.Keyword.Insert(Object("Key", Keyword.Key, "Command", Keyword.Command))
	XML_Save(XMLObject, Settings.ConfigPath "\Accessor.xml")
	
	DestroyIcon(Accessor.GenericIcons.Application)
	DestroyIcon(Accessor.GenericIcons.File)
	DestroyIcon(Accessor.GenericIcons.Folder)
	DestroyIcon(Accessor.GenericIcons.URL)
}
;Non blocking accessor window (can wait for closing in event system though)
CreateAccessorWindow(Action)
{
	global AccessorListView, Accessor, AccessorPlugins, AccessorOKButton
	start := A_TickCount
	if(AccessorGUINum := Accessor.GUINum)
	{
		gui %AccessorGUINum%:+LastFoundExist
		If(WinExist())
			return AccessorGUINum
	}
	AccessorGUINum:=10
	DetectHiddenWindows, On
    loop
	{
		; Window available?
		gui %AccessorGUINum%:+LastFoundExist
		If(!WinExist())
			break

		; Nothing available?
		if(AccessorGUINum=99)
		{
			MsgBox 262160,Too many windows, Unable to create Accessor window. GUI windows 10 to 99 are already in use.
			ErrorLevel=9999
			return ""
		}

		;-- Increment window
		AccessorGUINum++
	}
	active := WinExist("A") ;Active window for fastfolder purposes later
	Gui, %AccessorGUINum%: Default
	Gui, Destroy 
	Gui, Add,Edit, w800 y10 -Multi gAccessorEdit vAccessorEdit hwndHWNDEdit
	Gui, Add,ListView, w800 y+10 AltSubmit 0x8 -Multi R15 NoSortHdr gAccessorListView vAccessorListView hwndHWNDListView, #| |Title|Path| | |
	Gui, Add,Button, y10 x+10 w75 gAccessorOK vAccessorOKButton, OK
	Gui, Add,Button,y+8 w75 gAccessorCancel, Cancel
	Gui, Add,StatusBar
	Gui, -MinimizeBox -MaximizeBox +LabelAccessor +AlwaysOnTop -Caption +Border
	Action.tmpGuiNum := AccessorGUINum
	Accessor.WindowTitle := "7Plus Accessor"
	Gui, Show,, % Accessor.WindowTitle
	Accessor.GUINum := AccessorGUINum
	Accessor.hwndEdit := hwndEdit
	Accessor.hwndListView := hwndListView	
	Accessor.PreviousWindow := Active
	Accessor.LauncherHotkey := Action.LauncherHotkey
	Accessor.History.Index := 1
	Accessor.History[1] := ""
	for index, AccessorPlugin in AccessorPlugins
		AccessorPlugin.OnAccessorOpen(Accessor)
	;Check if a plugin set a custom filter
	GuiControlGet, Filter, , AccessorEdit
	if(!Filter)
		FillAccessorList()
	OnMessage(0x06,"WM_ACTIVATE")
	old := OnMessage(0x100)
	OnMessage(0x100, "")
	Accessor.OldKeyDown := old
	;OnMessage(0x100, "Accessor_WM_KEYDOWN")
	;return Gui number to indicate that the Accessor box is still open
	return AccessorGUINum
}

;This is the main function that populates the Accessor list.
FillAccessorList()
{
	global AccessorPlugins,Accessor,AccessorListView,AccessorEdit
	if(Accessor.SuppressListViewUpdate)
	{
		Accessor.SuppressListViewUpdate := false
		return
	}
	LastFilter := Accessor.LastFilter
	guinum := Accessor.GUINum
	if(!guinum)
		return
	Gui, %guinum%: Default
	Gui, ListView, AccessorListView	
	GuiControlGet, BaseFilter, , AccessorEdit
	
	for Index, Keyword in Accessor.Keywords
	{
		;if filter starts with keyword and ends directly after it or has a space after it
		if(InStr(BaseFilter, Keyword.Key) = 1 && (strlen(BaseFilter) = strlen(Keyword.Key) || InStr(BaseFilter, " ") = strLen(Keyword.Key) + 1))
		{
			Filter := StringReplace(BaseFilter, Keyword.Key, Keyword.Command)
			;Code below treats the ${1}-${9} placeholders that may be used with URLs (mostly for Accessor keywords, e.g. to launch a search engine url with a specific query).
			;~ for index, Parameter in Parameters
				;~ Filter := StringReplace(Filter, "${" index "}", Parameter)
			UsingKeyword := true
			break
		}
	}
	
	Filter := Filter ? Filter : BaseFilter
	
	;Parse parameters. They are split by spaces. Quotes (" ") can be used to treat multiple words as one parameter. The first parameter is the Filter variable without the options.
	Parameters := Array()
	p0 := Parse(Filter, "q"")1 2 3 4 5 6 7 8 9 10", p1, p2, p3, p4, p5, p6, p7, p8, p9, p10)
	
	Loop % min(p0, 10)
		Parameters.Insert(A_Index - 1, p%A_Index%) ;Store parameters with offset of -1, so p1 will become p0 since it isn't a real parameter but rather the keyword
	
	if(UsingKeyword)
	{
		if(InStr(Filter, "${1}"))
			Filter := p1 ;If atleast one placeholder is used, all parameters will be inserted 
		;Code below treats the ${1}-${10} placeholders that may be used with keywords, e.g. to launch a search engine url with a specific query.
		for Index, Parameter in Parameters
		{
			if(Index = 0)
				continue
			;If this is the last placeholder used in the query, insert all parameters into it so queries with spaces become possible for the last placeholder
			if(InStr(Filter, "${" Index "}") && !InStr(Filter, "${" (Index+1) "}"))
			{
				CollectedParameters := Parameter
				Loop % Parameters.MaxIndex() - Index
					CollectedParameters .= " " Parameters[A_Index + Index]
				Filter := StringReplace(Filter, "${" Index "}", CollectedParameters, "ALL")
				UsingPlaceholder := true
				break
			}
			else if(InStr(Filter, "${" Index "}"))
			{
				Filter := StringReplace(Filter, "${" Index "}", Parameter, "ALL")
				UsingPlaceholder := true
			}
			else
				break
		}
		if(UsingPlaceholder)
			Parameters := Array() ;Clear parameters since they are now integrated in the query
	}
	Accessor.ListViewReady := false
	GuiControl, -Redraw, AccessorListView
	LV_Delete()
	if(Accessor.ImageListID)
		IL_Destroy(Accessor.ImageListID)
	Accessor.ImageListID := IL_Create(20,5,1) ; Create an ImageList so that the ListView can display some icons
	IconCount := 4
	ImageList_ReplaceIcon(Accessor.ImageListID, -1, Accessor.GenericIcons.Application)
	ImageList_ReplaceIcon(Accessor.ImageListID, -1, Accessor.GenericIcons.Folder)
	ImageList_ReplaceIcon(Accessor.ImageListID, -1, Accessor.GenericIcons.URL)
	ImageList_ReplaceIcon(Accessor.ImageListID, -1, Accessor.GenericIcons.File)
	Accessor.List := Array()
	;Find out if we are in a single plugin context, and add only those items
	for index, Plugin in AccessorPlugins
	{
		if(Plugin.Enabled && SingleContext := ((Plugin.Settings.Keyword && Filter && KeywordSet := InStr(Filter, Plugin.Settings.Keyword " ") = 1) || Plugin.IsInSinglePluginContext(Filter, LastFilter)))
		{
			Accessor.SingleContext := Plugin.Type
			Filter := strTrim(Filter, Plugin.Settings.Keyword " ")
			Plugin.FillAccessorList(Accessor, Filter, LastFilter, IconCount, KeywordSet, Parameters)
			break
		}
	}
	;If we aren't, let all plugins add the items we want according to their priorities
	if(!SingleContext)
	{
		Accessor.SingleContext := false
		Pluginlist := ""
		Loop % AccessorPlugins.MaxIndex()
			Pluginlist .= (Pluginlist ? "," : "") A_Index
		Sort, Pluginlist, F AccessorPrioritySort D`,
		Loop, Parse, Pluginlist, `,
		{
			Plugin := AccessorPlugins[A_LoopField]
			if(Plugin.Enabled && !Plugin.KeywordOnly && StrLen(Filter) >= Plugin.MinChars)
				Plugin.FillAccessorList(Accessor, Filter, LastFilter, IconCount, False, Parameters)
		}
	}
	;Now that items are added, add them to the listview
	LV_SetImageList(Accessor.ImageListID, 1) ; Attach the ImageLists to the ListView so that it can later display the icons
	Loop % Accessor.List.MaxIndex()
	{
		AccessorListEntry := Accessor.List[A_Index]
		AccessorPlugin := AccessorPlugins.GetItemWithValue("Type", AccessorListEntry.Type)
		AccessorPlugin.GetDisplayStrings(AccessorListEntry, Title := AccessorListEntry.Title, Path := AccessorListEntry.Path, Detail1 := AccessorListEntry.Detail1, Detail2 := AccessorListEntry.Detail2)
		LV_Add("Icon" AccessorListEntry.Icon, "", A_Index, Title, Path, Detail1, Detail2)
	}
	Accessor.ListViewReady := true
	Accessor.LastFilter := Filter
	Accessor.LastParameters := Parameters
	selected := LV_GetNext()
	if(!selected)
		LV_Modify(1,"Select")
	LV_ModifyCol()
	LV_ModifyCol(1, "Auto") ; icon column
    LV_ModifyCol(2, 0) ; hidden column for row number    
    LV_ModifyCol(3, "460") ;Col_3_w) ; resize title column
	; SendMessage, 0x1000+29, 2, 0,, % "ahk_id " Accessor.hwndListView ; LVM_GETCOLUMNWIDTH is 0x1000+29
	; Width_Column_3 := ErrorLevel
	; If Width_Column_3 > 430
		; LV_ModifyCol(3, 430) ; resize title column
    ; LV_ModifyCol(4, "Auto") ; exe
    ; SendMessage, 0x1000+29, 3, 0,, % "ahk_id " Accessor.hwndListView ; LVM_GETCOLUMNWIDTH is 0x1000+29
	; Width_Column_4 := ErrorLevel
	; If Width_Column_4 > 200
	LV_ModifyCol(4, 170) ; resize title column
	LV_ModifyCol(5, 55)
	LV_ModifyCol(6, "AutoHdr") ; OnTop
	; LV_ModifyCol(7, 0) ; OnTop
	GuiControl, +Redraw, AccessorListView
}

AccessorPrioritySort(First,Second)
{
	global AccessorPlugins
	return AccessorPlugins[First].Priority = AccessorPlugins[Second].Priority ? First > Second ? 1 : First < Second ? -1 : 0 : AccessorPlugins[First].Priority < AccessorPlugins[Second].Priority ? 1 : -1
}
AccessorEdit:
; Critical, 1000 ;Doesn't seem to have any effect, tried various combinations
Accessor.NeedsUpdate := true
SetTimer, AcccessorEditQueue, -1
return
AcccessorEditQueue:
if(Accessor.NeedsUpdate)
{
	Accessor.NeedsUpdate := false
	AccessorEditEvents()
}
return

AccessorListView:
AccessorListViewEvents()
;~ SetTimer, AccessorListViewEvents, -10
return

;~ AccessorListViewEvents:
;~ AccessorListViewEvents()
;~ return

AccessorEditEvents()
{
	global Accessor, AccessorPlugins, AccessorEdit
	;~ static Running = 0
	;~ if(Running)
	;~ {
		;~ SetTimer, AccessorEditEvents, -10
		;~ return
	;~ }
	;~ Running := true
	;~ Critical ;Make critical so that it isn't called while it's still running
	AccessorListEntry := AccessorGetSelectedListEntry()
	GuiControlGet, Filter, , AccessorEdit
	
	NeedsUpdate := 1
	for index, AccessorPlugin in AccessorPlugins ;Check if single context plugin requests an update
	{
		if(AccessorPlugin.Enabled && SingleContext := ((AccessorPlugin.Settings.Keyword && Filter && strStartsWith(Filter, AccessorPlugin.Settings.Keyword)) || AccessorPlugin.IsInSinglePluginContext(Filter, Accessor.LastFilter)))
		{
			Accessor.SingleContext := AccessorPlugin.Type
			NeedsUpdate := AccessorPlugin.EditEvents(AccessorListEntry, Filter, LastFilter)
			break
		}
	}
	if(!SingleContext)
		Accessor.SingleContext := false
	if(!NeedsUpdate) ;Check if any plugin requests an update
		for index, AccessorPlugin in AccessorPlugins
		{
			if(AccessorPlugin.Enabled && !AccessorPlugin.KeywordOnly)
				NeedsUpdate := AccessorPlugin.EditEvents(AccessorListEntry, Filter, LastFilter)
			if(NeedsUpdate)
				break
		}
	;~ outputdebug % "cycle " Accessor.History.CycleHistory " filter " Filter
	if(!Accessor.History.CycleHistory)
		Accessor.History[1] := Filter
	else
		Accessor.History.CycleHistory := 0
	if(NeedsUpdate)
		FillAccessorList()
	;~ Critical, Off
	;~ Running := false
}
AccessorListViewEvents()
{
	global AccessorPlugins, AccessorOKButton
	if(IsObject(AccessorListEntry := AccessorGetSelectedListEntry()))
	{
		AccessorPlugin := AccessorPlugins.GetItemWithValue("Type", AccessorListEntry.Type)
		if(AccessorPlugin.Enabled)
		{
			handled := AccessorPlugin.ListViewEvents(AccessorListEntry)
			GUIControlGet, name, , AccessorOKButton
			if(name != AccessorPlugin.OKName)
				GuiControl,,AccessorOKButton, % AccessorPlugin.OKName
		}
	}
	if(!handled && A_GUIEvent = "DoubleClick")
	{
		AccessorOK()
		AccessorClose()
	}
	return
}

AccessorOK:
AccessorOK()
AccessorClose:
AccessorClose()
return
AccessorEscape:
AccessorClose()
return
AccessorCancel:
AccessorClose()
return
AccessorOK()
{
	global Accessor, AccessorPlugins, AccessorEdit
	if(IsObject(AccessorListEntry := AccessorGetSelectedListEntry()))
	{
		GuiControlGet, Filter, , AccessorEdit
		Accessor.History.Insert(2, Filter)
		while(Accessor.History.MaxIndex() > 10)
			Accessor.History.Delete(11)
		if(AccessorPlugin := AccessorPlugins.GetItemWithValue("Type", AccessorListEntry.Type))
			AccessorPlugin.PerformAction(Accessor, AccessorListEntry)
	}
}
AccessorClose()
{
	global Accessor, AccessorPlugins
	WasCritical := A_IsCritical
	Critical
	if(Accessor.GUINum)
	{
		Gui, % Accessor.GUINum ": Default"
		GUI, ListView, AccessorListView
		for index, AccessorPlugin in AccessorPlugins
			AccessorPlugin.OnAccessorClose(Accessor)
		Accessor.LastFilter := ""
		OnMessage(0x100, Accessor.OldKeyDown) ; Restore previous KeyDown handler
		Accessor.GUINum := 0
		Gui, Destroy
	}
	if(!WasCritical)
		Critical, Off
}
WM_ACTIVATE(wParam,lParam)
{
	global Accessor
	if(wParam=0 && lParam=WinExist(Accessor.WindowTitle " ahk_class AutoHotkeyGUI"))
		AccessorClose()
	return 0
}
#if Accessor.GUINum
Tab::Down
*Up::AccessorUp()
*Down::AccessorDown()
#if
AccessorUp()
{
	global Accessor,AccessorEdit,AccessorListView
	Gui, % Accessor.GUINum ": Default"
	GUI, ListView, AccessorListView
	if(GetKeyState("Control", "P"))
	{
		if(Accessor.History.MaxIndex() >= Accessor.History.Index + 1 && Accessor.History.Index < 10)
		{
			Accessor.History.Index := Accessor.History.Index + 1
			Accessor.History.CycleHistory := 1
			GuiControl,, AccessorEdit, % Accessor.History[Accessor.History.Index]
			SendMessage, 0xC1, -1, , Edit1, A ; EM_LINEINDEX (Gets index number of line)
			CaretTo := ErrorLevel
			SendMessage, 0xB1, 0, CaretTo, Edit1, A ;EM_SETSEL
		}
	}
	else if(LV_GetCount() > 0)
	{
		selected := LV_GetNext()
		selected := Mod(selected - 1 + LV_GetCount() - 1, LV_GetCount()) + 1
		LV_Modify(selected,"Select Vis")
	}
}
AccessorDown()
{
	global Accessor,AccessorEdit,AccessorListView
	Gui, % Accessor.GUINum ": Default"
	GUI, ListView, AccessorListView
	if(GetKeyState("Control", "P"))
	{
		if(Accessor.History.Index > 1)
		{
			Accessor.History.Index := Accessor.History.Index - 1
			Accessor.History.CycleHistory := 1
			GuiControl,, AccessorEdit, % Accessor.History[Accessor.History.Index]
			SendMessage, 0xC1, -1, , Edit1, A ; EM_LINEINDEX (Gets index number of line)
			CaretTo := ErrorLevel
			SendMessage, 0xB1, 0, CaretTo, Edit1, A ;EM_SETSEL
		}
	}
	else if(LV_GetCount() > 0)
	{
		selected := LV_GetNext()
		selected := Mod(selected,LV_GetCount()) + 1
		LV_Modify(selected,"Select Vis")
	}
}
#if Accessor.GUINum && IsControlActive("Edit1")
PgUp::
PostMessage, 0x100, 0x21, 0,SysListView321,A
return
PgDn::
PostMessage, 0x100, 0x22, 0,SysListView321,A
return
AppsKey::
PostMessage, 0x100, 0x5D, 0,SysListView321,A
return
#if

#if Accessor.GUINum && (!Accessor.SingleContext || !AccessorPlugins.GetItemWithValue("Type", Accessor.SingleContext).HandlesEnter)
Enter::
NumpadEnter::
AccessorOK()
AccessorClose()
return
#if

#if Accessor.GUINum && !Edit_TextIsSelected("","ahk_id " Accessor.HwndEdit)
^c::
AccessorCopy()
return
#if
AccessorCopy()
{
	global AccessorPlugins
	if(AccessorListEntry := AccessorGetSelectedListEntry())
	{
		AccessorPlugin := AccessorPlugins.GetItemWithValue("Type", AccessorListEntry.Type)
		if(IsFunc(AccessorPlugin.OnCopy))
			AccessorPlugin.OnCopy(AccessorListEntry)
		else
			AccessorCopyField("Path", AccessorListEntry)
	}
}

AccessorContextMenu:
AccessorContextMenu()
return
AccessorContextMenu()
{
	global AccessorPlugins
	if(A_GuiControl = "AccessorListView")
	{
		selected := A_EventInfo
		if(!selected)
			return
		
		Menu, AccessorMenu, add, 1,AccessorOK
		Menu, AccessorMenu, DeleteAll
		if(AccessorListEntry := AccessorGetSelectedListEntry())
		{
			if(AccessorPlugin := AccessorPlugins.GetItemWithValue("Type", AccessorListEntry.Type))
			{
				AccessorPlugin.SetupContextMenu(AccessorListEntry)
				Menu, AccessorMenu, Show
			}
		}
	}
}

AccessorGetSelectedListEntry()
{
	global Accessor, AccessorListView
	if(!Accessor.GUINum)
		return
	Gui, % Accessor.GUINum ": Default"
	Gui, ListView, AccessorListView
	selected := LV_GetNext()
	LV_GetText(id,selected,2)
	return id ? Accessor.List[id] : ""
}
;Some generic functions and labels that may be used by different plugins
AccessorOpenExplorer:
AccessorOpenExplorer()
return
AccessorOpenCMD:
AccessorOpenCMD()
return
AccessorRunWithArgs:
AccessorRunWithArgs()
return
AccessorRunAsAdmin:
AccessorRunAsAdmin()
return
AccessorRun:
AccessorRun()
return
AccessorExplorerContextMenu:
AccessorExplorerContextMenu()
return
AccessorCopyPath:
AccessorCopyField("Path")
return
AccessorCopyTitle:
AccessorCopyField("Title")
return
AccessorCopyField(Field = "Path", AccessorListEntry = "")
{
	if(!IsObject(AccessorListEntry))
		AccessorListEntry := AccessorGetSelectedListEntry()
	if(IsObject(AccessorListEntry))
		Clipboard := AccessorListEntry[Field]
}
AccessorExplorerContextMenu(AccessorListEntry = "")
{
	if(!IsObject(AccessorListEntry))
		AccessorListEntry := AccessorGetSelectedListEntry()
	if(AccessorListEntry.Path)
		ShellContextMenu(AccessorListEntry.Path)
}
AccessorRun(AccessorListEntry = "")
{
	if(!IsObject(AccessorListEntry))
		AccessorListEntry := AccessorGetSelectedListEntry()
	if(AccessorListEntry.Path)
	{
		WorkingDir := GetWorkingDir(AccessorListEntry.Path)
		ProgramLauncherAddToCache(AccessorListEntry)
		Run(Quote(AccessorListEntry.Path) (AccessorListEntry.args ? " " AccessorListEntry.args : ""), WorkingDir)
		AccessorClose()
	}
}
AccessorRunAsAdmin(AccessorListEntry = "")
{
	if(!IsObject(AccessorListEntry))
		AccessorListEntry := AccessorGetSelectedListEntry()
	if(AccessorListEntry.Path)
	{
		ProgramLauncherAddToCache(AccessorListEntry)
		WorkingDir := GetWorkingDir(AccessorListEntry.Path)
		Run(AccessorListEntry.Path, WorkingDir, "", 0)
		AccessorClose()
	}
}
AccessorRunWithArgs(AccessorListEntry = "")
{
	if(!IsObject(AccessorListEntry))
		AccessorListEntry := AccessorGetSelectedListEntry()
	if(AccessorListEntry.Path)
	{
		ProgramLauncherAddToCache(AccessorListEntry)
		Event := new CEvent()
		Event.Name := "Run with arguments"
		Event.Temporary := true
		Event.Actions.Insert(new CInputAction())
		Event.Actions[1].Text := "Enter program arguments"
		Event.Actions[1].Title := "Enter program arguments"
		Event.Actions.Insert(new CRunAction())
		Event.Actions[2].Command := """" AccessorListEntry.Path """ ${Input}"
		Event.Actions[2].WorkingDirectory := GetWorkingDir(AccessorListEntry.Path)
		EventSystem.TemporaryEvents.RegisterEvent(Event)
		Event.TriggerThisEvent()
		AccessorClose()
	}
}
AccessorOpenExplorer(AccessorListEntry = "")
{
	if(!IsObject(AccessorListEntry))
		AccessorListEntry := AccessorGetSelectedListEntry()
	if(AccessorListEntry)
	{
		if(!InStr(FileExist(AccessorListEntry.Path),"D"))
			Run(A_WinDir "\explorer.exe /Select," AccessorListEntry.Path)
		else
			Run(A_WinDir "\explorer.exe /n,/e," AccessorListEntry.Path)
		AccessorClose()
	}
}
AccessorOpenCMD(AccessorListEntry = "")
{
	if(!IsObject(AccessorListEntry))
		AccessorListEntry := AccessorGetSelectedListEntry()
	if(AccessorListEntry)
	{
		path := AccessorListEntry.Path
		if(!InStr(FileExist(path),"D"))
			SplitPath, path,, path
		Run("cmd.exe /k cd /D """ path """", GetWorkingDir(AccessorListEntry.Path))
		AccessorClose()
	}
}
GUI_EditAccessorPlugin(TemporaryPlugin,GoToLabel="")
{
	static sTemporaryPlugin, sOriginalPlugin, result, PluginGUI, GuiNum
	global AccessorPlugins
	if(GoToLabel = "")
	{
		;Don't show more than once
		if(sTemporaryPlugin)
			return
		GuiNum := GetFreeGuiNum(4)
		if(!GuiNum)
			return
		sTemporaryPlugin := TemporaryPlugin.DeepCopy()
		sOriginalPlugin := AccessorPlugins.GetItemWithValue("Type", sTemporaryPlugin.Type)
		if(!sOriginalPlugin)
			return
		result := ""
		PluginGUI := object("x",38,"y",80)
		if(CSettingsWindow && !CSettingsWindow.IsDestroyed)
			CSettingsWindow.Disable()
		Gui, %GuiNum%:Default
		Gui, +LabelEditAccessorPlugin +OwnerCSettingsWindow1 +ToolWindow +OwnDialogs
		width := 500
		height := 460
		;Gui, 4:Add, Button, ,OK
		x := PluginGUI.x
		y := Height - 34
		Gui, Add, Button, gEditAccessorPluginHelp x%x% y%y% w70 h23, &Help
		x := Width - 174
		Gui, Add, Button, gEditAccessorPluginOK x%x% y%y% w70 h23, &OK
		x := Width - 94
		Gui, Add, Button, gEditAccessorPluginCancel x%x% y%y% w80 h23, &Cancel
		;Fill tabs
		x := 40
		y := 18
		
		Gui, Add, Text, x%x% y%y%, % sOriginalPlugin.Description
		
		x := 28
		y += 40 + 4
		w := width - 54
		h := height - 110
		Gui, Add, GroupBox, x%x% y%y% w%w% h%h%, Options
		sOriginalPlugin.ShowSettings(sTemporaryPlugin.Settings, PluginGUI)
		Gui, Show, w%width% h%height%, Edit Plugin
		
		Gui, %GuiNum%:+LastFound
		WinGet, EditAccessorPlugin_hWnd,ID
		DetectHiddenWindows, Off
		Critical, Off
		loop
		{
			sleep 250
			IfWinNotExist ahk_id %EditAccessorPlugin_hWnd% 
				break
		}
		sTemporaryPlugin := ""
		sOriginalPlugin := ""
		PluginGUI := ""
		GuiNum := 0
		if(CSettingsWindow && !CSettingsWindow.IsDestroyed)
			CSettingsWindow.Enable()
		return result
	}
	else if(GoToLabel = "EditAccessorPluginOK")
	{
		outputdebug ok %guinum%
		sOriginalPlugin.SaveSettings(sTemporaryPlugin.Settings, PluginGUI)
		SubmitControls(sTemporaryPlugin.Settings, PluginGUI)
		Gui, %GuiNum%:Submit, NoHide
		outputdebug % " keyword: " sTemporaryPlugin.Settings.Keyword " Default: " sTemporaryPlugin.DefaultKeyword
		if(sTemporaryPlugin.Settings.Keyword = "" && sOriginalPlugin.DefaultKeyword != "")
			sTemporaryPlugin.Settings.Keyword := sOriginalPlugin.DefaultKeyword
		result := sTemporaryPlugin
		Gui 1:+LastFoundExist
		IfWinExist		
			Gui, 1:-Disabled
		Gui, %GuiNum%:Destroy
		return
	}
	else if(GoToLabel = "EditAccessorPluginClose")
	{
		Gui 1:+LastFoundExist		
		result := ""
		IfWinExist		
			Gui, 1:-Disabled
		Gui,%GuiNum%: Cancel
		Gui,%GuiNum%: destroy
		Gui 1:+LastFoundExist
		IfWinExist		
			Gui, 1:Default
		return
	}
	else if(GoToLabel = "EditAccessorPluginHelp")
		run % "http://code.google.com/p/7plus/wiki/docsAccessor" sTemporaryPlugin.Type ",,UseErrorLevel"
} 
EditAccessorPluginOK:
outputdebug ok
GUI_EditAccessorPlugin("","EditAccessorPluginOK")
return

EditAccessorPluginClose:
EditAccessorPluginEscape:
EditAccessorPluginCancel:
GUI_EditAccessorPlugin("","EditAccessorPluginClose")
return

EditAccessorPluginHelp:
GUI_EditAccessorPlugin("","EditAccessorPluginHelp")
return