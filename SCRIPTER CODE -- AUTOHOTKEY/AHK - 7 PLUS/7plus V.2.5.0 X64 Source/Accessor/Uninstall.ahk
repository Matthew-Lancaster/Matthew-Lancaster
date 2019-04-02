Accessor_Uninstall_Init(ByRef Uninstall, PluginSettings)
{
	Uninstall.Settings.Keyword := "Uninstall"
	Uninstall.DefaultKeyword := "Uninstall"
	Uninstall.KeywordOnly := True
	Uninstall.MinChars := 0
	Uninstall.OKName := "Uninstall"
	Uninstall.Description := "This plugin lets you uninstall programs or remove the uninstall entries from the list."
	Uninstall.Settings.FuzzySearch := PluginSettings.HasKey("FuzzySearch") ? PluginSettings.FuzzySearch : 0
	Uninstall.HasSettings := True
}
Accessor_Uninstall_ShowSettings(Uninstall, PluginSettings, PluginGUI)
{
	AddControl(PluginSettings, PluginGUI, "Edit", "Keyword", "", "", "Keyword:")
	AddControl(PluginSettings, PluginGUI, "Checkbox", "FuzzySearch", "Use fuzzy search (slower)", "", "")
}
Accessor_Uninstall_IsInSinglePluginContext(Uninstall, Filter, LastFilter)
{
}
Accessor_Uninstall_GetDisplayStrings(Uninstall, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
	Path := AccessorListEntry.UninstallString
	Detail1 := "Uninstall"
}
Accessor_Uninstall_OnAccessorOpen(Uninstall, Accessor)
{
	Uninstall.List := Array()
}
Accessor_Uninstall_OnAccessorClose(Uninstall, Accessor)
{
	if(IsObject(Uninstall.List))
	for index, ListEntry in Uninstall.List
		if(ListEntry.Icon != Accessor.GenericIcons.Application)			
			DestroyIcon(ListEntry.Icon)
}
Accessor_Uninstall_OnExit(Uninstall)
{
}
Accessor_Uninstall_FillAccessorList(Uninstall, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet)
{
	;Lazy loading
	if(Uninstall.List.MaxIndex() = 0)
	{
		Loop, HKLM , SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, 2, 0
		{
			GUID := A_LoopRegName ;Note: This is not always a GUID but can also be a regular name. It seems that MSIExec likes to use GUIDs
			RegRead, DisplayName, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%GUID%, DisplayName
			; outputdebug displayname %displayname%
			RegRead, UninstallString, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%GUID%, UninstallString
			RegRead, InstallLocation, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%GUID%, InstallLocation
			RegRead, DisplayIcon, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%GUID%, DisplayIcon
			; outputdebug orig displayicon %displayicon%
			if(RegexMatch(DisplayIcon,".+,\d*"))
			{
				; outputdebug match
				Number := strTrim(SubStr(DisplayIcon, InStr(DisplayIcon,",",0,0) + 1), " ")
				DisplayIcon := strTrim(SubStr(DisplayIcon, 1, InStr(DisplayIcon,",",0,0) - 1), " ")
			}
			if(!Number)
				Number := 0
			; outputdebug displayicon %displayicon% number %number%
			if(FileExist(DisplayIcon))
				hIcon := ExtractAssociatedIcon(Number, DisplayIcon, iIndex)
			else
				hIcon := Accessor.GenericIcons.Application
			; outputdebug hicon %hicon%
			if(DisplayName)
				Uninstall.List.Insert(Object("GUID", GUID, "DisplayName", DisplayName, "UninstallString", UninstallString, "InstallLocation", InstallLocation, "Icon", hIcon))
		}
	}
	FuzzyList := Array()
	Loop % Uninstall.List.MaxIndex()
	{
		x := 0
		if(x := (Filter = "" || InStr(Uninstall.List[A_Index].DisplayName,Filter)) || y := (Uninstall.Settings.FuzzySearch && FuzzySearch(Uninstall.List[A_Index].DisplayName,Filter) < 0.4))
		{
			ImageList_ReplaceIcon(Accessor.ImageListID, -1, Uninstall.List[A_Index].Icon)
			IconCount++
			if(x)
				Accessor.List.Insert(Object("Title",Uninstall.List[A_Index].DisplayName,"Path",Uninstall.List[A_Index].InstallLocation, "UninstallString", Uninstall.List[A_Index].UninstallString, "Type","Uninstall", "Icon", IconCount))
			else
				FuzzyList.Insert(Object("Title",Uninstall.List[A_Index].DisplayName,"Path",Uninstall.List[A_Index].InstallLocation, "UninstallString", Uninstall.List[A_Index].UninstallString, "Type","Uninstall", "Icon", IconCount))
		}
	}
}
Accessor_Uninstall_PerformAction(Uninstall, Accessor, AccessorListEntry)
{
	Run(AccessorListEntry.UninstallString)
}
Accessor_Uninstall_ListViewEvents(Uninstall, AccessorListEntry)
{
}
Accessor_Uninstall_EditEvents(Uninstall, AccessorListEntry, Filter, LastFilter)
{
	return true
}
Accessor_Uninstall_SetupContextMenu(Uninstall, AccessorListEntry)
{
	Menu, AccessorMenu, add, Uninstall,AccessorOK
	Menu, AccessorMenu, Default,Uninstall
	Menu, AccessorMenu, add, Remove uninstall entry, RemoveUninstallEntry
	if(AccessorListEntry.Path)
	{
		Menu, AccessorMenu, add, Open path in explorer, AccessorOpenExplorer
		Menu, AccessorMenu, add, Open path in CMD,AccessorOpenCMD
	}
}

RemoveUninstallEntry:
RemoveUninstallEntry()
return
RemoveUninstallEntry()
{
	global Accessor, AccessorListView, AccessorPlugins
	GUINum := Accessor.GUINum
	Gui, %GUINum%: Default
	Gui, ListView, AccessorListView
	selected := LV_GetNext()
	if(!selected)
		return
	LV_GetText(id,selected,2)
	GUID := Accessor.List[id].GUID
	RegDelete, HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%GUID%
	Uninstall := AccessorPlugins.GetItemWithValue("Type", "Uninstall")
	Uninstall.List.Delete(Uninstall.List.FindKeyWithValue("GUID", GUID))
	FillAccessorList()
}