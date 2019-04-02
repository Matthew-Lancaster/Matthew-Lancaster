Accessor_FastFolders_Init(ByRef FastFolders, PluginSettings)
{
	FastFolders.Settings.Keyword := "FastFolders"
	FastFolders.DefaultKeyword := "FastFolders"
	FastFolders.KeywordOnly := false
	FastFolders.MinChars := 1
	FastFolders.OKName := "Open Folder"
	FastFolders.Settings.FuzzySearch := PluginSettings.HasKey("FuzzySearch") ? PluginSettings.FuzzySearch : 0
	FastFolders.Description := "Access the stored FastFolders by typing a part of a folder name."
	FastFolders.HasSettings := True
}
Accessor_FastFolders_ShowSettings(FastFolders, PluginSettings, PluginGUI)
{
	AddControl(PluginSettings, PluginGUI, "Edit", "Keyword", "", "", "Keyword:")
	AddControl(PluginSettings, PluginGUI, "Edit", "BasePriority", "", "", "Base Priority:")
	AddControl(PluginSettings, PluginGUI, "Checkbox", "FuzzySearch", "Use fuzzy search (slower)", "", "")
}
Accessor_FastFolders_IsInSinglePluginContext(FastFolders, Filter, LastFilter)
{
	return false
}
Accessor_FastFolders_GetDisplayStrings(FastFolders, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
	Detail1 := "FastFolder"
}
Accessor_FastFolders_OnAccessorOpen(FastFolders, Accessor)
{
	FastFolders.Priority := FastFolders.Settings.BasePriority
}
Accessor_FastFolders_OnAccessorClose(FastFolders, Accessor)
{
}
Accessor_FastFolders_OnExit(FastFolders)
{
}
Accessor_FastFolders_FillAccessorList(FastFoldersPlugin, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet)
{	
	global FastFolders
	FuzzyList := Array()
	for index, FastFolder in FastFolders
		if(FastFolder.Path)
		{
			if(InStr(FastFolder.Title,Filter) || InStr(FastFolder.Path,Filter))
				Accessor.List.Insert(Object("Title",FastFolder.Title,"Path",FastFolder.Path,"Type","FastFolders", "Icon", 2)) ;Use generic folder icon
			else if(FastFoldersPlugin.Settings.FuzzySearch && FuzzySearch(FastFolder.Title,Filter) < 0.4)
				FuzzyList.List.Insert(Object("Title",FastFolder.Title,"Path",FastFolder.Path,"Type","FastFolders", "Icon", 2)) ;Use generic folder icon
		}
	Accessor.List.extend(FuzzyList)
}
Accessor_FastFolders_PerformAction(FastFolders, Accessor, AccessorListEntry)
{
	if(WinGetClass("ahk_id " Accessor.PreviousWindow) = "CabinetWClass" || WinGetClass("ahk_id " Accessor.PreviousWindow) = "ExploreWClass")
		ShellNavigate(AccessorListEntry.Path,Accessor.PreviousWindow)
	else
		Run(A_WinDir "\explorer.exe /n,/e," AccessorListEntry.Path)
}
Accessor_FastFolders_ListViewEvents(FastFolders, AccessorListEntry)
{
}
Accessor_FastFolders_EditEvents(FastFolders, AccessorListEntry, Filter, LastFilter)
{
	return true
}
Accessor_FastFolders_SetupContextMenu(FastFolders, AccessorListEntry)
{
	Menu, AccessorMenu, add, Open path in explorer,AccessorOK
	Menu, AccessorMenu, Default,Open path in explorer
	Menu, AccessorMenu, add, Open path in CMD,AccessorOpenCMD
	Menu, AccessorMenu, add, Copy Path (CTRL+C), AccessorCopyPath
	Menu, AccessorMenu, add, Explorer context menu, AccessorExplorerContextMenu
}