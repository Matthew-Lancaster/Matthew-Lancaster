Accessor_Run_Init(ByRef Run, PluginSettings)
{
	Run.Settings.Keyword := "Run"
	Run.DefaultKeyword := "Run"
	Run.KeywordOnly := false
	Run.MinChars := 1
	Run.OKName := "Run command"
	Run.Description := "This plugin tries to execute the entered text directly. You can press CTRL+Enter`n in the Accessor window to execute this action even if it is not selected.`nCTRL+SHIFT+Enter will run it with admin permissions."
	Run.HasSettings := True
}

Accessor_Run_ShowSettings(Run, PluginSettings, PluginGUI)
{
	AddControl(PluginSettings, PluginGUI, "Edit", "Keyword", "", "", "Keyword:")
	AddControl(PluginSettings, PluginGUI, "Edit", "BasePriority", "", "", "Base Priority:")
}

Accessor_Run_IsInSinglePluginContext(Run, Filter, LastFilter)
{
	return false
}

Accessor_Run_GetDisplayStrings(Run, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
	Detail1 := "Run command"
}

Accessor_Run_OnAccessorOpen(Run, Accessor)
{
	Run.Priority := Run.Settings.BasePriority
}

Accessor_Run_OnAccessorClose(Run, Accessor)
{
}

Accessor_Run_OnExit(Run)
{
}

Accessor_Run_FillAccessorList(RunPlugin, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet)
{
	Accessor.List.Insert(Object("Title",Filter,"Path", Filter,"Type", "Run", "Icon", 1)) ;Use generic icon
}

Accessor_Run_PerformAction(Run, Accessor, AccessorListEntry)
{
	AccessorRun()
}

Accessor_Run_EditEvents(Run, AccessorListEntry, Filter, LastFilter)
{
	return true
}

Accessor_Run_SetupContextMenu(Run, AccessorListEntry)
{
	Menu, AccessorMenu, add, Run as user, AccessorRun
	Menu, AccessorMenu, add, Run as admin, AccessorRunAsAdmin
	Menu, AccessorMenu, Default, Run as user
}

#if (Accessor.GUINum)
^Enter::
^NumpadEnter::
AccessorRun(Accessor.List.GetItemWithValue("Type", "Run"))
AccessorClose()
return
+Enter::
+NumpadEnter::
AccessorRunAsAdmin(Accessor.List.GetItemWithValue("Type", "Run"))
AccessorClose()
return
#if