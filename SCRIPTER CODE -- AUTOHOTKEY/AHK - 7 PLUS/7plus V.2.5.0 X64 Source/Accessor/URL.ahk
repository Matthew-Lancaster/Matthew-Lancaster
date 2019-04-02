Accessor_URL_Init(ByRef URL, PluginSettings)
{
	URL.KeywordOnly := false
	URL.MinChars := 3
	URL.OKName := "Open URL"
	URL.Description := "This plugin allows to open URLs in the browser and also has a history function."
	URL.HasSettings := true
	URL.Settings.UseHistory := PluginSettings.HasKey("UseHistory") ? PluginSettings.UseHistory : 1
	URL.Settings.MaxHistoryLen := PluginSettings.HasKey("MaxHistoryLen") ? PluginSettings.MaxHistoryLen : 1
	URL.Settings.SaveHistoryOnExit := PluginSettings.HasKey("SaveHistoryOnExit") ? PluginSettings.SaveHistoryOnExit : 1
	URL.History := Array()
	if(!FileExist(Settings.ConfigPath "\History.xml"))
		return
	FileRead, xml, % Settings.ConfigPath "\History.xml"
	XMLObject := XML_Read(xml)
	if(IsObject(XMLObject))
	{
		;Convert empty and single arrays to real array
		if(!IsObject(XMLObject.List) || !XMLObject.List.MaxIndex())
			XMLObject.List := IsObject(XMLObject.List) ? Array(XMLObject.List) : Array()		
		
		Loop % min(XMLObject.List.MaxIndex(), URL.Settings.MaxHistoryLen)
		{
			XMLObjectListEntry := XMLObject.List[A_Index]
			HistoryURL := XMLObjectListEntry.URL
			URL.History.Insert(Object("URL",HistoryURL))
		}
	}
}
Accessor_URL_ShowSettings(URL, PluginSettings, PluginGUI)
{	
	AddControl(PluginSettings, PluginGUI, "Checkbox", "UseHistory", "Use history", "", "")
	AddControl(PluginSettings, PluginGUI, "Checkbox", "SaveHistoryOnExit", "Save history on exit", "", "")
	AddControl(PluginSettings, PluginGUI, "Edit", "MaxHistoryLen", "", "", "History length:","Clear history","Accessor_URL_ClearHistory")
}
Accessor_URL_ClearHistory:
Accessor_URL_ClearHistory()
return
Accessor_URL_ClearHistory()
{
	global AccessorPlugins
	URLPlugin := AccessorPlugins.GetItemWithValue("Type","URL")
	URLPlugin.History := Array()
}
Accessor_URL_IsInSinglePluginContext(URL, Filter, LastFilter)
{
	return IsURL(Filter)
}
Accessor_URL_GetDisplayStrings(URL, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
	; Title := AccessorListEntry.titleNoFormatting
	; Path := AccessorListEntry.visibleUrl
	; Detail1 := AccessorListEntry.Detail1
	; Detail2 := AccessorListEntry.Detail2
}
Accessor_URL_OnAccessorOpen(URL, Accessor)
{
}
Accessor_URL_OnAccessorClose(URL, Accessor)
{
}
Accessor_URL_OnExit(URL)
{
	FileDelete, % Settings.ConfigPath "\History.xml"
	if(!URL.Settings.SaveHistoryOnExit)
		return
	XMLObject := Object("List",Array())
	Loop % URL.History.MaxIndex()
		XMLObject.List.Insert(Object("URL",URL.History[A_Index].URL))
	XML_Save(XMLObject, Settings.ConfigPath "\History.xml")
}
Accessor_URL_FillAccessorList(URL, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet, Parameters)
{
	if(!CouldBeURL(Filter))
		return
	
	Accessor.List.Insert(Object("Title",Filter,"Path", "Open URL", "Type","URL", "Detail1", "URL", "Detail2", "","Icon", 3))
	
	if(URL.Settings.UseHistory)
		for index, HistoryEntry in URL.History
			if(InStr(HistoryEntry.URL, Filter) && HistoryEntry.URL != Filter && CouldBeURL(HistoryEntry.URL))
				Accessor.List.Insert(Object("Title", HistoryEntry.URL, "Path", "Open URL", "Type", "URL", "Detail1", "URL", "Detail2", "","Icon", 3, "History", true))
}
Accessor_URL_PerformAction(URLPlugin, Accessor, AccessorListEntry)
{
	global AccessorEdit
	if(AccessorListEntry.Title)
	{
		if(URLPlugin.Settings.UseHistory)
		{			
			if(index := URLPlugin.History.FindKeyWithValue("URL",AccessorListEntry.Title)) ;Move existing items to the top
				URLPlugin.History.Delete(index)
			URLPlugin.History.Insert(Object("URL", AccessorListEntry.Title)) ;Add entered item to the top
			if(URLPlugin.History.MaxIndex() > URLPlugin.Settings.MaxHistoryLen) ;Make sure history len is not exceeded
				URLPlugin.History.Delete(1)
		}
		outputdebug % "append new len: " URLPlugin.History.MaxIndex()
		url := (!InStr(AccessorListEntry.Title, "://") ? "http://" : "") AccessorListEntry.Title
		run %url%,,UseErrorLevel
	}
	return
}
Accessor_URL_ListViewEvents(URL, AccessorListEntry)
{
}
Accessor_URL_EditEvents(URL, AccessorListEntry, Filter, LastFilter)
{
}
Accessor_URL_SetupContextMenu(URL, AccessorListEntry)
{
	Menu, AccessorMenu, add, Open URL, AccessorOK
}