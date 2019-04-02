Accessor_Weather_Init(ByRef Weather, PluginSettings)
{
	Weather.Settings.Keyword := "weather"
	Weather.DefaultKeyword := "weather"
	Weather.KeywordOnly := false ;This is actually true, but Accessor_Google_IsInSinglePluginContext needs to be called every time so it is handled manually here
	Weather.MinChars := 0
	Weather.OKName := "Copy"
	Weather.Description := "This plugin shows the weather information for a location. Just type ""[Keyword] location""."
	Weather.HasSettings := True
}
Accessor_Weather_ShowSettings(Weather, PluginSettings, PluginGUI)
{
	AddControl(PluginSettings, PluginGUI, "Edit", "Keyword", "", "", "Keyword:")
}
Accessor_Weather_IsInSinglePluginContext(Weather, Filter, LastFilter)
{
	if(strStartsWith(Filter, Weather.Settings.Keyword " "))
	{
		if(!Weather.Cleared)
		{
			Weather.Cleared := true
			FillAccessorList()
		}
		return true
	}
	Weather.Cleared := false
	return false
}
Accessor_Weather_GetDisplayStrings(Weather, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
	; Title := AccessorListEntry.titleNoFormatting
	; Path := AccessorListEntry.visibleUrl
	; Detail1 := AccessorListEntry.Detail1
	; Detail2 := AccessorListEntry.Detail2
}
Accessor_Weather_OnAccessorOpen(Weather, Accessor)
{
	Weather.List := Array()
	Weather.Cleared := false
}
Accessor_Weather_OnAccessorClose(Weather, Accessor)
{
	for index, ListEntry in Weather.List
		if(ListEntry.Icon)
			DestroyIcon(ListEntry.Icon)
}
Accessor_Weather_OnExit(Weather)
{
}
Accessor_Weather_FillAccessorList(Weather, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet)
{
	if(!KeywordSet)
		return
	Loop % Weather.List.MaxIndex()
	{
		ImageList_ReplaceIcon(Accessor.ImageListID, -1, Weather.List[A_Index].Icon)
		IconCount++
		Accessor.List.Insert(Object("Title",Weather.List[A_Index].Title,"Path",Weather.List[A_Index].Path, "Type","Weather", "Detail1", Weather.List[A_Index].Detail1, "Detail2", Weather.List[A_Index].Detail2,"Icon", IconCount))
	}
}
Accessor_Weather_PerformAction(Weather, Accessor, AccessorListEntry)
{
	Clipboard := AccessorListEntry.Title
	return
}
Accessor_Weather_ListViewEvents(Weather, AccessorListEntry)
{
}
Accessor_Weather_EditEvents(Weather, AccessorListEntry, Filter, LastFilter)
{
	outputdebug edit events
	SetTimer, QueryWeatherResult, -100
	return false
}
Accessor_Weather_OnCopy(Weather, AccessorListEntry)
{
	AccessorCopyField("Title", AccessorListEntry)
}
QueryWeatherResult:
QueryWeatherResult()
return
QueryWeatherResult()
{
	global AccessorPlugins, AccessorEdit, Accessor
	WeatherPlugin := AccessorPlugins.GetItemWithValue("Type", "Weather")
	if(!Accessor.GUINum)
		return
	Gui, % Accessor.GUINum ": Default"
	GuiControlGet, Filter, , AccessorEdit
	if(!strStartsWith(Filter, WeatherPlugin.Settings.Keyword " "))
		return
	Filter := strTrim(Filter, WeatherPlugin.Settings.Keyword " ")
	URL := uriEncode("http://www.google.com/ig/api?weather=") uriEncode(Filter, 1) ;"&rsz=8"
	FileDelete, %A_Temp%\7plus\WeatherQuery.xml
	URLDownloadToFile, %URL%, %A_Temp%\7plus\WeatherQuery.xml
	FileRead, WeatherQuery, %A_Temp%\7plus\WeatherQuery.xml
	Loop % WeatherPlugin.List.MaxIndex()
		if(WeatherPlugin.List[A_Index].Icon)
			DestroyIcon(WeatherPlugin.List[A_Index].Icon)
	WeatherPlugin.List := Array()
	
	Loop 5
		pos%A_Index% := 0
	RegexMatch(WeatherQuery, "i)<city data=""(.*?)""/>",city,1)
	if(!city1) ;No results
		return
	pos1 := RegexMatch(WeatherQuery, "i)<condition data=""(.*?)""/>",condition,pos1+1)
	RegexMatch(WeatherQuery, "i)<temp_c data=""(.*?)""/>",temp_c,1)
	RegexMatch(WeatherQuery, "i)<humidity data=""(.*?)""/>",humidity,1)
	RegexMatch(WeatherQuery, "i)<icon data=""(.*?)""/>",icon,1)
	name := SubStr(icon1, InStr(icon1, "/", 0, 0) + 1)
	
	if(!FileExist(A_Temp "\7plus\" name))
		URLDownloadToFile, http://google.com%icon1%, %A_Temp%\7plus\%name%		
	pBitmap := Gdip_CreateBitmapFromFile(A_Temp "\7plus\" name)
	hIcon := Gdip_CreateHICONFromBitmap(pBitmap)
	WeatherPlugin.List.Insert(Object("Title", "Now: " condition1 ", " temp_c1 "°C, " humidity1, "Path","Weather in " city1, "Icon", hIcon ))
	Loop
	{
		pos1 := RegexMatch(WeatherQuery, "i)<condition data=""(.*?)""/>",condition,pos1+1)
		pos2 := RegexMatch(WeatherQuery, "i)<low data=""(.*?)""/>",low,pos2+1)
		pos3 := RegexMatch(WeatherQuery, "i)<high data=""(.*?)""/>",high,pos3+1)
		pos4 := RegexMatch(WeatherQuery, "i)<day_of_week data=""(.*?)""/>",day_of_week,pos4+1)
		pos5 := RegexMatch(WeatherQuery, "i)<icon data=""(.*?)""/>",icon,pos5+1)
		
		if(pos1 && pos2 && pos3 && pos4 && pos5 && condition1 && low1 && high1 &&day_of_week1 && icon1)
		{
			name := SubStr(icon1, InStr(icon1, "/", 0, 0) + 1)
			if(!FileExist(A_Temp "\7plus\" name))
				URLDownloadToFile, http://google.com%icon1%, %A_Temp%\7plus\%name%	
			pBitmap := Gdip_CreateBitmapFromFile(A_Temp "\7plus\" name)
			hIcon := Gdip_CreateHICONFromBitmap(pBitmap)
			low1 := Round((5/9)*(low1-32)) ;Convert °F to °C
			high1 := Round((5/9)*(high1-32)) ;Convert °F to °C
			WeatherPlugin.List.Insert(Object("Title", day_of_week1 ": " condition1 ", Low: " low1 "°C, high: " high1 "°C", "Path", "Weather in " city1, "Icon", hIcon ))
		}
		else
			break
	}
	FillAccessorList()
} 
Accessor_Weather_SetupContextMenu(Weather, AccessorListEntry)
{
	Menu, AccessorMenu, add, Open page,AccessorOK
	Menu, AccessorMenu, Default,Open page
	Menu, AccessorMenu, add, Copy (CTRL+C), AccessorCopyTitle
}