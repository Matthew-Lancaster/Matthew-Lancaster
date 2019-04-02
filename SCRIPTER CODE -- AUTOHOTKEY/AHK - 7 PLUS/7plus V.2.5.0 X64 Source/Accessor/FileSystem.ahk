Accessor_FileSystem_Init(ByRef FileSystem, PluginSettings)
{
	FileSystem.KeywordOnly := false
	FileSystem.MinChars := 2
	FileSystem.HandlesEnter := true
	FileSystem.DefaultKeyword := ""
	FileSystem.Description := "Browse the file system by typing a path. `nUse Tab for switching through matching entries and enter to enter a folder.`nApplications launched through this method are directly added to Program Launcher cache."
	FileSystem.Settings.UseIcons := PluginSettings.HasKey("UseIcons") ? PluginSettings.UseIcons : 0
	FileSystem.HasSettings := True
}
Accessor_FileSystem_ShowSettings(FileSystem, PluginSettings, PluginGUI)
{
	AddControl(PluginSettings, PluginGUI, "Checkbox", "UseIcons", "Use exact icons (much slower)", "", "")
}
Accessor_FileSystem_IsInSinglePluginContext(FileSystem, Filter, LastFilter)
{
	Filter := ExpandPathPlaceholders(Filter)
	SplitPath, Filter, name, dir,,,drive
	if((x := InStr(dir, ":") ) != 0 && x != 2) ;Colon may only be drive separator
		return false
	return dir != "" && !InStr(Filter, "://") ;Don't match URLs
}
Accessor_FileSystem_GetDisplayStrings(FileSystem, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
	if(InStr(FileExist(AccessorListEntry.Path),"D"))
		Detail1 := "Folder"
	else
		Detail1 := "File"
}
Accessor_FileSystem_OnAccessorOpen(FileSystem, Accessor)
{
}
Accessor_FileSystem_OnAccessorClose(FileSystem, Accessor)
{
}
Accessor_FileSystem_ListViewEvents(FileSystem, AccessorListEntry)
{
	IsDirectory := InStr(FileExist(AccessorListEntry.Path),"D")
	if(A_GUIEvent = "DoubleClick" && IsDirectory) ;Go into directories
		Accessor_FileSystem_OnEnter()
	if(IsDirectory)
		FileSystem.OKName := "Open Folder"
	else
		FileSystem.OKName := "Open File"
}
Accessor_FileSystem_EditEvents(FileSystem, AccessorListEntry, Filter, LastFilter)
{
	return true
}
Accessor_FileSystem_FillAccessorList(FileSystem, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet)
{
	Filter := ExpandPathPlaceholders(Filter)
	outputdebug filter1 %filter%
	SplitPath, filter, name, dir,,,drive
	if(dir)
	{
		if(FileSystem.AutocompletionString)
			name := FileSystem.AutocompletionString
		Loop %dir%\*%name%*, 1, 0
		{
			if(FileSystem.Settings.UseIcons)
			{
				hIcon := ExtractAssociatedIcon(0, A_LoopFileFullPath, iIndex)
				ImageList_ReplaceIcon(Accessor.ImageListID, -1, hIcon)
				DestroyIcon(hIcon)
				Icon := ++IconCount
			}
			else
			{
				if(InStr(FileExist(A_LoopFileFullPath), "D"))
					Icon := 2
				else if A_LoopFileExt in exe,bat,cmd
					Icon := 1
				else
					Icon := 4
			}
			Accessor.List.Insert(Object("Title",A_LoopFileName,"Path",A_LoopFileFullPath,"Type","FileSystem", "Icon", Icon))
		}
	}
}
Accessor_FileSystem_PerformAction(FileSystem, Accessor, AccessorListEntry)
{
	AccessorRun()
}
Accessor_FileSystem_OnExit(FileSystem)
{
}
#if (Accessor.GUINum && Accessor.SingleContext = "FileSystem")
Tab::
Accessor_FileSystem_OnTab()
return
#if
#if (Accessor.GUINum && Accessor.SingleContext = "FileSystem")
Enter::
NumpadEnter::
Accessor_FileSystem_OnEnter()
return
#if
Accessor_FileSystem_OnTab()
{
	global AccessorEdit, Accessor, AccessorListView, AccessorPlugins
	FileSystem := AccessorPlugins.GetItemWithValue("Type", "FileSystem")
	Gui, % Accessor.GUINum ": Default"
	Gui, ListView, AccessorListView	
	GuiControlGet, Filter, , AccessorEdit
	Filter := ExpandPathPlaceholders(Filter)
	SplitPath, Filter, name, dir,,,drive
	LV_GetText(first,1,4)
	if(LV_GetCount() = 1 && InStr(FileExist(first),"D")) ;Go into folder if there is only one entry
	{
		LV_GetText(newname,1,3)
		FileSystemSetFolder(newname)
		return
	}
	if(selected && !FileSystem.AutocompletionString)
		return
	else
	{
		if(name)
		{
			if(!Accessor.AutocompletionString)
				Accessor.AutocompletionString := name
			AutocompletionString := Accessor.AutocompletionString
			Loop %dir%\*%AutocompletionString%*,1,0
			{
				if(A_Index = 1)
					first := A_LoopFileName
				if(A_LoopFileName = name)
				{
					usenext := true
					continue
				}
				if(usenext || (A_Index = 1 && name = AutocompletionString))
				{
					newname := A_LoopFileName
					break
				}
			}				
			Accessor.SuppressListViewUpdate := 1
		}
		else 
			return 0
	}
	if(!newname)
		newname := first
	GuiControl, ,AccessorEdit,%dir%\%newname%
	
	hwndEdit := Accessor.hwndEdit
	SendMessage, 0xC1, -1,,, AHK_ID %hwndEdit%  ; EM_LINEINDEX (Gets index number of line)
	CaretTo := ErrorLevel
	outputdebug caretto %carretto%
	SendMessage, 0xB1, CaretTo, CaretTo,, AHK_ID %hwndEdit% ;EM_SETSEL
	Loop % LV_GetCount()
	{
		LV_GetText(text,A_Index,3)
		if(text = newname)
		{
			LV_Modify(A_Index,"Select Vis")
			break
		}
	}
	return 1
}
Accessor_FileSystem_OnEnter()
{
	AccessorListEntry := AccessorGetSelectedListEntry()
	if(!IsObject(AccessorListEntry))
		return
	if(InStr(FileExist(AccessorListEntry.Path),"D"))
		FileSystemSetFolder(AccessorListEntry.Title)
	else
	{
		AccessorOK()
		AccessorClose()
	}
}
Accessor_FileSystem_OnKeyDown(FileSystem, wParam, lParam, Filter, selected, AccessorListEntry)
{
	global AccessorEdit, Accessor
	if(wParam = 9)
	{
		
	}
	else
		WindowSwitcher.AutocompletionString := ""
	if(wParam = 13 && selected && InStr(FileExist(AccessorListEntry.Path),"D")) ;Enter on Folders
	{
		FileSystemSetFolder(AccessorListEntry.Title)
		return 1
	}
	
	return 0
}
Accessor_FileSystem_SetupContextMenu(FileSystem, AccessorListEntry)
{
	path := AccessorListEntry.Path
	SplitPath, path, name, dir, ext
	if(InStr(FileExist(path), "D"))
	{
		Menu, AccessorMenu, add, Open path in explorer,AccessorOpenExplorer
		Menu, AccessorMenu, add, Open path in CMD,AccessorOpenCMD
		Menu, AccessorMenu, Default, Open path in explorer
	}
	else
	{
		if(InStr("exe,cmd,bat",ext))
		{
			Menu, AccessorMenu, add, Run program, AccessorOK
			Menu, AccessorMenu, Default, Run program
			Menu, AccessorMenu, add, Run program with arguments,AccessorRunWithArgs
		}
		else
		{
			Menu, AccessorMenu, add, Open document,AccessorOK
			Menu, AccessorMenu, Default,Open document
		}
		Menu, AccessorMenu, add, Open path in explorer,AccessorOpenExplorer
		Menu, AccessorMenu, add, Open path in CMD,AccessorOpenCMD
	}
	Menu, AccessorMenu, add, Copy Path (CTRL+C), AccessorCopyPath
	Menu, AccessorMenu, add, Explorer context menu, AccessorExplorerContextMenu
}
FileSystemSetFolder(subfolder)
{
	global Accessor, AccessorEdit
	outputdebug set folder to %subfolder%
	GUINum := Accessor.GUINum
	Gui, %GUINum%: Default
	GUI, ListView, AccessorListView
	GuiControlGet, Filter, , AccessorEdit
	Filter := ExpandPathPlaceholders(Filter)
	SplitPath, Filter, name, dir,,,drive
	if(dir)
	{
		GuiControl, ,AccessorEdit,%dir%\%subfolder%\			
		hwndEdit := Accessor.hwndEdit
		SendMessage, 0xC1, -1,,, AHK_ID %hwndEdit%  ; EM_LINEINDEX (Gets index number of line)
		CaretTo := ErrorLevel
		SendMessage, 0xB1, CaretTo, CaretTo,, AHK_ID %hwndEdit% ;EM_SETSEL
	}
}