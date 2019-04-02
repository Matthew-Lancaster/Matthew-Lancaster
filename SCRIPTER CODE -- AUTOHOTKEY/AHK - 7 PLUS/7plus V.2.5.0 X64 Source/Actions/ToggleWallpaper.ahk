Class CToggleWallpaperAction Extends CAction
{
	static Type := RegisterType(CToggleWallpaperAction, "Change desktop wallpaper")
	static Category := RegisterCategory(CToggleWallpaperAction, "System")
	
	Execute(Event)
	{
		ToggleWallpaper()
	}
	
	DisplayString()
	{
		return "Toggle wallpaper (mouse needs to be on desktop)"
	}
}

ToggleWallpaper()
{
	global
	if(!IsMouseOverDesktop() || A_OSVersion != "WIN_7")
		return false
	ShellContextMenu("Desktop",1)
	return true
}