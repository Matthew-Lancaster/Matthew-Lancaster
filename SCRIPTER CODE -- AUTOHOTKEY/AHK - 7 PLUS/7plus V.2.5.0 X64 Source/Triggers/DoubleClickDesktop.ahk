Class CDoubleClickDesktopTrigger Extends CTrigger
{
	static Type := RegisterType(CDoubleClickDesktopTrigger, "Double click on desktop")
	static Category := RegisterCategory(CDoubleClickDesktopTrigger, "Hotkeys")
	
	Matches(Filter)
	{
		return true ;type is checked elsewhere
	}
}

#MaxThreadsPerHotkey 2
#if (Vista7 && IsWindowUnderCursor("WorkerW")) || (!Vista7 && IsWindowUnderCursor("ProgMan"))
~LButton::DoubleClickDesktop()
#if
#MaxThreadsPerHotkey 1

DoubleClickDesktop()
{
	CurrentDesktopFiles:=GetSelectedFiles()
	if(IsDoubleClick() && CurrentDesktopFiles = "")
	{
		Trigger := new CDoubleClickDesktopTrigger()
		EventSystem.OnTrigger(Trigger)
	}
	Return
}