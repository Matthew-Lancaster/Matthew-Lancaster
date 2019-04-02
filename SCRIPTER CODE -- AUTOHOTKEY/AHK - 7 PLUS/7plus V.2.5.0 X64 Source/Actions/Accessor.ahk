Class CAccessorAction Extends CAction
{
	static Type := RegisterType(CAccessorAction, "Show Accessor")
	static Category := RegisterCategory(CAccessorAction, "Window")
	static FlashingWindows := 1
	
	__New()
	{
	}
	Execute(Event)
	{
		if(!this.tmpGuiNum)
		{		
			if(this.FlashingWindows)
			{
				result:=FlashingWindows(this) ;Since FlashingWindows function also uses an object value called FlashingWindows, it can straightly use this action here
				if(result)
					return 1
			}
			result := CreateAccessorWindow(this)
			return result > 1 ? 1 : (result = 1 ? 1 : 0)
		}
		else
		{
			GuiNum := this.tmpGuiNum
			Gui,%GuiNum%:+LastFound 
			WinGet, hwnd,ID
			DetectHiddenWindows, Off
			If(WinExist("ahk_id " hwnd)) ;window not closed yet, need more processing time
				return -1
			else
				return 1 ;window closed, all fine
		}
	}

	DisplayString()
	{
		return "Show accessor"
	}

	GuiShow(ActionGUI, GoToLabel = "")
	{
		this.AddControl(ActionGUI, "Checkbox", "FlashingWindows", "Activate flashing windows first")
	}

	OnExit()
	{
		global Accessor
		Accessor.OnExit()
	}
}
