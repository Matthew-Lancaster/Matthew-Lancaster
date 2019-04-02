Class CRunOrActivateAction Extends CAction
{
	static Type := RegisterType(CRunOrActivateAction, "Run a program or activate it")
	static Category := RegisterCategory(CRunOrActivateAction, "System")
	static _ImplementsRun := ImplementRunInterface(CRunOrActivateAction)
	
	Execute(Event)
	{
		if(this.tmpPid)
			return this.RunExecute(Event)
		else
		{
			Process, Exist, %name%
			if(Errorlevel != 0)
				WinActivate ahk_pid %ErrorLevel%
			else
				return this.RunExecute(Event)
		}
		return 1
	}
	
	DisplayString()
	{
		return "Run or activate " this.Command
	}
	
	GuiShow(GUI)
	{
		this.AddControl(GUI, "Text", "Desc", "This action will run a program or activate it if it is already running.")
		this.RunGUIShow(GUI)
	}
	
	GuiSubmit(GUI)
	{
		Base.GuiSubmit(GUI)
		this.RunGUISubmit(GUI)
	}
}