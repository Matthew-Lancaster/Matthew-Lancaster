;Generic Run interface for subevents. They can implement this interface like this:
;static _ImplementsRun := ImplementRunInterface(CSubEvent)
;It's important to use a "_" or "tmp" at the start of the name to mark this property as temporary so it won't be saved.
ImplementRunInterface(Run)
{	
	Run.WaitForFinish := false
	Run.RunAsAdmin := false
	Run.Command := "cmd.exe"
	Run.WorkingDirectory := ""
	if(Run.HasKey("__Class"))
	{
		Run.RunExecute := Func("Run_Execute")
		Run.RunDisplayString := Func("Run_DisplayString")
		Run.RunGUIShow := Func("Run_GUIShow")
		Run.RunGUISubmit := Func("Run_GUISubmit")
	}
}

Run_Execute(SubEvent, Event)
{
	if(!SubEvent.tmpPid)
	{
		command := Event.ExpandPlaceholders(SubEvent.Command)
		WorkingDirectory := Event.ExpandPlaceholders(SubEvent.WorkingDirectory)
		if(SubEvent.WaitForFinish)
		{
			SubEvent.tmpPid := Run(command, WorkingDirectory, "", !SubEvent.RunAsAdmin)
			if(SubEvent.tmpPid) ;If retrieved properly
				return -1
			MsgBox Waiting for %command% failed!
			return 0
		}
		else
			Run(command, WorkingDirectory, "", !SubEvent.RunAsAdmin)
	}
	else
	{
		pid := SubEvent.tmpPid
		Process, Exist, %pid%
		if(ErrorLevel)
			return -1
	}
	return 1
}

Run_DisplayString(SubEvent)
{
	return "Run " SubEvent.Command
}

Run_GuiShow(SubEvent, GUI, GoToLabel = "")
{
	if(GoToLabel = "")
	{
		SubEvent.tmpRunGUI := GUI
		SubEvent.AddControl(GUI, "Text", "Text", "Enclose paths with spaces in quotes and append parameters in command field.")
		SubEvent.AddControl(GUI, "Edit", "Command", "", "", "Command:","Browse", "Action_Run_Browse", "Placeholders", "Action_Run_Placeholders")
		SubEvent.AddControl(GUI, "Edit", "WorkingDirectory", "", "", "Working Dir:","Browse", "Action_Run_Browse_WD", "Placeholders", "Action_Run_Placeholders_WD")
		SubEvent.AddControl(GUI, "Checkbox", "WaitForFinish", "Wait for finish", "", "")
		SubEvent.AddControl(GUI, "DropDownList", "RunAsAdmin", "-1: Current permissions|0: Standard User|1: Elevated", "", "Run as admin")
	}
	else if(GoToLabel = "Browse")
		SubEvent.SelectFile(SubEvent.tmpRunGUI, "Command", "Select File", "", 1)
	else if(GoToLabel = "Placeholders")
		ShowPlaceholderMenu(SubEvent.tmpRunGUI, "Command")
	else if(GoToLabel = "Browse_WD")
		SubEvent.Browse(SubEvent.tmpRunGUI, "WorkingDirectory", "Select working directory", "", 1)
	else if(GoToLabel = "Placeholders_WD")
		ShowPlaceholderMenu(SubEvent.tmpRunGUI, "WorkingDirectory")
}
Run_GUISubmit(SubEvent, GUI)
{
	SubEvent.Remove("tmpRunGUI")
}
Action_Run_Browse:
GetCurrentSubEvent().RunGuiShow("", "Browse")
return
Action_Run_Placeholders:
GetCurrentSubEvent().RunGuiShow("", "Placeholders")
return
Action_Run_Browse_WD:
GetCurrentSubEvent().RunGuiShow("", "Browse_WD")
return
Action_Run_Placeholders_WD:
GetCurrentSubEvent().RunGuiShow("", "Placeholders_WD")
return

Run(Target, WorkingDir = "", Mode = "", NonElevated=-1) 
{
	;run as current user
	if(!Vista7 || NonElevated = -1 || (!A_IsAdmin && NonElevated) || (A_IsAdmin && !NonElevated))
	{
		Run, %Target% , %WorkingDir%, %Mode% UseErrorLevel, v
		if(A_LastError)
			Msgbox Error launching %Target%
		Return, v
	}
	
	;Run under explorer process as normal user
	if(A_IsAdmin && NonElevated)
		return RunAsUser(Target, WorkingDir, Mode)
	
	;Split command and argument
	SplitCommandLine(Target, Args)
	
	;Show UAC prompt and run elevated
	if(!A_IsAdmin && !NonElevated)
	{		
		If(RunAsAdmin(Target, args, WorkingDir)) ;UAC prompt confirmed
			return 0
	}
	;Still here, error
	Msgbox Error launching %Target%
}

SplitCommandLine(ByRef Target, ByRef Args)
{
	;Split command and argument
	if(InStr(Target, """")=1 && InStr(Target, """",false,3)) ;command has quotes, split it
	{
		Args := SubStr(Target,InStr(Target, """", false, 3) + 2)
		Target := SubStr(Target, 2,InStr(Target, """", false, 3) - 2)
	}
	else if(InStr(Target, " ")) ;look for spaces after the command, e.g. "C:\Program Files\bla.exe -arg"
	{
		Args := SubStr(Target, InStr(Target, " ", false) + 1)
		Target := SubStr(Target, 1, InStr(Target, " ", false) - 1)
	}
	else
		Args := "" ;Single Command
}

GetWorkingDir(Command)
{
	WorkingDir := ""
	if(Exists := FileExist(Command) && !InStr(Exists, "D"))
		SplitPath, Command,, WorkingDir
	return WorkingDir
}
RunAsUser(Command, WorkingDir, Options)
{
	result := DllCall("Explorer.dll\CreateProcessMediumIL", Str, Command, Str, WorkingDir, Str, Options, "UInt")
	if(A_LastError = 740) ;ERROR_ELEVATION_REQUIRED
	{
		Run, %Command% , %WorkingDir%, %Mode% UseErrorLevel, v
		if(A_LastError)
			Msgbox Error launching %Target%
		Return, v
	}
}
RunAsAdmin(target, args, WorkingDir)
{
	uacrep := DllCall("shell32\ShellExecute", uint, 0, str, "RunAs", str, target, str, args, str, WorkingDir, int, 1)
	return uacrep = 42 ;UAC dialog confirmed
}