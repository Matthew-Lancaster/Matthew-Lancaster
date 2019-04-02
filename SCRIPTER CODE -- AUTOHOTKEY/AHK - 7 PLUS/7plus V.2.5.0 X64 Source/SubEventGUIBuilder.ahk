;Adds a row of controls to a GUI using a value associated with ValueObject. Control handles are stored in GUI, and can be stored back and have the controls delete by CSubEvent.GUISubmit
AddControl(ValueObj, GUI, type, name, text = "", glabel = "", description = "", Button1Text = "", Button1gLabel = "", Button2Text = "", Button2gLabel = "", Tooltip = "")
{
	x := GUI.x
	y := GUI.y
	w := 200
	if(description != "")
	{
		y += 4
		Gui, Add, Text, x%x% y%y% hwndDesc_%name%, %description%
		GUI["Desc_" name] := Desc_%name%
		x += 70
		y -= 4
	}
	if(type = "Text")
	{
		y += 4
		if(description = "") ;No description, more space for text
			w += 100
		if(!Button1Text)
		{
			w += 60
			if(!Button2Text)
				w += 60
		}
		Gui, Add, Text, x%x% y%y% w%w% hwndText_%name%, %text%
		if(GUI.HasKey("Text_" name))
			Msgbox Subevent GUI control already exists: %name%
		GUI["Text_" name] := Text_%name%
		if(description = "")
			w -= 100
		ControlGetPos,,,,h,,% "ahk_id " Text_%name%
		y += h + 7
		if(Tooltip)
			AddToolTIp(Text_%name%, Tooltip)
	}
	else if(type = "Checkbox")
	{
		if(ValueObj[name] = 1)
			Gui, Add, Checkbox, x%x% y%y% hwndCheck_%name% Checked, %text%
		else
		{
			Gui, Add, Checkbox, x%x% y%y% hwndCheck_%name%, %text%
			if(ValueObj[name] != 0)
				Msgbox SubEvent.AddControl(%type%,%name%, %text%, %description%) has wrong checkbox value!
		}
		y += 30		
		if(GUI.HasKey("Check_" name))
			Msgbox Subevent GUI control already exists: %name%
		GUI["Check_" name] := Check_%name%
		if(Tooltip)
			AddToolTIp(Check_%name%, Tooltip)
	}
	else if(type = "Edit")
	{
		y += 1
		text := ValueObj[name]
		options = x%x% y%y% w%w% hwndEdit_%name% -Multi R1
		options .= InStr(name, "password") ? " Password" : ""
		Gui, Add, Edit, %options% , %text%
		y -= 1
		y += 30
		if(GUI.HasKey("Edit_" name))
			Msgbox Subevent GUI control already exists: %name%
		GUI["Edit_" name] := Edit_%name%
		if(Tooltip)
			AddToolTIp(Edit_%name%, Tooltip)
	}
	else if(type = "Button")
	{
		Gui, Add, Button, x%x% y%y% w%w% hwndButton_%name% g%gLabel% r1 -Wrap, %text%
		y += 30
		if(GUI.HasKey("Button_" name))
			Msgbox Subevent GUI control already exists: %name%
		GUI["Button_" name] := Button_%name%
		if(Tooltip)
			AddToolTIp(Button_%name%, Tooltip)
	}
	else if(type = "DropDownList" || type = "ComboBox")
	{
		;Select event dropdownlist. This only works while the settings window is open.
		if(strStartsWith(text, "TriggerType:"))
		{
			Triggertype := SubStr(text, 13)
			text := ""
			Loop % SettingsWindow.Events.MaxIndex()
			{
				if(!TriggerType || SettingsWindow.Events[A_Index].Trigger.Type = TriggerType)
					text .= SettingsWindow.Events[A_Index].ID ": " SettingsWindow.Events[A_Index].Name "|"
			}
		}
		;Construct options
		Loop, Parse, text, |
		{
			if(A_LoopField)
				if(InStr(A_LoopField, ": ")) ;if list entries start with "\d+: " or similar, it is sufficient to store \d+ in the assigned variable to select its item
					text1 .= A_LoopField (SubStr(A_LoopField, 1, InStr(A_LoopField, ": ") - 1) = ValueObj[name] ? "||" : "|")
				else
					text1 .= A_LoopField (A_LoopField = ValueObj[name] ? "||" : "|")
		}
		if(!strEndsWith(text1,"||"))
			text1 := SubStr(text1, 1, -1)
		if(type = "DropDownList")
		{
			options = x%x% y%y% w%w% hwndDropDown_%name%
			if(gLabel != "")
				options .= " g" gLabel
			Gui, Add, DropDownList, %options%, %text1%
			if(GUI.HasKey("DropDown_" name))
				Msgbox Subevent GUI control already exists: %name%
			GUI["DropDown_" name] := DropDown_%name%
			if(Tooltip)
				AddToolTIp(DropDown_%name%, Tooltip)
		}
		else if(type = "ComboBox")
		{
			options = x%x% y%y% w%w% hwndComboBox_%name%
			if(gLabel != "")
				options .= " g" gLabel
			Gui, Add, ComboBox, %options%, %text1%
			if(GUI.HasKey("ComboBox_" name))
				Msgbox Subevent GUI control already exists: %name%
			GUI["ComboBox_" name] := ComboBox_%name%
			if(!InStr(text1, "||"))
				ControlSetText, , % ValueObj[name], % "ahk_id " ComboBox_%name%
			if(Tooltip)
				AddToolTIp(ComboBox_%name%, Tooltip)
		}
		y += 30
	}	
	else if(type = "Time")
	{
		text := ValueObj[name]
		outputdebug time %text%
		Gui, Add, DateTime, x%x% y%y% hwndTime_%name% Choose20100101%text%, Time
		y += 30		
		if(GUI.HasKey("Time_" name))
			Msgbox Subevent GUI control already exists: %name%
		GUI["Time_" name] := Time_%name%
		if(Tooltip)
			AddToolTIp(Time_%name%, Tooltip)
	}
	if(Button1Text != "")
	{				
		x += 210
		y := GUI.y
		w := 70
		if(Button2Text != "")
			Gui, Add, Button, x%x% y%y% w%w% hwndButton1_%name% g%Button1gLabel% r1 -Wrap, %Button1Text%
		else
			Gui, Add, Button, x%x% y%y% hwndButton1_%name% g%Button1gLabel% r1 -Wrap, %Button1Text%
		y += 30
		GUI["Button1_" name] := Button1_%name%
		if(Button2Text != "")
		{		
			x += 76
			y := GUI.y
			Gui, Add, Button, x%x% y%y% w%w% hwndButton2_%name% g%Button2gLabel% r1 -Wrap, %Button2Text%
			y += 30
			GUI["Button2_" name] := Button2_%name%
		}
	}
	GUI.y := y
}

;Stores the values of the controls in GUI in their associated field in ValueObj. This function is used in combination with AddControl().
SubmitControls(ValueObj, GUI)
{
	;Loop over all controls added to GUI, and store their results and delete them
	enum := GUI._newEnum()
	while enum[key,value]
	{
		if(strStartsWith(key, "Desc_") || strStartsWith(key, "Text_") || strStartsWith(key, "Button"))
		{
			WinKill, ahk_id %value%
		}
		else if(strStartsWith(key, "Check_"))
		{
			name := SubStr(key,7)
			ControlGet, Checked, Checked, , ,ahk_id %value%
			ValueObj[name] := Checked
			outputdebug save %name% %checked%
			WinKill, ahk_id %value%
		}
		else if(strStartsWith(key, "Edit_"))
		{
			name := SubStr(key, 6)
			ControlGetText, text, , ahk_id %value%
			ValueObj[name] := text
			outputdebug save %name% %text%
			WinKill, ahk_id %value%
		}
		else if(strStartsWith(key, "DropDown_"))
		{
			name := SubStr(key, 10)
			ControlGetText, text, , ahk_id %value%
			if(InStr(text, ": ")) ;If the selection starts with "\d+: " or similar, (\d+) is returned instead
				text := SubStr(text, 1, InStr(text, ": ") - 1)
			outputdebug save %name% %text%
			ValueObj[name] := text
			WinKill, ahk_id %value%
		}
		else if(strStartsWith(key, "ComboBox_"))
		{
			name := SubStr(key, 10)
			ControlGetText, text, , ahk_id %value%
			if(InStr(text, ": ")) ;If the selection starts with "\d+: " or similar, (\d+) is returned instead
				text := SubStr(text, 1, InStr(text, ": ") - 1)
			outputdebug save %name% %text%
			ValueObj[name] := text
			WinKill, ahk_id %value%
		}
		else if(strStartsWith(key, "Time_"))
		{
			name := SubStr(key, 6)
			ControlGetText, text, , ahk_id %value%
			StringReplace, text, text, :,,All
			outputdebug store %text% in %name%
			ValueObj[name] := text
			WinKill, ahk_id %value%
		}
	}
}

;Shows a browse dialog and shows the result in the GUI control associated with "name". This function is used in combination with AddControl()
Browse(Subevent, GUI, name, Title="Select Folder",Options=0, Quote=0)
{
	Gui +OwnDialogs
	path:=COMObjCreate("Shell.Application").BrowseForFolder(0, Title, Options).Self.Path
	if(path!="")
	{
		enum := GUI._newEnum()
		while enum[key,value]
		{
			if(InStr(key,"_" name) && !InStr(key, "Button1_") && !InStr(key, "Button2_") && !InStr(key, "Desc_"))
			{
				if(Quote)
					path := Quote(path)
				ControlSetText, , %path%, ahk_id %value%
				break
			}
		}
	}
}
;Shows a file selection dialog and shows the result in the GUI control associated with "name". This function is used in combination with AddControl()
SelectFile(SubEvent, GUI, name, Title = "Select File", Filter = "", Quote=0, options = 3)
{
	Gui +OwnDialogs
	FileSelectFile, path , %options%, , %Title%, %Filter%
	if(path != "")
	{
		enum := GUI._newEnum()
		while enum[key,value]
		{
			if(InStr(key,"_" name) && !InStr(key, "Button1_") && !InStr(key, "Button2_") && !InStr(key, "Desc_"))
			{
				if(Quote)
					path := Quote(path)
				ControlSetText, , %path%, ahk_id %value%
				break
			}
		}
	}
}