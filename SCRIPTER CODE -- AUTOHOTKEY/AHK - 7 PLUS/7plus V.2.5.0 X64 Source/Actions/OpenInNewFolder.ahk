Class COpenInNewFolderAction Extends CAction
{
	static Type := RegisterType(COpenInNewFolderAction, "Open folder in new window / tab")
	static Category := RegisterCategory(COpenInNewFolderAction, "Explorer")
	static Action := "Tab in Background"
	
	Execute(Event)
	{
		OpenInNewFolder(this)
	}

	DisplayString()
	{
		return "Open explorer folder under mouse in new window/tab"
	}

	GuiShow(GUI)
	{
		this.AddControl(GUI, "Text", "Desc", "This action opens the explorer folder under the mouse in a new window, tab or in a tab without activating it.")
		this.AddControl(GUI, "DropDownList", "Action", "Tab|Tab in Background|Window", "", "Open in new:")
	}
}

;Opens the folder under the mouse in a new window or tab
OpenInNewFolder(Action)
{
 	if(!(hwnd := WinActive("ahk_group ExplorerGroup"))||!IsMouseOverFileList())
 		return false
	selected:=GetSelectedFiles(0)
	Send {LButton}
	Sleep 100
	if(InStr(FileExist(undermouse:=GetSelectedFiles()), "D"))
		dir:=true
	if(select!=selected)
		SelectFiles(selected,1,0,0)
	if(!dir)
		return false
	if(Action.Action = "Tab" && Settings.Explorer.Tabs.UseTabs)
		CreateTab(hwnd,undermouse, 1)
	else if(Action.Action = "Tab in background" && Settings.Explorer.Tabs.UseTabs)
		CreateTab(hwnd,undermouse, 0)
	else
		Run(A_WinDir "\explorer.exe /n,/e," undermouse)
	return true
}