Class CFastFoldersMenuAction Extends CAction
{
	static Type := RegisterType(CFastFoldersMenuAction, "Fast Folders menu")
	static Category := RegisterCategory(CFastFoldersMenuAction, "Fast Folders")
	
	Execute(Event)
	{
		if(!this.tmpShowing)
		{
			this.tmpShowing := true
			FastFolderMenu()
		}
		else if(!IsContextMenuActive()) ;Menu closed
		{
			this.tmpShowing := false
			return 1
		}
		return -1 ;Waiting for menu to close
	}

	DisplayString()
	{
		return "Show Fast Folders menu"
	}
}