Class CMenuItemTrigger Extends CTrigger
{
	static Type := RegisterType(CMenuItemTrigger, "Menu item clicked")
	static Category := RegisterCategory(CMenuItemTrigger, "System")
	static Menu := "MenuName"
	static Name := "Menu entry"
	static Submenu := ""

	Enable()
	{
		if(this.Menu = "Tray")
			BuildMenu("Tray")
	}
	Disable()
	{
		if(this.Menu = "Tray")
			BuildMenu("Tray")
	}

	Delete()
	{
		if(this.Menu = "Tray")
			BuildMenu("Tray")
	}

	Matches(Filter)
	{
		return false ;This event is triggered through a CTriggerTrigger when appropriate
	}

	DisplayString()
	{
		return "MenuItem " this.Name " in " this.Menu
	}

	GuiShow(TriggerGUI)
	{
		Menus := Array()
		Loop % SettingsWindow.Events.MaxIndex()
		{
			if(SettingsWindow.Events[A_Index].Trigger.Is(CMenuItemTrigger) && Menus.indexOf(SettingsWindow.Events[A_Index].Trigger.Menu) = 0)
			{
				Menus.Insert(SettingsWindow.Events[A_Index].Trigger.Menu)
				MenuString .= (Menus.MaxIndex() = 1 ? "" : "|") SettingsWindow.Events[A_Index].Trigger.Menu
			}
		}
		
		this.AddControl(TriggerGUI, "ComboBox", "Menu", MenuString, "", "Menu:","","","","","The name of the menu to which this item gets added. If it does not exist, it is created here. You can later reference this menu name for submenus.")
		this.AddControl(TriggerGUI, "Edit", "Name", "", "", "Name:","","","","","The name of the menu item")
		this.AddControl(TriggerGUI, "Edit", "Submenu", "", "", "Submenu:","","","","","If you specify a menu name here, this event will not be able to get triggered. Instead it will show all MenuItem entries that use the same value as menu name as submenu.")
	}
}


MenuItemHandler:
MenuItemTriggered(A_ThisMenu, A_ThisMenuItem, A_ThisMenuItemPos)
return

MenuItemTriggered(menu, item, pos)
{
	if(menu = "Tray")
		pos -= A_IsCompiled ? 3 : 10
	index := 1
	Loop % EventSystem.Events.MaxIndex()
	{
		if(EventSystem.Events[A_Index].Trigger.Is(CMenuItemTrigger) && EventSystem.Events[A_Index].Trigger.Menu = menu)
		{
			if(index = pos)
			{
				Trigger := new CTriggerTrigger()
				Trigger.TargetID := EventSystem.Events[A_Index].ID
				EventSystem.OnTrigger(Trigger)				
				return
			}
			index++
		}
	}
	outputdebug MenuItem: Event not found! menu: %menu%, item: %item%, pos: %pos%
}