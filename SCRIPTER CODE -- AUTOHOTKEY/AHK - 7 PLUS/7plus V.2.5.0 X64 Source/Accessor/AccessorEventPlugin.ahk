Accessor_EventPlugin_Init(ByRef EventPlugin, PluginSettings)
{
	EventPlugin.Settings.Keyword := "Event"
	EventPlugin.DefaultKeyword := "Event"
	EventPlugin.KeywordOnly := false
	EventPlugin.MinChars := 2
	EventPlugin.OKName := "Execute"
	EventPlugin.Description := "Makes it possible to implement an Accessor function by using the event system.`nThe parameters after the keyword of an event can be accessed`nthrough the ${Acc1} - ${Acc9} placeholders."
	EventPlugin.HasSettings := True
}
Accessor_EventPlugin_ShowSettings(EventPlugin, PluginSettings, PluginGUI, GoToLabel = "")
{
	AddControl(PluginSettings, PluginGUI, "Edit", "Keyword", "", "", "Keyword:")
	AddControl(PluginSettings, PluginGUI, "Edit", "BasePriority", "", "", "Base Priority:")
}
Accessor_EventPlugin_GetDisplayStrings(EventPlugin, AccessorListEntry, ByRef Title, ByRef Path, ByRef Detail1, ByRef Detail2)
{
	Title := AccessorListEntry.Event.Trigger.Title
	Path := AccessorListEntry.Event.Trigger.Path
	Detail1 := AccessorListEntry.Event.Trigger.Detail1
	Detail2 := AccessorListEntry.Event.Trigger.Detail2
}
Accessor_EventPlugin_OnAccessorOpen(EventPlugin, Accessor)
{
	EventPlugin.Priority := EventPlugin.Settings.BasePriority
}
Accessor_EventPlugin_FillAccessorList(EventPlugin, Accessor, Filter, LastFilter, ByRef IconCount, KeywordSet, Parameters)
{
	for index, Event in EventSystem.Events
	{
		if(Event.Trigger.Is(CAccessorTrigger) && Filter = Event.Trigger.Keyword)
			Accessor.List.Insert(Object("Title", Event.Trigger.Title, "Path", Event.Trigger.Path, "Event", Event, "Type", "EventPlugin", "Icon", 1))
	}
}
Accessor_EventPlugin_PerformAction(EventPlugin, Accessor, AccessorListEntry)
{
	ScheduledEvent := AccessorListEntry.Event.TriggerThisEvent()
	for index, Parameter in Accessor.LastParameters
		ScheduledEvent.Placeholders["Acc" index] := Parameter
}