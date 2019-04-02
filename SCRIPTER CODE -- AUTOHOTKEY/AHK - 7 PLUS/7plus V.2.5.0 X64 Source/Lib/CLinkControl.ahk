/*
Class: CLinkControl
A link control that can be used for hyperlinks. It uses HTML-like markup language for links, e.g. <a href="http://www.url.com">URL Text</a>.
href tags are attempted to be executed directly, while id tags are solely handled in script.

This control extends <CControl>. All basic properties and functions are implemented and documented in this class.
*/
Class CLinkControl Extends CControl
{
	__New(Name, Options, Text, GUINum)
	{
		Base.__New(Name, Options, Text, GUINum)
		this.Type := "Link"
		this._.Insert("Events", ["Click", "DoubleClick"])
	}
	/*
	Event: Introduction
	To handle control events you need to create a function with this naming scheme in your window class: ControlName_EventName(params)
	The parameters depend on the event and there may not be params at all in some cases.
	Additionally it is required to create a label with this naming scheme: GUIName_ControlName
	GUIName is the name of the window class that extends CGUI. The label simply needs to call CGUI.HandleEvent().
	For better readability labels may be chained since they all execute the same code.
	Instead of using ControlName_EventName() you may also call <CControl.RegisterEvent> on a control instance to register a different event function name.
	
	Event: Click(URLOrID, Index)
	Invoked when the user clicked on a link.
	*/
	HandleEvent(Event)
	{
		this.CallEvent("Click", Event.ErrorLevel, Event.EventInfo)
	}
}