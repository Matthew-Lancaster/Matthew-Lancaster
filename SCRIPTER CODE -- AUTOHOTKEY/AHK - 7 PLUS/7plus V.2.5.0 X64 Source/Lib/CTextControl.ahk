/*
Class: CTextControl
A static text control that can also be used as hyperlink.

This control extends <CControl>. All basic properties and functions are implemented and documented in this class.
*/
Class CTextControl Extends CControl
{
	__New(Name, Options, Text, GUINum)
	{
		Base.__New(Name, Options, Text, GUINum)
		this.Type := "Text"
		this._.Insert("ControlStyles", {Center : 0x1, Left : 0, Right : 0x2, Wrap : -0xC})
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
	
	Event: Click()
	Invoked when the user clicked on the control.
	
	Event: DoubleClick()
	Invoked when the user double-clicked on the control.
	*/
	HandleEvent(Event)
	{
		this.CallEvent(Event.GUIEvent = "DoubleClick" ? "DoubleClick" : "Click")
	}
}