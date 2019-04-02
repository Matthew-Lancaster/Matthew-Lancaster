/*
Class: CEditControl
An edit control.

This control extends <CControl>. All basic properties and functions are implemented and documented in this class.
*/
Class CEditControl Extends CControl
{
	__New(Name, Options, Text, GUINum)
	{
		Base.__New(Name, Options, Text, GUINum)
		this.Type := "Edit"
		this._.Insert("ControlStyles", {Center : 0x1, LowerCase : 0x10, Number : 0x2000, Multi : 0x4, Password : 0x20, ReadOnly : 0x800, Right : 0x2, Uppercase : 0x8, WantReturn : 0x1000})
		this._.Insert("ControlMessageStyles", {ReadOnly: {Message: 0xCF, On: {W: 1}, Off: {W: 0}}})
		this._.Insert("Events", ["TextChanged"])
		this._.Insert("Messages", {0x200 : "KillFocus", 0x100 : "SetFocus" }) ;Used for automatically registering message callbacks
	}
	/*
	Function: AddUpDown
	Adds an UpDown control to this text field. This function needs to be called immediately after adding the edit control to the window.
	
	Parameters:
		Min - The minimum value of the UpDown control.
		Max - The maximum value of the UpDown control.
	*/
	AddUpDown(Min, Max)
	{
		WM_USER := 0x0400
		UDM_SETBUDDY := WM_USER + 105
		;If this edit control belongs to a tab, set the correct tab first and unset it afterwards
		if(this.hParentControl && CGUI.GUIList[this.GUINum].Controls[this.hParentControl].Type = "Tab")
			Gui, % this.GUINum ":Tab", % this.TabNumber, % CGUI.GUIList[this.GUINum].Controls[this.hParentControl]._.TabIndex
		
		Gui, % this.GUINum ":Add", UpDown, Range%Min%-%Max% hwndhUpDown, % this.Text
		
		if(this.hParentControl && CGUI.GUIList[this.GUINum].Controls[this.hParentControl].Type = "Tab")
			Gui, % this.GUINum ":Tab"
		hwnd := this.hwnd
		;~ SendMessage, UDM_SETBUDDY, hwnd, 0,, % "ahk_id " hwnd
		this._.UpDownHwnd := hUpDown
		this._.Min := Min
		this._.Max := Max
	}
	;~ HWND CreateControl(const ustring& classname,const HWND hParent,DWORD extstyle, const HINSTANCE hInst,DWORD dwStyle, const RECT& rc,const int id)
	;~ {
		;~ dwStyle|=WS_CHILD|WS_VISIBLE;
		;~ DllCall("CreateWindowEx", "UINT", extStyle, "Str", classname, "Str", "", "UINT", dwStyle, "UINT", x, "UINT", y, "UINT", w, "UINT", h, "PTR", this.hwnd, "PTR", 0, "PTR", 0, "UINT", 0
		;~ return CreateWindowEx(extstyle,           //extended styles
							  ;~ classname.c_str(),  //control 'class' name
							  ;~ 0,                  //control caption
							  ;~ dwStyle,            //wnd style
							  ;~ rc.left,            //position: left
							  ;~ rc.top,             //position: top
							  ;~ rc.right,           //width
							  ;~ rc.bottom,          //height
							  ;~ hParent,            //parent window handle
							  ;~ //control's ID
							  ;~ reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
							  ;~ hInst,              //instance
							  ;~ 0);                 //user defined info
	;~ }
	__Get(Name)
    {
		;~ if(Name != "GUINum" && !CGUI.GUIList[this.GUINum].IsDestroyed)
		;~ {
			;~ if(Name = "Text" && this._.UpDownHwnd) ;Use text from UpDown control if possible
				;~ GuiControlGet, Value, % this.GUINum ":", % this.ClassNN
		;~ }
		;~ if(Value)
			;~ return Value
	}
	
	/*
	Property: Min
	If AddUpDown() has been called befored, the minimum value can be changed here.
	
	Property: Max
	If AddUpDown() has been called befored, the maximum value can be changed here.
	*/
	__Set(Name, Params*)
	{
		if(Name != "GUINum" && !CGUI.GUIList[this.GUINum].IsDestroyed)
		{
			;Fix completely weird __Set behavior. If one tries to assign a value to a sub item, it doesn't call __Get for each sub item but __Set with the subitems as parameters.
			Value := Params.Remove()
			if(Params.MaxIndex())
			{
				Params.Insert(1, Name)
				Name := Params.Remove()
				return (this[Params*])[Name] := Value
			}
			if(this._.UpDownHwnd && this._.HasKey({Min : "Min", Max : "Max"}[Name]))
			{
				SendMessage, 0x400 + 111, this._.Min := (Name = "Min" ? Value : this._.Min), this._.Max := (Name = "Max" ? Value : this._.Max),,% "ahk_id " this._.UpDownHwnd
				return Value
			}
		}
	}
	
	/*
	Event: Introduction
	To handle control events you need to create a function with this naming scheme in your window class: ControlName_EventName(params)
	The parameters depend on the event and there may not be params at all in some cases.
	Additionally it is required to create a label with this naming scheme: GUIName_ControlName
	GUIName is the name of the window class that extends CGUI. The label simply needs to call CGUI.HandleEvent().
	For better readability labels may be chained since they all execute the same code.
	Instead of using ControlName_EventName() you may also call <CControl.RegisterEvent> on a control instance to register a different event function name.
	
	Event: TextChanged()
	Invoked when the text of the control is changed.
	*/
	HandleEvent(Event)
	{
		this.CallEvent("TextChanged")
	}
}