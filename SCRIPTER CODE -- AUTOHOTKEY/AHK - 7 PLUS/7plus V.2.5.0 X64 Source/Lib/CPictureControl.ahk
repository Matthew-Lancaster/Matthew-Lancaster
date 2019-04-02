/*
Class: CPictureControl
A picture control.

This control extends <CControl>. All basic properties and functions are implemented and documented in this class.
*/
Class CPictureControl Extends CControl
{
	__New(Name, Options, Text, GUINum)
	{
		base.__New(Name, Options, Text, GUINum)
		this.Type := "Picture"
		this._.Picture := Text
		this._.Insert("Picture", "Text")
		this._.Insert("ControlStyles", {Center : 0x200, ResizeImage : 0x40})
		this._.Insert("Events", ["Click", "DoubleClick"])
	}
	
	/*
	Property: Picture
	The picture can be changed by assigning a filename to this property.
	If the picture was set by providing a hBitmap in <SetImageFromHBitmap>, this variable will be empty.
	
	Property: PictureWidth
	The width of the currently displayed picture in pixels.
	
	Property: PictureHeight
	The height of the currently displayed picture in pixels.
	*/
	__Get(Name)
    {
		if(Name != "GUINum" && !CGUI.GUIList[this.GUINum].IsDestroyed)
		{
			DetectHidden := A_DetectHiddenWindows
			DetectHiddenWindows, On
			if(Name = "Picture")
				Value := this._.Picture
			else if(Name = "PictureWidth" || Name = "PictureHeight")
			{
				if(!this._.HasKey(Name))
				{
					pBitmap := Gdip_CreateBitmapFromFile(this._.Picture) ;Note: This does not work if the user specifies icon number or other special properties in picture path since these are handled as separate parameters here.
					Gdip_GetImageDimensions(pBitmap, w, h)
					this._.PictureWidth := w
					this._.PictureHeight := h
					Gdip_DisposeImage(pBitmap)
				}
				Value := this._[Name]
			}
			if(!DetectHidden)
				DetectHiddenWindows, Off
			if(Value != "")
				return Value
		}
	}
	
	__Set(Name, Params*)
	{
		if(!CGUI.GUIList[this.GUINum].IsDestroyed)
		{
			;Fix completely weird __Set behavior. If one tries to assign a value to a sub item, it doesn't call __Get for each sub item but __Set with the subitems as parameters.
			Value := Params.Remove()
			if(Params.MaxIndex())
			{
				Params.Insert(1, Name)
				Name := Params.Remove()
				return (this[Params*])[Name] := Value
			}
			DetectHidden := A_DetectHiddenWindows
			DetectHiddenWindows, On
			Handled := true
			if(Name = "Picture")
			{
				Gui, % this.GUINum ":Default"
				GuiControl,, % this.ClassNN, %Value%
				this._.Picture := Value
				this._.Remove("PictureWidth")
				this._.Remove("PictureHeight")
			}
			else
				Handled := false
		if(!DetectHidden)
			DetectHiddenWindows, Off
		if(Handled)
			return Value
		}
	}
	/*
	Function: SetImageFromHBitmap
	Sets the image of this control.
	
	Parameters:
		hBitamp - The bitmap handle to which the picture of this control is set
	*/
	SetImageFromHBitmap(hBitmap, Path = "")
	{
		SendMessage, 0x172, 0x0, hBitmap,, % "ahk_id " this.hwnd
		DllCall("gdi32\DeleteObject", "PTR", ErrorLevel)
		this._.Remove("PictureWidth")
		this._.Remove("PictureHeight")
		this._.Picture := Path
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
	*/
	HandleEvent(Event)
	{
		this.CallEvent(Event.GUIEvent = "DoubleClick" ? "DoubleClick" : "Click")
	}
	

}