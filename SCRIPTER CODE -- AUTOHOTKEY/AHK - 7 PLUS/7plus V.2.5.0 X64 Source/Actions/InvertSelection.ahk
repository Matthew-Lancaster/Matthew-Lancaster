Class CInvertSelectionAction Extends CAction
{
	static Type := RegisterType(CInvertSelectionAction, "Invert file selection")
	static Category := RegisterCategory(CInvertSelectionAction, "Explorer")
	
	Execute(Event)
	{
		InvertSelection(WinExist("A"))
		return 1
	} 

	DisplayString(Action)
	{
		return "Invert selection of active explorer window"
	}
}