Class CMouseOverTabButton Extends CCondition
{
	static Type := RegisterType(CMouseOverTabButton, "Mouse over tab button")
	static Category := RegisterType(CMouseOverTabButton, "Mouse")
	
	Evaluate()
	{
		return IsMouseOverTabButton()
	}
	DisplayString()
	{
		return "Mouse is over Explorer tab button"
	}
}