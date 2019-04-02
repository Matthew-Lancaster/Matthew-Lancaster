#InstallKeybdHook

if ( WaitForAnyKey( 10000 ) )
	MsgBox, <any key> pressed
else
	MsgBox, [WaitForAnyKey] timeout expired
ExitApp

WaitForAnyKey( p_timeout )
{
	start := A_TickCount

	loop,
	{
		if ( idle > A_TimeIdlePhysical )
			return, true
		else if ( A_TickCount-start >= p_timeout )
			return, false
	
		idle := A_TimeIdlePhysical

		Sleep, 1
	}
}