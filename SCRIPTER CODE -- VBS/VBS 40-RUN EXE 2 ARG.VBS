' -------------------------------------------------------------------
' [ Sunday 05:05:50 Am_03 March 2019 ]
' -------------------------------------------------------------------
' THIS FILE BECAME USE BY
' LOT OF AUTOHOTKEYS TRAY MENU ITEM
' AND THEY USE IT HERE THIS ONE
' C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03_PROCESS_KILLER.BAT
' ORIGINAL AT HERE LOCATION 
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
' AND MOVED HERE MAYBE 
' -------------------------------------------------------------------

Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

Set args = Wscript.Arguments
For Each arg In args
	XF0 = XF0 +""""+ arg + """ "
	XF0 = Arg 
	EXIT FOR
Next

MSGBOX XF0,vbMsgBoxSetForeground,Wscript.ScriptName 

'If Sent-to is Used Sometime Additional Quote Surround are User
'--------------------------------------------------------------
XF0 = UCASE(Replace(XF0,"""""",""""))

MSGBOX XF0,vbMsgBoxSetForeground,Wscript.ScriptName 


' WSHShell.Run "CMD /K """ + "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03_PROCESS_KILLER.BAT" + """ , ShowWindow, DontWaitUntilFinished

' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03_PROCESS_KILLER.BAT""" , ShowWindow, DontWaitUntilFinished


WSHShell.Run XF0 , ShowWindow, DontWaitUntilFinished

'WSHShell.Run XF0+ " " + ""+XF2_VOID+"" , ShowWindow, WaitUntilFinished
