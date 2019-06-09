' -------------------------------------------------------------------
' [ Sunday 05:05:50 Am_03 March 2019 ]
' -------------------------------------------------------------------
' THIS FILE BECAME USE BY
' LOT OF AUTOHOTKEYS TRAY MENU ITEM
' AND THEY USE IT HERE THIS ONE
' C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT
' ORIGINAL AT HERE LOCATION 
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
' AND MOVED HERE MAYBE 
' -------------------------------------------------------------------

Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

Set args = Wscript.Arguments
For Each arg In args
	XF0 = XF0 + arg + " "
Next

'If Sent-to is Used Sometime Additional Quote Surround are User
'--------------------------------------------------------------
XF0 = UCASE(Replace(XF0,"""",""))

' WSHShell.Run "CMD /K """ + "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT" + """ , ShowWindow, DontWaitUntilFinished

' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT""" , ShowWindow, DontWaitUntilFinished


WSHShell.Run """"+XF0+"""" + " " + ""+XF2_VOID+"" , ShowWindow, DontWaitUntilFinished
