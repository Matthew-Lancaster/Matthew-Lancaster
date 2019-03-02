' -------------------------------------------------------------------
' [ Saturday 14:52:10 Pm_02 March 2019 ]
' -------------------------------------------------------------------
' THIS FILE BECAME USE BY
' LOT OF AUTOHOTKEYS TRAY MENU ITEM
' AND THEY USE IT HERE THIS ONE
' C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\BAT_03_PROCESS_KILLER.BAT
' ORIGINAL AT HERE LOCATION 
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 39-KILL PROCESS.VBS
' AND MOVED HERE MAYBE 
' -------------------------------------------------------------------


const DontWaitUntilFinished = false, ShowWindow = 1, DontShowWindow = 0, WaitUntilFinished = true

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

Set args = Wscript.Arguments
For Each arg In args
	XF0 = XF0 + arg+" "
Next

'If Sent-to is Used Sometime Additional Quote Surround are User
'--------------------------------------------------------------
XF0 = UCASE(Replace(XF0,"""",""))

MSGBOX XF0

' WSHShell.Run "CMD /K """ + "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\BAT_03_PROCESS_KILLER.BAT" + """ , ShowWindow, DontWaitUntilFinished

' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\BAT_03_PROCESS_KILLER.BAT""" , ShowWindow, DontWaitUntilFinished


WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\BAT_03_PROCESS_KILLER.BAT""" + " " + ""+XF0+"" , ShowWindow, DontWaitUntilFinished
