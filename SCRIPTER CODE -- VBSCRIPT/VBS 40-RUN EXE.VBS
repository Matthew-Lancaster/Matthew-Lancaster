' -------------------------------------------------------------------
' [ Sunday 05:05:50 Am_03 March 2019 ]
' -------------------------------------------------------------------
' THIS FILE BECAME USE BY
' LOT OF AUTOHOTKEYS TRAY MENU ITEM
' AND THEY USE IT HERE THIS ONE
' C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03_PROCESS_KILLER.BAT
' ORIGINAL AT HERE LOCATION 
' C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 39-KILL PROCESS.VBS
' AND MOVED HERE MAYBE 
' -------------------------------------------------------------------

Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

Set objArgs = Wscript.Arguments
' -------------------------------------------------------------------
IF objArgs.COUNT=0 THEN 
	SCRIPT_NAME=WScript.ScriptFullName
	SCRIPT_NAME=REPLACE(SCRIPT_NAME , WScript.ScriptName , VBCRLF + VBCRLF + WScript.ScriptName)
	IMSGBOX=""
	IMSGBOX=IMSGBOX+ "NOTHING TO RUN NONE PARAMETER PASS ON TO ME -- MSGBOX FROM"
	IMSGBOX=IMSGBOX+ VBCRLF + VBCRLF + VBCRLF 
	IMSGBOX=IMSGBOX+ SCRIPT_NAME
	MSGBOX IMSGBOX
	Wscript.QUIT
END IF
' -------------------------------------------------------------------

Set objArgs = Wscript.Arguments
For Each strArg In objArgs

	PARAM = PARAM + """"+REPLACE(strArg,"*"," ")+"""" +" "
	
Next

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run PARAM , ShowWindow, DontWaitUntilFinished



' -------------------------------------------------------------------
' -------------------------------------------------------------------
Wscript.Arguments.Item(0)

' WSHShell.Run "CMD /K """ + "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03_PROCESS_KILLER.BAT" + """ , ShowWindow, DontWaitUntilFinished

' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03_PROCESS_KILLER.BAT""" , ShowWindow, DontWaitUntilFinished

' WSHShell.Run """"+XF0+"""" + " " + ""+XF2_VOID+"" , ShowWindow, DontWaitUntilFinished
