Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

' -------------------------------------------------------------------
' WORK TODAY
' -------------------------------------------------------------------
' Wed 18-Sep-2019 21:21:07
' Wed 18-Sep-2019 22:55:17
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' HOW TO OPEN U A LAUNCHER FOR VBSCRIPT TO PASS THE ARGUMENT THE PROPER METHOD
' THE APP THE WANT USER HERE
' SEND ARGUMENT AND ENCODED IN SAME WAY AS URL WAY
' WITH A %20 FOR SPACE
' THE AUGMENT ARE DIVIDED BY SPACE
' AND UNABLE SPLIT ARRAY THEM IN WITHOUT A LOAD OF SPLIT LINE 
' WHERE EACH SPACE IS IN THE LINE
' I WAS JUST ONTO METHOD LOT PEOPLE USE FOR ADD WILDCARD FOR SPACE
' AND I TAKEN URL FORMAT
' IT SEEM THE CODE HERE IS FROM DEMO ON THE INTERNET
' AND I HAD HARD WORK SEARCHER
' AND NOTHING BETTER THAN THIS AROUND
' JUST A PATHETIC GET THE ARGS 
' AND THEN COMPLAIN FOR HALF AND HOUR
'
' USE STAR METHOD THAN URL METHOD SPACE TO STAR ENCODE
' 
' -------------------------------------------------------------------
' SOME SORT OF DISAPPEAR TRICK TO SUPPLY THE CODE
' -------------------------------------------------------------------
' [ Wednesday 21:28:50 Pm_18 September 2019 ]
' -------------------------------------------------------------------

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
	
	MSGBOX PARAM
Next


Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run PARAM , ShowWindow, DontWaitUntilFinished





' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Wscript.Arguments.Item(0)


' --------------------------------------------------------------
' If Sent-to is Used Sometime Additional Quote Surround are User
' --------------------------------------------------------------
'XF0 = UCASE(Replace(XF0,"""",""))

' WSHShell.Run "CMD /K """ + "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT" + """ , ShowWindow, DontWaitUntilFinished

' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\BAT_03_PROCESS_KILLER.BAT""" , ShowWindow, DontWaitUntilFinished

