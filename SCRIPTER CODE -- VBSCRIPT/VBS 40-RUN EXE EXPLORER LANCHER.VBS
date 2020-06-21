Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

' -------------------------------------------------------------------
' WORK TODAY
' -------------------------------------------------------------------
' Wed 18-Sep-2019 21:21:07
' Wed 18-Sep-2019 22:55:17
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' EXAMPLE USER OF VISUAL BASIC 6
' FROM THIS CODER -- 
' --
' D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer Loader\Shell Explorer Loader.exe
' -------------------------------------------------------------------
' Private Sub MNU_GC_PROFILE_Click(Index As Integer)
    ' Me.WindowState = vbMinimized
    ' XX = MNU_GC_PROFILE(Index).Caption
    ' SHELL_LINE_1 = "C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 40-RUN EXE EXPLORER LANCHER.VBS"
    ' SHELL_LINE_2 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ' SHELL_LINE_3 = "--profile-directory=""Profile " + Trim(Str(Index + 1)) + """"
    ' ' "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 4"
    ' SHELL_LINE_2 = Replace(SHELL_LINE_2, " ", "*")
    ' SHELL_LINE_3 = Replace(SHELL_LINE_3, " ", "*")
    ' Dim WSHShell
    ' Set WSHShell = CreateObject("WScript.Shell")
        ' l = """" + SHELL_LINE_1 + """" + " " + SHELL_LINE_2 + " " + SHELL_LINE_3
        ' Debug.Print l
        ' WSHShell.Run l, ShowWindow_2, DontWaitUntilFinished
    ' Set WSHShell = Nothing
' End
' End Sub
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

' -------------------------------------------------------------------
' THIS FILE HERE
' -------------------------------------------------------------------
' C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 40-RUN EXE EXPLORER LANCHER.VBS
' IS LINKER ON
' D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer Loader\VBS 40-RUN EXE EXPLORER LANCHER.VBS
' -------------------------------------------------------------------
' NEW BEST WAY TO PASS ARGUMENT ON VBSCRIPT FROM ANY APP WANT COMMAND LINE TO ANOTHER
' -------------------------------------------------------------------
' TPYE OF THING SEEN TO GET ARGS INTO CODE FROM OTHER APP AROUND 2003
' AND NEVER SUNK IN TO PRODUCE WHEN COMPILED IN EXE APP EXAMPLE NEVER EXPLAIN
' AS I TALK
' THE NET ONLY SHOW STANDARD WAY TO GET ARGUMENT PARAMETER 
' THE CODE HERE ORIGINAL DESIGN FOR USE BY HERE
' TONS OF REQUEST AND NOBODY COME BACK TO 
' FINAL RESULT ABLE MAKE THING MUCH MORE PRETTY
' NONE MORE FLOAT ABOUT WITH DOG EARS QUOTATION MARK
' AND SOME QUOTATION NEST BIRDY ARE NEAR IMPOSABLE 
' GIVE UP LEVEL LITERALLY ARE IMPOSABLE
' -------------------------------------------------------------------
' D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer Loader\Shell Explorer Loader.exe
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Fri 20-Sep-2019 17:20:29
' Fri 20-Sep-2019 17:53:00 	-- TIME TO PATCH THIS END BIT IN NOTES AND COPY OVER 
' 							-- LIKE SNAIL PACE
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' 0 = ARG FROM OTHER APP WILL GET HERE
' -------------------------------------------------------------------
' 1ST ARG WILL BE APP
' 2ND ARG WILL BE PATH SOMETHING
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
	
	' MSGBOX PARAM
Next

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run PARAM , ShowWindow, DontWaitUntilFinished

' -------------------------------------------------------------------
' -------------------------------------------------------------------
Wscript.Arguments.Item(0)


' --------------------------------------------------------------
' If Sent-to is Used Sometime Additional Quote Surround are User
' --------------------------------------------------------------
'XF0 = UCASE(Replace(XF0,"""",""))

