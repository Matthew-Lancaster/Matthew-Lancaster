'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 42-PSTART SYNC NETWORK.VBS
'# __ 
'# __ VBS 42-PSTART SYNC NETWORK.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Fri 22-Mar-2019 10:31:08 ]
'# END FINAL TIME [ Fri 22-Mar-2019 11:50:00 ]
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
' FOUND SNIPPET ON MICROSOFT WEBSITE WHILE LOOKUP ANSWER DURING CODER
' LOOKED INTERESTING
' FIREWALL PERMISSION ONLY ALLOW LOOK AT SELF MACHINE OUT OF 7 INCLUDE ME
' WORTH A SORTER LATER
'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' -------------------------------------------------------------------
'# START     TIME [ Fri 22-Mar-2019 10:31:08 ]
'# END FINAL TIME [ Fri 22-Mar-2019 11:50:00 ]
'#-------------------------------------------------------------------

' -------------------------------------------------------------------
' SOURCE GIVEN
' -------------------------------------------------------------------
' vbscript - how to check vbs script in windows is running or not? - Stack Overflow
' https://stackoverflow.com/questions/7849699/how-to-check-vbs-script-in-windows-is-running-or-not
' -------------------------------------------------------------------
' [ Wednesday 13:08:00 Pm_20 March 2019 ]
' -------------------------------------------------------------------

' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF IS WORKER
' -------------------------------------------------------------------
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

i1 = 0 ' ANY PROGRAM WSCRIPT 
i2 = 0 ' MY  PROGRAM WSCRIPT 
FOR EACH objProcess IN colProcesses
    IF NOT (IsNull(objProcess.CommandLine )) THEN
		strScriptName = Trim(Right(objProcess.CommandLine,Len(objProcess.CommandLine) - InstrRev(objProcess.CommandLine,"\")))
		strScriptName = Left(strScriptName, Len(strScriptName) - 1)                             
		IF strScriptName = Wscript.ScriptName THEN
			i2=i2+1
		else
			i1=i1+1
		END IF
		IF i1 > 0 THEN 
				' Wscript.Quit
		END IF  
		IF i2 > 1 THEN 
				Wscript.Quit
		END IF  
    END IF  
NEXT

Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

'--------------------------------------------------------
' AFTER THE EFFORT 
' VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
' IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
' HERE VBS DOES SINGLE 
'--------------------------------------------------------

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

' Dim Array_FileName()
' Dim Array_NETPATH_01()
' Dim Array_NETPATH_02()
' Dim Array_PSTART_COUNTER()
' ReDim Preserve Array_FileName(50)
' ReDim Preserve Array_NETPATH_01(50)
' ReDim Preserve Array_NETPATH_02(50)
' ReDim Preserve Array_PSTART_COUNTER(50)

' CALL READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY

' FOR RI=1 TO UBOUND(Array_FileName)
	' Array_FileName(RI)=Array_FileName(RI)+"_01_C_DRIVE\PStart\PStart.xml"
' NEXT


' ----
' Can I Read an Entire Text File Rather Than Reading It Line-by-Line? | Scripting
' https://devblogs.microsoft.com/scripting/can-i-read-an-entire-text-file-rather-than-reading-it-line-by-line/
' ----

Set objFSO = CreateObject("Scripting.FileSystemObject")

FILEN_1="C:\NETWORK_COMPUTER_NAME.txt"
Set objTextFile = objFSO.OpenTextFile (FILEN_1, ForReading)
strText = objTextFile.ReadAll
objTextFile.Close
arrComputers = Split(strText, vbCrLf)

For RI= 0 TO UBOUND(arrComputers)
	IF RI=0                           THEN arrComputers(RI)=""
	IF RI=1                           THEN arrComputers(RI)=""
	IF RI=2                           THEN arrComputers(RI)=""
	IF RI=3                           THEN arrComputers(RI)=""
	IF RI=4                           THEN arrComputers(RI)=""
	IF arrComputers(RI)="BTHUB"       THEN arrComputers(RI)=""
	IF arrComputers(RI)="NAS-QNAP-ML" THEN arrComputers(RI)=""
NEXT

' ALL COMPUTER BUT MY OWN IS NOT ACCESSIBLE -- winmgmts

For Each strComputer in arrComputers
	IF strComputer<>"" THEN 
		MSGBOX strComputer,vbMsgBoxSetForeground,Wscript.ScriptName 
		Set objWMIService = GetObject ("winmgmts:\\" & strComputer & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery ("Select * From Win32_OperatingSystem")
		For Each objItem in ColItems

			' Wscript.Echo strComputer & ": " & objItem.Caption
			MSGBOX strComputer & ": " & objItem.Caption,vbMsgBoxSetForeground,Wscript.ScriptName 
		Next
	END IF
Next



WScript.Quit(0)

'--------------------------------------------------------------------
'END STOP -----------------------------------------------------------
'--------------------------------------------------------------------

Sub READ_PSTART_FOR_BIGGER()
    
    Dim SET_GO
	FOR RI=1 TO UBOUND(Array_FileName)
		'\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE\PStart\PStart.xml
		FILEN_1=Array_FileName(RI)
		Set File_1 = fso.OpenTextFile(FILEN_1, ForReading)
		Do Until File_1.AtEndOfStream
			NET_PATH_1 = File_1.ReadLine+VBCRLF
			
			SET_GO = FALSE
			If InStr(NET_PATH_1, "<execution>") > 0 Then SET_GO = TRUE
			If SET_GO = True Then
				NET_PATH_2 = File_1.ReadLine+VBCRLF
				NET_PATH_3 = File_1.ReadLine+VBCRLF
				NET_PATH_4 = File_1.ReadLine+VBCRLF
				If InStr(NET_PATH_2, "<count>") > 0 Then
					Array_PSTART_COUNTER(RI)=NET_PATH_1+NET_PATH_2+NET_PATH_3+NET_PATH_4
				END IF
				EXIT DO
			End If
		Loop
		File_1.Close
	NEXT
End Sub


Sub READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY()
    
    Dim SET_GO
    
    ArrayCount = 0

	FILEN_1="C:\NETWORK_COMPUTER_NAME.txt"
	Set File_1 = fso.OpenTextFile(FILEN_1, ForReading)
	Do Until File_1.AtEndOfStream
		NET_PATH = File_1.ReadLine
        
        SET_GO = True
        If InStr(NET_PATH, "BTHUB") > 0 Then SET_GO = False
        If InStr(NET_PATH, "NAS-QNAP-ML") > 0 Then SET_GO = False
        If SET_GO = True Then
            ArrayCount = ArrayCount + 1
            Array_NETPATH_01(ArrayCount) = NET_PATH
            Array_NETPATH_02(ArrayCount) = Replace(NET_PATH, "-", "_")
            ELEMENT1 = "\\"
            ELEMENT2 = Array_NETPATH_01(ArrayCount)
            ELEMENT3 = "\"
            ELEMENT4 = Array_NETPATH_02(ArrayCount)
            ELEMENT5 = FileName_2
            ELEMENT7 = "_" + NET_PATH

            Array_FileName(ArrayCount) = ELEMENT1 + ELEMENT2 + ELEMENT3 + ELEMENT4 + ELEMENT5' + ELEMENT7
        End If
	Loop
	File_1.Close
    
	ReDim Preserve Array_FileName(ArrayCount)
	ReDim Preserve Array_NETPATH_01(ArrayCount)
	ReDim Preserve Array_NETPATH_02(ArrayCount)
	ReDim Preserve Array_PSTART_COUNTER(ArrayCount)

End Sub




Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

' MsgBox i_TIME
' MsgBox Format(Now, "YYYY-MMM-DD -- HH:MM:SS - HH:MM:SS-AM/PM DDD")
' 2019-Mar-18 -- 14:00:40 - 02:00:40-PM - Sun
