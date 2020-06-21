'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 52-GOODSYNC GOOGLE DRIVE INSYNC HIKVISION DELETE EMPTY FOLDER.VBS
'# __ 
'# __ VBS 52-GOODSYNC GOOGLE DRIVE INSYNC HIKVISION DELETE EMPTY FOLDER.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Sat 25-May-2019 13:08:22 ]
'# END FINAL TIME [ Sat 25-May-2019 13:44:18 ]
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' -------------------------------------------------------------------
'# START     TIME [ Sat 25-May-2019 13:08:22 ]
'# END FINAL TIME [ Sat 25-May-2019 13:44:18 ]
'#-------------------------------------------------------------------

' -------------------------------------------------------------------
' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF ARE
' SOURCE GIVEN
' -------------------------------------------------------------------
' vbscript - how to check vbs script in windows is running or not? - Stack Overflow
' https://stackoverflow.com/questions/7849699/how-to-check-vbs-script-in-windows-is-running-or-not
' -------------------------------------------------------------------
' [ Wednesday 13:08:00 Pm_20 March 2019 ]
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' NEW CODE INCORPORATE BOTH SET CODE NOW
' TEMPORARY RUNNER BY THIS SCRIPT LEARN ASSHOLE MODIFY GOODSYNC
' NEW SCRIPT MOVE FOLDER INTO CORRECT AREA ALSO AS WELL
' HERE DO THE DELETE EMPTY
' -------------------------------------------------------------------



Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
FILE_NAME="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 52-INSYNC HIKVISION COPY DUPLICATE FOLDER.VBS"
'--------------------------------------------------------------
WSHShell.Run """"+FILE_NAME+"""" + " " + ""+PARAM+"" , ShowWindow, DontWaitUntilFinished
'--------------------------------------------------------------

'WScript.Quit(0)
'--------------------------------------------------------------------
'END STOP -----------------------------------------------------------
'--------------------------------------------------------------------



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

Dim objFolder

FOLDER_STRING_1="D:\0 CLOUD\GD-INSYNC\rub.rim@gmail.com\snapshot ("
FOLDER_STRING_2="D:\0 CLOUD\GD-INSYNC\rub.rim@gmail.com\snapshot ("
FOLDER_STRING_3="C:\Users\MATT 04\Insync\rub.rim@gmail.com\Google Drive\snapshot ("

ON ERROR RESUME NEXT
FOR RV=1 TO 3
	FOR RO=2 TO 20
		XXC=CSTR(RO)
		IF RV=1 THEN FOLDER_STRING=FOLDER_STRING_1+XXC+")"
		IF RV=2 THEN FOLDER_STRING=FOLDER_STRING_2+XXC+")"
		IF RV=3 THEN FOLDER_STRING=FOLDER_STRING_3+XXC+")"

		IF FSO.FOLDERExists(FOLDER_STRING) THEN
			I1=-1
			I2=-1
			ERR.CLEAR 
			Set objFolder = FSO.GetFolder(FOLDER_STRING)
			I1=objFolder.SubFolders.Count
			I2=objFolder.Files.Count
			
			If I1=0 AND I2=0 AND ERR.NUMBER=0 THEN  
				MSGBOX "HJ"
				fso.DeleteFolder objFolder.Path
				IF FSO.FOLDERExists(FOLDER_STRING) THEN
					fso.DeleteFolder FOLDER_STRING
				END IF
			END IF
		END IF
	NEXT
NEXT



FILE_NAME="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 52-GOODSYNC GOOGLE DRIVE INSYNC HIKVISION DELETE EMPTY FOLDER_2.VBS"
'--------------------------------------------------------------
WSHShell.Run """"+FILE_NAME+"""" + " " + ""+PARAM+"" , ShowWindow, DontWaitUntilFinished
'--------------------------------------------------------------


WScript.Quit(0)

'--------------------------------------------------------------------
'END STOP -----------------------------------------------------------
'--------------------------------------------------------------------





Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

' MsgBox i_TIME
' MsgBox Format(Now, "YYYY-MMM-DD -- HH:MM:SS - HH:MM:SS-AM/PM DDD")
' 2019-Mar-18 -- 14:00:40 - 02:00:40-PM - Sun




' SOURCE CREDIT DELETE EMPTY
' ----
' Deleting Empty Folders – Random Musings of Jeremy Jameson
' https://blogs.msdn.microsoft.com/jjameson/2009/11/03/deleting-empty-folders/
' ----