'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 28-GOODSYNC JOB WORK COUNT.VBS
'# __ 
'# __ VBS 28-GOODSYNC JOB WORK COUNT.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Tue 12-Jun-2018 13:55:53 ]
'# END       TIME [ Tue 12-Jun-2018 14:03:00 ] 8 MINUTES LEARNING _COUNTER NUMERIC WORKING
'# END FINAL TIME [ Tue 12-Jun-2018 19:00:00 ] 5 HOUR-ING RESTING AFTER THAT _SCRIPT LIST MADE
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
'# ------------------------------------------------------------------
' COUNT THE TOTAL JOBS WORKING OF GOODSYNC
' -------------------------------------------------------------------


Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

'--------------------------------------------------------
'AFTER THE EFFORT 
'VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
'IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
'HERE VBS DOES SINGLE 
'--------------------------------------------------------

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

'C:\Users\MATT 04\AppData\Roaming\GoodSync\jobs-groups-options.tic

FILEN_1="C:\Users\"&CUR_USER&"\AppData\Roaming\GoodSync\jobs-groups-options.tic"

Set File_1 = fso.OpenTextFile(FILEN_1, ForReading)
TT_EDITOR = File_1.ReadAll

SET fso = NOTHING

' COUNT THE JOBS WORKING
' ----------------------
X = 1
Do
    T1 = InStr(X, TT_EDITOR, "[n:")
    If T1 > 0 Then
        Counter_VAR = Counter_VAR + 1
    End If

    X = T1 + 2

	IF T1=0 THEN EXIT DO
    T2 = InStr(T1+2, TT_EDITOR, "[n:")
	IF T2=0 AND T1>0 THEN T2=LEN(TT_EDITOR)
    T4 = InStr(T1, TT_EDITOR, "k:1:@|")
	
	IF T4<T2 and T4>0 THEN 
		BLANK_JOBBER=FALSE
		Counter_DUD_JOB_SET = Counter_DUD_JOB_SET + 1
	ELSE
		BLANK_JOBBER=TRUE
	END IF

	T1 = InStr(X-2, TT_EDITOR, "[n:")

	T2 = InStr(T1+3, TT_EDITOR, ":")
	N3 = Mid(TT_EDITOR, T1+3, (T2 - T1)-3)
	N4 = Mid(TT_EDITOR, T2+1, CLNG(N3))
	N5 = LPad(CSTR(Counter_VAR),"0",4)+"--"

	IF BLANK_JOBBER=TRUE THEN REAL_OR_BLANK="0  "
	IF BLANK_JOBBER=FALSE THEN REAL_OR_BLANK="x--"
	
	' SELECTIVE FILTER
	' ---------------------------------------------------------------
	SET_GO=FALSE
	IF INSTR(N4,"HDD 0")>0 THEN SET_GO=TRUE
	IF INSTR(N4,"HDD 4")>0 THEN SET_GO=TRUE
	IF INSTR(N4,"HDD A")>0 THEN SET_GO=TRUE
	IF INSTR(N4,"HDD H")>0 THEN SET_GO=TRUE
	IF INSTR(N4,"HDD S")>0 THEN SET_GO=TRUE
	
	IF SET_GO=TRUE THEN
		I_X=I_X+N5+REAL_OR_BLANK+N4+VBCRLF
	ELSE
		I_X=I_X+N5+REAL_OR_BLANK+N4+VBCRLF
	END IF
	
Loop Until T1 = 0



i_TIME=""
i_TIME=i_TIME+LPad(Year(Now), "0", 4)+"-"
i_TIME=i_TIME+LPad(Month(Now), "0", 2)+"-"
i_TIME=i_TIME+LPad(Day(Now), "0", 2)+"-"
i_TIME=i_TIME+"-"
i_TIME=i_TIME+LPad(Hour(Now), "0", 2)+"-"
i_TIME=i_TIME+LPad(Minute(Now), "0", 2)+"-"
i_TIME=i_TIME+LPad(Second(Now), "0", 2)


I=""
I=I+"----------------------------------------------------------------" + VBCRLF 
I=I+"GoodSync Job Script" + VBCRLF 
I=I+"----------------------------------------------------------------" + VBCRLF 
I=I+i_TIME + VBCRLF 
I=I+"----------------------------------------------------------------" + VBCRLF 
I=I+"Total Job Count ="  + VBCRLF + CSTR(Counter_VAR) + VBCRLF
I=I+"Blank Job Set Count ="  + VBCRLF + CSTR(Counter_DUD_JOB_SET) + VBCRLF 
I=I+"Job Set Minus Blank Job Set Count ="  + VBCRLF + CSTR(Counter_VAR-Counter_DUD_JOB_SET)+ VBCRLF 
I=I+"----------------------------------------------------------------" + VBCRLF 
I=I+"KEY SYSTEM" + VBCRLF 
I=I+"X = Blank Jobber" + VBCRLF 
I=I+"0 = Worker Jobber" + VBCRLF 
I=I+"----------------------------------------------------------------" + VBCRLF 
I=I+"Script Text Generator" + VBCRLF 
I=I+MID(Wscript.ScriptFullName,3) + VBCRLF 
I=I+"----------------------------------------------------------------" + VBCRLF 

I=I+ VBCRLF 

I_X = I + I_X


'Wscript.Echo "Script name: " & Wscript.ScriptName

i_TIME_2=""
i_TIME_2=i_TIME_2+LPad(Year(Now), "0", 4)+"-"
i_TIME_2=i_TIME_2+LPad(Month(Now), "0", 2)+"-"
'i_TIME_2=i_TIME_2+UCASE(MID(MonthName(Month(Now)),1,3))+")-"
i_TIME_2=i_TIME_2+LPad(Day(Now), "0", 2)

SCRIPT_FOLDER_1=MID(Wscript.ScriptFullName,1,INSTRREV(Wscript.ScriptFullName,"\"))
SCRIPT_FOLDER_2=SCRIPT_FOLDER_1+"GOODSYNC JOB WORKER"
SCRIPT_NAME_1=SCRIPT_FOLDER_2+"\GOODSYNC JOB WORKER GS-DESKTOP "+ComputerName+"_"+CUR_USER+"_"+i_TIME_2+".TXT"
SCRIPT_NAME_2=SCRIPT_FOLDER_2+"\GOODSYNC JOB WORKER GS-DESKTOP "+ComputerName+"_"+CUR_USER+"_"+".TXT"

Set oFSO = CreateObject("Scripting.FileSystemObject")

If Not oFSO.FolderExists(SCRIPT_FOLDER_2) Then
	Set objFolder = oFSO.CreateFolder(SCRIPT_FOLDER_2)
End If

Set fso = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists(SCRIPT_NAME_1) Then 
	FSO.DeleteFile SCRIPT_NAME_1
END IF
If FSO.FileExists(SCRIPT_NAME_2) Then 
	FSO.DeleteFile SCRIPT_NAME_2
END IF

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFile = FSO.CreateTextFile(SCRIPT_NAME_1)
objFile.Write I_X
SET fso = NOTHING

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFile = FSO.CreateTextFile(SCRIPT_NAME_2)
objFile.Write I_X
SET fso = NOTHING

MSGBOX I_X

WScript.Quit(0)
'--------------------------------------------------------------------
'END STOP -----------------------------------------------------------
'--------------------------------------------------------------------

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function
