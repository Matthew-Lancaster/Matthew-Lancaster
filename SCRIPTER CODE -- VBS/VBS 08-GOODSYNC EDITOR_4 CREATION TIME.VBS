Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

'MsgBox "Done Total " & CHR(&H64) & " Proecesses"


' ---------------------------------------------------------------------
' WORK TODAY
' ADD COD FOR SELECTIVE FILTER _ ONLY FILTER SOME JOB FOR ACTION TAKEN
' ---------------------------------------------------------------------
' Tue 27-Nov-2018 19:07:29
' Tue 27-Nov-2018 20:24:00
' ---------------------------------------------------------------------


'--------------------------------------------------------
'AFTER THE EFFORT 
'VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
'IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
'HERE VBS DOES SINGLE 
'--------------------------------------------------------

CUR_USER = CreateObject("WScript.Network").UserName

' C:\Users\MATT 04\AppData\Roaming\GoodSync\jobs-groups-options.tic

FILEN_1="C:\Users\"&CUR_USER&"\AppData\Roaming\GoodSync\jobs-groups-options.tic"
FILEN_2="C:\Users\"&CUR_USER&"\AppData\Roaming\GoodSync\jobs-groups-options - Copy.tic"

FSO.COPYFILE FILEN_1, FILEN_2

DIM TT_EDITOR
DIM Counter_VAR
DIM Counter_Change_VAR

Set File_1 = fso.OpenTextFile(FILEN_2, ForReading)
TT_EDITOR = File_1.ReadAll


' CALL USUAL_MAIN
CALL USUAL_MAIN_2

' CALL ROUTINE_WITH_ACTION_ON_SPECIFIC_JOB



MsgBox "Done Total " & Counter_VAR & " Proecesses"+VBCRLF+"Total Changes "& Counter_Change_VAR


Set File_2 = fso.OpenTextFile("C:\Users\"&CUR_USER&"\AppData\Roaming\GoodSync\jobs-groups-options.tic", ForWriting)
File_2.Write TT_EDITOR


WScript.Quit(0)



SUB ROUTINE_WITH_ACTION_ON_SPECIFIC_JOB:

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
	
	I_X=I_X+N5+REAL_OR_BLANK+N4+VBCRLF
	
	' SPECIFY THE JOB TO FILTER THE SELECTION
	' ---------------------------------------------------------------
	SET_GO=FALSE
	IF INSTR(N4,"HDD 0")>0 THEN SET_GO=TRUE
	IF INSTR(N4,"HDD 4")>0 THEN SET_GO=TRUE
	IF INSTR(N4,"HDD A")>0 THEN SET_GO=TRUE
	IF INSTR(N4,"HDD H")>0 THEN SET_GO=TRUE
	IF INSTR(N4,"HDD S")>0 THEN SET_GO=TRUE
	
	IF SET_GO=TRUE THEN
		I_X=I_X+N5+REAL_OR_BLANK+N4+VBCRLF

		T5 = InStr(X-2, TT_EDITOR, VBLF)
		T8 = InStr(X-2, TT_EDITOR, "|==")
		T9 = InStr(X-2, TT_EDITOR, "]"+CHR(&HA5)+"=")

		IF T8<T5 THEN
			' FILES CHANGED MORE THAN A PERCENT AND THEN PREVENT SYNCE 
			' IF NOT = TO 90 THEN SET TO = 90
			' -------------------------------------------------------
			T1 = InStr(X, TT_EDITOR, "|==")
			If T1 > 0 Then
				
				T2 = InStr(T1+1, TT_EDITOR, "|")
				N3 = Mid(TT_EDITOR, T1, (T2 - T1))+"|"
				N4 = Replace(N3, "|==","")
				N4 = Replace(N4, "|","")
				' MSGBOX N3
				
				IF LEN(N3)<=8 THEN
					IF N4<>"80" THEN 
						TT_EDITOR_2=Mid(TT_EDITOR, X, T5-X)
						TT_EDITOR_3=Mid(TT_EDITOR, X, T5-X)
						TT_EDITOR_3 = Replace(TT_EDITOR_3, N3, "|=="+"80"&"|",1,1)
						TT_EDITOR = Replace(TT_EDITOR, TT_EDITOR_2,TT_EDITOR_3,1,1)
						Counter_Change_VAR = Counter_Change_VAR + 1
					END IF
				END IF
			END IF

			IF T9<T5 THEN
				' CHECK BOX FOR _ IF THIS MANY FILE CHANGED % __ FLAG THE BIT
				'------------------------------------------------------------
				T1 = InStr(X, TT_EDITOR, "]"&CHR(&HA5)&"=1")
				If T1 > 0 Then
					TT_EDITOR_2=Mid(TT_EDITOR, X, T5-X)
					TT_EDITOR_3=Mid(TT_EDITOR, X, T5-X)
					TT_EDITOR_3 = Replace(TT_EDITOR_3, "]"&CHR(&HA5)&"=1", "]"&CHR(&HA5)&"=0",1,1)
					TT_EDITOR = Replace(TT_EDITOR, TT_EDITOR_2,TT_EDITOR_3,1,1)
					Counter_Change_VAR = Counter_Change_VAR + 1
				End If
			End If
		END IF
	END IF
	
Loop Until T1 = 0

END SUB


SUB USUAL_MAIN_2

	' CHANGE SCHEDULE HOUR FROM 2 TO 4
	' --------------------------------
	X = 1
	Do
		T1 = InStr(X, TT_EDITOR, "]d=4")
		If T1 > 0 Then
			TT_EDITOR = Replace(TT_EDITOR, "]d=4", "]d=5",1,1)
			Counter_VAR = Counter_VAR + 1
		End If
		X = T1 + 2
	Loop Until T1 = 0

	Counter_Change_VAR=Counter_VAR
	
END SUB



SUB USUAL_MAIN:

	' COPY ATTRIBUTES YES
	' -------------------
	X = 1
	Do
		T1 = InStr(X, TT_EDITOR, "|.=0")
		If T1 > 0 Then
			TT_EDITOR = Replace(TT_EDITOR, "|.=0", "|.=1",1,1)
			Counter_VAR = Counter_VAR + 1
		End If
		X = T1 + 2
	Loop Until T1 = 0


	' PARALLEL COPY FILE 10
	' IF NOT = TO 0 THEN SET TO = 10
	'---------------------
	X = 1
	Do
		T1 = InStr(X, TT_EDITOR, "|"&CHR(&HCD))
		If T1 > 0 Then
			
			T2 = InStr(T1, TT_EDITOR, "|"&CHR(&HCE))
			N3 = Mid(TT_EDITOR, T1, (T2 - T1))&"|"&CHR(&HCE)
			N4 = Replace(N3, "|"&CHR(&HCD)&"=","")
			N4 = Replace(N4, "|"&CHR(&HCE),"")

			IF LEN(N3)<=8 THEN 
				IF N4<>"0" AND N4<>"10" AND 1=2 THEN 
					TT_EDITOR = Replace(TT_EDITOR, N3, "|"&CHR(&HCD)&"=10"&"|"&CHR(&HCE),1,1)
					Counter_VAR = Counter_VAR + 1
					'MSGBOX N4
					'WScript.Quit(0)
				END IF
			END IF
		End If
		X = T1 + 2
	Loop Until T1 = 0

	' PARALLEL COPY FILE 10 __ FLAG THE BIT
	'-------------------------------------
	X = 1
	Do
		T1 = InStr(X, TT_EDITOR, "|"&CHR(&HD2)&"=0")
		If T1 > 0 Then
			TT_EDITOR = Replace(TT_EDITOR, "|"&CHR(&HD2)&"=0", "|"&CHR(&HD2)&"=1",1,1)
			Counter_VAR = Counter_VAR + 1
		End If
		X = T1 + 2
	Loop Until T1 = 0

	' WAIT FOR LOCKS TO CLEAR 10
	' SET TO = 10
	'---------------------
	X = 1
	Do
		T1 = InStr(X, TT_EDITOR, "|"&CHR(&HCE))
		If T1 > 0 Then
			
			T2 = InStr(T1, TT_EDITOR, "|"&CHR(&HD0))
			N3 = Mid(TT_EDITOR, T1, (T2 - T1))&"|"&CHR(&HD0)
			N4 = Replace(N3, "|"&CHR(&HCE)&"=","")
			N4 = Replace(N4, "|"&CHR(&HD0),"")

			IF LEN(N3)<=8 THEN 
				IF N4<>"10" THEN 
					TT_EDITOR = Replace(TT_EDITOR, N3, "|"&CHR(&HCE)&"=10"&"|"&CHR(&HD0),1,1)
					Counter_VAR = Counter_VAR + 1
				END IF
			END IF
		End If
		X = T1 + 2
	Loop Until T1 = 0

	' WAIT FOR LOCKS TO CLEAR __ FLAG THE BIT
	'----------------------------------------
	X = 1
	SET_BIT_1="1"
	SET_BIT_2="0"
	Do
		T1 = InStr(X, TT_EDITOR, "|#="+SET_BIT_1)
		If T1 > 0 Then
			TT_EDITOR = Replace(TT_EDITOR, "|#="+SET_BIT_1, "|#="+SET_BIT_2,1,1)
			Counter_VAR = Counter_VAR + 1
		End If
		X = T1 + 2
	Loop Until T1 = 0



	' AUTO CLEAR TREE AFTER YES
	' -------------------------
	X = 1
	Do
		T1 = InStr(X, TT_EDITOR, "|i=0")
		If T1 > 0 Then
			TT_EDITOR = Replace(TT_EDITOR, "|i=0", "|i=1",1,1)
			Counter_VAR = Counter_VAR + 1
		End If
		X = T1 + 2
	Loop Until T1 = 0

	' TURN OFF NOTIFICATIONS
	' ----------------------
	X = 1
	Do
		T1 = InStr(X, TT_EDITOR, "|j=1")
		If T1 > 0 Then
			TT_EDITOR = Replace(TT_EDITOR, "|j=1", "|j=0",1,1)
			Counter_VAR = Counter_VAR + 1
		End If
		X = T1 + 2
	Loop Until T1 = 0


	' CHANGE SCHEDULE HOUR FROM 2 TO 4
	' --------------------------------
	X = 1
	Do
		T1 = InStr(X, TT_EDITOR, "]d=2")
		If T1 > 0 Then
			TT_EDITOR = Replace(TT_EDITOR, "]d=2", "]d=4",1,1)
			Counter_VAR = Counter_VAR + 1
		End If
		X = T1 + 2
	Loop Until T1 = 0

END SUB


'--------------------------------------------------------------------

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function