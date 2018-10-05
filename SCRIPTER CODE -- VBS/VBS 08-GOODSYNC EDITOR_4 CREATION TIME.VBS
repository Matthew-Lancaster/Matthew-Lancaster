Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

'MsgBox "Done Total " & CHR(&H64) & " Proecesses"


'--------------------------------------------------------
'AFTER THE EFFORT 
'VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
'IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
'HERE VBS DOES SINGLE 
'--------------------------------------------------------

CUR_USER = CreateObject("WScript.Network").UserName

FILEN_1="C:\Users\"&CUR_USER&"\AppData\Roaming\GoodSync\jobs-groups-options.tic"
FILEN_2="C:\Users\"&CUR_USER&"\AppData\Roaming\GoodSync\jobs-groups-options - Copy.tic"

FSO.COPYFILE FILEN_1, FILEN_2

Set File_1 = fso.OpenTextFile(FILEN_2, ForReading)
TT_EDITOR = File_1.ReadAll

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
			IF N4<>"0" AND N4<>"10" THEN 
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

MsgBox "Done Total " & Counter_VAR & " Proecesses"

'WScript.Quit(0)

Set File_2 = fso.OpenTextFile("C:\Users\"&CUR_USER&"\AppData\Roaming\GoodSync\jobs-groups-options.tic", ForWriting)

File_2.Write TT_EDITOR

