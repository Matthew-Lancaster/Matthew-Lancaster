' ----------------------------------------------------------------
' __ MATTHEW LANCASTER 
' __ MATT.LAN@BTINTERNET.COM
' ----------------------------------------------------------------
' __ FROM ---- Wed 16-May-2018 02:10:55
' __ TO   ---- Wed 16-May-2018 03:32:00 __ 1 HOUR 22 MINUTE  _ ONTO NEXT VBSCRIPT FROM BAT
' __ TO   ---- Wed 16-May-2018 04:38:00 __ 2 HOUR 38 MINUTE  _ FORMATING DONE
' __ TO   ---- Wed 16-May-2018 05:05:00 __ 3 HOUR            _ LAST BIT 
' ----------------------------------------------------------------


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

DIM GETUSERNAME
DIM GETCOMPUTERNAME
DIM COMPUTERNAME
DIM WSHSHELL
SET WSHSHELL = CreateObject("WScript.Shell")
GETUSERNAME = CreateObject("WScript.Network").USERNAME
GETCOMPUTERNAME = CreateObject("WScript.Network").COMPUTERNAME
COMPUTERNAME = CreateObject("WScript.Network").COMPUTERNAME
' ------------------------------------------------------------------
' ANOTHER EXAMPLE
' ------------------------------------------------------------------
' COMPUTERNAME = WSHSHELL.ExpandEnvironmentStrings("%computername%")
' ------------------------------------------------------------------


Set args = Wscript.Arguments
For Each arg In args
	XF0 = XF0 + arg
Next

'If Sent-to is Used Sometime Additional Quote Surround are User
'--------------------------------------------------------------
XF0 = Replace(XF0,"""","")

'XF0="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 33-INSTALLED PROGRAMS AND LOCATION\INSTALLED_PROGRAMS_AND_LOCATION__2018_05_16__04_51_20.TXT"

FILEN_1=XF0

'WScript.Quit(0)

'MSGBOX FILEN_1

Set File_1 = fso.OpenTextFile(FILEN_1, ForReading,0,0)
TT_EDITOR = File_1.ReadAll

X = 1
Do
    T1 = InStr(X, TT_EDITOR, ":\")
    If T1 > 0 Then
		VAR_1=MID(TT_EDITOR,T1-1,3)
		VAR_2=MID(TT_EDITOR,T1-2,1)
		
		IF VAR_2=" " THEN 
			VAR_1=MID(TT_EDITOR,1,T1-2)
			VAR_2=MID(TT_EDITOR,T1-1)
			VAR_3=VBCRLF
			TT_EDITOR=VAR_1+VAR_3+VAR_2
			Counter_VAR = Counter_VAR + 1
		END IF
    End If
    X = T1 + 4
Loop Until T1 = 0


X = 1
Do
    T1 = InStr(X, TT_EDITOR, ":\")
	IF T1>0 THEN 
		T2 = InStr(T1, TT_EDITOR, VBCRLF)
		VAR_4=MID(TT_EDITOR,T2-1,1)
    END IF
	If T1 > 0 AND VAR_4 =" " Then
		VAR_1=MID(TT_EDITOR,1,T2)
		VAR_2=MID(TT_EDITOR,T2+1)
		VAR_3=STRING(70,"-")+VBCRLF
		TT_EDITOR=VAR_1+VAR_3+VAR_2
		Counter_VAR = Counter_VAR + 1
	END IF
    X = T1 + 4
Loop Until T1 = 0

Set File_2 = fso.OpenTextFile(FILEN_1, ForWriting)
File_2.Write TT_EDITOR

WScript.Quit(0)
