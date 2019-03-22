
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

Dim Array_FileName()
Dim Array_NETPATH_01()
Dim Array_NETPATH_02()
ReDim Preserve Array_FileName(50)
ReDim Preserve Array_NETPATH_01(50)
ReDim Preserve Array_NETPATH_02(50)

CALL READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY


FOR RI=1 TO UBOUND(Array_FileName)
	Array_FileName(RI)=Array_FileName(RI)+"_01_C_DRIVE\PStart\PStart.xml"
NEXT

FOR RI=1 TO UBOUND(Array_FileName)
	MSGBOX Array_FileName(RI)
	EXIT FOR
NEXT

CALL READ_PSTART_FOR_BIGGER

' Wscript.Echo i_TIME

WScript.Quit(0)

'--------------------------------------------------------------------
'END STOP -----------------------------------------------------------
'--------------------------------------------------------------------

Sub READ_PSTART_FOR_BIGGER()
    
    Dim SET_GO
    
	FOR RI=1 TO UBOUND(Array_FileName)

	'\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_1-ASUS-X5DIJ_01_C_DRIVE\PStart\PStart.xml
	'\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE\PStart\PStart.xml
	
		FILEN_1=Array_FileName(RI)
		MSGBOX Array_FileName(RI)
		Set File_1 = fso.OpenTextFile(FILEN_1, ForReading)
		Do Until File_1.AtEndOfStream
			NET_PATH_1 = File_1.ReadLine
			
			SET_GO = FALSE
			If InStr(NET_PATH_1, "<execution>") > 0 Then SET_GO = TRUE
			If SET_GO = True Then
				NET_PATH_2 = File_1.ReadLine
				NET_PATH_3 = File_1.ReadLine
				NET_PATH_4 = File_1.ReadLine

				MSGBOX NET_PATH_2
				EXIT DO
				
			End If
		Loop
		File_1.Close

	NEXT
		
End Sub




Sub READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY()
    
    ReDim Preserve Array_FileName(50)
    ReDim Preserve Array_NETPATH_01(50)
    ReDim Preserve Array_NETPATH_02(50)
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

End Sub





Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function



' MsgBox i_TIME
' MsgBox Format(Now, "YYYY-MMM-DD -- HH:MM:SS - HH:MM:SS-AM/PM DDD")
' 2019-Mar-18 -- 14:00:40 - 02:00:40-PM - Sun
