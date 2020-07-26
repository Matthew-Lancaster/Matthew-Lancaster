Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true


' FILE PATH SEARCH STRING
' "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT"
' $(AllLocalDrives)\

' $(AllLocalDrives)\;C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT




Set objArgs = Wscript.Arguments
For Each strArg In objArgs

	' PARAM = PARAM + """"+REPLACE(strArg,"*"," ")+"""" +" "
	PARAM = PARAM + """"+strArg+"""" +" "
	
Next

'MSGBOX PARAM

IF PARAM="" THEN 
	Wscript.QUIT
END IF


' "VB VB-NT 01 EXE 4G 1X _001_ EXE" "--RESULT--" "" "--ERROR--" "1" "--CONFLICT--" "0" "--LOGPATH--" "D:\GOODSYNC\Profile\_mirrors_\gstps-matt-lan-btinternet-com@4-asus-gl522vw.matt-lan-btinternet-com.goodsync-D-VB6-VB-NT\2020-0726-082045-4-ASUS-GL522VW-VB VB-NT 01 EXE 4G 1X _001_ EXE.log" 

XX=INSTR(PARAM,"--LOGPATH--")
IF XX>0 THEN 
	LOGPATH=MID(PARAM,XX+LEN("--LOGPATH--")+2)
	LOGPATH=REPLACE(LOGPATH,"""","")
END IF






DIM NETNAME(7)

i = i + 1: NETNAME(i) = "1-ASUS-X5DIJ"
i = i + 1: NETNAME(i) = "2-ASUS-EEE"
i = i + 1: NETNAME(i) = "3-LINDA-PC"
i = i + 1: NETNAME(i) = "4-ASUS-GL522VW"
i = i + 1: NETNAME(i) = "5-ASUS-P2520LA"
i = i + 1: NETNAME(i) = "7-ASUS-GL522VW"
i = i + 1: NETNAME(i) = "8-MSI-GP62M-7RD"



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

SET FILE_1 = FSO.OpenTextFile(LOGPATH)

Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run """C:\Program Files (x86)\Notepad++\notepad++.exe""" + " " + """"+LOGPATH+""""
Set WSHShell = Nothing

Do Until FILE_1.AtEndOfStream
	READ_LINE=FILE_1.ReadLine
	
	IF MM=0 THEN 
		I="Cannot delete file: Access is denied."
		MM=INSTR(READ_LINE,I)
	END IF
	IF MM=0 THEN 
		I="The process cannot access the file because it is being used by another process."
		MM=INSTR(READ_LINE,I)
	END IF
	I="Copy Over '"
	M1=INSTR(READ_LINE,I)
	IF M1>0 THEN 
		FOR R=1 TO 7
			MSGBOX NETNAME(R)+VBCR+READ_LINE
			IF INSTR(UCASE(NETNAME(R)),UCASE(READ_LINE))>0 THEN 
				COMPUTERNAME_1=UCASE(NETNAME(R))
			END IF
		NEXT 
		FOR R=1 TO 7
			IF INSTR(UCASE(NETNAME(R)),UCASE(READ_LINE))>0 AND NETNAME(R)<>COMPUTERNAME_1 THEN 
				COMPUTERNAME_2=UCASE(NETNAME(R))
			END IF
		NEXT 
	END IF


	IF MM>0 THEN 
		MSGBOX READ_LINE+VBCR+COMPUTERNAME_1+VBCR+COMPUTERNAME_2
	END IF
Loop



Wscript.QUIT




X = 1
Do
    T1 = InStr(X, TT_EDITOR, "|M:0")
    
    If T1 > 0 Then
        
        T2 = InStr(T1, TT_EDITOR, "|i=")
        N3 = Mid(TT_EDITOR, T1, (T2 - T1) + Len("|i="))
        
        REASSEMBLE_VAR = "|M:0:|U=0|B:0:|J:0:|i="
        TT_EDITOR = Replace(TT_EDITOR, N3, REASSEMBLE_VAR)
        
        Counter_VAR = Counter_VAR + 1
    End If
    
    X = T1 + 4
Loop Until T1 = 0




' D:\GOODSYNC\Profile\_mirrors_\gstps-matt-lan-btinternet-com@4-asus-gl522vw.matt-lan-btinternet-com.goodsync-D-VB6-VB-NT\2020-0726-082045-4-ASUS-GL522VW-VB VB-NT 01 EXE 4G 1X _001_ EXE.log


' Cannot delete file: Access is denied




Wscript.QUIT




Public Function isValidFileName( ByVal strPath, ByVal strFileName )

	On Error Resume Next
	
	' bIsValid = False
	' IF EXTNAME=".LOG" AND ISNUMERIC(F4) THEN bIsValid = TRUE
	' IF EXTNAME=".LOG" AND F8="GOODSYNC" THEN bIsValid = TRUE
	' IF EXTNAME=".LOG" AND F8="SENDSMTP" THEN bIsValid = TRUE

	bIsValid = TRUE
	
	' IF strFileName="FILE_LOCATOR_DUMMY_FILE_TO_RUN_SCRIPT_DO_OTHER.txt" THEN
		' bIsValid = FALSE

		' ' -----------------------------------------------------------
		' ' RUN FROM BAT 
		' ' -----------------------------------------------------------
		' ' XF0="C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.BAT"
		' ' Dim WSHShell
		' ' Set WSHShell = CreateObject("WScript.Shell")
			' ' WSHShell.Run """"+XF0+"""" , 1, TRUE
		' ' EXIT FUNCTION
		
		' ' -----------------------------------------------------------
		' ' RUN DIRECT POWERSHELL
		' ' -----------------------------------------------------------
		' FILE_NAME="""C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1"""
		' Dim WSHShell
		' Set WSHShell = CreateObject("WScript.Shell")
		' WSHShell.run("powershell -file "+FILE_NAME), 0, TRUE
		' EXIT FUNCTION
	' END IF
	
	
	
	IF INSTR(strPath,"_saved_")>0 THEN    bIsValid = FALSE
	IF INSTR(strPath,"_history_")>0 THEN  bIsValid = FALSE
	IF INSTR(strPath,"_gsdata_")>0 THEN   bIsValid = FALSE

	' IF
		' EXTNAME=UCASE(MID(strFileName,INSTRREV(strFileName,".")))
		' ' CHECK FOR "._GS" ".GSS" ".GSL" ".LOG"
	' END IF
'	IF INSTR(strPath,"\TEMP\")>0 THEN 
		DIM COMPNAME(7)
		COMPNAME(1)="1-ASUS-X5DIJ"
		COMPNAME(2)="2-ASUS-EEE"
		COMPNAME(3)="3-LINDA-PC"
		COMPNAME(4)="4-ASUS-GL522VW"
		COMPNAME(5)="5-ASUS-P2520LA"
		COMPNAME(6)="7-ASUS-GL522VW"
		COMPNAME(7)="8-MSI-GP62M-7RD"
		FOR R=1 TO 7
			IF INSTR(strPath + strFileName,COMPNAME(R))>0 THEN   
				bIsValid = FALSE
				EXIT FOR
			END IF
		NEXT 
		FOR R=1 TO 7
			IF INSTR(strPath + strFileName,REPLACE(COMPNAME(R),"-","_"))>0 THEN   
				bIsValid = FALSE
				EXIT FOR
			END IF
		NEXT 
		IF INSTR(strPath + strFileName,"#NFS_EX")>0 THEN      bIsValid = FALSE
		IF INSTR(strFileName,"VBSCRIPT")>0 THEN    
		IF INSTR(strFileName,".TXT")>0 THEN   
			bIsValid = FALSE
		END IF
		END IF
		
		
		
		' EXTNAME=UCASE(MID(strFileName,INSTRREV(strFileName,".")))
		' CHECK FOR ".EXE" -- "insync-upd"
'	END IF
	IF bIsValid = True THEN 
		IF INSTR(strFileName,".gdcache")>0    THEN bIsValid = False
		IF INSTR(strFileName,"insync-upd")>0  THEN bIsValid = False
		IF EXTNAME="._GS" THEN                     bIsValid = False
		IF EXTNAME=".GSS" THEN                     bIsValid = False
		IF EXTNAME=".GSL" THEN                     bIsValid = False
		IF EXTNAME=".LOG" THEN                     bIsValid = False
		'IF EXTNAME=".EXE" THEN                     bIsValid = False
	END IF
	
	isValidFileName = bIsValid
	' isValidFileName = NOT bIsValid -- INVERT THEM

End Function


PUBLIC FUNCTION TEST_ROUTINE_NOT_RUN_TWO_OF_SOMETHING_IS_COUNTER_HERE_MIGHT_WORK
	bIsValid = TRUE
	' C:\0 00 LINK SET QUICKER MOVER\# 0\FILE_LOCATOR_DUMMY_FILE_TO_RUN_SCRIPT_DO_OTHER\
	
	' IF SearchParms.LookIn="C:\0 00 LINK SET QUICKER MOVER\# 0\;$(AllLocalDrives)\;=C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT" THEN
	
	IF INSTR(SearchParms.LookIn,"C:\0 00 LINK SET QUICKER MOVER\# 0\")>0  THEN
		bIsValid = FALSE
		isValidFileName = bIsValid
		EXIT FUNCTION
	END IF

	isValidFileName = bIsValid
	EXIT FUNCTION
	
	' strCustom = 
	
END FUNCTION
