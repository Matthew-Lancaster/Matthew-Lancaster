'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 54-CAMERA RENAME MP4.VBS
'# __ 
'# __ VBS 54-CAMERA RENAME MP4.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Sat 10-Aug-2019 10:18:00 ]
'# END FINAL TIME [ Sat 10-Aug-2019 12:04:00 ] NEAR 2 HOUR
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
' IT GOT AN ARRAY LIKE HERE
'
' I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4"
' I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2019 CyberShot HX60V_#"'
'
' IT IS DONE IN SUB RECURSIVE 
' ALL THE FILE ARE READ BEFORE SO NONE PROBLEM 
'
' IT THE MODIFY VERSION COPY OF 
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 29-COPY CAMERA PHOTO IMAGES.VBS
'
' IT RENAME VIDEO FILE .MP4 AND .TXT SO EXTENSION ARE UPPER CASE
'
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' SESSION 002
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Sat 10-Aug-2019 10:18:00
' Sat 10-Aug-2019 12:04:00
' -------------------------------------------------------------------


Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

DIM CUR_USER
DIM ComputerName

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

DIM JPG_BEEN_DONE
DIM MP4_BEEN_DONE

DIM STOP_PROGRESS
DIM IX_ERROR_MESSENGER

DIM DRIVE_SPACE_INFO

DIM objStartFolder
DIM objFolder
DIM I_COUNT_1
DIM I_COUNT_2
DIM VAR_PATH_SCANNER
DIM IMAGE_POINTER_NAME

DIM objSuperFolder

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO_2 = CreateObject("Scripting.FileSystemObject")

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")


Dim arr1
Dim oDict ' Key: Full filename, Item: Filename

Dim FileModDate
DIM FileModDate_3

'FOLDER_PATH_VAR="D:\DSC\2011"
'FOLDER_PATH_VAR="D:\DSC\z compare\NIKON TOTAL\DCIM"
'FOLDER_PATH_VAR="\"

' IF FSO.FolderExists(FOLDER_PATH_VAR)=FALSE THEN 
	' MSGBOX "FOLDER NOT EXIST"
	' Wscript.QUIT
' END IF


'Call the Search subroutine to start the recursive search.
'Search objFSO.GetFolder(FOLDER_PATH_VAR)
'---------------------------------------------------------
'FindWantedFiles objFSO.GetFolder(FOLDER_PATH_VAR)
'DisplayFindings
'---------------------------------------------------------

DIM R_COUNT

DIM FOLDER_PATH_INDEX

NUMBER_OF_ITEM = 100

DIM FOLDER_PATH()
DIM CAPITAL_FOLDER()
DIM FOLDER_PATH_VAR()
DIM FOLDER_PATH_IGNOR_DATE_INDEX()

REDIM FOLDER_PATH(NUMBER_OF_ITEM)
REDIM CAPITAL_FOLDER(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_VAR(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_IGNOR_DATE_INDEX(NUMBER_OF_ITEM)

DIM FOLDER_PATH_AVOID_ITEM()
REDIM PRESERVE FOLDER_PATH_AVOID_ITEM(100)
NUMBER_COUNTER = 0
I=NUMBER_COUNTER
I=0
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="_gsdata_"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="System Volume Information"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="\F.A"                      ' DASH CAMERA
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".THM"                      ' SONY
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".AmazonDriveConfig"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=""
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=""
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=""
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=""

FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_AVOID_ITEM)
	IF FOLDER_PATH_AVOID_ITEM(R_COUNT_2)<>"" THEN NUMBER_COUNTER=R_COUNT_2
NEXT
REDIM PRESERVE FOLDER_PATH_AVOID_ITEM(NUMBER_COUNTER)


DIM FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE()
REDIM PRESERVE FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(100)
NUMBER_COUNTER = 0
NUMBER_COUNTER = NUMBER_COUNTER + 1
I=NUMBER_COUNTER

I=0
'I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".JPG"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".AVI"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".MOV"          ' DASH CAMERA
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".MP4"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".TXT"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".FLV"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".URL"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".WMV"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".FFmpeg"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".ADV"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE)
	IF FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2)<>"" THEN 
		NUMBER_COUNTER=R_COUNT_2
	END IF
NEXT
REDIM PRESERVE FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(NUMBER_COUNTER)
	
DIM FOLDER_PATH_STORE

DIM EXPLORER_LOADED
DIM SHOW_EXPLORER_FLAG

DIM I
I=0
' I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4"
' I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2019 CyberShot HX60V_#"
' I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2011 CyberShot H70____#"
' I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2011 CyberShot HX5V__#"
' I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2015 CyberShot HX60V_#"
' I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2016 CyberShot HX60V_#"
' I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2017 CyberShot HX60V_#"
' I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2018 CyberShot HX60V_#"
' I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2018 CyberShot HX60V\MP4"
' I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2017 CyberShot HX60V\MP4"
' I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2016 CyberShot HX60V\MP4"
' I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2015 CyberShot HX60V\MP4"
' I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2011 CyberShot H70 __\MP4"
I=I+1:FOLDER_PATH(I)="D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\BUNKER.COM"


I=0
I=I+1:CAPITAL_FOLDER(I)="D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\BUNKER.COM"


FOR R_COUNT=1 TO UBOUND(CAPITAL_FOLDER)
	IF CAPITAL_FOLDER(R_COUNT)<>"" THEN 
		DIMSIZE=R_COUNT
	END IF
NEXT

REDIM PRESERVE CAPITAL_FOLDER(DIMSIZE)

FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)
	IF FOLDER_PATH(R_COUNT)<>"" THEN 
		DIMSIZE=R_COUNT
	END IF
NEXT

REDIM PRESERVE FOLDER_PATH(DIMSIZE)

FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)

	I_COUNT_1=0
	I_COUNT_2=0

	WORK_GO_ER=TRUE
	FOLDER_PATH_INDEX=CSTR(R_COUNT)+" "+FOLDER_PATH(R_COUNT)
	FOLDER_PATH_STORE=MID(FOLDER_PATH(R_COUNT),1,2)

	IF FOLDER_PATH_STORE<>"" THEN 
		FOLDE_POINTER=R_COUNT
		FOLDER_PATH_STORE=FOLDER_PATH(R_COUNT)
		MAIN
	END IF

NEXT


SET objFSO = NOTHING
Wscript.QUIT
'-------------------------------------------------------------------
'-------------------------------------------------------------------


Sub Main
    Set oDict = CreateObject("Scripting.Dictionary")
	
	IF FSO.FolderExists(FOLDER_PATH_STORE)=FALSE THEN 
		' MSGBOX "Folder NOT Exist ERROR"+VBCRLF+VBCRLF+FOLDER_PATH_STORE
		EXIT SUB
	END IF

    ' Call Recursive Sub
    FindWantedFiles(objFSO.GetFolder(FOLDER_PATH_STORE))
    ' Display All Findings from Dictionary object

    DisplayFindings

	ShowDriveInfo(MID(FOLDER_PATH_STORE,1,2))
	
	I=""
	I=I+"DONE RENAME COUNT: " 
	I=I+CSTR(I_COUNT_1)
	I=I+" of " + CSTR(I_COUNT_2) + VBCRLF
	'I=I+DRIVE_SPACE_INFO
	I=I+FOLDER_PATH_STORE
	
	IF SHOW_EXPLORER_FLAG=TRUE OR I_COUNT_1>0 THEN 
		MSGBOX I
	END IF
	
    Set oDict = Nothing
End Sub


function ShowDriveInfo(drvPath)
	strComputer = "." 

	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery( _
		"SELECT * FROM Win32_LogicalDisk where DeviceID='"+drvPath+"'",,48) 

	For Each objItem in colItems 
		if len(objItem.VolumeName)>0 then
			S=S+"-----------------------------------------------------------" & vbCrLf
			S=S+"VolumeName:" & vbTab & objItem.VolumeName  & vbCrLf
			S=S+"-----------------------------------------------------------" & vbCrLf
			S=S+"Drive Letter 1st Found for Match Volume Name: " & drvPath & "\" & vbCrLf
			S=S+"-----------------------------------------------------------" & vbCrLf
			S=S+"FreeSpace:" & vbTab
			S=S+FormatNumber((CDbl(objItem.FreeSpace)/1024/1024/1024))  & " GB" & vbCrLf
			S=S+"Size:" & vbTab & vbTab
			S=S+FormatNumber((CDbl(objItem.Size)/1024/1024/1024)) & " GB" & vbCrLf
			S=S+"Occupied Space:" & vbTab
			S=S+FormatNumber((CDbl(objItem.Size - objItem.FreeSpace)/1024/1024/1024)) & " GB"
			IF IX_ERROR_MESSENGER<>"" THEN 
				S=S+vbCrLf
				S=IX_ERROR_MESSENGER
			END IF
		END IF
	Next
	DRIVE_SPACE_INFO=S
END FUNCTION

Sub FindWantedFiles(sFolder)
    On Error Resume Next
    Dim oFDR, oItem, file
    ' Check if wanted files are in this folder
    For Each file In sFolder.Files
		oDict.Add FILE.path, file
    Next
	
	' Recurse into it's sub folders
	For Each oFDR In sFolder.SubFolders
		FindWantedFiles objFSO.GetFolder(oFDR.Path)
	Next
	
End Sub

Sub DisplayFindings()
    Dim oKeys, oKey
    oKeys = oDict.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
	
		IF FSO.FILEExists(VAR_PATH_SCANNER)=TRUE THEN 
			CALL PATH_SCANNER_COPY_CAMERA
		END IF
	Next
End Sub

SUB PATH_SCANNER_COPY_CAMERA
	
	DIM STRING_KEEPER1()
	DIM STRING_KEEPER2()

	I_COUNT_2=I_COUNT_2+1
	
	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	FILE_EXT=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"."))
	FILE_EXT=UCASE(FILE_EXT)
	FILE_VAR_OLD=FILE_VAR
	Set f = fso.GetFile(VAR_PATH_SCANNER)
	FileModDate = f.DateLastModified
	SET f=NOTHING
	
	' ' CONVERT THE CASE HERE RENAME 
	' ' ----------------------------
	NewFilename=VAR_PATH_SCANNER
	' FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE)
		' IF FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2)=FILE_EXT THEN 
			' EXT_1=LCASE(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2))
			' EXT_2=UCASE(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2))
			' IF INSTR(LCASE(NewFilename),EXT_1)>0 THEN 
				' NewFilename=REPLACE(NewFilename,EXT_1,EXT_2)
			' END IF
		' END IF
	' NEXT

	' CAPITAL_FOLDER RENAME
	' ----------------------------
	FOR R_COUNT_CAP=1 TO UBOUND(CAPITAL_FOLDER)
	FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE)
		IF FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2)=FILE_EXT THEN 
			IF INSTR(VAR_PATH_SCANNER,CAPITAL_FOLDER(R_COUNT_CAP))>0 THEN 
				PATH_VAR_2=MID(NewFilename,1,INSTRREV(NewFilename,"\"))
				FILE_VAR_2=MID(NewFilename,INSTRREV(NewFilename,"\")+1)
				' FILE_VAR_2=MID(FILE_VAR_2,1,INSTRREV(FILE_VAR_2,".")-1)
				' FILE_EXT_2=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"."))

				' THE EXTENSION WAS SORTED OUT BUT WILL LOWER CASE AGAIN AND REPEAT FUCTION AFTER AGIN

				FILE_VAR_2_OLD=FILE_VAR_2
				FILE_VAR_2=UCASE(MID(FILE_VAR_2,1,1))+MID(FILE_VAR_2,2)
				
				EE = LCase(FILE_VAR_2)
				EE = UCASE(MID(EE,1,1))+MID(EE,2)

				' USE AT END ROUTINE ARRAY
				' IS COLLECT VARIABLE STRING THAT WANT PRESERVE IN ORIGINAL STATE
				REDIM STRING_KEEPER1(100)
				REDIM STRING_KEEPER2(100)
				ISK=0
				String1=""
				String1=String1+" .com cdn cdn_hash cdn_ttl .com-content cdn_creation_time"
				String1=String1+" fb-1"
				String1=String1+" "
				String1=TRIM(String1)
				String_ARRAY=Split(String1," ",-1,1)
				I=UBOUND(String_ARRAY)
				MSGBOX I
				' FOR THE TIME WHEN A SPACE IN FRONT -- ADD TO THE ARRAY
				REDIM PRESERVE String_ARRAY(I+100)
				I=I+1:String_ARRAY(I)=" A."
				MSGBOX I
				FOR R3 = 1 TO UBOUND(String_ARRAY)
					IF String_ARRAY(R3)="" THEN EXIT FOR
				NEXT
				REDIM PRESERVE String_ARRAY(R3)

				For EACH R3 IN String_ARRAY
					STR2=R3
					IF INSTR(FILE_VAR_2,STR2)>0 THEN
						T1=1
						DO
							X=INSTR(T1,FILE_VAR_2,STR2)
							IF X>0 THEN 
								ISK=ISK+1
								STRING_KEEPER1(ISK)=STR2
								STRING_KEEPER2(ISK)=X
							END IF
							IF X=0 THEN EXIT DO
							T1=X+1
						LOOP UNTIL T1=0
					END IF
				NEXT
				REDIM PRESERVE STRING_KEEPER1(ISK)
				REDIM PRESERVE STRING_KEEPER2(ISK)
				
				For r1 = 2 To LEN(EE)
					ED1=MID(EE,r1,1)
					ED2=MID(EE,r1-1,1)
					IF ISNUMERIC(ED2) THEN 
						IF LCase(ED1)<> UCase(ED1) THEN 
							EE = Replace(EE, ED2+ED1, UCase(ED2+ED1))
						END IF
					END IF
				Next

				STR1=" -().\/_"""+VBCR+VBLF+"0123456789"
				STRY="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
				For r1 = 1 To LEN(STR1)
					For r2 = 1 To LEN(STRY)
						STR2=MID(STR1,r1,1)
						STR3=MID(STRY,r2,1)
						EE = Replace(EE, STR2 + LCase(STR3), STR2 + UCase(STR3))
					Next
				Next

				String1=""
				String1=String1+" _ADV _HD"
				String1=String1+" .JPG .AVI .MOV .MP4 .TXT .FLV .URL .WMV .FFmpeg"
				String1=TRIM(String1)
				String_ARRAY=Split(String1," ",-1,1)
				For r1 = 1 To LEN(STR1)
					For r2 = 1 To LEN(STR1)
						For EACH R3 IN String_ARRAY
							STR2=R3
							STR3=MID(STR2,1,1)+UCASE(MID(STR2,2,1))+LCASE(MID(STR2,3))
							EE = Replace(EE, STR3 , STR2)
						NEXT
					Next
				Next

				' HERE A LATER ONE CLEAR UP AFTER TYPE THING
				' TEST IF SEARCH CHAR AND THEN NUMBER AFTER
				'
				' THESE ARE WHAT PUT LOWER CASE
				'
				STR1=" -().\/_"""+VBCR+VBLF
				String1=LCASE(" A AND IN OF GO BY AN")
				For r2 = 1 To LEN(STRY)
					STR3=MID(STRY,r2,1)
					String1 = Replace(String1, " "+ LCase(STR3), " " + UCase(STR3))
				Next
				String1=TRIM(String1)
				String_ARRAY=Split(String1," ",-1,1)
				For r1 = 1 To LEN(STR1)
					For r2 = 1 To LEN(STR1)
						For EACH R3 IN String_ARRAY
							STR2=MID(STR1,r1,1)
							STR3=MID(STR1,r2,1)
							EE = Replace(EE, STR2 + R3 + STR3, LCASE(STR2 + R3 + STR3))
						NEXT
					Next
				Next

				' HERE A LATER AGAIN -- CLEAR UP AFTER TYPE THING
				' THESE ARE PRESERVE STRING
				String1=""
				String1=String1+" MILF CD POV AndEager FilthyGhettoD MyThickBlackA"
				String1=String1+" BF MMF GF WOW"
				String1=String1+" "
				String1=TRIM(String1)
				
				String_ARRAY=Split(String1," ",-1,1)
				For r1 = 1 To LEN(STR1)
					For r2 = 1 To LEN(STR1)
						For EACH R3 IN String_ARRAY
							STR2=R3
							STR3=UCASE(MID(STR2,1,1))+LCASE(MID(STR2,2))
							EE = Replace(EE, STR3, STR2)
							 ' IF INSTR(LCASE(EE),LCASE("xvideos.com")) THEN
								' MSGBOX EE+VBCRLF+STR2+VBCRLF+STR3
							' END IF
						NEXT
					Next
				Next

				' HERE A LATER AGAIN -- CLEAR UP AFTER TYPE THING
				String1=""
				String1=String1+" .com"
				String1=String1+" .cdn"
				String1=TRIM(String1)
				String_ARRAY=Split(String1," ",-1,1)
				For EACH R3 IN String_ARRAY
					STR2=R3
					STR3=MID(STR2,1,1)+UCASE(MID(STR2,2,1))+LCASE(MID(STR2,3))
					EE = Replace(EE, STR3 , STR2)
				NEXT


				' HERE A LATER AGAIN -- CLEAR UP AFTER TYPE THING
				String1=""
				String1=String1+" xvideos.com"
				String1=String1+" "
				String1=String1+" "
				String1=TRIM(String1)
				
				String_ARRAY=Split(String1," ",-1,1)
				For EACH R3 IN String_ARRAY
					STR2=LCASE(R3)
					STR3=UCASE(MID(STR2,1,1))+LCASE(MID(STR2,2))
					EE = Replace(EE, STR3, STR2)
					' IF INSTR(LCASE(EE),LCASE("xvideos.com")) THEN
						' MSGBOX EE+VBCRLF+STR2+VBCRLF+STR3
					' END IF
				NEXT

				' FROM THE TOP ARRAY
				' IS COLLECT VARIABLE STRING THAT WANT PRESERVE IN ORIGINAL STATE
				For ISK = 1 TO UBOUND(STRING_KEEPER1)
					G1=LEN(STRING_KEEPER1(ISK))
					G2=STRING_KEEPER2(ISK)
					EF1=EE
					EE=MID(EE,1,G2-1)+STRING_KEEPER1(ISK)+MID(EE,G1+G2)
				NEXT
				
				FILE_VAR_2=EE
			
				NewFilename=PATH_VAR_2+FILE_VAR_2
				
				' CONVERT THE CASE FOR EXTENSION HERE RENAME AGAIN
				' ------------------------------------------------
				FOR R_COUNT_3=1 TO UBOUND(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE)
					EXT_1=LCASE(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_3))
					EXT_1="."+UCASE(MID(EXT_1,2,1))+MID(EXT_1,3)
					EXT_2=FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_3)
					IF INSTR(NewFilename,EXT_1)>0 THEN 
						NewFilename=REPLACE(NewFilename,EXT_1,EXT_2)
					END IF
				NEXT

				FILE_VAR_3=MID(NewFilename,INSTRREV(NewFilename,"\")+1)
				
				' TEST DEBUG
				IF NewFilename<>VAR_PATH_SCANNER THEN 
					MSGBOX PATH_VAR+VBCRLF+VBCRLF+FILE_VAR_OLD+VBCRLF+VBCRLF+FILE_VAR_3
				END IF
				' Wscript.QUIT

			END IF
		END IF
	NEXT
	NEXT
	
	
	PATH_VAR_2=MID(NewFilename,1,INSTRREV(NewFilename,"\"))
	
	NewFilename=REPLACE(NewFilename,"\\","\")

	If Not FSO.FolderExists(PATH_VAR_2) Then
		MSGBOX "PATH NOT EXIST ERROR"+VBCRLF+VBCRLF+PATH_VAR_2
		' BuildFullPath PATH_VAR_2
		' CreateFolderTree(NewPATH)
	End If
	
	ON ERROR RESUME NEXT
	ERR.CLEAR
	
	IF NewFilename<>VAR_PATH_SCANNER THEN 
		If FSO.FileExists(NewFilename) Then
			' MSGBOX "PATH FILE ALREADY EXIST RENAME ERROR"+VBCRLF+VBCRLF+NewFilename
		END IF
	END IF

	
	
	IF NewFilename<>VAR_PATH_SCANNER THEN 
		' Wscript.QUIT
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
			
		I_COUNT_1=I_COUNT_1+1
	END IF
	
	IF ERR.NUMBER>0 THEN
		IX=""+VBCRLF
		IX=IX+"ERROR COPY FILE _ TRY REDO ANOTHER"+VBCRLF+VBCRLF
		IX=IX+VAR_PATH_SCANNER +VBCRLF+VBCRLF
		IX=IX+PATH_FILE_VAR    +VBCRLF+VBCRLF
		IX=IX+ERR.DESCRIPTION  +VBCRLF+VBCRLF
		IX_ERROR_MESSENGER=IX
		STOP_PROGRESS=TRUE
		EXIT SUB
	END IF
	
END SUB

Sub BuildFullPath(ByVal FullPath)
	If Not fso.FolderExists(FullPath) Then
		BuildFullPath fso.GetParentFolderName(FullPath)
		fso.CreateFolder FullPath
	End If
End Sub

Function GetDriveLetterFromVolumeName( volumeName )
Dim volumes, volume

    ' Unless we found a matching volume, an empty string will be the returned value
    GetDriveLetterFromVolumeName=""

    ' Ask WMI for the list of volumes with the requested label
    Set volumes = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") _ 
                  .ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE Label='" & volumeName & "'")

    ' If exist an matching volume, get its drive letter
	If volumes.Count > 0 Then 
        For Each volume In volumes 
            GetDriveLetterFromVolumeName = volume.DriveLetter
            Exit For
        Next 
    End If

End Function

' NOT ROUTINE USER HERE THIS CODE SCRIPT
Sub DeleteEmptyFolders(folder)
    Dim subfolder
	ON ERROR RESUME NEXT
	ERR.CLEAR
    For Each subfolder in folder.SubFolders
		DeleteEmptyFolders subfolder
    Next

    If folder.SubFolders.Count = 0 And folder.Files.Count = 0 Then

		ON ERROR RESUME NEXT
		ERR.CLEAR
		fso.DeleteFolder folder.Path, TRUE
		IF ERR.NUMBER>0 THEN 
			ERR_COUNTER = ERR_COUNTER +1
			ERROR_WITH_SOME = folder.Path & VBCRLF & ERR.DESCRIPTION

			ERR.CLEAR
			objFolder.DELETE, TRUE
			IF ERR.NUMBER=0 THEN ERR_COUNTER = ERR_COUNTER - 1
		END IF
    End If
End Sub


'------------------------------------------------
Sub Search(Folder)
'------------------------------------------------
	ON ERROR RESUME NEXT
	
	ERR.CLEAR
	for each file in Folder.Files
		IF ERR.NUMBER=0 THEN 
			VAR_PATH_SCANNER=file.PATH
			CALL PATH_SCANNER_RENAME
		END IF
	NEXT

	'Find EACH SUBFOLDER.
	For Each subFolder In Folder.SubFolders
		'Call the Search subroutine to start the recursive search on EACH SUBFOLDER.
		Search objFSO.GetFolder(subFolder.Path)
	Next

End Sub
		
		


'---------
' REFERENCE REFERENCE
'---------
' How Can I Get a List of All the Files in a Folder and Its Subfolders? – Hey, Scripting Guy! Blog
' https://blogs.technet.microsoft.com/heyscriptingguy/2004/10/20/how-can-i-get-a-list-of-all-the-files-in-a-folder-and-its-subfolders/
' ----

'------------------------------------------------------------------------
' THIS PROJECT USING THIS CODE
' BECAUSE A MOVE FILE HAPPEN MUST GATHER WHOLE FILE LIST BEFORE ACT ON IT
' AND IT WAS DEBUGED IN THE NEXT URL LINK LINE
'------------------------------------------------------------------------
'----
' windows - VBS to Search for Multiple Files Recursively in C:\Users - Stack Overflow
' https://stackoverflow.com/questions/25658004/vbs-to-search-for-multiple-files-recursively-in-c-users
'------------------------------------------------------------------------
'------------------------------------------------------------------------
' vbscript - Read files in subfolders - Stack Overflow
' https://stackoverflow.com/questions/16905278/read-files-in-subfolders
'----
'------------------------------------------------------------------------

' ----
' vbscript - Read files in subfolders - Stack Overflow
' https://stackoverflow.com/questions/16905278/read-files-in-subfolders
' ----