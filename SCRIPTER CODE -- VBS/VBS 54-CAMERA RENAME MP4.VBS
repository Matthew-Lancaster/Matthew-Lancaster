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
DIM FOLDER_PATH_VAR()
DIM FOLDER_PATH_IGNOR_DATE_INDEX()

REDIM FOLDER_PATH(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_VAR(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_IGNOR_DATE_INDEX(NUMBER_OF_ITEM)

DIM FOLDER_PATH_AVOID_ITEM()
REDIM PRESERVE FOLDER_PATH_AVOID_ITEM(100)
NUMBER_COUNTER = 0
I=NUMBER_COUNTER
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

I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".JPG"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".AVI"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".MOV"          ' DASH CAMERA
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".MP4"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".TXT"
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
I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4"
I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2019 CyberShot HX60V_#"
I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2011 CyberShot H70____#"
I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2011 CyberShot HX5V__#"
I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2015 CyberShot HX60V_#"
I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2016 CyberShot HX60V_#"
I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2017 CyberShot HX60V_#"
I=I+1:FOLDER_PATH(I)="D:\VI_ DSC ME\2010+Sony\2018 CyberShot HX60V_#"
I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2018 CyberShot HX60V\MP4"
I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2017 CyberShot HX60V\MP4"
I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2016 CyberShot HX60V\MP4"
I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2015 CyberShot HX60V\MP4"
I=I+1:FOLDER_PATH(I)="D:\DSC\2015+Sony\2011 CyberShot H70 __\MP4"



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
	
	I_COUNT_2=I_COUNT_2+1
	
	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)

	Set f = fso.GetFile(VAR_PATH_SCANNER)
	FileModDate = f.DateLastModified
	SET f=NOTHING
	
	NewFilename=REPLACE(VAR_PATH_SCANNER,".mp4",".MP4")
	NewFilename=REPLACE(NewFilename,".txt",".TXT")

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


' NOT ROUTINE USER HERE THIS CODE SCRIPT
SUB PATH_SCANNER_RENAME_4

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)

	NewFilename=FILE_VAR
	NewFilename=REPLACE(NewFilename,"__","_")
	NewFilename=REPLACE(NewFilename," _ ","_")
	NewFilename=REPLACE(NewFilename,"_ ","_")
	NewFilename=REPLACE(NewFilename," _","_")
	NewFilename=REPLACE(NewFilename,".jpg",".JPG")
	Extension=ucase(mid(NewFilename,len(NewFilename)-3))
	
	NewFilename=PATH_VAR+NewFilename
	
	SET_GO=FALSE
	if Extension=".JPG" then 
		SET_GO=TRUE
	end	if
	
	if VAR_PATH_SCANNER <> NewFilename AND SET_GO=TRUE then 
		'msgbox NewFilename
		I_COUNT_1=I_COUNT_1+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		IF ERR.NUMBER>0 THEN 
			MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename
		END IF
		'Wscript.QUIT
	end if

END SUB


' NOT ROUTINE USER HERE THIS CODE SCRIPT
SUB PATH_SCANNER_RENAME_3

	NewFilename=REPLACE(VAR_PATH_SCANNER,".JPG",".jpg")
	
	if VAR_PATH_SCANNER <> NewFilename then 
		I_COUNT_1=I_COUNT_1+1
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		'msgbox NewFilename
		'Wscript.QUIT
	end if

END SUB

' NOT ROUTINE USER HERE THIS CODE SCRIPT
SUB PATH_SCANNER_RENAME_2

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)

	SET_GO=TRUE 
	IF INSTR(UCASE(VAR_PATH_SCANNER),".JPG")=0 THEN SET_GO=FALSE
	'IF MID(FILE_VAR,1,4)<>"2010" THEN  SET_GO=FALSE

	NewFilename=FILE_VAR

	IF LEN(FILE_VAR)>= 12 AND SET_GO=TRUE THEN SET_GO=FALSE
	
	ON ERROR RESUME NEXT
	if SET_GO=TRUE then
		NewFilename_COUNTER=CLNG(MID(NewFilename,4,INSTR(NewFilename,".")-4))
		IF ERR.NUMBER >0 THEN 
			msgbox "ERROR CLNG WITH "+VBCRLF+VBCRLF+NewFilename
		END IF
	END IF

	'IF NewFilename_COUNTER>= 761 THEN SET_GO=FALSE
	
	IF SET_GO=TRUE THEN
		'MSGBOX VAR_PATH_SCANNER
		NewFilename=FILE_VAR

		if FSO.FileExists(PATH_VAR+NewFilename)=TRUE THEN 
			'msgbox "FILE ALREADY EXISTS "+VBCRLF+VBCRLF+NewFilename
			'Wscript.QUIT
			X_D=0
			DO
				FOR R=1 TO 500
					X_D=X_D+1
					NewFilename="DSC"+CSTR(X_D)+".jpg"
					NewFilename_VAR_2=MID(NewFilename,4)
					 
					IF LEN(NewFilename)<12 THEN 
						X_C=1
						DO
							IF LEN(NewFilename)<12 THEN 
								NewFilename="DSC"+STRING(X_C,"0")+NewFilename_VAR_2
								X_C=X_C+1
							END IF
						LOOP UNTIL LEN(NewFilename)=12
					end if
					if FSO.FileExists(PATH_VAR+NewFilename)=FALSE THEN EXIT FOR
				NEXT
			LOOP UNTIL FSO.FileExists(PATH_VAR+NewFilename)=FALSE

		END IF

		NewFilename=PATH_VAR+NewFilename

		if VAR_PATH_SCANNER <> NewFilename then 
			msgbox NewFilename
			Wscript.QUIT
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename

			I_COUNT_1=I_COUNT_1+1
		END IF
	END IF
	
	
END SUB

' NOT ROUTINE USER HERE THIS CODE SCRIPT
SUB PATH_SCANNER_RENAME

	SET_GO=FALSE
	IF INSTR(UCASE(VAR_PATH_SCANNER),"- SONY DSC-H7 -") THEN SET_GO=TRUE
	IF INSTR(UCASE(VAR_PATH_SCANNER),".JPG")=0 THEN SET_GO=FALSE
	
	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	
	'IF MID(UCASE(FILE_VAR),1,4)="2010" THEN SET_GO=TRUE
	
	SET_GO=TRUE 
	IF INSTR(UCASE(VAR_PATH_SCANNER),".JPG")=0 THEN SET_GO=FALSE
	
	'IF MID(FILE_VAR,1,4)="DSCN" THEN  SET_GO=FALSE
	'IF MID(FILE_VAR,1,4)<>"2010" THEN  SET_GO=FALSE

	IF LEN(FILE_VAR)>= 12 AND SET_GO=TRUE THEN SET_GO=FALSE
	
	NewFilename=FILE_VAR
	NewFilename="DSC"+MID(NewFilename,12)
	
	ON ERROR RESUME NEXT
	if SET_GO=TRUE then
		NewFilename_COUNTER=CLNG(MID(NewFilename,4,INSTR(NewFilename,".")-4))
		IF ERR.NUMBER >0 THEN 
			msgbox "ERROR CLNG WITH "+VBCRLF+VBCRLF+NewFilename
		END IF
	END IF

	'IF NewFilename_COUNTER>= 761 THEN SET_GO=FALSE
	
	IF SET_GO=TRUE THEN 
		
		NewFilename_COUNTER_STRING=MID(NewFilename,INSTR(NewFilename,".")-4)
		NewFilename_COUNTER_STRING=CLNG(MID(NewFilename,4,INSTR(NewFilename,".")-4))
		NewFilename_COUNTER_STRING=CLNG(MID(NewFilename,4,INSTR(NewFilename,".")-4))
		NewFilename
		IF LEN(NewFilename)<12 THEN 
			X_C=1
			DO
				IF LEN(NewFilename)<12 THEN 
					NewFilename="DSC"+STRING(X_C,"0")+MID(NewFilename,4)
					X_C=X_C+1
				END IF
			LOOP UNTIL LEN(NewFilename)=12
		end if


		
		' IF LEN(FILE_VAR)<12 THEN 
			' X_C=1
			' DO
				' IF LEN(FILE_VAR)<12 THEN 
					' NewFilename="DSC"+STRING(X_C,"0")+MID(FILE_VAR,4)
					' X_C=X_C+1
				' END IF
			' LOOP UNTIL LEN(NewFilename)=12
			' NewFilename=PATH_VAR+NewFilename
		' end if
			
		'msgbox NewFilename
		'Wscript.QUIT
		
		if FSO.FileExists(PATH_VAR+NewFilename)=TRUE THEN 
			'msgbox "FILE ALREADY EXISTS "+VBCRLF+VBCRLF+NewFilename
			'Wscript.QUIT
			FOR R=1 TO 400
				NewFilename="DSC"+CSTR(R)+".jpg"
				IF LEN(NewFilename)<12 THEN 
					X_C=1
					DO
						IF LEN(NewFilename)<12 THEN 
							NewFilename="DSC"+STRING(X_C,"0")+MID(NewFilename,4)
							X_C=X_C+1
						END IF
					LOOP UNTIL LEN(NewFilename)=12
				end if
				 
				msgbox NewFilename
				 
				 if FSO.FileExists(PATH_VAR+NewFilename)=FALSE THEN EXIT FOR
			 NEXT
		END IF
		
		NewFilename=PATH_VAR+NewFilename
		
		SET_GO=FALSE
		if instr(ucase(NewFilename),".JPG")>0 THEN SET_GO=TRUE
		PATH_VAR_4=MID(NewFilename,1,INSTRREV(NewFilename,"\"))
		IF INSTR(VAR_PATH_SCANNER,PATH_VAR_4)=0 THEN SET_GO=FALSE
		'MSGBOX PATH_VAR_4
		'MSGBOX VAR_PATH_SCANNER
		IF SET_GO=TRUE THEN 
			msgbox NewFilename
			'Wscript.QUIT
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename
			I_COUNT_1=I_COUNT_1+1
		END IF
	END IF

END SUB


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