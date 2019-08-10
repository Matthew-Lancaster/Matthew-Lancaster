DIM objStartFolder
DIM objFolder
DIM I_COUNT
DIM VAR_PATH_SCANNER
DIM objSuperFolder

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO_2 = CreateObject("Scripting.FileSystemObject")

Dim arr1
Dim oDict ' Key: Full filename, Item: Filename

FOLDER_PATH="D:\DSC\2011"
FOLDER_PATH="D:\DSC\z compare\NIKON TOTAL\DCIM"
FOLDER_PATH="D:\DSC"

I_COUNT=0

IF FSO.FolderExists(FOLDER_PATH)=FALSE THEN 
	MSGBOX "FOLDER NOT EXIST"
	Wscript.QUIT
END IF

'Call the Search subroutine to start the recursive search.
'Search objFSO.GetFolder(FOLDER_PATH)
'---------------------------------------------------------
'FindWantedFiles objFSO.GetFolder(FOLDER_PATH)
'DisplayFindings
'---------------------------------------------------------

MAIN

SET objFSO = NOTHING
Wscript.QUIT
'-------------------------------------------------------------------
'-------------------------------------------------------------------

Sub Main
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oDict = CreateObject("Scripting.Dictionary")
    ' Call Recursive Sub
    FindWantedFiles(objFSO.GetFolder(FOLDER_PATH))
    ' Display All Findings from Dictionary object
    DisplayFindings

	MSGBOX "DONE RENAME COUNT _ " & CSTR(I_COUNT) & " of " & CSTR(oDict.count)

    Set fso = Nothing
    Set oDict = Nothing
End Sub

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
		CALL PATH_SCANNER_RENAME_4
    Next
End Sub


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
		I_COUNT=I_COUNT+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		IF ERR.NUMBER>0 THEN 
			MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename
		END IF
		'Wscript.QUIT
	end if

END SUB


SUB PATH_SCANNER_RENAME_3

	NewFilename=REPLACE(VAR_PATH_SCANNER,".JPG",".jpg")
	
	if VAR_PATH_SCANNER <> NewFilename then 
		I_COUNT=I_COUNT+1
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		'msgbox NewFilename
		'Wscript.QUIT
	end if

END SUB

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

			I_COUNT=I_COUNT+1
		END IF
	END IF
	
	
END SUB

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
			I_COUNT=I_COUNT+1
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
' REFERENCE
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