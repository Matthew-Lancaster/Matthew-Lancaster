DIM objStartFolder
DIM objFolder
DIM TEST_COUNTER
DIM I_COUNT

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

Set objFSO = CreateObject("Scripting.FileSystemObject")

DIM FIND_NAME_F(2)
DIM FIND_NAME_R(2)
I=0
I=I+1:FIND_NAME_F(I)="1-ASUS-EEE"
I=I+1:FIND_NAME_F(I)="2-ASUS-X5DIJ"
I=0
I=I+1:FIND_NAME_R(I)="2-ASUS-EEE"
I=I+1:FIND_NAME_R(I)="1-ASUS-X5DIJ"

DIM ARRAY(20)
I=0
I=I+1:ARRAY(I)="C:\RF"
I=I+1:ARRAY(I)="C:\0 00 LOGGERS TEXT"
I=I+1:ARRAY(I)="D:\VB6-EXE"
I=I+1:ARRAY(I)="C:\SCRIPTER"
I=I+1:ARRAY(I)="C:\TEMP"

I=I+1:ARRAY(I)="C:\DOWNLOADS"
I=I+1:ARRAY(I)="D:\VB6"
I=I+1:ARRAY(I)="D:\VB6 DUPE COMPARE"
I=I+1:ARRAY(I)="D:\VB6-ARC"
I=I+1:ARRAY(I)="D:\0 00 Art Loggers"

I=I+1:ARRAY(I)="D:\KAT MP3 RECORDER"
I=I+1:ARRAY(I)="D:\KAT MP3 RECORDER_ARRIVAL"
I=I+1:ARRAY(I)="C:\PStart"
I=I+1:ARRAY(I)="E:\"

I=I+1:ARRAY(I)=""
I=I+1:ARRAY(I)=""

' FOR R=2 TO 20:ARRAY(R)="":NEXT

DIM ARRAY_2(20000)
TEST_COUNTER=0

'------------------------------------------------
FOR R_LOOP=1 TO UBOUND(ARRAY)
IF ARRAY(R_LOOP)<>"" THEN 
	IF FSO.FolderExists(ARRAY(R_LOOP))=TRUE THEN 
		objStartFolder = ARRAY(R_LOOP)
		Set objFolder = objFSO.GetFolder(objStartFolder)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			FOR R_LOOP_2 = 1 TO UBOUND(FIND_NAME_F)
				IF instr(objFile.Name,FIND_NAME_F(R_LOOP_2))>0 THEN 
					I_COUNT=I_COUNT+1
					ARRAY_2(I_COUNT)=objFile.PATH
				END IF
			Next
		Next
		ShowSubfolders objFSO.GetFolder(objStartFolder)
	END IF
END IF
NEXT
SET objFSO = NOTHING

'------------------------------------------------
FOR R = I_COUNT TO 1 STEP -1
	FOR R_LOOP_2 = 1 TO UBOUND(FIND_NAME_F)
		STRIP_FILE_NAME=MID(ARRAY_2(R),INSTRREV(ARRAY_2(R),"\")+1)
		IF instr(STRIP_FILE_NAME,FIND_NAME_F(R_LOOP_2))>0 THEN 
			RENAME ARRAY_2(R)
		END IF
	Next
NEXT

'-------------------------------------------------------------------
'-------------------------------------------------------------------
MSGBOX "RENAME FILE AND SUBFOLDER _ DONE ON _ " & CSTR(TEST_COUNTER) 
'-------------------------------------------------------------------
Wscript.QUIT
'-------------------------------------------------------------------
'-------------------------------------------------------------------

'------------------------------------------------
SUB RENAME(FILE_PATH_IN)
'------------------------------------------------
	FILE_PATH_NAME=FILE_PATH_IN
	STRIP_FILE_NAME_1=MID(FILE_PATH_IN,INSTRREV(FILE_PATH_IN,"\")+1)
	O_FILE_PATH_NAME=""
	DO
		FOR R_LOOP_2 = 1 TO UBOUND(FIND_NAME_F)
			FIND_NAME_F_V1=FIND_NAME_F(R_LOOP_2)
			FIND_NAME_R_V2=FIND_NAME_R(R_LOOP_2)
			FILE_PATH_NAME=REPLACE(FILE_PATH_NAME,FIND_NAME_F_V1, FIND_NAME_R_V2)
		NEXT
		IF O_FILE_PATH_NAME=FILE_PATH_NAME THEN EXIT DO
		O_FILE_PATH_NAME=FILE_PATH_NAME
	LOOP UNTIL TRUE=FALSE

	STRIP_FILE_NAME_2=MID(FILE_PATH_NAME,INSTRREV(FILE_PATH_NAME,"\")+1)
		
	TEST_COUNTER=TEST_COUNTER+1
	ERROR_V=FALSE

	IF FSO.FileExists(FILE_PATH_IN)=TRUE THEN 
		IF FSO.FileExists(FILE_PATH_NAME)=TRUE THEN 
			WScript.Echo FILE_PATH_NAME & " ALREADY EXISTS"
			'FSO.DELETEFILE(FILE_PATH_NAME)
		END IF 
		ON ERROR RESUME NEXT
		ERR.CLEAR
		
		set oFILE = FSO.GetFile(FILE_PATH_IN)
		oFILE.Name = STRIP_FILE_NAME_2
		
		IF ERR.NUMBER > 0 THEN ERROR_V=TRUE:ERROR_S=ERR.DESCRIPTION
		ON ERROR GOTO 0
	END IF 
	
	IF FSO.FolderExists(FILE_PATH_IN)=TRUE THEN 
		IF FSO.FolderExists(FILE_PATH_NAME)=TRUE THEN 
			WScript.Echo FILE_PATH_NAME & " ALREADY EXISTS"
		END IF 
		ON ERROR RESUME NEXT
		ERR.CLEAR
		
		set oFldr = FSO.GetFolder(FILE_PATH_IN)
		oFldr.Name = STRIP_FILE_NAME_2
		
		IF ERR.NUMBER > 0 THEN ERROR_V=TRUE:ERROR_S=ERR.DESCRIPTION
		ON ERROR GOTO 0
	END IF 
	
	IF ERROR_V=TRUE THEN 
		WScript.Echo CINT(TEST_COUNTER) & " OF " & CINT(I_COUNT) & VBCRLF & FILE_PATH_IN & VBCRLF & "ERROR " & ERROR_S
	END IF 
	
END SUB

'------------------------------------------------
Sub ShowSubFolders(Folder)
'------------------------------------------------
	ON ERROR RESUME NEXT
	ERR.CLEAR
    For Each Subfolder in Folder.SubFolders
		IF ERR.NUMBER=0 THEN 
		FOR R_LOOP_2 = 1 TO UBOUND(FIND_NAME_F)
			IF instr(Subfolder.Path,FIND_NAME_F(R_LOOP_2))>0 THEN 
				I_COUNT=I_COUNT+1
				ARRAY_2(I_COUNT)=Subfolder.Path
			END IF
		Next
		END IF

		IF ERR.NUMBER=0 THEN 
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
		END IF 

		IF ERR.NUMBER=0 THEN 
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
			FOR R_LOOP_2 = 1 TO UBOUND(FIND_NAME_F)
				IF instr(objFile.Name,FIND_NAME_F(R_LOOP_2))>0 THEN 
					I_COUNT=I_COUNT+1
					ARRAY_2(I_COUNT)=objFile.PATH 
					'WScript.Echo objFile.PATH 
				END IF
			Next
        Next
		END IF 

		IF ERR.NUMBER=0 THEN 
			ShowSubFolders Subfolder
		END IF
    Next
End Sub
		

'---------
'REFERANCE
'---------
'How Can I Get a List of All the Files in a Folder and Its Subfolders? – Hey, Scripting Guy! Blog
'https://blogs.technet.microsoft.com/heyscriptingguy/2004/10/20/how-can-i-get-a-list-of-all-the-files-in-a-folder-and-its-subfolders/
'----