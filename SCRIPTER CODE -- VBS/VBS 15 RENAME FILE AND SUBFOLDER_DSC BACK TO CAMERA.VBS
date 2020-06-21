'#-------------------------------------------------------------------
' __ 
' __ CODE BY MATTHEW LANCASTER
' __ Matt.Lan@btinternet.com 
' __ 
'#-------------------------------------------------------------------
' Sat 08-Feb-2020 07:28:08
' Sat 08-Feb-2020 11:57:23 -- 4 HOUR 29 MINUTE

' -------------------------------------------------------------------
'#-------------------------------------------------------------------
'
'#-------------------------------------------------------------------
' Code Description
' VBScript
' -------------------------------------------------------------------
' IF EVER LOST CAMERA CARD DATA 
' AND BACKUP COMPUTER LIBRARY 
' BUT RENAME THEM ALL FANCY
' THE CAMERA WILL ONLY TAKE THEM BACK ON EXACT THE SAME
' IF PRESERVE AND THEN DSC NUMBER THAT WILL ABE ABLE
' HERE CODE ALLOW COMPUTER ALBUM PUT BACK IN CAMERA
' LIKE STREET CAMERA BEST TO REQUIRE SOME ON
' IT COME IN TWO PART
' RENAME FILE THAT IN ONE FOLDER
' USER ANOTHER EXTERNAL CODE THAT AS DESCRIBE IN CODE TEXT-ER
' THAT WILL PUT EACH FILE FOLD ACCORD TO DATE
' AND THEN THE 2ND PART RENAME FOLDER TO FORMAT CAMERA USER
' NOT EXACT SPECIFIC
' ONLY ORDERLY NUMBER -- THAT SLOT IN DATE SOMEWHERE
' USER DEFINE VARIABLE FOR THAT START PLACE
' THE SORT WILL BEGIN AND BACKWARD IN TIME DATE FOLDERIN_
' AND THEN CAMERA HITT BUTTON TO MAKE SORT SO THUMB EACH ONE
' NOT FOR VIDEO YET AS THEY LESS QUANTITY FOR MY CODIN_
' -------------------------------------------------------------------
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 001
' -------------------------------------------------------------------







' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF IS WORKER
' -------------------------------------------------------------------



DIM ORIGINAL_PATH
DIM EXPLORER_COUNTING

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

' -------------------------------------------------------------------
' HERE WAS WORK FOR A BIT 
' AND BELOW A BETTER CODE
' -------------------------------------------------------------------
Dim iMyPID
' -------------------------------------------------------------------
' iMyPID = GetObject("winmgmts:root\cimv2").Get("Win32_Process.Handle='" & CreateObject("WScript.Shell").Exec("mshta.exe").ProcessID & "'").ParentProcessId
' -------------------------------------------------------------------
' MSGBOX iMyPID
' -------------------------------------------------------------------

' -------------------------------------------------------------------
FOR EACH objProcess IN colProcesses
	IF NOT (IsNull(objProcess.ProcessID )) THEN
		iMyPID = objProcess.ProcessID
	END IF
NEXT
' -------------------------------------------------------------------
' MSGBOX "MY PROCESS NUMBER PID IS -- "+ CSTR(iMyPID),vbMsgBoxSetForeground 
' -------------------------------------------------------------------

Set args = Wscript.Arguments
For Each arg In args
	Wscript_Arguments = Wscript_Arguments + arg
Next

Dim strScriptName
Dim objProcess


i1 = 0 ' ANY PROGRAM WSCRIPT 
i2 = 0 ' MY  PROGRAM WSCRIPT 
FOR EACH objProcess IN colProcesses
    IF NOT (IsNull(objProcess.CommandLine )) THEN
		strScriptName = Trim(Right(objProcess.CommandLine,Len(objProcess.CommandLine) - InstrRev(objProcess.CommandLine,"\")))
        strScriptName = Left(strScriptName, Len(strScriptName) - 1)
		PID_Script = objProcess.ProcessID
		
		IF strScriptName = Wscript.ScriptName THEN
			i2=i2+1
		END IF
		IF strScriptName <> Wscript.ScriptName THEN
			i1=i1+1
		END IF
		
		IF i1 > 0 THEN ' ANY PROGRAM WSCRIPT 
			' Wscript.Quit
		END IF
		IF strScriptName = Wscript.ScriptName THEN ' MY PROGRAM WSCRIPT 
			' -------------------------------------------------------
			' KILL PROCESS NUMBER OF ALL OTHER INSTANCE 
			' HERE SCRIPT WHEN START
			' -------------------------------------------------------
			
			IF iMyPID<>PID_Script AND PID_Script>0 THEN 
			    ' MSGBOX CSTR(iMyPID) + VBCRLF + CSTR(PID_Script)
				objProcess.Terminate() 

			END IF
		END IF
    END IF
NEXT
                


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
I=I+1:ARRAY(I)=""
I=I+1:ARRAY(I)=""

' FOR R=2 TO 20:ARRAY(R)="":NEXT

I=0
I=I+1:ARRAY(I)="C:\DD"

DIM ARRAY_2(20000)
TEST_COUNTER=0


' ------------------------------------------------
FOR R_LOOP=1 TO 1 ' UBOUND(ARRAY)
IF ARRAY(R_LOOP)<>"" THEN 
	IF FSO.FolderExists(ARRAY(R_LOOP))=TRUE THEN 
		objStartFolder = ARRAY(R_LOOP)
		Set objFolder = objFSO.GetFolder(objStartFolder)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			FOR R_LOOP_2 = 1 TO UBOUND(FIND_NAME_F)
				IF instr(objFile.Name,"DSC")>0 THEN 
					I_COUNT=I_COUNT+1
					ARRAY_2(I_COUNT)=objFile.PATH
				END IF
			Next
		Next
		ShowSubfolders objFSO.GetFolder(objStartFolder)
	END IF
END IF
NEXT
' SET objFSO = NOTHING
' HERE GET ALL FILE

' AND PROCESS OPERATION 
'------------------------------------------------
FOR R = I_COUNT TO 1 STEP -1
	FOR R_LOOP_2 = 1 TO UBOUND(FIND_NAME_F)
		' -------------------------------------------------------
		' RENAME FUNCTION ROUTINE SET
		' -------------------------------------------------------
		RENAME ARRAY_2(R)
	Next
NEXT

' -------------------------------------------------------------------
' WHEN RUN HERE
' WANT TO RUN 
' D:\VB6\VB-NT\00_Send_To\Send To TouchDate\#Send To TouchDate.exe
' AND SELECTOR
' MENUBAR ITEM TOP ROW
' MOVE ALL FILE DATE FOLDER
' WHICH WILL PUT EACH FOLDER OF DATE AND MOVE FILE THEM DATE THERE
' AND CODE FOLOW
' WILL RENAME EACH FOLDER TO CAMERA STYLE
' NOT ABLE IN ONE SHOT
' WITHER REQUIRE ABOVE FIRST GO
' DO NOT ANY HARM -- GIVE IT RUN DOUBLE
' HAPPY RENAME FILE ABOVE WILL NOT BE REQUIRE AGAIN IF TWO
' -------------------------------------------------------------------
' Sat 08-Feb-2020 11:22:00
' -------------------------------------------------------------------

DIM DEST_FOLDER
DEST_FOLDER="80091200"

FOR R=1 TO UBOUND(ARRAY_2)
	ARRAY_2(R)=""
NEXT

'-------------------------------------------------------------------
'-------------------------------------------------------------------
MSGBOX "RENAME FILE AND SUBFOLDER _ DONE ON _ " & CSTR(TEST_COUNTER) ,vbMsgBoxSetForeground 
'-------------------------------------------------------------------

TEST_COUNTER=0
DIM ORIGINAL_FOLDER

R_LOOP=1
' ------------------------------------------------
IF ARRAY(R_LOOP)<>"" THEN 
	ORIGINAL_FOLDER=ARRAY(R_LOOP)
	IF FSO.FolderExists(ARRAY(R_LOOP))=TRUE THEN 
		objStartFolder = ARRAY(R_LOOP)
		Set objFolder = objFSO.GetFolder(objStartFolder)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			I_COUNT=I_COUNT+1
			ARRAY_2(I_COUNT)=objFile.PATH

			' IF NOT FILE ROOT FOLDER AND NOT SHOW HERE
			' NOT GET ROOT FOLDER IF FILE ARE ALL IN SUB FOLDER
			' -----------------------------------------
		Next
		ShowSubfolders objFSO.GetFolder(objStartFolder)
	END IF
END IF
SET objFSO = NOTHING

DIM O_FOLDER_NAME

' AND PROCESS OPERATION 
' ------------------------------------------------
FOR R = UBOUND(ARRAY_2) TO 0 STEP -1
	' -------------------------------------------------------
	' RENAME_FOLDER FUNCTION ROUTINE SET
	' -------------------------------------------------------
	RENAME_FOLDER ARRAY_2(R)
NEXT

'-------------------------------------------------------------------
'-------------------------------------------------------------------
MSGBOX "RENAME FILE AND SUBFOLDER _ DONE ON _ " & CSTR(TEST_COUNTER) ,vbMsgBoxSetForeground 
'-------------------------------------------------------------------
Wscript.QUIT
'-------------------------------------------------------------------
'-------------------------------------------------------------------

SUB RENAME_FOLDER(FILE_PATH_IN)
	IF FILE_PATH_IN="" THEN EXIT SUB

	' FILE AND FOLDER COME IN AND DON'T WANT FILE
	IF INSTR(FILE_PATH_IN,".")>0 THEN 
		EXIT SUB
	END IF
	
	IF O_FOLDER_NAME=FILE_PATH_IN THEN 
		EXIT SUB
	END IF 
	O_FOLDER_NAME=FOLDER_NAME
	
	
	DEST_FOLDER=CSTR(CLNG(DEST_FOLDER)-1)
	FILE_PATH_NAME=DEST_FOLDER
	STRIP_FILE_PATH_NAME=MID(FILE_PATH_NAME,INSTRREV(FILE_PATH_NAME,"\")+1)
	
	IF FSO.FolderExists(FILE_PATH_NAME)=TRUE THEN 
		WScript.Echo FILE_PATH_NAME & " ALREADY EXISTS"
	END IF 
	ON ERROR RESUME NEXT
	ERR.CLEAR
	
	' MSGBOX "01 "+VBCR+FILE_PATH_IN+VBCR+STRIP_FILE_PATH_NAME+VBCR
	set oFldr = FSO.GetFolder(FILE_PATH_IN)
	oFldr.Name = STRIP_FILE_PATH_NAME
	TEST_COUNTER=TEST_COUNTER+1
	
	IF ERR.NUMBER > 0 THEN ERROR_V=TRUE:ERROR_S=ERR.DESCRIPTION
	ON ERROR GOTO 0

	
END SUB	



'------------------------------------------------
SUB RENAME(FILE_PATH_IN)
'------------------------------------------------

	FOLDER_NAME=MID(FILE_PATH_IN,1, INSTRREV(FILE_PATH_IN,"\"))
	STRIP_FILE_NAME_1=MID(FILE_PATH_IN,INSTRREV(FILE_PATH_IN,"\")+1)
	O_FILE_PATH_NAME=""

	IF INSTR(STRIP_FILE_NAME_1,"DSC0")>0 THEN
	FILE_PATH_NAME=MID(STRIP_FILE_NAME_1,INSTR(STRIP_FILE_NAME_1,"DSC0"))
	IF INSTR(FILE_PATH_NAME,"_")>0 THEN
		FILE_PATH_NAME=MID(FILE_PATH_NAME,1,INSTR(FILE_PATH_NAME,"_")-1)+".JPG"
		' MSGBOX FILE_PATH_NAME
	END IF
	IF INSTR(FILE_PATH_NAME,"- Copy")=0 then
	IF LEN(FILE_PATH_NAME)>12 THEN 
		' MSGBOX "01 "+VBCR+"02 "+FILE_PATH_IN+VBCR+"03 "+FILE_PATH_NAME+VBCR+"04 "+STRIP_FILE_NAME_1
	END IF
	END IF
	END IF
	IF INSTR(FILE_PATH_NAME,"- Copy")>0 then FILE_PATH_NAME=""
		
	' MSGBOX FILE_PATH_NAME
		
	' RESTORE FOLDER NAME BACK IN
	' RENAME NOT REQUIRE
	' ---------------------------------------------------------------
	' FILE_PATH_NAME=FOLDER_NAME+FILE_PATH_NAME
	
	TEST_COUNTER=TEST_COUNTER+1
	ERROR_V=FALSE

	IF FILE_PATH_IN<>"" AND FILE_PATH_NAME<>"" THEN
		IF FSO.FileExists(FILE_PATH_IN)=TRUE THEN 
			IF FSO.FileExists(FILE_PATH_NAME)=TRUE THEN 
				' WScript.Echo FILE_PATH_NAME & " ALREADY EXISTS"
				' FSO.DELETEFILE(FILE_PATH_NAME)
			END IF 
			ON ERROR RESUME NEXT
			ERR.CLEAR
			
			IF FSO.FileExists(FILE_PATH_NAME)=FALSE THEN 
				' MSGBOX FILE_PATH_IN +VBCR+FILE_PATH_NAME
				set oFILE = FSO.GetFile(FILE_PATH_IN)
				oFILE.Name = FILE_PATH_NAME
				TEST_COUNTER=TEST_COUNTER+1
			END IF
			
			IF ERR.NUMBER > 0 THEN ERROR_V=TRUE:ERROR_S=ERR.DESCRIPTION
			ON ERROR GOTO 0
		END IF 
		
		IF INSTR(ERROR_S,"ERROR File already exists")=0 AND ERROR_V=TRUE THEN 
			ERROR_V=FALSE
		END IF
		IF ERROR_V=TRUE THEN 
			WScript.Echo CINT(TEST_COUNTER) & " OF " & CINT(I_COUNT) & VBCRLF & FILE_PATH_IN & VBCRLF & "ERROR " & ERROR_S
		END IF 
	END IF 
	
END SUB

'------------------------------------------------
Sub ShowSubFolders(Folder)
'------------------------------------------------
	ON ERROR RESUME NEXT
	ERR.CLEAR
    For Each Subfolder in Folder.SubFolders
		' IF LEN(Subfolder.Path)<=LEN(ORIGINAL_FOLDER)+2 THEN 
		'IF UCASE(ORIGINAL_FOLDER)=UCASE(Subfolder.Path) THEN 
			'MSGBOX "04 "+VBCR+UCASE(ORIGINAL_FOLDER)+VBCR+UCASE(Subfolder.Path)
		'END IF
		IF UCASE(ORIGINAL_FOLDER)<>UCASE(Subfolder.Path) THEN 
			IF ERR.NUMBER=0 THEN 
				I_COUNT=I_COUNT+1
				ARRAY_2(I_COUNT)=Subfolder.Path
			END IF
		END IF

		IF ERR.NUMBER=0 THEN 
			Set objFolder = objFSO.GetFolder(Subfolder.Path)
		END IF 

		IF ERR.NUMBER=0 THEN 
			Set colFiles = objFolder.Files
			For Each objFile in colFiles
				IF UCASE(ORIGINAL_FOLDER)<>UCASE(Subfolder.Path) THEN 
					I_COUNT=I_COUNT+1
					ARRAY_2(I_COUNT)=objFile.PATH 
				END IF
			Next
		END IF 

		IF ERR.NUMBER=0 THEN 
			ShowSubFolders Subfolder
		END IF
    Next
End Sub
		

'---------
'REFERENCE
'---------
'How Can I Get a List of All the Files in a Folder and Its Subfolders? – Hey, Scripting Guy! Blog
'https://blogs.technet.microsoft.com/heyscriptingguy/2004/10/20/how-can-i-get-a-list-of-all-the-files-in-a-folder-and-its-subfolders/
'----