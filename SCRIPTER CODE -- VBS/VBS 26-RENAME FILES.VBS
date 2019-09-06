DIM objStartFolder
DIM objFolder
DIM I_COUNT

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")

FOLDER_PATH="D:\0 00 VIDEO CCSS\CCSE BIRD IN ROOM FEW MINUTE"
FOLDER_PATH="\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\0 00 VIDEO CCSS\2018-07-26 _ BIRDS IN GARDEN HIKVISION"


IF FSO.FolderExists(FOLDER_PATH)=FALSE THEN 
	MSGBOX "FOLDER NOT EXIST",vbMsgBoxSetForeground 
	Wscript.QUIT
END IF

RENAME_SINGLE_PATH(FOLDER_PATH)

MSGBOX "DONE RENAME COUNT _ " & CSTR(I_COUNT),vbMsgBoxSetForeground 

SET objFSO = NOTHING
Wscript.QUIT
'-------------------------------------------------------------------
'-------------------------------------------------------------------


SUB RENAME_SINGLE_PATH(FOLDER_PATH)
	
	Set objSuperFolder = objFSO.GetFolder(FOLDER_PATH)
	Set objFolder = objFSO.GetFolder(objSuperFolder)

	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		SET_GO=FALSE
		IF UCASE(MID(objFile.NAME,1,2))="CH" THEN SET_GO=TRUE
		IF SET_GO=TRUE THEN 
			newFILE_1=UCASE(MID(objFile.NAME,1,4))
			
			newFilename=objFile.PATH
			newFilename=REPLACE(newFilename,"ch01_","")
			newFilename=REPLACE(newFilename,"ch02_","")
			newFilename=REPLACE(newFilename,"ch03_","")
			newFilename=REPLACE(newFilename,"ch04_","")
			newFilename=REPLACE(newFilename,"ch05_","")
			
			newFilename=REPLACE(newFilename,".mp4","_"+newFILE_1+".mp4")
			newFilename=REPLACE(newFilename,".txt","_"+newFILE_1+".txt")
			
			' msgbox objFile.PATH
			' msgbox newFilename
			
			Fso.MoveFile objFile.PATH, newFilename
			
			I_COUNT=I_COUNT+1
		END IF
	Next

END SUB


'------------------------------------------------
Sub ShowSubFolders(Folder)
'------------------------------------------------
	ON ERROR RESUME NEXT
	ERR.CLEAR
    For Each Subfolder in Folder.SubFolders
		IF ERR.NUMBER=0 THEN 
			SET_GO=FALSE
			IF instr(UCASE(Subfolder.NAME),".JPG")>0 THEN SET_GO=TRUE
			IF instr(Subfolder.PATH,"\DOC")>0 THEN SET_GO=FALSE
			IF instr(Subfolder.PATH,"\_gsdata_")>0 THEN SET_GO=FALSE
			IF SET_GO=TRUE THEN 
				I_COUNT=I_COUNT+1
			END IF
		END IF

		IF ERR.NUMBER=0 THEN 
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
		END IF 

		IF ERR.NUMBER=0 THEN 
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
			SET_GO=FALSE
			IF instr(UCASE(objFile.NAME),".JPG")>0 THEN SET_GO=TRUE
			IF instr(objFile.PATH,"\DOC")>0 THEN SET_GO=FALSE
			IF instr(objFile.PATH,"\_gsdata_")>0 THEN SET_GO=FALSE
			IF SET_GO=TRUE THEN 
				I_COUNT=I_COUNT+1
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