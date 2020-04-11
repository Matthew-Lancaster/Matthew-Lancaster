Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

	' Dim FSO
	' Set FSO = CreateObject("Scripting.FileSystemObject")
	' ' Const ForReading = 1, ForWriting = 2, ForAppending = 8
	' Set File_2 = FSO.OpenTextFile("C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT -  RETEKESS.TXT", 2)
	' File_2.Write strPath + strFileName + VBCRLF
	' wscript.quit
	
	
' FILE PATH SEARCH STRING
' "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT"
' $(AllLocalDrives)\

' $(AllLocalDrives)\;C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT

Public Function isValidFileName( ByVal strPath, ByVal strFileName)

	On Error Resume Next

	bIsValid = FALSE
	

	bIsValid = TRUE
	IF INSTR(strPath,"_gsdata_")>0 THEN bIsValid = FALSE
	
	isValidFileName = bIsValid
	
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set FSO_4 = CreateObject("Scripting.FileSystemObject")

	' Const ForReading = 1, ForWriting = 2, ForAppending = 8

	FIRST_DUMMY_FILE=strPath+strFileName
	
	' IF strPath="c:\" AND strFileName="Some_Text_File.txt" THEN 
	' IF strPath="c:\" THEN 
	' IF strFileName="Some_Text_File.txt" THEN 
	' IF FIRST_DUMMY_FILE="c:\Some_Text_File.txt" THEN 
	' IF strPath="" THEN 
	' IF FSO.FILEEXISTS(FILE_TEST)=FALSE THEN 

	
	FILE_TEST="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT -  RETEKESS.TXT"
	IF strPath="c:\" AND strFileName="Some_Text_File.txt" THEN 
		Set objFile = FSO.CreateTextFile(FILE_TEST,True)
		objFile.Close
	ELSE
		Set File = FSO.OpenTextFile(FILE_TEST, 8)
		File.Write strPath + strFileName + VBCRLF
	END IF
	
	
	
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
