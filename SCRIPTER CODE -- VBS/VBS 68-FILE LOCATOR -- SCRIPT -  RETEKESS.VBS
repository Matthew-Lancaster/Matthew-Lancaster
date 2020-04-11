Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true


' FILE PATH SEARCH STRING
' "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT"
' $(AllLocalDrives)\

' $(AllLocalDrives)\;C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT

Public Function isValidFileName( ByVal strPath, ByVal strFileName )

	On Error Resume Next

	bIsValid = FALSE
	

	bIsValid = TRUE
	IF INSTR(strPath,"_gsdata_")>0 THEN bIsValid = FALSE
	
	isValidFileName = bIsValid
	
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

	' Const ForReading = 1, ForWriting = 2, ForAppending = 8

	Set File_2 = fso.OpenTextFile("C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT -  RETEKESS.TXT", 8)
	File_2.Write strPath + strFileName
	
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
