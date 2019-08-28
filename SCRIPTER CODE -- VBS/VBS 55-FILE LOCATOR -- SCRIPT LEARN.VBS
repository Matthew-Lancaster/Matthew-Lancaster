
' An example of a NOT expression on a file's pathname
' using VBScript's built in Regular Expression object.

' Dim regEx
' Set regEx = new RegExp
' regEx.Pattern = SearchParms.FileNameCustomParm
' regEx.IgnoreCase = True

' Wscript.QUIT

Public Function isValidFileName( ByVal strPath, ByVal strFileName )

On Error Resume Next

	Dim bIsValid
'	bIsValid = True
' 	bIsValid = Not regEx.Test( strPath )

	IF INSTR(strPath,"1-ASUS-")>0 THEN bIsValid = True
	IF INSTR(strPath,"2-ASUS-")>0 THEN bIsValid = True
	IF INSTR(strPath,"3-LINDA-")>0 THEN bIsValid = True
	IF INSTR(strPath,"4-ASUS-")>0 THEN bIsValid = True
	IF INSTR(strPath,"5-ASUS-")>0 THEN bIsValid = True
	IF INSTR(strPath,"7-ASUS-")>0 THEN bIsValid = True
	IF INSTR(strPath,"8-MSI-")>0 THEN bIsValid = True
	IF INSTR(strPath,"C:\Documents and Settings")>0 THEN bIsValid = False
	IF INSTR(strPath + strFileName,"\_saved_")>0 THEN bIsValid = False
	IF INSTR(strPath + strFileName,"\Start Menu\")>0 THEN bIsValid = False
	IF INSTR(strPath + strFileName,"#NFS-")>0 THEN bIsValid = False
	IF INSTR(strPath + strFileName,"C:\0 00 SETTING INI SYNC-ER\")>0 THEN bIsValid = False
	
	
	isValidFileName = bIsValid

End Function

