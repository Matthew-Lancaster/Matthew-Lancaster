
' An example of a NOT expression on a file's pathname
' using VBScript's built in Regular Expression object.

' Dim regEx
' Set regEx = new RegExp
' regEx.Pattern = SearchParms.FileNameCustomParm
' regEx.IgnoreCase = True

' Wscript.QUIT

' DIM IPATH
' DIM strPath_2
' strPath_2="\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_02_D_DRIVE\CD"
' IPATH=MID(strPath_2,INSTR(17,strPath_2,"\"))
' MSGBOX IPATH

Public Function isValidFileName( ByVal strPath, ByVal strFileName )

On Error Resume Next

	Dim bIsValid
	DIM IPATH
	bIsValid = FALSE

	IPATH=MID(strPath,INSTR(17,strPath,"\"))
	IF INSTR(IPATH,"1-ASUS-")>0 THEN bIsValid = True
	IF INSTR(IPATH,"2-ASUS-")>0 THEN bIsValid = True
	IF INSTR(IPATH,"3-LINDA-")>0 THEN bIsValid = True
	IF INSTR(IPATH,"4-ASUS-")>0 THEN bIsValid = True
	IF INSTR(IPATH,"5-ASUS-")>0 THEN bIsValid = True
	IF INSTR(IPATH,"7-ASUS-")>0 THEN bIsValid = True
	IF INSTR(IPATH,"8-MSI-")>0 THEN bIsValid = True
	
	DIM M(20)
	I=0
	I=I+1:M(I)="\Documents and Settings"
	I=I+1:M(I)="\_saved_"
	I=I+1:M(I)="_MIRRORS_"
	I=I+1:M(I)="_HISTORY_"
	I=I+1:M(I)="\Start Menu\"
	I=I+1:M(I)="#NFS-"
	I=I+1:M(I)="C:\0 00 SETTING INI SYNC-ER\"
	I=I+1:M(I)="\TEMP\"
	I=I+1:M(I)="\Users\Administrator"
	I=I+1:M(I)="\RF\"
	I=I+1:M(I)="\USERS\"
	'I=I+1:M(I)="E:\"
	'I=I+1:M(I)=""

	FOR R=1 TO I
		M(R)=UCASE(M(R))
		IF M(R)<>"" THEN
			IF INSTR(UCASE(strPath + strFileName),M(R))>0 THEN bIsValid = False
		END IF
	NEXT
	
	
	isValidFileName = bIsValid

End Function

