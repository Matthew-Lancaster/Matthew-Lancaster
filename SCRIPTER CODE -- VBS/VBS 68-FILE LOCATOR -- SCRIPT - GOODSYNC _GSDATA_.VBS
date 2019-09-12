
Public Function isValidFileName( ByVal strPath, ByVal strFileName )

On Error Resume Next

	Dim bIsValid
	bIsValid = FALSE

	' DIM IPATH
	' IPATH=MID(strPath,INSTR(17,strPath,"\"))
	' IF INSTR(IPATH,"1-ASUS-")>0 THEN bIsValid = True
	' IF INSTR(IPATH,"2-ASUS-")>0 THEN bIsValid = True
	' IF INSTR(IPATH,"3-LINDA-")>0 THEN bIsValid = True
	' IF INSTR(IPATH,"4-ASUS-")>0 THEN bIsValid = True
	' IF INSTR(IPATH,"5-ASUS-")>0 THEN bIsValid = True
	' IF INSTR(IPATH,"7-ASUS-")>0 THEN bIsValid = True
	' IF INSTR(IPATH,"8-MSI-")>0 THEN bIsValid = True
	
	bIsValid = True
	
	strFileName
	EXTNAME=UCASE(MID(strFileName,INSTRREV(strFileName,".")))
	F4=UCASE(MID(strFileName,1,4))
	F8=UCASE(MID(strFileName,1,8))
	IF EXTNAME="._GS" THEN bIsValid = False
	IF EXTNAME=".GSS" THEN bIsValid = False
	IF EXTNAME=".GSL" THEN bIsValid = False
	IF EXTNAME=".LOG" THEN bIsValid = False

	' bIsValid = False
	' IF EXTNAME=".LOG" AND ISNUMERIC(F4) THEN bIsValid = TRUE
	' IF EXTNAME=".LOG" AND F8="GOODSYNC" THEN bIsValid = TRUE
	' IF EXTNAME=".LOG" AND F8="SENDSMTP" THEN bIsValid = TRUE
	
	
	' DIM M(20)
	' I=0
	' I=I+1:M(I)="_MIRRORS_"
	' I=I+1:M(I)="\_saved_"
	' I=I+1:M(I)="_HISTORY_"
	' I=I+1:M(I)="E:\"
	' I=I+1:M(I)=""
	
	' FOR R=1 TO I
		' M(R)=UCASE(M(R))
		' IF M(R)<>"" THEN
			' IF INSTR(UCASE(strPath + strFileName),M(R))>0 THEN bIsValid = False
		' END IF
	' NEXT
	
	
	isValidFileName = bIsValid

End Function

