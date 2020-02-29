
Public Function isValidFileName( ByVal strPath, ByVal strFileName )

On Error Resume Next
	
	' bIsValid = False
	' IF EXTNAME=".LOG" AND ISNUMERIC(F4) THEN bIsValid = TRUE
	' IF EXTNAME=".LOG" AND F8="GOODSYNC" THEN bIsValid = TRUE
	' IF EXTNAME=".LOG" AND F8="SENDSMTP" THEN bIsValid = TRUE

	bIsValid = FALSE
	
	IF INSTR(strPath,"_saved_")>0 THEN bIsValid = True
	IF INSTR(strPath,"_history_")>0 THEN 
		bIsValid = True
	END IF
	IF INSTR(strPath,"_gsdata_")>0 THEN 
		bIsValid = True
		EXTNAME=UCASE(MID(strFileName,INSTRREV(strFileName,".")))
		' CHECK FOR "._GS" ".GSS" ".GSL" ".LOG"
	END IF
	IF INSTR(strPath,"\TEMP\")>0 THEN 
		DIM COMPNAME(7)
		COMPNAME(1)="-1-ASUS-X5DIJ"
		COMPNAME(2)="-2-ASUS-EEE"
		COMPNAME(3)="-3-LINDA-PC"
		COMPNAME(4)="-4-ASUS-GL522VW"
		COMPNAME(5)="-5-ASUS-P2520LA"
		COMPNAME(6)="-7-ASUS-GL522VW"
		COMPNAME(7)="-8-MSI-GP62M-7RD"
		FOR R=1 TO 7
			IF INSTR(strPath,COMPNAME(R))>0 THEN   
				bIsValid = True
				EXIT FOR
			END IF
		NEXT 
		EXTNAME=UCASE(MID(strFileName,INSTRREV(strFileName,".")))
		' CHECK FOR ".EXE" -- "insync-upd"
	END IF
	IF bIsValid = True THEN 
		IF INSTR(strFileName,".gdcache")>0    THEN bIsValid = False
		IF INSTR(strFileName,"insync-upd")>0  THEN bIsValid = False
		IF EXTNAME="._GS" THEN                     bIsValid = False
		IF EXTNAME=".GSS" THEN                     bIsValid = False
		IF EXTNAME=".GSL" THEN                     bIsValid = False
		IF EXTNAME=".LOG" THEN                     bIsValid = False
		IF EXTNAME=".EXE" THEN                     bIsValid = False
	END IF
	
	isValidFileName = bIsValid

End Function

