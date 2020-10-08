#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#Persistent
; -------------------------------------------------------------------
; IT USER ExitFunc TO EXIT FROM #Persistent
; OR      Exitapp  TO EXIT FROM #Persistent
; Exitapp HAVE AR CALL ONTO ExitFunc
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; Register a function to be called on exit:
OnExit("ExitFunc")

; Register an object to be called on exit:
OnExit(ObjBindMethod(MyObject, "Exiting"))
; -------------------------------------------------------------------

;--------------------
#SingleInstance force
;--------------------

	SOUNDBEEP 4000,80

	; EXITAPP

	GOSUB GO_FOLDER

	EXITAPP

RETURN

F1::EXITAPP


GO_FOLDER:
{
	; GLOBAL FILE_SCRIPT
	; GLOBAL FILE_SCRIPT_COUNT

	FILE_SCRIPT_COUNT=0
	FILE_SCRIPT := Object()
	FOLDER_EXCLUDE := Object()

	; FILE_SCRIPT := []

	; -------------------------------------------------------------------
	; SET THE PATH OF PHOTO FOLDER _ WITHOUT RECURSING SUB-FOLDER IS GOOD IDEA
	; TO BE ABLE QUICKLY COUNT THEM VERIFY AND SEE AS ORDER TO GO ONTO FACEBOOK
	; AND NOT STRETCH MY CODE TOO MUCH ABOUT WANT RECURSING SUB-FOLDER 
	; SINGLE FOLDER ONLY AT THE MOMENT
	; -------------------------------------------------------------------
	
	; Example #3: Retrieve file names sorted by name (see next example to sort by date):
	FileList =  ; Initialize to be blank.
	Loop, Files, D:\*.*, D
		FileList = %FileList%%A_LoopFileName%`n
	Sort, FileList ;  R  ; The R option sorts in reverse order. See Sort for other options.

	Loop, parse, FileList, `n
	{
		if A_LoopField =  ; Ignore the blank item at the end of the list.
			continue
		FILE_SCRIPT[A_Index] := A_LoopField
		FILE_SCRIPT_COUNT := A_Index
		; MSGBOX % FILE_SCRIPT[A_Index]
		
	}	

	Loop % FILE_SCRIPT.MaxIndex()
	{
		FILE_NAME:=FILE_SCRIPT[A_Index]
		; StringUpper, FILE_NAME,FILE_NAME
		FILE_SCRIPT[A_Index]:=FILE_NAME
	}

	
	ACUM=0
	ACUM+=1
	FOLDER_EXCLUDE[ACUM] := "$RECYCLE.BIN"
	ACUM+=1
	FOLDER_EXCLUDE[ACUM] := "_gsdata_"
	ACUM+=1
	FOLDER_EXCLUDE[ACUM] := "System Volume Information"
	ACUM+=1
	FOLDER_EXCLUDE[ACUM] := "$GetCurrent"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "AWS"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "Documents and Settings"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "Program Files (x86)"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "ProgramData"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "RECOVERY"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "SADPLOG"
	; ACUM+=1
	; FOLDER_EXCLUDE[ACUM] := "windows"
	
	
	Loop % FOLDER_EXCLUDE.MaxIndex()
	{
		FILE_NAME:=FOLDER_EXCLUDE[A_Index]
		StringUpper, FILE_NAME,FILE_NAME
		FOLDER_EXCLUDE[A_Index]:=FILE_NAME
	}
	
	
	X_Index:=FILE_SCRIPT.MaxIndex()
	Loop % FILE_SCRIPT.MaxIndex()
	{
		X_Index-=1
		Loop % FOLDER_EXCLUDE.MaxIndex()
		{
			IF (InStr(FILE_SCRIPT[X_Index]"*", FOLDER_EXCLUDE[A_Index]"*"))
			{
				; msgbox % FILE_SCRIPT[X_Index]
				FILE_SCRIPT.RemoveAt(X_Index)
			}
		}	
	}	

	Loop % FILE_SCRIPT.MaxIndex()
	{
		VALUE:=FILE_SCRIPT[A_Index]
		; FileCreateDir, \\7-asus-gl522vw\7_asus_gl522vw_10_1_samsung_4tb_c\%VALUE%
		FileCreateDir, Q:\%VALUE%
		SOUNDBEEP 2000,20
		
	}

	SOUNDBEEP 2000,200

	; 31 FOLDER ADDED TO 4TB FROM D
	
}
RETURN
	



; FileCreateDir, \\8-msi-gp62m-7rd\8_msi_gp62m_7rd_04_4_samsung_4tb_hubic\D4\#
