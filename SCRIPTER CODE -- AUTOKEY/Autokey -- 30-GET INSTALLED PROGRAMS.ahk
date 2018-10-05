;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 21-AutoRun.ahk
;# __ 
;# __ Autokey -- 21-AutoRun.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Wed 06:54:00 Pm_16 May 2018 ]
;# TO          TIME [ Wed 06:54:00 Pm_16 May 2018 ]
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; 001 ---------------------------------------------------------------
; WORK ON BATCH FILE VERSION
; THEN POWERSHELL VERSION
; BOTH INCOMPLETE RESULTS COMPARED TO COMMERCIAL PROGRAM

; 002 ---------------------------------------------------------------
; WORK EXTRA HAD LEARNING ABOUT LISTVIEW ADD, INSERT, DELETE, MODIFY, AUTOSIZE COLUMN
; THE CODE WORKS BETTER STRIPPING EVERYTHING AT END TO NICE PRESENTED DATA
; -------------------------------------------------------------------
; FROM  Wed 30-May-2018 15:19:59
; TO    Wed 30-May-2018 18:00:00 __ 3 HOUR MINUS 20 MINUTE
; -------------------------------------------------------------------

; 003 ---------------------------------------------------------------
; ANOTHER LITTLE BASH AND TIDIED UP AND THEN DECIDE NOT WORKING GOOD ENOUGH
; THE REGISTRY IS NOT PULLING IN ALL THE REGISTRY
; SWITCHED OVER SUCCESSFULLY TO WORKING ON VB6
; FEW MINOR START_UP ISSUE OF THE VB6 CODE DOWNLOADED
; BUT WORKING AFTER MY LOOK
; SHOULDN'T TAKE LONG TO CATCH UP WITH WHERE I LEFT OFF HERE
; USING A BETTER ARRAY SYSTEM
; IN MY VB CODE AREA
; D:\VB6\VB-NT\00_Best_VB_01\READ REGISTRY KEY AND SUB\Get_Installed_App_Set.vbp
; -------------------------------------------------------------------
; FROM  Thu 31-May-2018 00:16:41
; TO    Thu 31-May-2018 01:48:00 __ 1 HOUR MINUS 30 MINUTE
; -------------------------------------------------------------------

; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force

DetectHiddenWindows, on

SoundBeep , 2000 , 100

; Allow the user to maximize or drag-resize the window:
;Gui,1: +Resize

Gui, Add, ListView, r20 w800 gMyListView, Name|Location
;GuiControl, +Report, MyListView

;Gui,1:Font,S13 CDefault,Lucida Console
;Gui,1:Add, ListView,backgroundteal csilver grid r20 w700 +hscroll altsubmit gLW2, Name|Location

;LV_ModifyCol()

VAR_SET=0
InstallLocation_VAR=""

;LV_ModifyCol()  ; Auto-size each column to fit its contents.
;LV_ModifyCol(1, "Text")  ; For sorting purposes, indicate that column 2 is an integer.
;LV_ModifyCol(1, "Sort")  ; For sorting purposes, indicate that column 2 is an integer.

; Display the window and return. The script will be notified whenever the user double clicks a row.
Gui, Show
;Gui, Submit, NoHide

COUNTER=

vIncludeSubkeys = 1
vRecurse = 1
;Idx = 0
;sbk = 0
;FileDelete, RegReadInstalled.txt
;fileAppend, `n-----HKLM-----`n, RegreadInstalled.txt

;Loop,HKLM, SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\, %vIncludeSubkeys%, %vRecurse%

Loop, Reg, HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\, KVR
{
	if A_LoopRegType = key
		value =
	else
	{
		RegRead, value
		if ErrorLevel
			value = *error*
	}

	
	value:=StrReplace(value, """" , "")

	
	;IfInString, A_LoopRegName, DisplayVersion
	;{
	;	InstallLocation_VAR=
	;	REG_VAR=%A_LoopRegKey%\%A_LoopRegSubKey%\%A_LoopRegName%
	;	LV_Add("", REG_VAR, value)
	;	VAR_SET=0
	;	REG_VAR_2=%A_LoopRegKey%\%A_LoopRegSubKey%\%A_LoopRegName%__%VALUE%
	;}
	
	REG_VAR_3=%A_LoopRegKey%\%A_LoopRegSubKey%\%A_LoopRegName%__%VALUE%

	IfInString, REG_VAR_3, DisplayName
	{
		InstallLocation_VAR=
		REG_VAR=%A_LoopRegKey%\%A_LoopRegSubKey%\%A_LoopRegName%
		LV_Add("", REG_VAR, value)
		VAR_SET=0
	}
	
	;IfInString, A_LoopRegName, InstallLocation
	IfInString, REG_VAR_3, InstallLocation
	IF value
		IF SubStr(value, 2 , 2)=":\"
		{
			REG_VAR=%A_LoopRegKey%\%A_LoopRegSubKey%\%A_LoopRegName%
			LV_Add("", REG_VAR, value)
			InstallLocation_VAR=%value%
			VAR_SET=1
		}
		
	IfInString, REG_VAR_3, InstallSource
	IF value
		IF SubStr(value, 2 , 2)=":\"
		{
			REG_VAR=%A_LoopRegKey%\%A_LoopRegSubKey%\%A_LoopRegName%
			LV_Add("", REG_VAR, value)
			VAR_SET=1
		}
	
	IfInString, REG_VAR_3, UninstallString
	IF value
	{
		SET_GO=0
		IF SubStr(value, 2 , 2)=":\"
			SET_GO=1
		IF (!InstallLocation_VAR)
			SET_GO=1
		IF SET_GO=1
		{
			REG_VAR=%A_LoopRegKey%\%A_LoopRegSubKey%\%A_LoopRegName%
			LV_Add("", REG_VAR, value)
			VAR_SET=1
		}
	}
	
	COUNTER+=1
	IF COUNTER<20
	{
		LV_ModifyCol()  ; Auto-size each column to fit its contents.
	}
	
}


LV_ModifyCol()  ; Auto-size each column to fit its contents.
LV_ModifyCol(1, "Text")  ; For sorting purposes, indicate that column 2 is an integer.
LV_ModifyCol(1, "Sort")  ; For sorting purposes, indicate that column 2 is an integer.


; WAS GOING TO BE STRIP ANYTHING WITH A \: BUT CHANGED MY MIND
;-------------------------------------------------------------
TotalRows := LV_GetCount()
RowNumber = %TotalRows%
O_Location=
Loop % TotalRows
{ 
	LV_GetText(Location, RowNumber,2)
	LV_GetText(Name, RowNumber,1)

	IfNotInString, Name, DisplayName
	{
		IF SubStr(Location, 2 , 2)=":\"
		IF (O_Location)
		IfInString, Location, %O_Location%
			{
			;MSGBOX %Location%`r`n%O_Location%
			;LV_Delete(RowNumber)
			O_Location=
			}

		O_Location=%Location%
	}
	RowNumber-=1
}

;STRIP ANYTHING WITH MsiExec.exe AWAY FROM LISTVIEW COLUMN
;---------------------------------------------------------
TotalRows := LV_GetCount()
RowNumber = %TotalRows%
Loop % TotalRows
{ 
	LV_GetText(Location, RowNumber,2)
	IfInString, Location, MsiExec.exe
		{
		LV_Delete(RowNumber)
		}
	RowNumber-=1
}

;STRIP THE BIG BIT OF REGISTRY PATH AWAY FROM LISTVIEW COLUMN
;------------------------------------------------------------
TotalRows := LV_GetCount()
RowNumber = 1
Loop % TotalRows
{ 
	LV_GetText(Name, RowNumber,1)
	
	VAR_REG_PATH=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
	
	IfInString, Name, %VAR_REG_PATH%
	{
		VAR_REG_PATH:=SubStr(Name,StrLen(VAR_REG_PATH)+2)
		LV_Modify(RowNumber,1,VAR_REG_PATH)
	}
	RowNumber+=1
}

LV_ModifyCol()  ; Auto-size each column to fit its contents.


; INSERT A ROW BETWEEN EACH ITEM
;-------------------------------
SET_DONE_FIRST=FALSE
TotalRows := LV_GetCount()
RowNumber = 1
OLD_SET_String=
Loop 
{ 
	LV_GetText(Name, RowNumber,1)

	SET_String:=SubStr(Name,1,InStr(Name,"\")-1)
	
	If OLD_SET_String<>%SET_String%
	{
		IF SET_DONE_FIRST=TRUE
		{
			LV_Insert(RowNumber)
			;RowNumber+=1
		}
		SET_DONE_FIRST=TRUE
	}
	RowNumber+=1
	
	OLD_SET_String=%SET_String%
	
	TotalRows := LV_GetCount()
	IF RowNumber>%TotalRows%
		{
		BREAK
	}
}



; INSERT ROW IF MISSING DISPLAY NAME WITH PATH ONLY EXIST BEFORE
;---------------------------------------------------------------
TotalRows := LV_GetCount()
RowNumber = 1
Loop
{ 
	LV_GetText(Name, RowNumber,1)
	if (!Name)
	{
		RowNumber+=1
		LV_GetText(Name, RowNumber,1)
		IfNotInString, Name, DisplayName
		IfInString, Name, UninstallString
		{
			SET_VAR=\DisplayName
			SET_UninstallString:=SubStr(Name,1,InStr(Name,"\")-1)
			SET_UninstallString=%SET_UninstallString%%SET_VAR%
			LV_Insert(RowNumber,1,SET_UninstallString)
			RowNumber+=1
		}
	}
	
	RowNumber+=1

	TotalRows := LV_GetCount()
	IF RowNumber>%TotalRows%
		BREAK
}

; SEPARATE THE \
;---------------
TotalRows := LV_GetCount()
RowNumber = 1
Loop % TotalRows
{ 
	LV_GetText(Name, RowNumber,1)
	
	VAR_REG_PATH:="  \  "
	StringReplace, VAR_REG_PATH, VAR_REG_PATH,Chr(39),,All
	;VAR_REG_PATH_2:=%A_SPACE% , %VAR_REG_PATH% , %A_SPACE%
	;VAR_REG_PATH_2:=
	IfInString, Name, \
	{
		VAR_REG_PATH_3:=SubStr(Name,1,InStr(Name,"\")-1)
		VAR_REG_PATH_4:=SubStr(Name,InStr(Name,"\")+1)
		VAR_REG_PATH_5=%VAR_REG_PATH_3%%VAR_REG_PATH%%VAR_REG_PATH_4%
		LV_Modify(RowNumber,1,VAR_REG_PATH_5)
	}
	RowNumber+=1
}

; DELETE ROWS THAT ARE DUPLICATE PATH MESSY
;------------------------------------------
TotalRows := LV_GetCount()
RowNumber = 1 ;%TotalRows%
OLD_Location_1=
Old_Name=
Loop
{ 
	LV_GetText(Location, RowNumber,2)
	LV_GetText(Name, RowNumber,1)
	
	IfNotInString, Name, , DisplayName
	IfNotInString, Old_Name, , DisplayName
	IF (OLD_Location_1)
	IfInString, Location, %OLD_Location_1%
	{
		IfNotInString, OLD_Name, DisplayName
		{
			LV_Delete(RowNumber)
			RowNumber-=3
		}
	}
	
	OLD_Location_1=%Location%
	OLD_Name=%Name%
	
	RowNumber+=1
	TotalRows := LV_GetCount()
	IF RowNumber>%TotalRows%
		BREAK
}


; REMOVE THE HASHING NAME AND PUT DISPLAY NAME
;---------------------------------------------
TotalRows := LV_GetCount()
RowNumber = 1
Loop
{ 
	LV_GetText(Name, RowNumber,1)
	LV_GetText(Location, RowNumber,2)
	IfInString, Name, DisplayName
	{
		{
			SET_UninstallString:=SubStr(Name,InStr(Name,"\"))
			
			;msgbox % SET_UninstallString
			
			;SET_UninstallString=%SET_UninstallString%%SET_VAR%
			
			;LV_Insert(RowNumber,1,SET_UninstallString)
			RowNumber+=1
		}
	}
	
	RowNumber+=1

	TotalRows := LV_GetCount()
	IF RowNumber>%TotalRows%
		BREAK
}



LV_ModifyCol()  ; Auto-size each column to fit its contents.

return


MyListView:
if A_GuiEvent = DoubleClick
{
    LV_GetText(RowText, A_EventInfo)  ; Get the text from the row's first field.
    ToolTip You double-clicked row number %A_EventInfo%. Text: "%RowText%"
}
return

GuiClose:  ; When the window is closed, exit the script automatically:
ExitApp