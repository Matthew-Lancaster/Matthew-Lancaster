;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 33-#0 Send To Date Of File In Front FileName.ahk
;# __ 
;# __ Autokey -- 33-#0 Send To Date Of File In Front FileName.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Fri 15-Jun-2018 00:40:00 ]
;# END         TIME [ Fri 15-Jun-2018 01:00:00 ] 20 MINUTE
;# __ 
;  =============================================================

; -------------------------------------------------------------------
; DESCRIPTION
; -------------------------------------------------------------------
; AUTO LAUNCH VB CODE AND CHECK IF ALREADY RUNNING 
; AND IN THE PROJECT IDE
; -------------------------------------------------------------------

; -------------------------------------------------------------------
; 001 ---------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Fri 15-Jun-2018 00:40:00
; TO   TIME __ Fri 15-Jun-2018 00:40:00 BEGINNER
; -------------------------------------------------------------------


;# ------------------------------------------------------------------
; Location Internet
;# ------------------------------------------------------------------
; Link to Folder of My AutoHotKeys Project Set
; https://drive.google.com/open?id=0BwoB_cPOibCPVmVYT1pKWUk4LVE
;# ------------------------------------------------------------------
; Link to Folder of My AutoHotKeys Project Set Dropbox
; https://www.dropbox.com/sh/h2ebk12dksaq7j3/AAD9Ow_SbBP33JKmuALRkO1_a?dl=0
;# ------------------------------------------------------------------
; REQUIRE CHECKING
; Link to This File On DropBox With Most Up to Date
; https://www.dropbox.com/s/w22p90h81uwni71/Autokey%20--%2021-AUTORUN.ahk?dl=0
;# ------------------------------------------------------------------


; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force

;SetStoreCapslockMode, off


DetectHiddenWindows, on
SetTitleMatchMode 2  ; Avoids specify the full path of the file below.

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

;--------------------------------------------------------------------
;AUTOHOTKEYS
;--------------------------------------------------------------------

FN_VAR:="D:\VB6\VB-NT\00_Send_To\Send To Date Of File In Front FileName\#0 Send To Date Of File In Front FileName.vbp"

WIN_VAR:="Send_To_Date_In_Front_FName - Microsoft Visual Basic [ ahk_class wndclass_desked_gsk"

IfExist, %FN_VAR%
	IFWinNotExist, %WIN_VAR%
	{
		SoundBeep , 2000 , 100
		Run, "%FN_VAR%", , MAX
	}
	ELSE
	{
		WinMaximize, %WIN_VAR%
		WinActivate, %WIN_VAR%
	}
	
EXITAPP