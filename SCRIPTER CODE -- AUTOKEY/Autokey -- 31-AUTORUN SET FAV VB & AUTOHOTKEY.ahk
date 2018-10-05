;  =============================================================
;# __ C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 31-AUTORUN SET FAV VB & AUTOHOTKEY.ahk
;# __ 
;# __ Autokey -- 31-AUTORUN SET FAV VB & AUTOHOTKEY.ahk
;# __ 
;# BY Matthew __ Matt.Lan@Btinternet.com __ 
;# __ 
;# START       TIME [ Sun 27-May-2018 04:48:00 ]
;# END         TIME [ Sun 27-May-2018 04:58:00 ] 10 MINUTE
;# LAST EDITOR TIME [ Sun 27-May-2018 04:58:00 ]
;# __ 
;  =============================================================

;# ------------------------------------------------------------------
;# ------------------------------------------------------------------

; 001 ---------------------------------------------------------------
; AUTOEXEC BOOT SCRIPT
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Sun 27-May-2018 04:48:00
; TO   TIME __ Sun 27-May-2018 04:48:00 10 MINUTE
; -------------------------------------------------------------------
; 002 ---------------------------------------------------------------
; -------------------------------------------------------------------
; -------------------------------------------------------------------
; FROM TIME __ Wed 30-May-2018 07:47:13
; TO   TIME __ Wed 30-May-2018 08:10:00 MAJOR CHANGE USE COMMAND LINE PARAM SAVE DUPLICATE CODE
; -------------------------------------------------------------------

;# ------------------------------------------------------------------
; Location OnLine
;# ------------------------------------------------------------------
; Link to Folder of My AutoHotKeys Project Set
; https://drive.google.com/open?id=0BwoB_cPOibCPVmVYT1pKWUk4LVE
;# ------------------------------------------------------------------
; Link to Folder of My AutoHotKeys Project Set Dropbox
; https://www.dropbox.com/sh/h2ebk12dksaq7j3/AAD9Ow_SbBP33JKmuALRkO1_a?dl=0
;# ------------------------------------------------------------------
; Link to This File On DropBox With Most Up to Date
; https://www.dropbox.com/s/w22p90h81uwni71/Autokey%20--%2021-AUTORUN.ahk?dl=0
;# ------------------------------------------------------------------


; SCRIPT BEGINNER ===================================================
#Warn
#NoEnv
#SingleInstance Force

DetectHiddenWindows, on
SetStoreCapslockMode, off

SoundBeep , 2000 , 100
SoundBeep , 2500 , 100

GOSUB MAIN_ROUTINE

SOUNDBEEP 1000,200
SOUNDBEEP 1500,200
; -------------------------------------------------------------------

Exitapp

RETURN



MAIN_ROUTINE:
	
FN_VAR:="C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 21-AutoRun.ahk"	
IfExist, %FN_VAR%
{
	SoundBeep , 2500 , 100
	Run, "%FN_VAR%" /RUN_FAVS
}

RETURN


