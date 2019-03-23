;Autokey Project to Use Enter Key on (Save As) Window Dialogue of Chrome.exe & Exit On Chrome Gone
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force
;------------------------------------------------------------------------------------
; Code to Press Enter Key on Save As Window in Chrome.exe & Exit When Chrome.exe Gone
;------------------------------------------------------------------------------------
; Thu 25 Aug 2016 01:47:51 -- Length Time 1 hr 45 Min -- Finish Code

;----------------------------------------------------------------------
;NOTE After a Little Work Test Method and Find Error Fix Through Help File
;Because Window 10 Won't Except Keyboard Send From The AutoKey Unless 
;It Set to Load As Administrator 
;By Set Properties on This Program in 64 Bit or (x86) 32 Bit -- Of the Main Program Folder
;C:\Program Files\AutoHotkey\AutoHotkey.exe
;Yes Fixed the Error Variable Increment -- Seem Like a C++ Method Style -- TEST_VAR += 1
;----------------------------------------------------------------------
;My code begin
;----------------------------------------------------------------------
SoundBeep , 1900 , 400
SetStoreCapslockMode, off


SetCapsLockState , On
SetNumLockState , Off
SetScrollLockState , Off