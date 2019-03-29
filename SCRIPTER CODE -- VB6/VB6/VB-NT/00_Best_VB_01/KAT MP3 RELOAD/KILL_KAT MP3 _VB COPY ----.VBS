Dim strPath
Dim lngSize

Dim fso
Dim fold
Dim files
Dim file
DIM COUNTER_UP_1
DIM COUNTER_UP_2
DIM VAR_FILE_NAME

FOR R = 1 TO 7

	INAME=""
	IF R=1 THEN INAME="D:\KAT MP3 RECORDER\KAT MP3 1-ASUS-EEE"
	IF R=2 THEN INAME="D:\KAT MP3 RECORDER\KAT MP3 1-ASUS-X5DIJ"
	IF R=3 THEN INAME="D:\KAT MP3 RECORDER\KAT MP3 3-LINDA-PC"
	IF R=4 THEN INAME="D:\KAT MP3 RECORDER\KAT MP3 4-ASUS-GL522VW"
	IF R=5 THEN INAME="D:\KAT MP3 RECORDER\KAT MP3 5-ASUS-P2520LA"
	IF R=7 THEN INAME="D:\KAT MP3 RECORDER\KAT MP3 7-ASUS-GL522VW"

	IF INAME<>"" THEN 

		strPath = INAME

		Set fso = CreateObject("scripting.filesystemobject")
		Set fold = fso.GetFolder(strPath)
		Set files = fold.files

		'2016-08-16-18-59-46.mp3

		For Each file In files
			COUNTER_UP_2 = COUNTER_UP_2 + 1
			XGO1=0
			lngSize = 400
			If file.Size <= lngSize Then XGO1=1
			'not equal to exactly then not want
			If LEN(file.NAME) <> 23 Then XGO1=0
			XGO2=0
			If INSTR(file.NAME,".MP3") = 0 Then XGO2= 1
			If INSTR(file.NAME,".WAV") = 0 Then XGO2= 1

			If XGO1 = 1 AND XGO2 = 1 Then
				'IF MSGB=0 THEN MSGBOX file.NAME:MSGB=1
				VAR_FILE_NAME=file.NAME
				COUNTER_UP_1 = COUNTER_UP_1 + 1
				'LOG.MESSAGE ("DELETE_ER") '& strPath & "UNDER SIZE")
				
				'EXIT FOR
				ON ERROR RESUME NEXT
				file.Delete 'True
				IF ERR.NUMBER > 0 THEN COUNTER_UP_1 = COUNTER_UP_1 - 1
				ON ERROR GOTO 0
			END IF
		Next
	END IF
NEXT

	MSGBOX COUNTER_UP_1 & " __ COUNT __ DELETE_ER __ OF __ " & COUNTER_UP_2 & VBCRLF & "D:\KAT MP3 RECORDER\ __ ALL MAIN FOLDER" & VBCRLF &  "UNDER SIZE __ " & lngSize & " __ "  & VBCRLF &   STRPATH&"\"&VAR_FILE_NAME & VBCRLF & "WAS THE LAST ONE OF WAV AND MP3 AND LEN FILENAME __ " & VBCRLF & "23 __ STANDARD SHORT_ER" 


REM 'C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe

REM ----
REM windows - Delete files less than a specific size - Stack Overflow
REM http://stackoverflow.com/questions/6764621/delete-files-less-than-a-specific-size
REM ----

REM ----
REM DELETE SIZE LESS THAN BATCH COMMMAND - Google Search
REM https://www.google.co.uk/search?q=DELETE+SIZE+LESS+THAN+BATCH+COMMMAND&rlz=1C1CHBD_en-GBGB721GB721&oq=DELETE+SIZE+LESS+THAN++BATCH+COMMMAND&aqs=chrome..69i57.45524j1j1&sourceid=chrome&ie=UTF-8
REM ----

REM LEARN SOME THIS 
REM ----
REM FileLocator Pro
REM file:///C:/Program%20Files/Mythicsoft/FileLocator%20Pro/help/en/index.html
REM ----