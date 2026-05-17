DIM I,FSO
DIM VF1
DIM VF2
DIM COUNTER

Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")

REM JUNK SEARCHING
REM strCurDir = WshShell.CurrentDirectory
REM WshShell.Run strCurDir & "\attribute.exe", 0
REM Set WshShell = Nothing

scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)


REM ----
REM How to Play beep sound in VBScript
REM http://www.visualbasicscript.com/How-to-Play-beep-sound-in-VBScript-m1328.aspx
REM --------
REM VBS SOUND BEEP - Google Search
REM https://www.google.co.uk/search?q=VBS+SOUND+BEEP&rlz=1C1CHBD_en-GBGB721GB721&oq=VBS+SOUND+BEEP&aqs=chrome..69i57.4066j1j1&sourceid=chrome&ie=UTF-8
REM ----

strSoundFile = "C:\WINDOWS\Media\Notify.wav"
Set objShell = CreateObject("Wscript.Shell")
strCommand = "C:\sndrec32 /play /close " & chr(34) & strSoundFile & chr(34)
objShell.Run strCommand, 0, True

REM ----
REM EXIT A VBS - Google Search
REM https://www.google.co.uk/search?q=EXIT+A+VBS&rlz=1C1CHBD_en-GBGB721GB721&oq=EXIT+A+VBS&aqs=chrome..69i57j69i65j69i60l2j0l2.2261j1j1&sourceid=chrome&ie=UTF-8
REM ----
REM ----
REM How Can I Stop a Script Partway Through? – Hey, Scripting Guy! Blog
REM https://blogs.technet.microsoft.com/heyscriptingguy/2004/08/30/how-can-i-stop-a-script-partway-through/
REM ----

'Wscript.Quit

Set FSO = CreateObject("scripting.filesystemobject")
Set fold = FSO.GetFolder(scriptdir)
Set files = fold.files
For Each file In files
	If INSTR(UCASE(file.NAME),".EXE")>0 Then 
		XYZ = scriptdir + "\" + file.NAME
	END IF
NEXT

PATH_FILE_NAME1 = XYZ
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

REM MSGBOX PATH_FILE_NAME1+VBCRLF+PATH_FILE_NAME2

ON ERROR RESUME NEXT
XTIME=INT(NOW)+TIMER+8
DO 
	COUNTER=COUNTER+1
	ERR.CLEAR
	Set F1 = FSO.GetFile(PATH_FILE_NAME2)
	'MSGBOX(F1.NAME & VBCRLF & PATH_FILE_NAME2)
	
	FSO.COPYFILE PATH_FILE_NAME1,PATH_FILE_NAME2
LOOP UNTIL ERR.NUMBER=0 OR INT(NOW)+TIMER>XTIME OR 1=1

IF ERR.NUMBER>0 THEN 
	
	MSGBOX "1 MINUTE __ COPY FAIL"+VBCRLF+"TRY RETRY"+VBCRLF+PATH_FILE_NAME2

END IF

REM ELSE

	'------------------------------------------------------------------
	'NOT WORK BUT LOOK UP LATER ---------------------------------------
	'RUN AS ADMIN -----------------------------------------------------
	REM Set objShell = CreateObject("Shell.Application")
	REM objShell.ShellExecute "cscript", """"+VF2+"""", "", "runas", 
	REM objShell.ShellExecute "cscript", VF2, "", "runas", 

	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run """"+PATH_FILE_NAME1+""""
	Set objShell = Nothing
	
REM END IF

ON ERROR GOTO 0
