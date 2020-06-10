' ----------------------------------------------------------------
' __ MATTHEW LANCASTER 
' __ MATT.LAN@BTINTERNET.COM
' ----------------------------------------------------------------
' __ FROM ---- Thu 24-May-2018 03:13:07
' __ TO   ---- Thu 24-May-2018 05:24:00 __ 2 HOUR 11 MINUTE
' ----------------------------------------------------------------

'-----------------------------------------------------------------
' GET OUR PUBLIC EXTERNAL IP
' COMPARE IT TO LAST CHECKED ON
' IF CHANGED WRITE TO A LOGGER
'-----------------------------------------------------------------

' ----
' Windows command that returns external IP - Super User
' https://superuser.com/questions/165986/windows-command-that-returns-external-ip
' ----

Option Explicit

DIM VAR_ResponseText
Dim http : Set http = CreateObject( "MSXML2.ServerXmlHttp" )
http.Open "GET", "http://myip.dnsomatic.com", False
ON ERROR RESUME NEXT
http.Send
IF ERR.NUMBER>0 THEN
	WScript.Quit(0)
END IF
'Wscript.Echo http.responseText   'or do whatever you want with it
VAR_ResponseText=http.responseText
IF ERR.NUMBER>0 THEN
	WScript.Quit(0)
END IF
ON ERROR GOTO 0
Set http = Nothing

'2018-05-19--18:08:33__7-ASUS-GL522VW__192.168.1.107 

DIM objShell
DIM TempDir
DIM objFSO
DIM outFile
DIM objFile
DIM objFolder
DIM ComputerName
DIM BASE_FOLDER
DIM MY_IP_FILE

Set objShell = Wscript.CreateObject("WScript.Shell")

' TempDir = objShell.ExpandEnvironmentStrings("%tmp%")
' SET TEMPDIR = CreateObject("Scripting.FileSystemObject").GETSPECIALFOLDER(2)

' MSGBOX TEMPDIR
' WScript.Quit(0)

ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

DIM SET_GO_1
DIM SET_GO_2
DIM SET_GO_3
Dim FSO
DIM MyIPLine
DIM i_TIME
Set FSO = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8

BASE_FOLDER="C:\C DRIVE ROOT LOGGER"
If Not FSO.FolderExists(BASE_FOLDER) Then
	Set objFolder = FSO.CreateFolder(BASE_FOLDER)
	Set objFolder = NOTHING		
End If

MY_IP_FILE=BASE_FOLDER+"\IP_LOGGER_00_MY_IP_"+ComputerName+".txt"

IF FSO.FILEEXISTS(MY_IP_FILE)=TRUE THEN 
	Set objFile = fso.OpenTextFile(MY_IP_FILE, ForReading)
	Do Until objFile.AtEndOfStream
		MyIPLine = objFile.ReadLine
	Loop
	objFile.Close
END IF


SET_GO_1 = FALSE
SET_GO_2 = FALSE
SET_GO_3 = FALSE
IF VAR_ResponseText<>MyIPLine THEN SET_GO_1 = TRUE
IF VAR_ResponseText<>"" THEN       SET_GO_2 = TRUE
IF LEN(VAR_ResponseText)<18 THEN   SET_GO_3 = TRUE

IF SET_GO_1 = TRUE AND SET_GO_2 = TRUE AND SET_GO_3 = TRUE THEN 

	'WScript.Echo MyIPLine
	
	Set objFile = fso.OpenTextFile(MY_IP_FILE, ForWriting,TRUE)
	objFile.WriteLine VAR_ResponseText
	objFile.Close

	i_TIME=""
	i_TIME=i_TIME+LPad(Year(Now), "0", 4)+"-"
	i_TIME=i_TIME+LPad(Month(Now), "0", 2)+"-"
	i_TIME=i_TIME+LPad(Day(Now), "0", 2)+"-"
	i_TIME=i_TIME+"-"
	i_TIME=i_TIME+LPad(Hour(Now), "0", 2)+"-"
	i_TIME=i_TIME+LPad(Minute(Now), "0", 2)+"-"
	i_TIME=i_TIME+LPad(Second(Now), "0", 2)
	
	Set objFile = fso.OpenTextFile(BASE_FOLDER+"\IP_LOGGER_00_"+ComputerName+".txt", ForAppending,TRUE)
	objFile.WriteLine i_TIME+"__"+ComputerName+"__"+VAR_ResponseText
	objFile.Close
	
	i_TIME=Wscript.ScriptFullName+vbcrlf+vbcrlf+"IP Address Has Changed"+vbcrlf+vbcrlf+i_TIME
	WScript.Echo i_TIME+"__"+ComputerName+"__"+VAR_ResponseText
END IF


Set FSO= Nothing
WScript.Quit(0)

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

' Year(Now)   = 2016
' Month(Now)  = 2
' Day(Now)    = 29
' Hour(Now)   = 13
' Minute(Now) = 2
' Second(Now) = 3
' LPad(Month(Date), "0", 2) 

' ----
' windows - vbscript output to console - Stack Overflow
' https://stackoverflow.com/questions/4388879/vbscript-output-to-console
' ----
DIM oWSH,vbsInterpreter

Set oWSH = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"

Call ForceConsole()
 
printf VAR_ResponseText

Function printf(txt)
WScript.StdOut.WriteLine txt
End Function

Function printl(txt)
WScript.StdOut.Write txt
End Function

Function scanf()
scanf = LCase(WScript.StdIn.ReadLine)
End Function

Function wait(n)
WScript.Sleep Int(n * 1000)
End Function

Function ForceConsole()
If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
	oWSH.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34) 
	WScript.Quit
End If
End Function

