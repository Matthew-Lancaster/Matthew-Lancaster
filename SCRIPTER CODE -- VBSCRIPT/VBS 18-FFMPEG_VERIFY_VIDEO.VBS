'#-------------------------------------------------------------------
' __ 
' __ CODE BY MATTHEW LANCASTER
' __ Matt.Lan@btinternet.com 
' __ 
'#-------------------------------------------------------------------
' Tue 17-Apr-2018 21:00:21 
' Wed 18-Apr-2018 00:53:40 _ 4 Hour-
' Wed 18-Apr-2018 01:40:00 _ 6 Hour Publishing Time on Google Blogger
' -------------------------------------------------------------------
' Tue 09-Oct-2018 15:20:35
' Tue 09-Oct-2018 18:20:00 _ 3 Hour Another Wave Session 002 Completer
'#-------------------------------------------------------------------
'
'#-------------------------------------------------------------------
' Code Description
' VBScript
' -------------------------------------------------------------------
' A Wrapper Code For a Program Called FFMPEG.EXE _ By _ Kyle Schwarz
'
' Has Complex Single Command Line & Useful When Required Worker Out 
' & Make Into Tool Code Wrapper
'
' Usage Put a Shortcut To This File in SendTo Folder & Right Click on 
' Any Video File 
'
' Got Video File and Want to Check It Integrity for Any Error in Frame or 
' Whatever to See if Perfect or Not 
' And If Damaged How Much Is Damage
' 
' Got Two Video File The Same But Checksum But Don't Match Use 
' This Code Wrapper For FFMPEG Tool To Find The Better Version
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 002
' -------------------------------------------------------------------
' TIDY CODE MAKE PRESENTABLE FOR SHARING ON-LINE
' TEST WITH LATEST VERSION FFMPEG.EXE
' AND SHOW WHERE TO FIND A COPY WHICH MUST BE USED WITH THE CODE IS 
' ONLY A WRAPPER FOR IT
' AS IT HAS A COMPLEX COMMAND LINE TO GET RIGHT WORKER
' PUT IN SENDTO FOLDER AND USE RIGHT CLICK IN EXPLORER TO SEND VIDEO TO PROCESSOR
'
' TIDY CODE UP TO DATE TO USE LATEST VERSION
' CLEAR AND ERROR SHOWING BATCH FILE AT END DELETE ITSELF AND THROWS AN 
' ANNOYING NOT CLEAR ERROR NOT WANTED
' 
' TIDY SEPARATE LONG LINE CODE SMALLER TO ENTER LATEST VERSION OF FFMPEG.EXE MORE EASILY-ER
'
' MAKE OUTPUT RESULT OF RUN MORE CLEARER
'
' QUICKLY ADD BATCH FILE CREATED TO QUESTION ASK IF WANT TO OPEN 
' CREATED OUTPUT PROCESS RESULT IN NOTEPAD++
'
' ALSO ADD NOT REQUIRED TO OPEN FILE LOOK AT RESULT IF EMPTY 
' & DELETE EMPTY RESULT FILE
'
' -------------------------------------------------------------------
' Tue 09-Oct-2018 15:20:35
' Tue 09-Oct-2018 18:20:00 3 Hour _ Update Make Sharable Proper
'                                   & Make Work With Latest Version
'#-------------------------------------------------------------------

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Idea Come Up for Code
' I Previously Coded It Before to Recursively Scan Sub Folder
' of Video
' I Use the FFMpeg and Support all Format WMV MPG AVI
' It Can Scan a Video File and Tell any Error in the Video Stream
' So You Learn about any Quality Loss
' Here I Converted it For Single USE It Does take a Time to Scan a 
' Whole Long Video Few Minute
' I Had Problem where two checksums for the same video were different
' And I Wanted to Find Out What was the Best One 
' Which Works Itself Well
' Usage for End User 
' Is Put the Code.VBS File in the Send-to Folder Or Your Menu Managed 
' One
' Do a Right Click on a Video File and See the Answer a Dos Box 
' Comes Up to Show when Finished
' It Has Been Created So You Can Run Two Multi-Tasking
'
' The Batch File So You Can See It Is Created in the Folder Of Passed 
' to Video the Batch File Delete itself After for Clean Up
' Leaving a Minor Error Messenger Behind It Can't Find Itself
' I Did Try and Pass the Command Line to Script Handler 
' To Save Creating a File But It was Too Hard Because of the 
' Pipe Arrows > >
' Got Over Problem Now By Redirect Error Output
' -------------------------------------------------------------------
' The FFmpeg Program is Available On-Line and I Included My Version 
' in a Folder with Scripts
' My Version is Quite older Now and Soon Be Upgrade
' Guess Still works the Same
' The FFMpeg Program Does a Lot More
' It is a Command line Driven Mpeg Modifier 
' The Help Screen is Not Much Help to Look at At
' Guess you can do all the Usual Things FFmpeging
' Like Split and Save Video to JPG Frames Maybe
' -------------------------------------------------------------------
' Zeranoe's FFmpeg Builds Home Page: 
' http://ffmpeg.zeranoe.com/builds/
' Version of My Use 2015-07-01 __ git-9c010ba
' Now Updated to Current 2018 Version
' -------------------------------------------------------------------
' Location Mine __ Folder and File GITHUB
' ----
' Matthew-Lancaster/VBS 18-FFMPEG_VERIFY_VIDEO.VBS at master · Matthew-Lancaster/Matthew-Lancaster
' https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20VBS/VBS%2018-FFMPEG_VERIFY_VIDEO.VBS
' ----
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' Originally Coded In VBasic6 and then for VBScript
' Attempted Batch Command Line Version From VB6 But Gave Up Getting 
' Command Line Right was Awkward
' -------------------------------------------------------------------

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

DIM objFSO
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Wscript.Echo "begin."
CALL BEGIN
'Wscript.Echo "end."
WScript.Quit 0
' -------------------------------------------------------------------
Sub BEGIN
	' ---------------------------------------------------------------
	Set args = Wscript.Arguments
	For Each arg In args
		XF0 = XF0 + arg
	Next

	' ---------------------------------------------------------------
	' TO USE PUT A SHORTCUT TO THIS PROGRAM .VBS IN _ SENDTO _ FOLDER 
	' AND RIGHT CLICK A VIDEO FILE TO HER
	' ---------------------------------------------------------------

	' ---------------------------------------------------------------
	' If Sent-to is Used Sometime Additional Quote Surround are User
	' ESPECIALLY IF CONTEXT MENU HAS BEEN ADDED IN REGISTRY 
	' ---------------------------------------------------------------
	XF0 = Replace(XF0,"""","")

	' ---------------------------------------------------------------
	' USE A HARDCODE-ED PATH IF WANT TO
	' ---------------------------------------------------------------
	' XF0="D:\DSC\2015-2018\2018 CyberShot HX60V\MP_ROOT\2018_08_30 Aug_Thu 20_25_49__MAQ00504.MP4"
	' ---------------------------------------------------------------

	IF XF0="" AND WScript.Arguments.Count = 0 THEN 
		Wscript.Echo "Not Any Command-line Arguments Were Given _ Exit without Do"
		exit sub
	End if

	IF objFSO.FileExists(XF0)=False THEN 
		Wscript.Echo XF0 & vbcrlf & "File Name Given in Command line Arguments Is Not Exist _ Exit without Do"
		exit sub
	End if

	XF1 = XF0 + ".txt"
	XF2 = Replace(XF1, ".txt", ".FFmpeg-Verify.txt")
	XF3 = Mid(XF0, 1, InStrRev(XF0, ".") - 1)
	XF3 = Mid(XF3, InStrRev(XF3, "\") + 1)
	XF4 = Mid(XF0, 1, InStrRev(XF0, "\")) + XF3 + "__FFMPEG_RUN_BATCH_" + timeStamp + ".BAT"

	' ---------------------------------------------------------------
	' POINT TO PATH WHERE YOUR GONNA USE __ FFMPEG.EXE 
	' COMMAND LINE VERSION
	' LATEST VERSION WORKER GOOD
	' HERE LOCATION
	' ----
	' Builds - Zeranoe FFmpeg
	' https://ffmpeg.zeranoe.com/builds/
	' ----
	' ---------------------------------------------------------------
	' DH = """C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 18-ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"" -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"
	' DH = """C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 18-ffmpeg-20181007-0a41a8b-win64-static\bin\ffmpeg.exe"" -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"
	' DH = """C:\PStart\# NOT INSTALL REQUIRED\ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"" -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"

	' ---------------------------------------------------------------
	' UP-TO-DATE VERSION 
	' ---------------------------------------------------------------
	' PATH_FFMPEG="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 18-ffmpeg-20181007-0a41a8b-win64-static\bin\ffmpeg.exe"
	
	' ---------------------------------------------------------------
	' SAME AS HERE COPIER
	' ---------------------------------------------------------------
	' C:\PStart\# NOT INSTALL REQUIRED\FFMPEG\ffmpeg-4.1.3-win64-static\bin\ffmpeg.exe
	' ---------------------------------------------------------------
	PATH_FFMPEG="C:\PStart\# NOT INSTALL REQUIRED\FFMPEG\ffmpeg.exe"
	' ---------------------------------------------------------------
	DH = """" + PATH_FFMPEG + """ -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"

	'-----------------------------
	'TRY TO RUN SCRIPTING FROM COMMAND LINE -- WRONG DEBUG IT WORK TO DO
	'DH = "CMD /K """"C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 18-ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"""" -v error -i """"" + XF0 + """""" + " -f null " + """""" + XF2 + """"""
	'-----------------------------
		
	TEMP = XF4
	Set MyFile = objFSO.OpenTextFile(TEMP, ForWriting, True, TristateFalse)
	MyFile.WriteLine "@CD\"
	MyFile.WriteLine DH
	MyFile.WriteLine "@ECHO."
	MyFile.WriteLine "@ECHO __ "
	MyFile.WriteLine "@ECHO __ SYSTEM AUDIO SOUND TO SIGNAL THE FINISH - FOR BIG VIDEO FILE"
	MyFile.WriteLine "@rundll32 user32.dll,MessageBeep"
	MyFile.WriteLine "@ECHO __ "
	' MyFile.WriteLine "@ECHO __ BATCH FILE DELETES ITSELF AFTER __ CLEANER"
	' MyFile.WriteLine "@ECHO __ ERROR ABOUT BATCH FILE NOT FOUND ENDING IS NORMAL DELETER HAPPEN"
	MyFile.WriteLine "@ECHO __ "
	MyFile.WriteLine "@ECHO __ EMPTY OUTPUT RESULT FILE INDICATE NONE ERROR IN VIDEO FILE"
	MyFile.WriteLine "@ECHO __ ERROR'S WOULD NORMALLY BE SHOWN IN GREAT DETAIL"
	MyFile.WriteLine "@ECHO __ "
	MyFile.WriteLine "@ECHO __ OUTPUT RESULT FILE WRITTEN IN SAME CURRENT DIRECTORY AS"
	MyFile.WriteLine "@ECHO __ PROCESSED VIDEO FILE UNDER EXTENSION .TXT _ HERE LOCATION"
	MyFile.WriteLine "@ECHO __ "
	MyFile.WriteLine "@ECHO __ " + XF2 
	MyFile.WriteLine "@ECHO __ "
	MyFile.WriteLine "@ECHO __ COMPLETER ___"
	MyFile.WriteLine "@ECHO __ "

	MyFile.WriteLine "@SET FILE_CONTAIN_DATA="

	' ---------------------------------------------------------------
	' Test Debug - Simulate Was an Error in Report to Test Ender Code
	' And Checker Code for Empty File Return
	' ---------------------------------------------------------------
	' MyFile.WriteLine "@ECHO TEST >>""" + XF2 +""""

	MyFile.WriteLine "@ECHO OFF"
	MyFile.WriteLine "(for /f usebackq^ eol^= %%a in ("""+XF2+""") do break) && echo __ OutPut File For Video Has Process Result Error Data || SET FILE_CONTAIN_DATA=FALSE"
	MyFile.WriteLine "@ECHO ON"

	MyFile.WriteLine "@IF ""%FILE_CONTAIN_DATA%"" == ""FALSE"" GOTO FILE_NOT_REQUIRED_AS_EMPTY"
	MyFile.WriteLine "@IF ""%FILE_CONTAIN_DATA%"" == """" GOTO FILE_HAS_DATA_NEXT_STEPPER"
	MyFile.WriteLine ":FILE_NOT_REQUIRED_AS_EMPTY"
	MyFile.WriteLine "@ECHO __ GOOD RESULT SHOWN BY PROCESS RESULT FILE BEING EMPTY"
	MyFile.WriteLine "@ECHO __ EMPTY RESULT OUTPUT FILE TEXT NOT ANY LONGER NEEDED AND WILL BE DELETER"
	MyFile.WriteLine "@ECHO __"
	MyFile.WriteLine "@ECHO __ GOOD RESULT PROCESS OF VIDEO FILE HAS NONE ERROR"
	MyFile.WriteLine "@ECHO __"
	MyFile.WriteLine "@DEL """ + XF2 + """"
	MyFile.WriteLine "@GOTO ENDE"
	MyFile.WriteLine ":FILE_HAS_DATA_NEXT_STEPPER"

	MyFile.WriteLine "@ECHO __ "
	MyFile.WriteLine "@ECHO __ DO YOU WANT OPEN IN NOTEPAD++ __  ?(Y/N)"
	MyFile.WriteLine "@SET INPUT="
	MyFile.WriteLine "@SET /P INPUT=Type input: %=%"
	MyFile.WriteLine "@IF ""%INPUT%"" == ""y"" GOTO YES"
	MyFile.WriteLine "@IF ""%INPUT%"" == ""Y"" GOTO YES"
	MyFile.WriteLine "@GOTO NOT"
	MyFile.WriteLine ":YES"
	MyFile.WriteLine "@START """" /MAX /HIGH ""C:\Program Files (x86)\Notepad++\notepad++.exe""" + " """+XF2+""""
	MyFile.WriteLine ":NOT"
	MyFile.WriteLine ":ENDE"

	MyFile.WriteLine "@DEL """ + TEMP + """"
	MyFile.Close

	' ---------------------------------------------------------------
	' Credit Due for Source Empty File in Batch Command
	' ---------------------------------------------------------------
	' ----
	' BATCH FILE CHECK FILE SIZE IS EMPTY - Google Search
	' https://www.google.co.uk/search?q=BATCH+FILE+CHECK+FILE+SIZE+IS+EMPTY&spell=1&sa=X&ved=0ahUKEwi2rYHE5vndAhWJKMAKHTt3Cw8QBQgrKAA&biw=1536&bih=551&dpr=1.25
	' --------
	' windows batch: how to check if a text file is blank (the file has an empty line so size is non zero) - Stack Overflow
	' https://stackoverflow.com/questions/33575962/windows-batch-how-to-check-if-a-text-file-is-blank-the-file-has-an-empty-line
	' ----
	' ---------------------------------------------------------------

	Dim WSHShell
	Set WSHShell = CreateObject("WScript.Shell")
	WSHShell.Run "CMD /K """ + TEMP + """ 2>NUL"
	Set WSHShell = Nothing

	' -----------------------------------------------------------------------------------
END SUB

Function timeStamp()
    Dim t 
    t = Now
    timeStamp = Year(t) & "-" & _
    Right("0" & Month(t),2)  & "-" & _
    Right("0" & Day(t),2)  & "_" & _  
    Right("0" & Hour(t),2) & _
    Right("0" & Minute(t),2) '    '& _    Right("0" & Second(t),2) 
End Function