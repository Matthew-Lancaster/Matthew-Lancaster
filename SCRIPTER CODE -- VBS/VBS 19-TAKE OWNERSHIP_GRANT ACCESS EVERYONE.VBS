' -----------------------------------------------------------------
' CODE BY MATTHEW LANCASTER 
' Matt.Lan@btinternet.com
' DATE BEGINNER AND ENDER --------------------------------------
' -----------------------------------------------------------------
' Sun 29-Apr-2018 01:25:44 
' Sun 29-Apr-2018 03:07:21
' -----------------------------------------------------------------
' -----------------------------------------------------------------
' Been Programmer ICALCS For a Long Time
' I Regularly have to use some script were it runs on 
' timer all the time
' Just Made a Script in VBA But Wanted for Usability in 
' VBSCRIPT to Show Off
' I Wrote in VB6 VBA And Could of Left it as That
' Extra Work Involve -----------
' 1. Find App Folder or Temp One
' 2. Write Output to File
' 3. Get the Now Time
' 4. check file names converted to folder name method okay
' 5. Shell The Command
' -----------------------------------------------------------------
' Location __ Folder and File DropBox
' https://www.dropbox.com/sh/3t4jp1lmwgak0o6/AAATryQH9xhqYjHmtkxBqUHpa?dl=0
' https://www.dropbox.com/s/uh4edw2oh2pwa9i/VBS%2018-FFMPEG_VERIFY_VIDEO.VBS?dl=0
' -----------------------------------------------------------------
' Location __ Folder and File GoogleDrive
' https://drive.google.com/open?id=1slYVkA1WIfjSy_8JNeAp7kX6aAdbznLH
' https://drive.google.com/open?id=1hOSeGECsG-WbJBzAmN3PtNATvtFmPfwe
' -----------------------------------------------------------------
' Search Description
' VBScript -- Put In Sendto Folder and Run Script on Folder
' Will Use ICACLS in 2 Passes to Set Ownership Everyone So You Have 
' Access Everywhere
' Owership is Set and then Grant Access
' Setting the Wrapper to Load a Batch File to Do Work Means You 
' Get Some Output on Screen to Check Over Job Working
' -----------------------------------------------------------------

' ----------------------------------------------------------------
' __ MATTHEW LANCASTER 
' __ MATT.LAN@BTINTERNET.COM
' ----------------------------------------------------------------
' __ ALL BATCH FILE COMMAND SETS 
' __ GOT AN IMPROVED ADMINISTRATOR LAUNCHER
' __ THAT DOESN'T GET STUCK IN A LOOP IF NOT SUCCESSFUL
' __ AT THE SACRIFICE OF COMMAND LINE ARGUMENTS BUT EASY TO 
' __ ADD CODE IF REQUIRE
' ----------------------------------------------------------------
' __ FROM ---- Thu 03-May-2018 01:31:51
' __ TO   ---- Thu 03-May-2018 03:35:00 __ 2 HOUR
' ----------------------------------------------------------------


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
'-------------------------------------
Sub BEGIN

    Set args = Wscript.Arguments
    For Each arg In args
        VAR_2_FOLDER = VAR_2_FOLDER + arg
    Next
    
    'If Sent-to is Used Sometime Additional Quote Surround are User
    '--------------------------------------------------------------
    VAR_2_FOLDER = Replace(VAR_2_FOLDER,"""","")

    IF objFSO.FileExists(VAR_2_FOLDER)=True THEN 
		' FOLDER ONLY
		' -----------
		VAR_2_FOLDER = Mid(VAR_2_FOLDER, 1, InStrRev(VAR_2_FOLDER, "\")-1)
	End if
	
	IF WScript.Arguments.Count = 0 THEN 
        Wscript.Echo "Not Any Command-line Arguments Were Given _ Exit without Do"
        exit sub
    End if
    
    IF objFSO.FolderExists(VAR_2_FOLDER)=False THEN 
        Wscript.Echo VAR_2_FOLDER & vbcrlf & "Folder Name Given in Command line Arguments Is Not Exist _ Exit without Do"
        exit sub
    End if

	result=Msgbox(VAR_2_FOLDER & VBCRLF & "Ownership Change Ready __ Are you sure?",vbYesNo+vbInformation, "")
	if result = vbno then WScript.Quit 0
	if result = vbcancle then WScript.Quit 0
	
	strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
	TEMP_FOLDER____ = objShell.ExpandEnvironmentStrings("%Temp%")
	
	BATCH_FILE_ = TEMP_FOLDER____ + "\" + strComputerName + "_" + timeStamp + ".BAT"
	SCRIPT_FILE = TEMP_FOLDER____ + "\" + strComputerName + "_" + timeStamp + ".VBS"

	'---------------------------------------------------------
	'IF THE CALLING PROGRAM HAS GOT ADMIN RIGHT
	'IT NOT REALLY REQUIRED TO REQUEST AGAIN
	'---------------------------------------------------------
	'WOULD OF BEEN MORE PROFESSIONAL TO DO IT THAT WAY AROUND
	'---------------------------------------------------------
	'SEEING AS THE LOOK OF IT WERE NOT ANY ADMIN RIGHT REQUIRE 
	'FOR VBS CODE FIRST PART OF PROGRAM
	'THEN MISTAKE WAS THERE
	'---------------------------------------------------------

	T1 = T1 + "@ECHO OFF" + vbCrLf
	T1 = T1 + "IF '%1%'=='/ADMIN' (" + vbCrLf
	T1 = T1 + "	    @REM :: Check for ADMIN Privileges Again" + vbCrLf
	T1 = T1 + "     >nul 2>&1 ""%SYSTEMROOT%\system32\cacls.exe"" ""%SYSTEMROOT%\system32\config\system""" + vbCrLf
	T1 = T1 + "	    @IF '%errorlevel%' GTR '2'  SET NOT_GOT_ADMIN=TRUE" + vbCrLf
	T1 = T1 + ")" + vbCrLf
	
	T1 = T1 + "@IF ""%NOT_GOT_ADMIN%""==""TRUE"" (" + vbCrLf
	T1 = T1 + "    ECHO AN ATTEMPT WAS MADE TO SET ADMIN BUT NOT SUCCESSFUL" + vbCrLf
	T1 = T1 + "    ECHO YOUR USER MUST BE A STANDARD USER OR SOMETHING LIKE HER" + vbCrLf
	T1 = T1 + "    ECHO SAVES GET STUCK IN A LOOP ASKING OR ALSO DELAY REQUIRING" + vbCrLf
	T1 = T1 + "    ECHO END" + vbCrLf
	T1 = T1 + "    PAUSE" + vbCrLf
	T1 = T1 + "    GOTO ENDE" + vbCrLf
	T1 = T1 + "	   EXIT /B 1" + vbCrLf
	T1 = T1 + ")" + vbCrLf

	T1 = T1 + "@Rem :: Check for ADMIN Privileges" + vbCrLf
	T1 = T1 + ">nul 2>&1 ""%SYSTEMROOT%\system32\cacls.exe"" ""%SYSTEMROOT%\system32\config\system""" + vbCrLf
	T1 = T1 + "if '%errorlevel%' GTR '2' (" + vbCrLf
	T1 = T1 + "    @Rem Get ADMIN Privileges" + vbCrLf
	T1 = T1 + "    echo Set UAC = CreateObject^(""Shell.Application""^) > ""%temp%\getadmin.vbs""" + vbCrLf
	T1 = T1 + "    echo UAC.ShellExecute ""%~s0"", ""/ADMIN"", """", ""runas"", 1 >> ""%temp%\getadmin.vbs""" + vbCrLf
	T1 = T1 + "    ""%temp%\getadmin.vbs""" + vbCrLf
	T1 = T1 + "    del ""%temp%\getadmin.vbs""" + vbCrLf
	T1 = T1 + "    GOTO ENDE" + vbCrLf
	T1 = T1 + "    EXIT /B" + vbCrLf
	T1 = T1 + ")" + vbCrLf
	T1 = T1 + "ECHO Got ADMIN Privileges" + vbCrLf
	T1 = T1 + ":ENDE" + vbCrLf
	
	T1 = T1 + "@CD\" + vbCrLf
	T1 = T1 + "@ECHO ON" + vbCrLf
	T1 = T1 + "SET PATH_TO_DO=" + vbCrLf
	T1 = T1 + "Set PATH_TO_DO=""" + VAR_2_FOLDER + "\*""" + vbCrLf
	T1 = T1 + "@ECHO %PATH_TO_DO%" + vbCrLf
	T1 = T1 + "ICACLS %PATH_TO_DO% /setowner Everyone /T /C /L /Q" + vbCrLf
	T1 = T1 + "ICACLS %PATH_TO_DO% /T /L /C /Q /Grant Everyone:F" + vbCrLf
	T1 = T1 + "@ECHO." + vbCrLf
	T1 = T1 + "@ECHO OFF" + vbCrLf
	T1 = T1 + "PAUSE" + vbCrLf
	T1 = T1 + "CMD /K """ + SCRIPT_FILE + """" + vbCrLf
	T1 = T1 + ":VERY_ENDER_BLOCK" + vbCrLf
	T1 = T1 + "EXIT" + vbCrLf

	T2 = T2 + "Set obj = CreateObject(""Scripting.FileSystemObject"")" + vbCrLf
	T2 = T2 + "obj.DeleteFile(""" + BATCH_FILE_ + """)" + vbCrLf
	T2 = T2 + "obj.DeleteFile(""" + SCRIPT_FILE + """)" + vbCrLf

	
	Set MyFile = objFSO.OpenTextFile(BATCH_FILE_, ForWriting, True, TristateFalse)
    MyFile.WriteLine T1
    MyFile.Close
	Set MyFile = objFSO.OpenTextFile(SCRIPT_FILE, ForWriting, True, TristateFalse)
    MyFile.WriteLine T2
    MyFile.Close

	Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run "CMD /K """ + BATCH_FILE_ + """"
    Set WSHShell = Nothing

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