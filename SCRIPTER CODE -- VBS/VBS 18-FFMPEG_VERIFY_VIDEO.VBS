' -----------------------------------------------------------------
' __ CODE BY MATTHEW LANCASTER __ Matt.Lan@btinternet.com
' __ DATE BEGINNER AND ENDER --------------------------------------
' Tue 17-Apr-2018 21:00:21 
' Wed 18-Apr-2018 00:53:40 4 Hour-
' Wed 18-Apr-2018 01:40:00 6 Hour Publishing Time on Google Blogger
' Originally Coded In  VBasic and then for VBScript
' -----------------------------------------------------------------
' -----------------------------------------------------------------
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
' The Batch File So You Can See It Is Created in the Folder Of Passed 
' to Video the Batch File Delete itself After for Clean Up
' Leaving a Minor Error Messenger Behind It Can't Find Itself
' I Did Try and Pass the Command Line to Script Handler 
' To Save Creating a File But It was Too Hard Because of the 
' Pipe Arrows > >
' -----------------------------------------------------------------
' the FFmpeg Program is Available On-Line and I Included My Version 
' in a Folder with Scripts
' My Version is Quite older Now and Soon Be Upgrade
' Guess Still works the Same
' The FFMpeg Program Does a Lot More
' It is a Command line Driven Mpeg Modifier 
' The Help Screen is Not Much Help to Look at At
' Guess you can do all the Usual Things FFmpeging
' Like Split and Save Video to JPG Frames Maybe
' -----------------------------------------------------------------
' Zeranoe's FFmpeg Builds Home Page: 
' <http://ffmpeg.zeranoe.com/builds/>
' Version of My Use 2015-07-01 __ git-9c010ba
' -----------------------------------------------------------------
' Variable Names Could Do with a Wordin Bit Sloppy 
' That One From Older Code
' -----------------------------------------------------------------
' Location Mine __ Folder and File DropBox
' https://www.dropbox.com/sh/3t4jp1lmwgak0o6/AAATryQH9xhqYjHmtkxBqUHpa?dl=0
' https://www.dropbox.com/s/uh4edw2oh2pwa9i/VBS%2018-FFMPEG_VERIFY_VIDEO.VBS?dl=0
' -----------------------------------------------------------------
' Location Mine __ Folder and File GoogleDrive
' https://drive.google.com/open?id=1slYVkA1WIfjSy_8JNeAp7kX6aAdbznLH
' https://drive.google.com/open?id=1hOSeGECsG-WbJBzAmN3PtNATvtFmPfwe
' -----------------------------------------------------------------
' Search Description
' VBScript -- Got Two Video File The Same But Checksum But Don't Match Use This Code Wrapper For FFMPEG Tool To Find The Better Version
' -----------------------------------------------------------------

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
'-------------------------------------
    Set args = Wscript.Arguments
    For Each arg In args
        XF0 = XF0 + arg
    Next
    
    'If Sent-to is Used Sometime Additional Quote Surround are User
    '--------------------------------------------------------------
    XF0 = Replace(XF0,"""","")
    
    IF WScript.Arguments.Count = 0 THEN 
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
    
    DH = """C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 18-ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"" -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"
    
    '-----------------------------
    'TRY TO RUN SCRIPTING -- WRONG DEBUG IT WORK TO DO
    'DH = "CMD /K """"C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 18-ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"""" -v error -i """"" + XF0 + """""" + " -f null " + """""" + XF2 + """"""
    '-----------------------------
        
    TEMP = XF4
    Set MyFile = objFSO.OpenTextFile(TEMP, ForWriting, True, TristateFalse)
    MyFile.WriteLine "@CD\"
    MyFile.WriteLine DH
    MyFile.WriteLine "@ECHO."
    MyFile.WriteLine "@ECHO __ SOUND TO SIGNAL FINISH"
    MyFile.WriteLine "@rundll32 user32.dll,MessageBeep"
    MyFile.WriteLine "@ECHO."
    MyFile.WriteLine "@ECHO __ BATCH FILE DELETES ITSELF AFTER __ CLEANER"
    MyFile.WriteLine "@ECHO __ MINOR ERROR THROWN"
    MyFile.WriteLine "@ECHO."
    MyFile.WriteLine "DEL """ + TEMP + """"
    MyFile.Close
        
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run "CMD /K """ + TEMP + """"
    Set WSHShell = Nothing
'-------------------------------------
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