-- Thu 03-May-18 00:06:05 --


-----


D:\VB6\VB-NT\#1 Documents\Programming Snippets.txt.bas
---------------------
Count = 011 -- Tue 10-May-2016 00:18:31
---------------------
Make Windows synchronize time more often - Google Chrome
---------------------
How to do it: A different scheduled task

The W32tm.exe tool, with /resync argument, triggers a clock sync, provided the Windows Time service is running. (If you want to test it manually, note that the command requires administrator rights and that the service may need to be started.)
---------------------
---------------------
Count = 012 -- Tue 10-May-2016 00:18:36
---------------------
Make Windows synchronize time more often - Google Chrome
---------------------
----
Make Windows synchronize time more often
http://www.pretentiousname.com/timesync/
----
---------------------



-- Sun 01-May-16 15:00:33 --

---------------------------------------
HERE FOR WINDOWS 10 CD ROM BURN DISABLE
---------------------------------------

C:\Users\Administrator\AppData\Local\Microsoft\Windows\Burn
or 
C:\Users*[your user name]*\AppData\Local\Microsoft\Windows\Burn
#2

----
Revisions to Why is desktop.ini "Ready to Be Written to the Disc" of any CD or DVD inserted under Windows 7? - Super User
http://superuser.com/posts/341207/revisions
----
1	
source  link
answered Sep 29 '11 at 12:29

Bite-me
291
Right click on the following folder: C:\Users\Administrator\AppData\Local\Microsoft\Win dows\Burn or C:\Users[your user name]\AppData\Local\Microsoft\Windows\Burn

Properties > Security TAB > Edit > Click on any user name > CHECK: Deny (to the right of Full Control) > OK > OK

Desktop.ini no longer shows up, I don't know what other consequences this will have other than maybe not directly burning to the disc, but most of us use third party burning software anyway.


#01 OF 03
POST ABOVE
----------
----
Revisions to Why is desktop.ini "Ready to Be Written to the Disc" of any CD or DVD inserted under Windows 7? - Super User
http://superuser.com/posts/341207/revisions
----

#02 OF 03
POST SOURCE FIND
----------------
----
dvd burning - Why is desktop.ini "Ready to Be Written to the Disc" of any CD or DVD inserted under Windows 7? - Super User
http://superuser.com/questions/275797/why-is-desktop-ini-ready-to-be-written-to-the-disc-of-any-cd-or-dvd-inserted-u
----

#03 OF 03
URL SEARCH TERM
---------------
----
DVD DRIVE WINDOWS 10 - DON'T WANT FILES TO BE WRITTEN TO DISK DESKTOP.INI - Google Search
https://www.google.co.uk/search?num=100&q=DVD+DRIVE+WINDOWS+10+-+DON%27T+WANT+FILES+TO+BE+WRITTEN+TO+DISK+DESKTOP.INI&spell=1&sa=X&ved=0ahUKEwj71IK6g7nMAhVIAsAKHffwBocQvwUIGigA&biw=1138&bih=495
----
-------- Sun 01-May-2016 15:08:29 ---------

-----
LOCK-UP GARAGE TO LET MILE OAK
£65 A MONTH
DEPOSIT REQUIRED £200
TELEPHONE BRIGHTON 01273419914



-02:15 12/08/2015

Option Explicit
Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type


Public Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type



Public Function GetShortName(ByVal sLongFileName As String) As String
' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/kb/q175512/
' ---> The original comments were by them :

       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
End Function

Public Function GetLongName(ByVal sShortName As String) As String
' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/default.aspx?scid=kb;EN-US;154822
' ---> The original comments were by them :
     Dim sLongName As String
     Dim sTemp As String
     Dim iSlashPos As Integer

     'Add \ to short name to prevent Instr from failing
     sShortName = sShortName & "\"

     'Start from 4 to ignore the "[Drive Letter]:\" characters
     iSlashPos = InStr(4, sShortName, "\")

     'Pull out each string between \ character for conversion
     While iSlashPos
       sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
       If sTemp = "" Then
         'Error 52 - Bad File Name or Number
         GetLongName = ""
         Exit Function
       End If
       sLongName = sLongName & "\" & sTemp
       iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
     Wend

     'Prefix with the drive letter
     GetLongName = Left$(sShortName, 2) & sLongName

   End Function







-


        Key = "PATH"
        'ScanPath.ListView1.ColumnHeaders , Item(Key), , 4100
        Set Item = ScanPath.ListView1.ColumnHeaders.Item(Key)
        Item.Width = 6000
        
        







-- Fri 19-Aug-11 16:03:15 --
1 GIG
1024 ^ 3
MEG
1024 ^ 2

-----
-- Tue 14-Jun-11 00:40:49 --
Hi Madam/Sir

Thank You 

Your Most Generous to Look
After My Purchase

This Item Had Been Checked Many Times At the Post Office

I was Giving All Pleasure For Your Service to Provide This

The Address Is Correct I May Change Room Word to Flat Instead From Now On 

Your Hardship to Return a Second is Grateful and Regretful

Yours Sincerely
Matthiaoso

-----
-- Tue 19-Apr-11 10:15:09 --
TT$ = "<abbr title=""Night Temperature in degrees"
w61 = InStr(w6, IP2$, TT$)
w62 = InStrRev(IP2$, ">", w61)
nighttemp$ = Mid$(IP2$, w62 + 1, w61 - w62 - 1) + "°C"


-----
-- Tue 23-Nov-10 04:08:50 --


Private Sub PATTERNTimer_Timer()
    
On Error GoTo ENDSUB
    
    MODTIMER = 20
    OTFX = Int(Timer) Mod MODTIMER
    
    If OTFX < MODTIMER / 2 And OTFX <> OTF Then
        OTF = Int(Timer) Mod 20
        SWAPFLAG = SWAPFLAG + 1
        If SWAPFLAG > 3 Then SWAPFLAG = 1
    End If

    'SWAPFLAG = 1
    'SWAPFLAG = 2
    'SWAPFLAG = 3

    RRCOZ = RRCOZ + 0.001
    RRVAR = 0.1
    
    If SWAPFLAG <> 2 Then RR = RR + RRVAR
    If SWAPFLAG = 2 Then RR = RR + RRVAR + (Cos(RRCOZ) / 3.5)
    
    OFFSETX = 8
    OFFSETY = 0
    
    TX = (Picture2.Width / 2) + OFFSETX
    TY = (Picture2.Height / 2)
    
    OM = 78 'Int(Picture2.Width / Screen.TwipsPerPixelX)
    OMX = 30
    If SWAPFLAG = 3 Or 8 = 8 Then
        RR3 = RR3 + 0.02
        OMX = (OM / 2.2) * Sin(RR3) + (OM / 2)
    End If
    
    'HALF EXAMPLE
    '------------------------
    '40 OUTER
    '20 SIZE
    'OUTER - SIZE / 2 = 30
    '=30 CIRCUMFERANCE
    
    'SIZE 20 IN 30 CIRCUMFERANCE
    'LOWER CIRCUMFERANCE BY HALF SIZE
    'TEST 30 - 10 = 20      PASS 2 = 20 +20
    
    'LOWER HALF EXAMPLE
    '------------------------
    '40 OUTER
    '10 SIZE
    'OUTER - SIZE / 2 = 15
    
    '1
    '--- MODIFY = CIRCUMFERANCE RAIL TO BE DOUBLE
    
    'RESULT ARRIVED AT
            
    'HIGHER
    'TOP
    'UPPER
            
    'TOP HALF EXAMPLE
    '------------------------
            
    
    
    
    TXD = OMX
    
    REMIANING_DIVEDE_DIAMETER = (OM - OMX) / 2
    Ri = REMIANING_DIVEDE_DIAMETER
    
    If OMX < OM Then
        Ri = REMIANING_DIVEDE_DIAMETER * 2
    Else
        Ri = REMIANING_DIVEDE_DIAMETER / 2
    End If
    REMIANING_DIVEDE_DIAMETER = Ri
    
    
    x = (Ri * Sin(RR) + TX)
    y = (Ri * Cos(RR) + TY)
    
    If SWAPFLAG = 3 Then
        y = (Ri * Cos(RR * 1.2) + TY)
    End If
    'Me.Picture2.Width =
    
    Me.Picture2.Cls
    Me.Picture2.DrawWidth = 3
    Me.Picture2.Circle (TX + OFFSETX, TY + OFFSETY), OM
    Me.Picture2.Circle (x + OFFSETX, y + OFFSETY), OMX

Exit Sub
ENDSUB:
    
    Me.Hide
    Stop
    Resume
    
End Sub



-----
-- Sat 20-Nov-10 22:43:15 --
Visual Basic Screen Saver

Author: 	Yannick Houbrix
Category: 	Screen/Graphics
Type: 	Snippets
Difficulty: 	Intermediate

Version Compatibility: 	 Visual Basic 6  

More information: Creating a screen saver in Visual Basic is very simple. You create one form and set the following properties:


Caption ""
ControlBox False
MinButton False
MaxButton False
BorderStyle 0 - None
WindowState 2 - Maximized
BackColor &H0

You then compile it to an .exe, change the extension to .scr and copy the file to the windows directory.

This package does all of this; in addition, it includes a settings dialog so you can configure the screen saver's look and feel before compiling. A sample .scr file is included in the package, but you can work with this code to make any type of screen saver you would like.

This code has been viewed 105909 times.

Instructions: Click the link below to download the code. Select 'Save' from the IE popup dialog. Once downloaded, open the .zip file from your local drive using WinZip or a comparable program to view the contents.

Code/DancingLine.zip 

-----
-- Thu 26-Aug-10 12:55:50 --
---------------------
Count = 035 -- Wed 25-Aug-2010 18:25:39
---------------------
In Internet Explorer -- 
WinTitle = http://www.devx.com/vb2themax/Tip/19033
URL = GetScreenSaverState, SetScreenSaverState - Enable or disable the screen saver
---------------------
GetScreenSaverState, SetScreenSaverState - Enable or disable the screen saver - Windows Internet Explorer
---------------------
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPIF_UPDATEINIFILE = &H1

Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

' return the Enabled state of the screen saver

Function GetScreenSaverState() As Boolean
    Dim result As Long
    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, result, 0
    GetScreenSaverState = (result <> 0)
End Function

' enable or disable the screen saver
'
' if second argument is true, it writes changes in user's profile
' returns True if the operation was successful, False otherwise

Function SetScreenSaverState(ByVal enabled As Boolean, _
    Optional ByVal PermanentChange As Boolean) As Boolean
    Dim fuWinIni As Long
    If PermanentChange Then
        fuWinIni = SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    End If
    SetScreenSaverState = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, enabled, _
        ByVal 0&, fuWinIni) <> 0
End Function



-----

    02:28 02 July 2010
    '-----------
    'CODE TO DETECT CHANGE OF DIR FROM SEARCH THEN ACT ON THAT
    'EXAMPLE HERE IS TO CHANGE SPEED OF SLIDESHOW THEN
    'RECALL LAST SPEED BACK WHEN NO LONGER IN DIR OR CHANGED
    'IT CAN DETECT A CHANGE ON CONSECUTIVE CHANGE TO ANOTHER NEXT
    'IN LINE ON ANOTHER SEARCH
    
    Dim DSG(2), ML
    DSG(1) = "0-ButterFly"
    DSG(2) = "AutoPix 0001-Cartoon"
    SliderOverRide2 = 0
    For ML = 1 To 2
        If InStr(A1 + B1, DSG(ML)) > 0 And OldML <> ML Then
    
            SliderOverRide1 = Slider1.Value
            If ML = 1 Then Slider1.Value = 10
            If ML = 2 Then Slider1.Value = 5
            OldML = ML
        End If
    
        If InStr(A1 + B1, DSG(ML)) > 0 Then
            SliderOverRide2 = ML
        End If
    Next
    
    If SliderOverRide2 <> OldML Then
        Slider1.Value = SliderOverRide1
        OldML = SliderOverRide2
    End If

    '-----------


--- Wed 10-Mar-10 10:25:31 --
NET USE \\55-88-HAPPY\IPC$ /U:ASUS\MATT_01 matto28

-----
-- Tue 23-Feb-10 16:25:48 --
---- Tue 23-Feb-2010 04:08:28p ----  ----
Get all Files a Folder and its SubFolders using Wildcards
Http://www.freevbcode.com/ShowCode.Asp?ID=5510



Public Function Get_All_Directory_Files_With_Wildcard_1ST_MP3_BIGGER_ZERO_SIZE(ByVal tfolder As String, _
 ByVal getsubdirs As Boolean, _
 ByVal wildcard As String) _
 As String

'made by Alexander Triantafyllou
'alextriantf@yahoo.gr
'Athens-Greece

'FILE AND FOLDER TYPE REQUIRES THIS
'Reference=*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#H:\WINNT\system32\scrrun.dll#Microsoft Scripting Runtime

Set fso = CreateObject("Scripting.FileSystemObject")
Dim objfile As File
Dim objfolder As Folder
Dim kokovar As Variant
Dim k As Long
Dim wildext As String
Dim wildexts As String
Dim wildfirst As String
Dim wildexte As String
Dim wildfirsts As String
Dim wildfirste As String
Dim examfirst As String
Dim examext As String
Dim afl_filetext As String

Kokovar = Split(wildcard, ",")

    If tfolder <> "" Then


        For Each objfile In fso.GetFolder(tfolder).Files
            'do the stuff we want with the files
            For k = 0 To UBound(kokovar)
         wildext = LCase(cutgetExtension(kokovar(k)))
         wildfirst = LCase(Mid(kokovar(k), 1, Len(kokovar(k)) - Len(wildext) - 1))
            
       If InStr(1, wildext, "*") = 0 Then
       wildexts = "888NONE888"
       wildexte = "888NONE888"
       Else
       wildexts = Mid(wildext, 1, InStr(1, wildext, "*") - 1)
       wildexte = Mid(wildext, InStr(1, wildext, "*") + 1, Len(wildext) - InStr(1, wildext, "*"))
       End If
       
        If InStr(1, wildfirst, "*") = 0 Then
       wildfirsts = "888NONE888"
       wildfirste = "888NONE888"
       Else
       wildfirsts = Mid(wildfirst, 1, InStr(1, wildfirst, "*") - 1)
       wildfirste = Mid(wildfirst, InStr(1, wildfirst, "*") + 1, Len(wildfirst) - InStr(1, wildfirst, "*"))
       End If
            
        examfirst = LCase(cutgetName(cutfilename(CStr(objfile))))
        examext = LCase(cutgetExtension(CStr(objfile)))
            
        If wildexts = "888NONE888" Then
'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do not have a wildcard neither on the beggining or the

'extension
        If examfirst = wildfirst And examext = wildext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
                
        Else
        
        
        '###############
        'we do have a wildcard in the beggining but not in
        'the extension
        If Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste And wildext = examext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
        
        End If
        
        Else
        'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do have a wildcard in the beggining but not in the
        'extension
        If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
        Else
        'we have a wildcard in both beggining and extension
        
            If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte _
         And Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
            End If
            
            End If
            
            'telos if
                    
            Next k
            
        Next

If getsubdirs Then

        For Each objfolder In fso.GetFolder(tfolder).SubFolders
           afl_filetext = afl_filetext & Get_All_Directory_Files_With_Wildcard_1ST_MP3_BIGGER_ZERO_SIZE(CStr(objfolder), getsubdirs, wildcard)
        
        Next
       
    End If
End If

Set fso = Nothing
'Get_All_Directory_Files_With_Wildcard = afl_filetext

End Function


Public Function Get_All_Directory_Files_With_Wildcard_ORIGINAL(ByVal tfolder As String, _
 ByVal getsubdirs As Boolean, _
 ByVal wildcard As String) _
 As String

'made by Alexander Triantafyllou
'alextriantf@yahoo.gr
'Athens-Greece

    
Dim objfile As File
Dim objfolder As Folder
Dim fso As New FileSystemObject
Dim kokovar As Variant
Dim k As Long
Dim wildext As String
Dim wildexts As String
Dim wildfirst As String
Dim wildexte As String
Dim wildfirsts As String
Dim wildfirste As String
Dim examfirst As String
Dim examext As String
Dim afl_filetext As String

Kokovar = Split(wildcard, ",")

    If tfolder <> "" Then


        For Each objfile In fso.GetFolder(tfolder).Files
            'do the stuff we want with the files
            For k = 0 To UBound(kokovar)
         wildext = LCase(cutgetExtension(kokovar(k)))
         wildfirst = LCase(Mid(kokovar(k), 1, Len(kokovar(k)) - Len(wildext) - 1))
            
       If InStr(1, wildext, "*") = 0 Then
       wildexts = "888NONE888"
       wildexte = "888NONE888"
       Else
       wildexts = Mid(wildext, 1, InStr(1, wildext, "*") - 1)
       wildexte = Mid(wildext, InStr(1, wildext, "*") + 1, Len(wildext) - InStr(1, wildext, "*"))
       End If
       
        If InStr(1, wildfirst, "*") = 0 Then
       wildfirsts = "888NONE888"
       wildfirste = "888NONE888"
       Else
       wildfirsts = Mid(wildfirst, 1, InStr(1, wildfirst, "*") - 1)
       wildfirste = Mid(wildfirst, InStr(1, wildfirst, "*") + 1, Len(wildfirst) - InStr(1, wildfirst, "*"))
       End If
            
        examfirst = LCase(cutgetName(cutfilename(CStr(objfile))))
        examext = LCase(cutgetExtension(CStr(objfile)))
            
        If wildexts = "888NONE888" Then
'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do not have a wildcard neither on the beggining or the

'extension
        If examfirst = wildfirst And examext = wildext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
                
        Else
        
        'we do have a wildcard in the beggining but not in
'the extension
        If Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste And wildext = examext Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
        
        End If
        
        Else
        'we do not have a wildcard in the extension
        If wildfirsts = "888NONE888" Then
        'we do have a wildcard in the beggining but not in the
'extension
        If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
        Else
        'we have a wildcard in both beggining and extension
        
            If Mid(examext, 1, Len(wildexts)) = wildexts And _
        Mid(examext, Len(wildext) - Len(wildexte) + 1, Len(wildexte)) = wildexte _
         And Mid(examfirst, 1, Len(wildfirsts)) = wildfirsts And _
        Mid(examfirst, Len(wildfirst) - Len(wildfirste) + 1, Len(wildfirste)) = wildfirste Then
        afl_filetext = afl_filetext + objfile + vbNewLine
        End If
            
            End If
            
            End If
            
            'telos if
                    
            Next k
            
        Next

If getsubdirs Then

        For Each objfolder In fso.GetFolder(tfolder).SubFolders
           afl_filetext = afl_filetext & Get_All_Directory_Files_With_Wildcard(CStr(objfolder), getsubdirs, wildcard)
        Next
       
    End If
End If

Set fso = Nothing
Get_All_Directory_Files_With_Wildcard = afl_filetext
End Function




Public Function cutfilename(ByVal fname As String) As String
Dim spos As Integer
Dim ffn As String
spos = InStrRev(fname, "\")
ffn = Mid(fname, spos + 1, Len(fname) - spos)
cutfilename = ffn

End Function

Public Function cutgetExtension(ByVal fname As String)
Dim spos As Integer
Dim koko As String

Spos = InStrRev(fname, ".")
If spos <> 0 Then
koko = Mid(fname, spos + 1, Len(fname) - spos)
End If

CutgetExtension = koko

End Function


Public Function cutgetName(ByVal fname As String)
Dim spos As Integer
Dim koko As String

Spos = InStrRev(fname, ".")
If spos <> 0 Then
koko = Mid(fname, 1, spos - 1)
End If
cutgetName = koko

End Function










-----
-- Fri 05-Feb-10 15:56:37 --
---- Fri 05-Feb-2010 03:53:52p ----  ----
VB - The simplest way to Create Shortcut - VBForums
Http://www.vbforums.com/showthread.php?t=234891

-----
-- Fri 05-Feb-10 12:56:54 --
---- Fri 05-Feb-2010 12:55:26p ----  ----
DevGuru WSH Method: WScript.CreateObject
Http://www.devguru.com/Technologies/wsh/quickref/wscript_CreateObject.html

-----
-- Sun 24-Jan-10 00:07:47 --

'PUT THIS IN DATE 1991 UPDATE WINDOW FASTER
        If chkRefreshListView.Value = vbUnchecked Then
            LockWindowUpdate ListView1.hWnd
        End If


-----
-- Tue 12-Jan-10 21:53:08 --

'BETTER


Public Enum SpecialFolderConstants
    spfldDesktop = &H0
    spfldInternet = &H1
    spfldPrograms = &H2
    spfldControls = &H3
    spfldPrinters = &H4
    spfldPersonal = &H5
    spfldFavorites = &H6
    spfldStartUp = &H7
    spfldRecent = &H8
    spfldSendTo = &H9
    spfldBitBucket = &HA
    spfldStartMenu = &HB
    spfldDesktopDirectory = &H10
    spfldDrives = &H11
    spfldNetwork = &H12
    spfldNethood = &H13
    spfldFonts = &H14
    spfldTemplates = &H15
    spfldCommonStartMenu = &H16
    spfldCommonPrograms = &H17
    spfldCommonStartup = &H18
    spfldCommonDesktopDirectory = &H19
    spfldAppData = &H1A
    spfldPrintHood = &H1B
    spfldAltStartup = &H1D                          '' DBCS
    spfldCommonAltStartup = &H1E                   '' DBCS
    spfldCommonFavorites = &H1F
    spfldInternetCache = &H20
    spfldCookies = &H21
    spfldHistory = &H22
    spfldWindows = &H24
    spfldWindowSystem = &H25
    
    '&H27& My Pictures
    '40 User Profile
    '43 Common Files
    '46 All Users Templates
    '47 Administrative Tools
    '49 Network Connections
    
End Enum

Sub Main()

Dim sSystemFolder As String
Dim sTempFolder As String
Dim sWindowsFolder As String


Set oFS = CreateObject("Scripting.FileSystemObject")
'Sys Object only uses 1 2 3 = 3 temp

' System Folder - Windows\System32
sSystemFolder = oFS.GetSpecialFolder(1)

' Windows Folder Path
sWindowsFolder = oFS.GetSpecialFolder(0)


'--------------------


'SPECIAL FOLDERS ARE NOW IN COMMON MODULE




'NOT BETTER



Private Const CSIDL_ADMINTOOLS = &H30
Private Const CSIDL_ALTSTARTUP = &H1D
Private Const CSIDL_APPDATA = &H1A
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_COMMON_ADMINTOOLS = &H2F
Private Const CSIDL_COMMON_ALTSTARTUP = &H1E
Private Const CSIDL_COMMON_APPDATA = &H23
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Private Const CSIDL_COMMON_DOCUMENTS = &H2E
Private Const CSIDL_COMMON_FAVORITES = &H1F
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_TEMPLATES = &H2D
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_COOKIES = &H21
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_HISTORY = &H22
Private Const CSIDL_INTERNET = &H1
Private Const CSIDL_INTERNET_CACHE = &H20
Private Const CSIDL_LOCAL_APPDATA = &H1C
Private Const CSIDL_MYMUSIC = &HD
Private Const CSIDL_MYPICTURES = &H27
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PRINTHOOD = &H1B
Private Const CSIDL_PROFILE = &H28
Private Const CSIDL_PROGRAM_FILES = &H26
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_SYSTEM = &H25
Private Const CSIDL_TEMPLATES = &H15
Private Const CSIDL_WINDOWS = &H24

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Public Type SHQUERYRBINFO
    cbSize As Long
    i64Size As ULARGE_INTEGER
    i64NumItems As ULARGE_INTEGER
End Type
    
    
    
    
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Public Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
'Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

    
    



Function SpecialFoldersAre()
    LP_RESULT_CALLBACK = messageBox("Start menu folder is located at : " + GetSpecialfolder(CSIDL_STARTMENU), vbInformation, "Start Menu Folder")
    LP_RESULT_CALLBACK = messageBox("Favorites folder is located at : " + GetSpecialfolder(CSIDL_FAVORITES), vbInformation, "Favorites Folder")
    LP_RESULT_CALLBACK = messageBox("Programs folder is located at : " + GetSpecialfolder(CSIDL_PROGRAMS), vbInformation, "Programs Menu Folder")
    LP_RESULT_CALLBACK = messageBox("Desktop folder is located at : " + GetSpecialfolder(CSIDL_DESKTOP), vbInformation, "Desktop Folder")
    End
End Function


Function GetSpecialfolder(ByVal CSIDL As Long) As String
    Dim R As Long
    Dim IDL As ITEMIDLIST
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If R = NOERROR Then
        Path$ = Space$(512)
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function











-----
-- Wed 30-Dec-09 16:11:05 --
Set FS = CreateObject("Scripting.FileSystemObject")

If FS.DRIVEEXISTS("m") = False Then End
Dim d
Set d = FS.GetDrive("M:\")
If d.ISREADY = False Then End

-----
-- Tue 29-Dec-09 20:49:45 --
DATA LOSS COST
If RSS = 1 Then ScanPath.txtPath.Text = "M:\WinRar Archives\"
If RSS = 2 Then ScanPath.txtPath.Text = "M:\00 Art\01 Loggers\Screen-App-Shots\"
If RSS = 3 Then ScanPath.txtPath.Text = "M:\00 Art\01 Loggers\Screen-Shots\"

THESE ON ICYBOX ICYBOX JUST KILLED CARELESSLY
WRONG POWER PLUG

SCREEN SHOTS BACK LOT MONTHS ABOUT 8
WINRARS ONLY 4 MONTHS




-----
-- Tue 17-Nov-09 14:58:17 --

'ERROR CODE
 'IN SUB USE
 Call Err_Dll(Err.LastDllError, "YOUR message", sLocation, "MORE message")

Option Explicit
Public Const ERROR_MOD_NOT_FOUND As Long = 126
Public Sub Err_Dll(ErrorNum As Long, ErrorDesc As String, Source As String, SubOrFunction As String)
    File_WriteTo "ERROR: " & ErrorNum & " at " & Source & "\" & SubOrFunction & " >>> " & ErrorDesc
End Sub
Public Sub Err_Vb(ErrorNum As Long, ErrorDesc As String, Source As String, SubOrFunction As String)
    File_WriteTo "VB ERROR: " & ErrorNum & " at " & Source & "\" & SubOrFunction & " >>> " & ErrorDesc
End Sub
Public Sub File_WriteTo(Text As String)
    '// Allways use this in programs:
 '   Open App.Path & "\PROGRAM.LOG" For Append As #1
 '       Print #1, Text
 '   Close #1
End Sub





    WinampHwnd = FindWindow("Winamp v1.x", vbNullString)
    messageResult = SendMessage(WinampHwnd, WM_COMMAND, Lowervolumeby1percent, ByVal XcNul)


-----
-- Thu 05-Nov-09 21:32:35 --
nice timeserver

Ntp.homehub.btopenworld.com



-----
-- Thu 05-Nov-09 21:02:34 --

'MATT-555ROIDS
'55-88-HAPPY

-----
-- Fri 02-Oct-09 22:48:42 --
################## FAULT FIX
OUTLOOK 2003 MENU BAR ONE PULL DOWN ITEM MISSING
MY CASE WITH TOOLS MENU JUST AFTER HITT CUSTOMISE BY ACCIDENT THEN FLICK GET RID OF IT

ANSWER ####

C:\Documents and Settings\Matt5\Application Data\Microsoft\Outlook
DELETE FILE IN HERE
OUTCMD.DAT




-----
-- Fri 21-Aug-09 07:05:03 --



Dim FileSpec, TT As Long
FileSpec = App.Path + "\" + App.EXEName + ".vbp"

If IsIDE = False And Dir$(FileSpec) <> "" Then
    TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec + """", vbMinimizedNoFocus)
    End
End If



Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End



-----
-- Fri 07-Aug-09 06:33:57 --
'Newest Code of Day
'I Adjusted this So Makes it Expand an Shrink Images to Fit Pciture Box
'before I only shrunk if too big
'enjoy

'sure there was another one I had today for you Peoples

'Private Sub Form_Load()
'Dim lPic As Picture
'Me.Picture1.AutoRedraw = True
'Set lPic = LoadPicture("C:\YourPicture.jpg") 'Use the correct path and filename here
'ResizePicture Me.Picture1, lPic
'End Sub

Public Sub ResizePicture(pBox As PictureBox, pPic As Picture)
Dim lWidth As Single, lHeight As Single
Dim lnewWidth As Single, lnewHeight As Single
'Clear the Picture in the PictureBox
pBox.Picture = Nothing
'Clear the Image in the Picturebox
pBox.Cls
'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
lWidth = pBox.ScaleX(pPic.Width, vbHimetric, pBox.ScaleMode)
lHeight = pBox.ScaleY(pPic.Height, vbHimetric, pBox.ScaleMode)
ox = Round(lWidth / Screen.TwipsPerPixelX, 0)
oy = Round(lHeight / Screen.TwipsPerPixelY, 0)
LBX(3) = Trim(Str(ox)) + " x " + Trim(Str(oy))
'Old Rems
'If image Width > pictureBox Width, resize Width
If lWidth > pBox.ScaleWidth Then
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight * (lnewWidth / lWidth) 'Risize Height keeping proportions
Else
    'if you want true size on small Pics then This
        'lnewWidth = lWidth 'If not, keep the original Width value

    'OR if you want expand small and shrink big then this
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight / (lWidth / lnewWidth) 'Risize Height keeping proportions

End If
    'Old Rems
    'If the image Height > The pictureBox Height, resize Height

If lHeight > pBox.ScaleHeight Then
    lnewHeight = pBox.ScaleHeight 'new Height = PB Height
    lnewWidth = lnewWidth * (lnewHeight / lHeight) 'Risize Width keeping proportions
Else
    'if you want true size on small then This only this line
        lnewHeight = lHeight 'If not, use the same value

    'OR if you want expand small and shrink big then this
    lnewWidth = lnewWidth / (lHeight / lnewHeight) 'Risize Width keeping proportions
End If

'add resized and centered to Picturebox
pBox.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, _
(pBox.ScaleHeight - lnewHeight) / 2, _
lnewWidth, lnewHeight

'Update the Picture with the new image if you need it
'Set pBox.Picture = pBox.Image
End Sub






-----
-- Fri 07-Aug-09 00:24:10 --
'Find your Drive if it changes Drive Letter Often

Dim d
On Local Error Resume Next
Set dc = FS.Drives
For Each d In dc
    dr = d.DriveLetter
    n = d.VolumeName
    If InStr(n, "RAM") > 0 Then Exit For
Next
On Local Error GoTo 0
RamDrive = dr


-----
-- Thu 06-Aug-09 01:24:48 --

'***********************************
'***********************************

'Put This Start Form Load
'It Takes Care of Messy Timers
'Make sure all your Timers have Timer as control Name Somewhere
'Careful you dont force call timers on form load
'If you got a prob it does happen as said messy timers


'***********************************
'***********************************

Dim CTCOL()
ReDim CTCOL(Me.Controls.Count)
For Each Control In Controls
    If InStr(LCase(Control.Name), "timer") > 0 Then
        If Control.Enabled = True Then
            I = I + 1
            CTCOL(i) = Control.Name
            Control.Enabled = False
        End If
    End If
Next
ReDim Preserve CTCOL(i)

'***********************************
'then this at end
'***********************************

'Wat a Job Now Re-Enable any timers that were enabled on start SWEET Safe
For Each Control In Controls
    If InStr(LCase(Control.Name), "timer") > 0 Then
        For I = 1 To UBound(CTCOL)
            If CTCOL(i) = Control.Name Then Control.Enabled = True
        Next
    End If
Next

'***********************************
'***********************************



-----
-- Wed 05-Aug-09 22:31:46 --

'----------------       Latest Version Of Save Chks
'#### REMEMBER SWITCHES

Sub zzLoad_Checks()
'Put This In Declarations
'Dim OCheckQ
'Very Nice Code Will Save all your Check Boxes an Menu Checks and Values you can filter
'some out -- works best as I Know

'if you use menu checks you have to set them yourself on clicks
'If Mnu_VB.Checked = True Then Mnu_VB.Checked = False: Exit Sub
'If Mnu_VB.Checked = False Then Mnu_VB.Checked = True

'3 Parts
' ---
'1 Load
'2 Save
'3 Timer ' This Chks for Changes using XOR Hash
'   Best way I know with Timer Hardly CPU Usage Unfriendly

'zzCheckTimer.Enabled = True
'Dont Have Timer Enabled on Form Load

'Call Timer to Run On Form Unload ----- call zzCheckTimer_Timer
'Then you could set timer slow like 10 Secs - I Do

'-------------------------------
Sub zzLoad_Checks()

ZzCheckTimer.Enabled = False

Dim Th()
ReDim Th(Me.Controls.Count * 3)

On Error Resume Next
I = 0
Open App.Path + "\0TextData\ChkSettings.txt" For Input As #1
Do
    Line Input #1, vv$
    Th(i) = vv$
    I = I + 1
    Line Input #1, vv$
    Th(i) = Val(vv$)
    I = I + 1
    Line Input #1, vv$
    Th(i) = Val(vv$)
    I = I + 1
Loop Until EOF(1)
Close #1
    
tit = i
For Each Control In Me.Controls
    Xx = 1
    
    ppi = LCase(Control.Name)
    If InStr(ppi, LCase("Combo")) Then Xx = 0
    If InStr(ppi, LCase("Check")) Then Xx = 0
    If InStr(ppi, LCase("Mnu")) Then Xx = 0
    If InStr(ppi, LCase("Menu")) Then Xx = 0
    
    
    xxd = -1
    For R = 0 To tit
        If Control.Name = Th(R) Then
            xxd = R: Exit For
        End If
    Next
    
    If xxd => 0 And Xx = 0 Then
        On Error Resume Next
        If Th(xxd + 1) <> 0 Then Control.Value = Th(xxd + 1)
        If Th(xxd + 2) <> 0 Then Control.Checked = Th(xxd + 2)
        'Dont Let People Play Around If you Want to Hard Code In Designer To Enable Chk This
        'This Lets Happen Eg If Th(xxd + 2) <> 0
        
        A1 = 0
        A1 = Control.Value
        B1 = 0
        B1 = Control.Checked
        OCheckQ = OCheckQ Xor (A1 + B1)
        On Error GoTo 0
    End If
Next
zzCheckTimer.Enabled = True
End Sub

'-------------------------------
Sub zzSave_Checks()
If zzCheckTimer.Enabled = False Then Exit Sub
Dim a1, B1 As Integer
On Error Resume Next

MkDir App.Path + "\" + GetComputerName
MkDir App.Path + "\" + GetComputerName + "\0TextData"
Open App.Path + "\0TextData\ChkSettings.txt" For Output As #1

For Each Control In Me.Controls
    Err.Clear
        A1 = 0
        A1 = Control.Value
        A2 = Err.Number
    Err.Clear
        B1 = 0
        B1 = Control.Checked
        B2 = Err.Number
    
    If A2 = 0 Or B2 = 0 Then
        Print #1, Control.Name
        Print #1, A1
        Print #1, B1
    End If
Next
Close #1

End Sub



'-------------------------------
Private Sub zzCheckTimer_Timer()

On Error Resume Next
For Each Control In Me.Controls
    Err.Clear
        A1 = 0
        A1 = Control.Value
    Err.Clear
        B1 = 0
        B1 = Control.Checked
        checkq = checkq Xor (A1 + B1)
Next

If checkq = OCheckQ Then Exit Sub
OCheckQ = checkq
Call zzSave_Checks

End Sub



-----
-- Tue 04-Aug-09 18:56:14 --

'AUTO RESIZE
'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

X = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0

Me.Width = (x + 80)
If mnu = 1 Then
    pluso = 720: pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (y + pluso)
Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2


-----
-- Mon 03-Aug-09 21:59:15 --

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function MoveWindow _
        Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long

Dim Rect8 As RECT

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Sub Code
I = GetWindowRect(Xxr, Rect8)
MoveWindow Xxr, 0, MeRyu4.Bottom, MeRyu4.Right, MeRyu5.Top - MeRyu4.Bottom, True
End Sub

-----
-- Mon 03-Aug-09 04:56:55 --

'This is how I sort a List view

'Project1 - Microsoft Visual Basic [design] - [Form1 (Code)]
'---------------------

'lvwDESCENDING
ListView1.SortOrder = lvwAscending
ListView1.SortKey = 0
ListView1.Sorted = True
ListView1.Sorted = False




-----
-- Sat 01-Aug-09 14:29:23 --

X = X + 1: ReadArray(X) = "(¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)"
X = X + 1: ReadArray(X) = "<º)))><¸..·´`·..¸><(((º>"
X = X + 1: ReadArray(X) = "-:¦:-•:*'''''*:•-:¦:-"


-----
-- Thu 30-Jul-09 21:35:25 --
Comercial Software Program Written -- Staff on holiday -- Retail Sells it - No More Updates Or extremely Slow

-----
-- Thu 30-Jul-09 21:25:05 --
   
   if you set these var in th VBP project file for speed optimization 
   
   just now I found 
   
	Ak1(1) = "AutoIncrementVer=1"
    Ak1(2) = "FavorPentiumPro(tm)=-1"
    Ak1(3) = "BoundsCheck=1"
    Ak1(4) = "FlPointCheck=-1"
    Ak1(5) = "FDIVCheck=-1"
    Ak1(6) = "UnroundedFP=0" '----------------------- This is a bad one to set I put back to zero
    
    getting first errors of this once in big prog

-----
-- Thu 30-Jul-09 10:14:51 --
'******* TIP of THE Day
'When you have a Auto Loader for your VB Programs
'Remember to Chdir First BEfore ther also Doubley Rember change Drive also
'Okay Okay Alright

ChDrive Mid$(A1$, 1, 2)
ChDir A1$
    
Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + A1$ + B1$ + """", vbNormalFocus


-----
-- Wed 29-Jul-09 09:26:47 --

'----------------WINAMP BASICS

Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112
Private Const ipc_isplaying As Long = 104
Private Const IPC_GETPLAYLISTFILE  As Long = 211
Private Const IPC_GETOUTPUTTIME  As Long = 105
Private Const IPC_GETINFO As Long = 126
Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WM_CLOSE = &H10
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
Private Const GW_HWNDNEXT = 2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wmessage As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Tf2 = FindWindow("Winamp v1.x", vbNullString)
messageResult = SendMessage(tf2, WM_USER, 0, ByVal ipc_isplaying)




-----
-- Wed 29-Jul-09 04:54:43 --

'Good Code to Launch VBP IN autorun IDE for Debug Intermitant Faults afetr you have started froming from exe
'Try to Make sure your project filename is same as EXe Easyier then
'Extra Process low Later if you want


Dim FileSpec,TT As Long
FileSpec = App.Path + "\" + App.EXEName + ".vbp" 
If IsIDE = False And Dir$(FileSpec) <> "" Then
    TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec+"""", vbMinimizedNoFocus)
    'Call KingMod.Process_Low(TT)
    End
End If

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************



-----
-- Wed 29-Jul-09 04:16:43 --

'Good Code -- Good Boy


'Dim this Public or How you Like -- sorry if not got them all -- Dim's

Dim InPutDate, TestDate, Year5, MonthStart, DayCount1, DayCountMonth, WholeYear1
Dim Month5, WeeksSinceYear, WeeksSinceInput, ResultYearDate, Month7



'Use this
InPutDate = DateValue("06-01-1997")
TestDate = DateValue("06-01-2009")
TestDate = Now
Call FindTimeInfo
TT$ = TT$ + "Tott Time In & Out Hoss -- " + ResultYearDate + " Years" + Str(Month7) + "M" + Str(DayCountMonth) + "D" + Str(WeeksSinceInput) + " Weeks" + vbCrLf



Sub FindTimeInfo()

'InPutDate = DateValue("01-12-2008")
'TestDate = DateValue("31-05-2009")

Year5 = 0
For r5 = Year(InPutDate) + 1 To Year(TestDate) + 1
    If DateSerial(r5, Month(InPutDate), Day(InPutDate)) < Now Then Year5 = Year5 + 1
Next
For r5 = 1 To -2 Step -1
    xx = DateSerial(Year(TestDate), Month(TestDate) + r5, Day(InPutDate))
    MonthStart = xx
    If xx <= TestDate Then Exit For
Next
Month5 = 0
Month7 = 0
For r5 = 1 To 10000
    xx = DateSerial(Year(InPutDate), Month(InPutDate) + r5, Day(InPutDate))
    If Year(xx) <> oxx Then Month7 = 0
    oxx = Year(xx)
    If xx <= TestDate + 1 Then Month5 = Month5 + 1: Month7 = Month7 + 1
    If xx >= TestDate + 1 Then Exit For
Next

For r5 = Year(TestDate) To 1 Step -1
    If DateSerial(r5, Month(InPutDate), Day(InPutDate)) < Now Then
        DayCount1 = DateDiff("d", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        DayCountMonth = DateDiff("d", MonthStart, TestDate)
        WholeYear1 = DateDiff("d", DateSerial(r5, Month(InPutDate), Day(InPutDate)), DateSerial(r5 + 1, Month(InPutDate), Day(InPutDate)))
        'Month5 = DateDiff("m", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        WeeksSinceYear = DateDiff("w", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate) - 1
        WeeksSinceInput = DateDiff("w", InPutDate, TestDate) - 1
        ResultYearDate = Format$(Year5 + DayCount1 / WholeYear1, ".000")
        Exit For
    End If
Next

End Sub




-----
-- Sun 16-Nov-08 19:19:32 --

Declare'--------
Public Nx!, Ny!, Xxr47

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
             
Public Sub FindCursor2(x, y)
Dim P As POINTAPI
GetCursorPos P
x = P.x
y = P.y
End Sub




'-------------------------------------------

This is Cool Code
'Set any window in VB IDE you work on so soon as you hoover mouse over it always maxi to right size save you big job
'try to do with a click on get current hwnd but that dont work

'Put it in the Timer loop
'Wanted this for ages Just sorting other Peoples code made me why theirs not mine mine all setup okay manually


Declare'--------
Public Nx!, Ny!, Xxr47

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
             



'Short be right my little text pad
'see if it becomes a problem an you want to minimize one

'-----------------

Sub Timer_timer

'VbCode window To maximumon Detect
Dim mWnd As Long
FindCursor Nx, Ny
mWnd = WindowFromPoint(Nx, Ny)

Xxrs = GetActiveWindowClass(mWnd) 'Mwnd will get reset here if none
'Dont Foget Think Thas All
If GetWindowTitle(mWnd) = "Immediate" Then mWnd = 0
If GetWindowTitle(mWnd) = "Locals" Then mWnd = 0
If GetWindowTitle(mWnd) = "Watches" Then mWnd = 0

'Yes we need if starts on form also maybemore
If xxrs = "VbaWindow" Or xxrs = "DesignerWindow" And mWnd > 0 And Xxr47 <> mWnd And mWnd > 0 Then
    Xxr47 = mWnd
    
    ShowWindow mWnd, SW_MAXIMIZE
End If

End Sub

Public Sub FindCursor2(x, y)
Dim P As POINTAPI
GetCursorPos P
x = P.x
y = P.y
End Sub



'In Bas Mod Maybe

Function GetActiveWindowClass(ReturnParent As Long) As String
    GetActiveWindowClass = GetWindowClass(ReturnParent)
End Function

Function GetWindowClass(ByVal hwnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hwnd, sText, 255)
    sText = Left$(sText, Ret)
   GetWindowClass = sText
End Function

Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function






'--------------------------------------

Excellent Code Set the File Time To ANything you want touch date Pro Keep you file time after rotate JPg's



Option Explicit

'THIS TOGGLES A SMALL PASSWORD FEATURE THAT I EMBEDDED
Const MPassword = False

'THIS IS THE CODE THAT MUST BE ENTERED (FINISHED BY PRESSING HASH)
'WHEN THE PASSWORD TOGGLE IS TRUE
Const MSequence = 22515

Dim MEntered As Variant

 Private Type FILETIME
     dwLowDate  As Long
     dwHighDate As Long
 End Type
 
 Private Type SYSTEMTIME
     wYear      As Integer
     wMonth     As Integer
     wDayOfWeek As Integer
     wDay       As Integer
     wHour      As Integer
     wMinute    As Integer
     wSecond    As Integer
     wMillisecs As Integer
 End Type
 
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
  
Private Declare Function CreateFile Lib "kernel32" Alias _
   "CreateFileA" (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) _
   As Long

Private Declare Function LocalFileTimeToFileTime Lib _
     "kernel32" (lpLocalFileTime As FILETIME, _
      lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
   (ByVal hFile As Long, ByVal MullP As Long, _
    ByVal NullP2 As Long, lpLastWriteTime _
    As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
    "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
    As FILETIME) As Long
    
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Public Function SetFileDateTime(ByVal FileName As String, _
  ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
If Dir(FileName) = "" Then Exit Function
If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

LRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

LFileHnd = CreateFile(FileName, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)
    
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
SetFileDateTime = lRet > 0

End Function






---------------------------------
All ways - to set focus
AppActivate "CID Run Me"
.setfocus
'Rft = Putfocus(Rf)
'Rft = Putfocus(Me.hwnd)
 GetForegroundWindow <> Xxr Then 
 SetForegroundWindow Xxr
 
 Public Declare Function Putfocus _
        Lib "user32" _
        Alias "SetFocus" _
        (ByVal hwnd As Long) As Long

--- Best well get an set has double feature

***** PutFocus dont work as ever when setforegroundwindow - Does
 
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

----------------------------------------------------------------------
----------------------------------------------------------------------
--- or always on top
----------------------------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags
End Function
----------------------------------------------------------------------
----------------------------------------------------------------------



'----------------------------------------



'Credits:
'The GetFolder() code was sourced from VB.NET (Brad Martinez & Randy Birch)
'##############################################################################################
'Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'---------------


Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const IDLE_PRIORITY_CLASS As Long = &H40
Public Const HIGH_PRIORITY_CLASS As Long = &H80
Public Const REALTIME_PRIORITY_CLASS As Long = &H100

Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)


Public Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean

'--------------------------------------------- use these for *******************
Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const IDLE_PRIORITY_CLASS As Long = &H40
Public Const HIGH_PRIORITY_CLASS As Long = &H80
Public Const REALTIME_PRIORITY_CLASS As Long = &H100

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000

Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean
'---------------------------------------------






CID_Run_Me_Pro - Microsoft Visual Basic [design]

Public Sub Process_Low(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    If Idle_Timer_Proc < Now Then
        SetPriorityClass hProcess, BELOW_NORMAL_PRIORITY_CLASS
    Else
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    End If
End Sub







Public Function GetFileFromHwnd(lngHwnd) As String

'messageBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

StrFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function


Public Function GetFileFromProc(lngProcess) As String

'messageBox getfilefromhwnd(Me.hwnd)
'Dim lngProcess&, hProcess&, bla&, C&
Dim hProcess&, bla&, C&
Dim strFile As String
Dim x

StrFile = String$(256, 0)
'x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromProc = Left(strFile, C)

End Function






'-------------------------------------------------------------------


Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long


Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
		If IsIconic(lngHwnd) <> 0 Then
			GetWindowState = vbMinimized
		ElseIf IsZoomed(lngHwnd) <> 0 Then
			GetWindowState = vbMaximized
		End If
    End If
End Function




-----
-- Sat 15-Nov-08 02:56:06 --


    'older than 3 h "-to3h"
    '-ri0:10 slow it down
    '-ep1 Exculde base folder from names
    '-inul no error message
    '-cfg- no profile usage
    '-ibck in background
    '-Hp encrypt filenames also
    
ScanPath.cboDate.ListIndex = 0
ScanPath.DTPicker1(0).Value = Now - 20 ' any newer than this date


ScanPath.cboSize.ListIndex = 2 ' -- Select Files Greater than
ScanPath.cboSizeType(0).ListIndex = 2 'select Mega Bytes
ScanPath.txtSize(0).Text = "2" ' ---- For our List Not Allowed over 5 Meg

'This how to get first to columns in a listview1
For we = ScanPath.ListView1.ListItems.Count To 4 Step -1
	A1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(1)
	B1$ = ScanPath.ListView1.ListItems.Item(1)
Next




-----


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************









Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(Tx$)
        
    qq = Now + TimeSerial(0, 1, 30)
    Do
        DoEvents
        ii = FileInUse(Tx$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or qq < Now
    
    If ii = True Then
        messageBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub

'--------------------------------
Better -- Not Better

Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function IsFileAlreadyOpen(FileName As String) As Boolean
    Dim hFile As Long
    Dim lastErr As Long
    hFile = -1
    lastErr = 0
    hFile = lOpen(FileName, &H10)
    If hFile = -1 Then
        lastErr = Err.LastDllError
    Else
        lClose (hFile)
    End If
    isFileAlreadyOpen = (hFile = -1) And (lastErr = 32)
End Function

Function IsFileOpenDelay(Tx$)
	QQ = Now + TimeSerial(0, 10, 0)
    Do
        DoEvents
        ii = IsFileAlreadyOpen(Tx$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or QQ < Now
    If ii = True Then
        messageBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 10 Min"
        Stop
    End If

End Function




'----------------------------


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long



Property Let WindowVisible(hwnd As Long, New_Value As Boolean)
    If hwnd = 0 Then Exit Property
     
    Call ShowWindow(hwnd, IIf(New_Value, SW_SHOW, SW_HIDE))

End Property

Property Get WindowVisible(hwnd As Long) As Boolean

    If hwnd = 0 Then Exit Property
    WindowVisible = (IsWindowVisible(hwnd) > 0)

End Property






'-----------------
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'right order

Function LastModifiedToCurrent(dufile As String)
    Dim lngHandle As Long
    lngHandle = CreateFile(dufile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If TouchFileTimes(lngHandle, ByVal 0&) = 0 Then
        'messageBox "Error while changing file dates. Access Denied." + vbCrLf + duFile, vbCritical, "Modify Error"
    LastModifiedToCurrent = False
    Else
        'messageBox "File date changed successfully.", vbInformation, "Modify Successful"
    LastModifiedToCurrent = True
    End If
    CloseHandle lngHandle
End Function
'---------------------------




/n  Opens a new single-pane Window for the default selection.
This is usually the root of the drive on which Windows is installed. 

/e Starts Windows Explorer using its default view. 

/e, <object> Starts Windows Explorer using its default view with the focus on the specified folder. 

/root, <object> Opens a window view of the specified object. 

/select, <object> Opens a window view with the specified folder, file or
program selected. 

/e /select, FOLDER AND FILE 

Shell "Explorer.exe /e, " + App.Path + "\0TextData", vbNormalFocus
Shell "Explorer.exe /select, " + A1+B1, vbNormalFocus

'--------------------------------------------------------------------
'---------------------
'For example, use switch -to15d to process files older than 15 days and -to2h30m to process files older than 2 hours 15 minutes.
'-to1d
    
'Switch -TB<date>  -  process files modified before the specified date       'Oldest Then Stuff
'Switch -TA<date>  -  process files modified after the specified date         'Newsest stuff
'Switch -TN<time>  -  process files newer than the specified time
'For example, use switch -tn15d to process files newer than 15 days and -tn2h30m to process files newer than 2 hours 30 minutes.
    
'Switch -TO<time>  -  process files older than the specified time

'Switch -RI<p>[:<s>]  -  set priority and sleep time
'If <p> is 0, WinRAR uses the default task priority. <p> equal to 1 sets the lowest possible priority, 15 - the highest possible.
'Sleep time <s> is a value from 0 to 1000 (milliseconds). This is a period of time that WinRAR gives back to system after every
'read or write operation while compressing or extracting. Non-zero <s> May be useful if you need to reduce system load even
'more than can be achieved with <p> parameter.
'execute WinRAR with default priority and 10 ms sleep time:
'WinRAR a -ri0:10 backup *.*
    
'Switch -X@<listfile>  -  exclude files using a specified list file
'Exclude files that are too big
'WinRAR a -x@list.txt bin *.exe
    
'-s Switch -S  -  create solid archive
'-ep Switch -EP  -  exclude paths from names
'-ep1 Exculde base folder from names
'-inul no error message
'-cfg- no profile usage
'-ibck in background
'-Hp encrypt filenames also
'-y answer yes to all dont stop
'-df  'Del afterwards'
'-x*\temp\*
'-x*\temp\*
'-r Enables subfolder recursion
'-sv indepentdent volumes


'------------------------------

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long


'------------------------------
'Center Screen
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

'center screen ONCE

If Me.WindowState <> 1 Then Exit Sub
DIM VARCENTER
IF VARCENTER=TRUE THEN EXIT SUB
Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
Me.Left = Screen.Width / 2 - Me.Width / 2 '+400
VARCENTER=TRUE


'Taskbar at the Bottom
'---------------------
Dim RectTaskbar As RECT
HwndTask = FindWindow("Shell_TrayWnd", vbNullString)
HwndResult = GetWindowRect(HwndTask, RectTaskbar)
' CENTER ON SCREEN COMPENSATE TASKBAR AT BOTTOM
'----------------------------------------------
If Screen.Height - RectTaskbar.Top * Screen.TwipsPerPixelY < 2500 Then
	Me.Left = Screen.Width / 2 - Me.Width / 2
	Me.Top = (Screen.Height - ((RectTaskbar.Bottom - RectTaskbar.Top) * Screen.TwipsPerPixelY)) / 2 - Me.Height / 2
Else
	Me.Left = Screen.Width / 2 - Me.Width / 2
	Me.Top = Screen.Height / 2 - Me.Height / 2
End If


'------------------------------

Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.Getfile(txtFilename)
a = F.Size

'------------------------------

Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Public Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long


'---------------------------------------------

'Simple Copy File
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile "", ""

Note - Do we have to spec filename on destination ?


'Simple Find Date
Dim Filespec, Adate1
Set fs = CreateObject("Scripting.FileSystemObject")
Filespec = App.Path + "\"
Set F = fs.GetFile(Filespec)
Adate1 = F.datelastmodified

'More-------

Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFile(Path + Filename)
adate1 = f.datelastmodified
asize1 = f.Size


More------

Public FS

'Put this in start form no need to keep creating 
Set fs = CreateObject("Scripting.FileSystemObject")

Set f = fs.GetFile(Path + Filename)
adate1 = f.datelastmodified
asize1 = f.Size



'-----------------------------------------------------

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GWL_WNDPROC = (-4)
Private Const IDANI_OPEN = &H1
Private Const IDANI_CLOSE = &H2
Private Const IDANI_CAPTION = &H3
Private Const WM_USER = &H400

Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function



Function FindWinPart() As Long

Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String, Huge

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
Do While test_hwnd <> 0
    
        ash$ = GetWindowTitle(test_hwnd)
        If ash$ <> "" And InStr(ash$, "Enter password for") > 0 Then
            huge = test_hwnd
        End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop
FindWinPart = False
If huge > 0 Then FindWinPart = huge
'If huge > 0 Then FindWinPart = true

End Function


'-------------------------

Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function


'MATT-555ROIDS
'55-88-HAPPY

'--------------------------

Grab Exe From Hwnd Full Path
'Public cProcesses As New clsCnProc
    '#### Functions/Consts used for GetHWnd() (part of Convert())
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

'in code
filexxx$ = GetFileFromHwnd(CurProcHwnd)
'***********

Public Function GetFileFromHwnd(lngHwnd) As String

'messageBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

StrFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameExA(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function




'----------------------------------------------------------


Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect _
   As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim Rect2 As RECT

'-------------------------------------------------------
'Sort a ListView
'This sorts on one column then the data is still kept sorted that column then the next column sorted got like double sort

            ListView1.SortOrder = lvwAscending
            ListView1.SortKey = 0
            ListView1.Sorted = True
            ListView1.SortKey = 1
            ListView1.Sorted = True
            ListView1.Sorted = False


'-------------------------------------------------------------


**** Start Of
---------------------
Count = 160 -- Fri 17-Jul-2009 01:37:21
---------------------
In Internet Explorer -- 
WinTitle = http://www.scottandmichelle.net/scott/code/index.html
URL = Scott's Sample Code
---------------------
Scott's Sample Code - Windows Internet Explorer
---------------------



ShellAndWait 

Description: 
 This function mimics the VB Shell function, however it waits for the completion of the shelled function to finish, which can be very useful sometimes. 
  
Code: 
 Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal _
   dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal _
   dwProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
   As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal _
   dwMilliseconds As Long)

Private Const INFINITE = &HFFFF
Private Const SYNCHRONIZE = &H100000
Private Const WAIT_TIMEOUT = &H102

Public Sub ShellAndWait(PathName, Optional WindowStyle As _
   VbAppWinStyle = vbMinimizedFocus, Optional bDoEvents As _
   Boolean = False)

    Dim dwProcessID As Long
    Dim hProcess As Long
    
    dwProcessID = Shell(PathName, WindowStyle)
    
    If dwProcessID = 0 Then
        Exit Sub
    End If
    
    hProcess = OpenProcess(SYNCHRONIZE, False, dwProcessID)
    
    If hProcess = 0 Then
        Exit Sub
    End If
    
    If bDoEvents Then
        Do While WaitForSingleObject(hProcess, 100) = WAIT_TIMEOUT
            DoEvents
        Loop
    Else
        WaitForSingleObject hProcess, INFINITE
    End If
    
    CloseHandle hProcess
    
End Sub

 
  
Sample Usage: 
  
     Dim nStart As Date
    nStart = Now

    ShellAndWait "notepad", vbNormalFocus, True
    
    messagebox "You spent " & DateDiff("s", nStart, Now) & _
       " second(s) in notepad.", vbCritical

 
'----------------------------------------------------------


Web Server 

Description: 
 clsWebServer is a basic web server. Since it's almost completed encapsulated in these four classes, it's very easy to drop into an existing project with minimal new code. All you need to do is setup some options, and handle the page requests. The class also supports basic authentication. 
  
Code: 
 '----------------------------------------------------------------------
'Start of Class Module: clsWebServer

Option Explicit

Private m_nPortNumber As Long
Private m_sAuthUser As String
Private m_sAuthPass As String
Private m_sAuthRealm As String
Private m_bAuth As Boolean
Private m_frm As Form

Public Event GetPage(sPath As String, wsr As clsWebServerResponse)

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const BASE64_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklm" & _
                             "nopqrstuvwxyz0123456789+/"

Private WithEvents wsk As Winsock

Private m_wpMain As clsWebServerParent
Private WithEvents m_wpMainWatcher As clsWebServerParent
Private m_cWinsocks As Collection
Private m_nID As Long
Private WithEvents m_Timer As Timer

Public Property Set ParentForm(frmNew As Form)

    Set m_frm = frmNew

End Property

Public Property Let PortNumber(nPort As Long)

    m_nPortNumber = nPort
    
End Property

Public Property Let AuthMode(bAuth As Boolean)

    m_bAuth = bAuth

End Property

Public Property Get AuthMode() As Boolean

    AuthMode = m_bAuth

End Property

Public Property Let AuthRealm(sAuthRealm As String)

    m_sAuthRealm = sAuthRealm

End Property

Public Property Get AuthRealm() As String

    AuthRealm = m_sAuthRealm

End Property

Public Property Let AuthUser(sAuthUser As String)

    m_sAuthUser = sAuthUser

End Property

Public Property Let AuthPass(sAuthPass As String)

    m_sAuthPass = sAuthPass

End Property

Public Property Get AuthString() As String

    AuthString = "Basic " & EncodeBase64(m_sAuthUser & ":" & m_sAuthPass)

End Property

Public Sub Start()

    Debug.Assert Not m_frm Is Nothing
    Debug.Assert m_nPortNumber > 0

    Dim sName As String

    Set m_wpMain = New clsWebServerParent
    Set m_wpMainWatcher = m_wpMain
    Set m_cWinsocks = New Collection
    sName = "wsk_" & Format(GetTickCount) & "_0"
    Set m_Timer = m_frm.Controls.Add("VB.Timer", sName & "_timer", m_frm)
    m_Timer.Enabled = True
    m_Timer.Interval = 60000

    Set wsk = m_frm.Controls.Add("MSWinsock.Winsock.1", sName, m_frm)
    
    wsk.LocalPort = m_nPortNumber
    wsk.Listen
    
End Sub

Private Sub Class_Initialize()

    m_nPortNumber = 80
    m_bAuth = False
    m_sAuthRealm = "Protected content"

End Sub

Private Sub m_Timer_Timer()

    Dim obj As clsWebServerClient
    Dim nMinTick As Long
    Dim I As Long
    
    nMinTick = GetTickCount - 60000
    For I = m_cWinsocks.Count To 1 Step -1
        Set obj = m_cWinsocks.Item(i)
        If obj.LastTick < nMinTick Then
            obj.wsk.Close
            m_wpMainWatcher_Closed obj
        End If
    Next

End Sub

Private Sub m_wpMainWatcher_Closed(wsk As clsWebServerClient)

    Dim obj As clsWebServerClient
    Dim I As Long
    For I = m_cWinsocks.Count To 1 Step -1
        Set obj = m_cWinsocks.Item(i)
        If obj Is wsk Then
            m_cWinsocks.Remove i
            Exit For
        End If
    Next

End Sub

Private Sub m_wpMainWatcher_GetID(nID As Long)

    m_nID = m_nID + 1
    nID = m_nID

End Sub

Private Sub m_wpMainWatcher_GetPage(sPath As String, wsr As _
   clsWebServerResponse)

    RaiseEvent GetPage(sPath, wsr)

End Sub

Private Sub wsk_ConnectionRequest(ByVal requestID As Long)

    Dim obj As clsWebServerClient
    Set obj = New clsWebServerClient
    obj.Start m_frm, m_wpMain, Me
    
    obj.wsk.Accept requestID
    
    m_cWinsocks.Add obj

End Sub

Function EncodeBase64(sToEncode As String)
  Dim I As Long
  Dim cOut As String
  
  For I = 1 To Len(sToEncode) Step 3
    Dim nGroup As Long
    Dim sGroup As String
    
    nGroup = &H10000 * Asc(Mid(sToEncode, i, 1)) + _
      &H100 * Asc(Mid(sToEncode, I + 1, 1) & Chr(0)) + _
      Asc(Mid(sToEncode, I + 2, 1) & Chr(0))
    
    sGroup = Oct(nGroup)
    sGroup = String(8 - Len(sGroup), "0") & sGroup
    
    EncodeBase64 = EncodeBase64 & _
      Mid(BASE64_CHARS, CLng("&o" & Mid(sGroup, 1, 2)) + 1, 1) + _
      Mid(BASE64_CHARS, CLng("&o" & Mid(sGroup, 3, 2)) + 1, 1) + _
      Mid(BASE64_CHARS, CLng("&o" & Mid(sGroup, 5, 2)) + 1, 1) + _
      Mid(BASE64_CHARS, CLng("&o" & Mid(sGroup, 7, 2)) + 1, 1)
    
  Next
  
  Select Case Len(sToEncode) Mod 3
    Case 1: '8 bit final
      EncodeBase64 = Left(EncodeBase64, Len(EncodeBase64) - 2) + "=="
    Case 2: '16 bit final
      EncodeBase64 = Left(EncodeBase64, Len(EncodeBase64) - 1) + "="
  End Select
End Function

'End of Class Module: clsWebServer
'----------------------------------------------------------------------

'----------------------------------------------------------------------
'Start of Class Module: clsWebServerClient

Option Explicit

Public WithEvents wsk As Winsock
Public LastTick As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private m_sData As String
Private m_wpMain As clsWebServerParent
Private m_wpSettings As clsWebserver
Private m_frm As Form
Private m_sGetPath As String
Private m_sGetAuth As String
Private m_sOut As String

Public Sub Start(frm As Form, wp As clsWebServerParent, wpm As _
   clsWebserver)

    LastTick = GetTickCount

    Dim nID As Long
    Dim sName As String

    Set m_wpMain = wp
    Set m_wpSettings = wpm
    
    m_wpMain.TriggerGetID nID
    sName = "wsk_" & Format(GetTickCount) & "_" & Format(nID)

    Set wsk = frm.Controls.Add("MSWinsock.Winsock.1", sName, frm)
    
End Sub

Private Sub wsk_Close()

    m_wpMain.TriggerClosed Me

End Sub

Private Sub wsk_DataArrival(ByVal bytesTotal As Long)

    Dim sData As String
    wsk.GetData sData, vbString
    
    m_sData = m_sData & Replace(sData, vbCr, "")

    Do While InStr(1, m_sData, vbLf)
        ParseLine Mid(m_sData, 1, InStr(1, m_sData, vbLf) - 1)
        m_sData = Mid(m_sData, InStr(1, m_sData, vbLf) + 1)
    Loop

End Sub

Private Sub ParseLine(sLine As String)

    If Mid(sLine, 1, 4) = "GET " Then
        Dim vSplit As Variant
        vSplit = Split(sLine, " ")
        m_sGetPath = vSplit(1)
    ElseIf Mid(sLine, 1, 15) = "Authorization: " Then
        m_sGetAuth = Mid(sLine, 16)
    ElseIf sLine = "" Then
        Dim sOut As String
        Dim bShowPage As Boolean
        bShowPage = False
        
        If m_wpSettings.AuthMode Then
            If m_sGetAuth = m_wpSettings.AuthString Then
                bShowPage = True
            Else
                sOut = "HTTP/1.1 401 Authorization Required" & vbNewLine
                sOut = sOut & "WWW-Authenticate: Basic realm=""" & _
                   m_wpSettings.AuthRealm & """" & _
                   vbNewLine

                sOut = sOut & "Content-Type: text/html" & vbNewLine
                sOut = sOut & "" & vbNewLine
                sOut = sOut & _
                   "<html><head><title>Sorry</title></head><body>" & _
                   vbNewLine

                sOut = sOut & "You entered the wrong " _
                   & "password</body></html>" & _
                   vbNewLine

            End If
        Else
            bShowPage = True
        End If
                
        If bShowPage Then
            Dim wsr As clsWebServerResponse
            Set wsr = New clsWebServerResponse
            m_wpMain.TriggerGetPage m_sGetPath, wsr
            
            sOut = "HTTP/1.1 200 OK" & vbNewLine
            sOut = sOut & "Content-type: " & wsr.ContentType & vbNewLine _
               & vbNewLine

            sOut = sOut & wsr.Page
        End If
        
        m_sGetPath = ""
        m_sGetAuth = ""
        m_sOut = sOut
        SendData
    End If

End Sub

Private Sub SendData()

    If m_sOut = "" Then
        wsk.Close
        m_wpMain.TriggerClosed Me
        Exit Sub
    End If

    If Len(m_sOut) > 4096 Then
        wsk.SendData Mid(m_sOut, 1, 4096)
        m_sOut = Mid(m_sOut, 4097)
    Else
        wsk.SendData m_sOut
        m_sOut = ""
    End If
    
End Sub

Private Sub wsk_SendComplete()

    SendData

End Sub

'End of Class Module: clsWebServerClient
'----------------------------------------------------------------------

'----------------------------------------------------------------------
'Start of Class Module: clsWebServerParent

Option Explicit

Public Event Closed(wsk As clsWebServerClient)
Public Event GetID(nID As Long)
Public Event GetPage(sPath As String, wsr As clsWebServerResponse)

Public Sub TriggerGetPage(sPath As String, wsr As clsWebServerResponse)
    RaiseEvent GetPage(sPath, wsr)
End Sub

Public Sub TriggerClosed(wsk As clsWebServerClient)
    RaiseEvent Closed(wsk)
End Sub

Public Sub TriggerGetID(nID As Long)
    RaiseEvent GetID(nID)
End Sub

'End of Class Module: clsWebServerParent
'----------------------------------------------------------------------

'----------------------------------------------------------------------
'Start of Class Module: clsWebServerResponse

Option Explicit

Private m_sPage As String
Private m_sContentType As String

Public Property Get Page() As String

    Page = m_sPage

End Property

Public Property Let Page(sNewPage As String)

    m_sPage = sNewPage

End Property

Public Property Get ContentType() As String

    ContentType = m_sContentType

End Property

Public Property Let ContentType(sNewContentType As String)

    m_sContentType = sNewContentType

End Property

Public Sub LoadPageFromFile(sFile As String)

    Dim nFile As Long
    nFile = FreeFile
    Open sFile For Binary As #nFile
    m_sPage = Input(LOF(nFile), #nFile)
    Close #nFile
    
End Sub

Private Sub Class_Initialize()

    m_sContentType = "text/html"

End Sub

'End of Class Module: clsWebServerResponse
'----------------------------------------------------------------------

 
  
Sample Usage: 
  
 Option Explicit

Private WithEvents m_webMain As clsWebserver

Private Sub Form_Load()

    Set m_webMain = New clsWebserver
    Set m_webMain.ParentForm = Me
    m_webMain.PortNumber = 80
    m_webMain.Start

End Sub

Private Sub m_webMain_GetPage(sPath As String, wsr As _
   clsWebServerResponse)

    wsr.Page = "<html><head><title>Sample</title></head><body>"
    wsr.Page = wsr.Page & "This is a sample page<br>Path = ["
    wsr.Page = wsr.Page & sPath & "]</body></html>"
    
End Sub

 


'----------------------------------------


Winamp Class 

Description: 
 This class provides a thin wrapper around many of the WM_WA_IPC functions. Use with caution.

7/12/1999: Added the command button messages.

12/10/2000: Added PressButton, which exposes more WinAmp functionality.

5/7/2004: Updated the Playlist functions to deal with changes in recent WinAmp versions. 
  
Code: 
 '----------------------------------------------------------------------
' Start: clsWinamp

Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, _
   lParam As Any) As Long
   
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
   (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Function PostMessageLng Lib "user32" Alias "PostMessageA" _
   (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Private Const WM_COMMAND = &H111
Private Const WM_COPYDATA = &H4A

Private Const WINAMP_BUTTON1 = 40044
Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WINAMP_BUTTON4 = 40047
Private Const WINAMP_BUTTON5 = 40048

Private Const WM_USER = &H400
Private Const WM_WA_IPC = WM_USER

Private Const IPC_GETVERSION = 0
Private Const IPC_PLAYFILE = 100
Private Const IPC_DELETE = 101
Private Const IPC_STARTPLAY = 102
Private Const IPC_CHDIR = 103
Private Const IPC_ISPLAYING = 104
Private Const IPC_GETOUTPUTTIME = 105
Private Const IPC_JUMPTOTIME = 106
Private Const IPC_WRITEPLAYLIST = 120

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Enum enumPlay
    playPlaying = 1
    playPaused = 3
    playStoped = 0
End Enum

Public Enum enumButtons
    PreviousTrackButton = 40044
    NextTrackButton = 40048
    PlayButton = 40045
    PauseUnpauseButton = 40046
    StopButton = 40047
    StopAfterCurrentTrack = 40147
    FadeoutAndStop = 40157
    FastForward5Seconds = 40148
    FastRewind5Seconds = 40144
    StartOfPlaylist = 40154
    GoToEndOfPlaylist = 40158
    OpenFileDialog = 40029
    OpenUrlDialog = 40155
    OpenFileInfoBox = 40188
    SetTimeDisplayModeToElapsed = 40037
    SetTimeDisplayModeToRemaining = 40038
    TogglePreferencesScreen = 40012
    OpenVisualizationOptions = 40190
    OpenVisualizationPlugInOptions = 40191
    ExecuteCurrentVisualizationPlugIn = 40192
    ToggleAboutBox = 40041
    ToggleTitleAutoscrolling = 40189
    ToggleAlwaysOnTop = 40019
    ToggleWindowshade = 40064
    TogglePlaylistWindowshade = 40266
    ToggleDoublesizeMode = 40165
    ToggleEq = 40036
    TogglePlaylistEditor = 40040
    ToggleMainWindowVisible = 40258
    ToggleMinibrowser = 40298
    ToggleEasymove = 40186
    RaiseVolumeBy1Perc = 40058
    LowerVolumeBy1Perc = 40059
    ToggleRepeat = 40022
    ToggleShuffle = 40023
    OpenJumpToTimeDialog = 40193
    OpenJumpToFileDialog = 40194
    OpenSkinSelector = 40219
    ConfigureCurrentVisualizationPlugIn = 40221
    ReloadTheCurrentSkin = 40291
    CloseWinamp = 40001
    MovesBack10TracksInPlaylist = 40197
    ShowTheEditBookmarks = 40320
    AddsCurrentTrackAsABookmark = 40321
    PlayAudioCd = 40323
    LoadAPresetFromEq = 40253
    SaveAPresetToEqf = 40254
    OpensLoadPresetsDialog = 40172
    OpensAutoLoadPresetsDialog = 40173
    LoadDefaultPreset = 40174
    OpensSavePresetDialog = 40175
    OpensAutoLoadSavePreset = 40176
    OpensDeletePresetDialog = 40178
    OpensDeleteAnAutoLoadPresetDialog = 40180
End Enum

Private m_sWinampDir As String
Private m_cPlaylist As Collection

Private Function WAWin() As Long
    WAWin = FindWindow("Winamp v1.x", vbNullString)
    Debug.Assert WAWin <> 0
End Function

Public Property Get Version() As Long
    Version = SendMessage(WAWin, WM_WA_IPC, 0, ByVal IPC_GETVERSION)
End Property

Public Sub AddPlaylist(ByVal sFile As String)

    Dim hWnd As Long
    hWnd = WAWin
    
    If Version > &H1700 Then
        Dim cds As COPYDATASTRUCT
        Dim sFileAscii As String
        sFileAscii = StrConv(sFile, vbFromUnicode)
        cds.dwData = IPC_PLAYFILE
        cds.lpData = StrPtr(sFileAscii)
        cds.cbData = Len(sFile) + 1
        SendMessage hWnd, WM_COPYDATA, 0, ByVal VarPtr(cds)
    Else
        sFile = sFile & vbNullChar
        While sFile <> ""
            PostMessageLng hWnd, WM_WA_IPC, Asc(Mid(sFile, 1, 1)), IPC_PLAYFILE
            sFile = Mid(sFile, 2)
        Wend
    End If
    
End Sub

Private Sub AddPlaylist2(Filename As String)
End Sub

Public Sub DeletePlaylist()
    PostMessage WAWin, WM_WA_IPC, 0, ByVal IPC_DELETE
End Sub

Public Sub StartPlay()
    PostMessage WAWin, WM_WA_IPC, 0, ByVal IPC_STARTPLAY
End Sub

Public Sub ChDir(ByVal sDir As String)

    Dim hWnd As Long
    hWnd = WAWin
    sDir = sDir & vbNullChar
    
    While sDir <> ""
        PostMessage hWnd, WM_WA_IPC, Asc(Mid(sDir, 1, 1)), _
            ByVal IPC_CHDIR
            
        sDir = Mid(sDir, 2)
    Wend

End Sub

Public Property Get IsPlaying() As enumPlay
    IsPlaying = SendMessage(WAWin, WM_WA_IPC, 0, ByVal IPC_ISPLAYING)
End Property

Public Property Get SongPos() As Long
    SongPos = SendMessage(WAWin, WM_WA_IPC, 0, ByVal IPC_GETOUTPUTTIME)
End Property

Public Property Get SongLen() As Long
    SongLen = SendMessage(WAWin, WM_WA_IPC, 1, ByVal IPC_GETOUTPUTTIME)
End Property

Public Sub JumpToTime(ByVal nTime As Long)
    PostMessage WAWin, WM_WA_IPC, nTime, ByVal IPC_JUMPTOTIME
End Sub

Public Property Get WinampDir() As String
    WinampDir = m_sWinampDir
End Property

Public Property Let WinampDir(sNew As String)
    m_sWinampDir = sNew
End Property

Public Function RefreshPlaylist() As Long
    RefreshPlaylist = SendMessage(WAWin, WM_WA_IPC, 0, _
                       ByVal IPC_WRITEPLAYLIST)
    
    Dim nFile As Byte
    Dim sLine As String
    nFile = FreeFile
    Open m_sWinampDir & "\winamp.m3u" For Input As #nFile
    
    Do While m_cPlaylist.Count > 0
        m_cPlaylist.Remove 1
    Loop
    
    While Not EOF(nFile)
        Line Input #nFile, sLine
        m_cPlaylist.Add sLine
    Wend
    
    Close #nFile
    
End Function

Public Function PlaylistCount() As Long
    PlaylistCount = m_cPlaylist.Count
End Function

Public Function PlaylistSong(nIndex As Long) As String
    Debug.Assert nIndex >= 1
    Debug.Assert nIndex <= m_cPlaylist.Count
    PlaylistSong = m_cPlaylist(nIndex)
End Function

Private Sub Class_Initialize()
    Set m_cPlaylist = New Collection
    m_sWinampDir = "c:\program files\winamp\"
End Sub

Public Sub CommandPrevSong()
    SendMessage WAWin, WM_COMMAND, WINAMP_BUTTON1, ByVal 0
End Sub

Public Sub CommandPlay()
    SendMessage WAWin, WM_COMMAND, WINAMP_BUTTON2, ByVal 0
End Sub

Public Sub CommandPause()
    SendMessage WAWin, WM_COMMAND, WINAMP_BUTTON3, ByVal 0
End Sub

Public Sub CommandStop()
    SendMessage WAWin, WM_COMMAND, WINAMP_BUTTON4, ByVal 0
End Sub

Public Sub CommandNextSong()
    SendMessage WAWin, WM_COMMAND, WINAMP_BUTTON5, ByVal 0
End Sub

Public Sub PressButton(buttonNum As enumButtons)
    SendMessage WAWin, WM_COMMAND, buttonNum, ByVal 0
End Sub

' End: clsWinamp
'----------------------------------------------------------------------

 
  
Sample Usage: 
  
     Dim WA As New clsWinamp
    Dim nSong As Long
    
    Debug.Print "Winamp version: " & Hex(WA.Version)
    WA.AddPlaylist "C:\Mp3's\Superman.mp3"
    WA.StartPlay
    nSong = WA.RefreshPlaylist
    Debug.Print "Current Song: " & WA.PlaylistSong(nSong)

 
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------



Font 

Description: 
 This is a replacement for using the Font properties of a PictureBox and Printing text on them. Unlike the built in VB functions, this class allows you to control the rotation and antialiased of text. It's also a good jumping off point for using other Font APIs directly. 
  
Code: 
 '------------- Class: clsFont

Option Explicit

Private Declare Function GetTextExtentPoint Lib "gdi32" Alias _
    "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As _
    String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
    "CreateFontIndirectW" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
    ByVal nIndex As Long) As Long

Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const LF_FACESIZE = 32
Private Const NONANTIALIASED_QUALITY = 3
Private Const LOGPIXELSY = 90

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * LF_FACESIZE
End Type

Private m_nSize As Double
Private m_szFontName As String
Private m_bNotAntialiased As Boolean
Private m_nEscapement As Long
Private m_pctBox As PictureBox

Public Property Get PctBox() As PictureBox
    Set PctBox = m_pctBox
End Property

Public Property Set PctBox(pctNewBox As PictureBox)
    Set m_pctBox = pctNewBox
End Property

Public Property Get FontSize() As Double
    FontSize = m_nSize
End Property

Public Property Let FontSize(ByVal nNewSize As Double)
    m_nSize = nNewSize
End Property

Public Property Get FontName() As String
    FontName = m_szFontName
End Property

Public Property Let FontName(ByVal szNewFontName As String)
    m_szFontName = szNewFontName
End Property

Public Property Get NotAntialiased() As Boolean
    NotAntialiased = m_bNotAntialiased
End Property

Public Property Let NotAntialiased(ByVal bNewNotAntialiased As Boolean)
    m_bNotAntialiased = bNewNotAntialiased
End Property

Public Property Get Escapement() As Long
    Escapement = m_nEscapement
End Property

Public Property Let Escapement(ByVal nNewEscapement As Long)
    m_nEscapement = nNewEscapement
End Property

Private Sub SetupFont(hdc As Long, font As LOGFONT)

    font.lfEscapement = m_nEscapement
    font.lfFaceName = m_szFontName
    font.lfHeight = -Int((m_nSize * _
        GetDeviceCaps(hdc, LOGPIXELSY) / 72#) + 0.5)
    font.lfQuality = IIf(m_bNotAntialiased, NONANTIALIASED_QUALITY, 0)

End Sub

Public Sub PrintText(ByVal sText As String)
    
    Dim font As LOGFONT
    Dim prevFont As Long
    Dim hFont As Long
    
    SetupFont m_pctBox.hdc, font
    
    hFont = CreateFontIndirect(font)
    
    prevFont = SelectObject(m_pctBox.hdc, hFont)
    
    m_pctBox.Print sText
    
    SelectObject m_pctBox.hdc, prevFont
    DeleteObject hFont
    
End Sub

Public Sub PrintTextLine(ByVal sText As String)
    
    Dim font As LOGFONT
    Dim prevFont As Long
    Dim hFont As Long
    
    SetupFont m_pctBox.hdc, font
    
    hFont = CreateFontIndirect(font)
    
    prevFont = SelectObject(m_pctBox.hdc, hFont)
    
    m_pctBox.Print sText;
    
    SelectObject m_pctBox.hdc, prevFont
    DeleteObject hFont
    
End Sub

Public Function GetTextHeight(ByVal sText As String) As Long
    
    Dim font As LOGFONT
    Dim prevFont As Long
    Dim hFont As Long
    
    SetupFont m_pctBox.hdc, font
    
    hFont = CreateFontIndirect(font)
    
    prevFont = SelectObject(m_pctBox.hdc, hFont)
    
    Dim sSize As Size
    GetTextExtentPoint m_pctBox.hdc, sText, Len(sText), sSize
    
    GetTextHeight = m_pctBox.ScaleY(sSize.cy, vbPixels, _
        m_pctBox.ScaleMode)
    
    SelectObject m_pctBox.hdc, prevFont
    DeleteObject hFont
    
End Function

Public Function GetTextWidth(ByVal sText As String) As Long
    
    Dim font As LOGFONT
    Dim prevFont As Long
    Dim hFont As Long
    
    SetupFont m_pctBox.hdc, font
    
    hFont = CreateFontIndirect(font)
    
    prevFont = SelectObject(m_pctBox.hdc, hFont)
    
    Dim sSize As Size
    GetTextExtentPoint m_pctBox.hdc, sText, Len(sText), sSize
    
    GetTextWidth = m_pctBox.ScaleX(sSize.cx, _
        vbPixels, m_pctBox.ScaleMode)
    
    SelectObject m_pctBox.hdc, prevFont
    DeleteObject hFont
    
End Function

Public Sub PrintCenterText(ByVal sText As String)

    m_pctBox.CurrentX = (m_pctBox.Width - GetTextWidth(sText)) / 2
    m_pctBox.CurrentY = (m_pctBox.Height - GetTextHeight(sText)) / 2
    
    PrintText sText

End Sub

Private Sub Class_Initialize()

    m_nSize = 10
    m_szFontName = "Times New Roman"
    
End Sub

    'Dim myfont As clsFont
    'Set myfont = New clsFont
    'Set myfont.PctBox = Picture1
    'myfont.Escapement = 450
    'myfont.PrintCenterText "Hello World!"

'------------- End of class: clsFont

 
  
Sample Usage: 
  
     Dim myfont As clsFont
    Set myfont = New clsFont
    Set myfont.PctBox = Picture1
    myfont.Escapement = 450
    myfont.PrintCenterText "Hello World!"

 
'----------------------------------------------------------
'----------------------------------------------------------
'----------------------------------------------------------
'----------------------------------------------------------



AutoSize 

Description: 
 This class automatically resizes all controls on a form to keep them the same size and position, relative the changing dimensions of the form. For instance, if the user resizes the form to twice it's size, the class will automatically double each of the control's widths and heights, double the font size, and reposition the controls so they stay in the same relative positions. I would recommend changing the font of each of the controls you'll be using to a TrueType font to avoid granularity in the fonts as they're resized. 
  
Code: 
 '---------------------------------------------------------------------
' Begin of Class Module: clsAutoSize

Option Explicit

Private Declare Function CopyMemory Lib "kernel32" Alias _
   "RtlMoveMemory" (dest As Any, src As Any, ByVal length As Long) _
   As Long

Private WithEvents m_frmParent As Form 'The parent form to watch

Private Type SizeAndPos ' Used to track the size and position of
    LocLeft As Single   ' child controls
    LOcTop As Single
    SizeWidth As Single
    SizeHeight As Single
    FontSize As Single
End Type

Private m_colOrig As Collection ' Collection of sizes for controls,
' actually contains a string for each control that's passed through
' PackUDT/UnPackUDT to see its values

Private m_sizeOrig As SizeAndPos 'The size of the parent form

Private m_nLastPercX As Double ' Last percentage of change of the
Private m_nLastPercY As Double ' parent form, initially 1.0

Private Function PackUDT(tmp As SizeAndPos) As String
    ' PackUDT and UnPackUDT are a little cheat to copy UDT's to and
    ' from strings.  This makes it possible to stuff the UDTs into a
    ' collection, without having to use a dummy class

    PackUDT = Space(Len(tmp))
    
    CopyMemory ByVal PackUDT, tmp, Len(tmp)

End Function

Private Sub UnPackUDT(tmp As SizeAndPos, sIn As String)

    CopyMemory tmp, ByVal sIn, Len(tmp)

End Sub

Public Sub Initialize(frmParent As Form)

    ' Initialize all the member variables
    m_nLastPercX = 1
    m_nLastPercY = 1
    Set m_colOrig = New Collection
    Set m_frmParent = frmParent
    
    Dim ctrl As Control
    Dim tmp As SizeAndPos
    Dim sTemp As String
    
    On Error Resume Next
    
    ' Loop through each control in the controls on the parent form,
    ' and record the starting size and position.  Making sure to
    ' handle controls that don't have a property, such as FontSize.
    
    For Each ctrl In m_frmParent.Controls
        ' Since tmp is used in each iteration, make sure and reset it
        ' to sane values before using it again
        tmp.LocLeft = 0
        tmp.LocTop = 0
        tmp.SizeHeight = 0
        tmp.SizeWidth = 0
        tmp.FontSize = 0

        ' Attempt to read each property, if there's an error,
        ' just put -1 in the UDT
        Err.Clear
        tmp.LocLeft = ctrl.Left
        If Err Then
            tmp.LocLeft = -1
            Err.Clear
        End If
        
        tmp.LocTop = ctrl.Top
        If Err Then
            tmp.LocTop = -1
            Err.Clear
        End If
        
        tmp.SizeHeight = ctrl.Height
        If Err Then
            tmp.SizeHeight = -1
            Err.Clear
        End If
        
        tmp.SizeWidth = ctrl.Width
        If Err Then
            tmp.SizeWidth = -1
            Err.Clear
        End If
        
        tmp.FontSize = ctrl.FontSize
        If Err Then
            tmp.FontSize = -1
            Err.Clear
        End If
        
        'Store the UDT for later use
        m_colOrig.Add PackUDT(tmp), NameOfControl(ctrl)
    Next

    'Make note of the form's starting size
    m_sizeOrig.FontSize = m_frmParent.FontSize
    m_sizeOrig.SizeHeight = m_frmParent.Height
    m_sizeOrig.SizeWidth = m_frmParent.Width

End Sub

Private Function NameOfControl(ByVal ctrl As Control) As String
    ' Simple little function that returns the name of a control,
    ' adding on the index of the control if it has one. Used as
    ' the key in m_colOrig

    On Error Resume Next
    
    NameOfControl = "Control:" & ctrl.Name
    NameOfControl = NameOfControl & ":"
    NameOfControl = NameOfControl & Format(ctrl.Index)
    If Err Then
        Err.Clear
        NameOfControl = NameOfControl & "-1"
    End If

End Function

Private Sub m_frmParent_Resize()
    ' Because m_frmParent is declared WithEvents, this will get
    ' called when the parent form is resized, just like Form_Resize

    If m_frmParent.WindowState = vbMinimized Then
        'Don't do anything if the form is minimized
        Exit Sub
    End If

    Dim nPercX As Double
    Dim nPercY As Double
    Dim tmp As SizeAndPos
    Dim sTemp As String
    Dim ctrl As Control
    
    ' What's the percentage of change from the original?
    nPercX = m_frmParent.Width / m_sizeOrig.SizeWidth
    nPercY = m_frmParent.Height / m_sizeOrig.SizeHeight
    
    ' If the percentage of change exactly matches the last seen
    ' values, don't bother to do anything
    If nPercX = m_nLastPercX And nPercY = m_nLastPercY Then
        Exit Sub
    End If
    
    m_nLastPercX = nPercX
    m_nLastPercY = nPercY
    
    On Error Resume Next
    
    ' Loop through each of the controls, and change it's size and
    ' position to mirror the new size of the form.  Gracefully handle
    ' controls that have appeared since we initialized, or controls
    ' that can't be moved or resized for whatever reason.
    For Each ctrl In m_frmParent.Controls
        
        Err.Clear
        sTemp = m_colOrig.Item(NameOfControl(ctrl))
        If Err Then
            Err.Clear
        Else
            UnPackUDT tmp, sTemp
            
            If tmp.FontSize <> -1 Then
            ' Font size is an odd one.  Should we use the percentage of
            ' change for the width or the height?  I decided to simply
            ' ues the lesser, whichever that may be, but the situtation
            ' might call for something different.
                ctrl.FontSize = tmp.FontSize * _
                   IIf(nPercX < nPercY, nPercX, nPercY)
            End If
            
            ' Here we're protected by an On Error, so if the property
            ' is not -1, just try to change it, if it fails, nothing
            ' we can do, so simply move on to the next property.
            
            If tmp.LocLeft <> -1 Then
                ctrl.Left = tmp.LocLeft * nPercX
            End If
            If tmp.LocTop <> -1 Then
                ctrl.Top = tmp.LocTop * nPercY
            End If
            If tmp.SizeHeight <> -1 Then
                ctrl.Height = tmp.SizeHeight * nPercY
            End If
            If tmp.SizeWidth <> -1 Then
                ctrl.Width = tmp.SizeWidth * nPercX
            End If
        End If
    Next

End Sub

' End of Class Module: clsAutoSize
'---------------------------------------------------------------------


 
  
Sample Usage: 
  
 Option Explicit

Private m_AutoSize As clsAutoSize

Private Sub Form_Load()
    
    ' To fully appreciate this example, place a good number of controls on
    ' the example form, and after it's loaded and running, resize the form.
    
    Set m_AutoSize = New clsAutoSize
    m_AutoSize.Initialize Me
    
End Sub

 


'------------------------------------------------------------------
'------------------------------------------------------------------
'------------------------------------------------------------------
'------------------------------------------------------------------


BrowseForFolder 

Description: 
 This displays the browse for folder common dialog via API's, thus allowing VB easy access to it. 
  
Code: 
 Option Explicit

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal _
   lpString1 As String, ByVal lpString2 As String) As Long
   
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As _
   BrowseInfo) As Long
   
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal _
   pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(hWndOwner As Long, sPrompt As _
   String) As String

    Dim nNull As Integer
    Dim lpIDList As Long
    Dim nResult As Long
    Dim sPath As String
    Dim bi As BrowseInfo

    bi.hWndOwner = hWndOwner
    bi.lpszTitle = lstrcat(sPrompt, "")
    bi.ulFlags = BIF_RETURNONLYFSDIRS

    lpIDList = SHBrowseForFolder(bi)
    
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        nResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        nNull = InStr(sPath, vbNullChar)
        If nNull Then
            sPath = Left$(sPath, nNull - 1)
        End If
    End If

    BrowseForFolder = sPath

End Function

 
  
Sample Usage: 
  
     Debug.Print BrowseForFolder(Me.hWnd, "Please select a folder:")
 
'------------------------------------------------
'------------------------------------------------
'------------------------------------------------
'------------------------------------------------
'------------------------------------------------



Color Button 

Description: 
 This code usings a windows hook to process onwer drawing of a command button. To use the sample code, add two buttons, Command1 and Command2 to a standard form, and add the code to the Form_Load procedure. 
  
Code: 
 '------------------------
' Module: modHook

Option Explicit

Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

Public Const TRANSPARENT = 1
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20
Public Const PS_SOLID = 0
Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)
Public Const ODS_SELECTED = &H1
Public Const ODS_DISABLED = &H4
Public Const ODS_FOCUS = &H10

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long
   
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
   (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal message As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
   
Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest _
   As Any, src As Any, ByVal length As Long) As Long
   
Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect _
   As RECT) As Long
   
Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect _
   As RECT) As Long
   
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As _
   Long, ByVal y As Long) As Long
   
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) _
   As Long
   
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As _
   RECT, ByVal hBrush As Long) As Long
   
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) _
   As Long

Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal _
   crColor As Long) As Long
   
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal _
   nBkMode As Long) As Long
   
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc _
   As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect _
   As RECT, ByVal wFormat As Long) As Long
   
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal _
   nWidth As Long, ByVal crColor As Long) As Long
   
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal _
   hObject As Long) As Long
   
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x _
   As Long, ByVal y As Long) As Long
   
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x _
   As Long, ByVal y As Long, lpPoint As Any) As Long
   
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long

Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
   ByVal x As Long, ByVal y As Long) As Long


Public Const BS_OWNERDRAW = &HB&
Public Const WM_DRAWITEM = &H2B
Public Const WM_NCDESTROY = &H82

Private m_hWndParent As Long
Private m_lpPrevWndProc As Long
Private m_bHooked As Boolean
Private m_colButtons As New Collection

Public Sub DefineColor(frmParent As Form, cmdButton As CommandButton, _
   Optional colorFore As Long = vbButtonText, Optional colorBack _
   As Long = vbButtonFace, Optional colorDisable As Long = _
   vbButtonShadow, Optional nBevel As Long = 2)

    Debug.Assert cmdButton.Style = 1
       'Command button /must/ be set to Style = 1 - vbGraphical

    Dim tmp As New clsButtonInfo
    tmp.colorBack = IIf(colorBack < 0, _
       GetSysColor(Val("&H" & Mid(Hex(colorBack), 2))), colorBack)
       
    tmp.colorFore = IIf(colorFore < 0, _
       GetSysColor(Val("&H" & Mid(Hex(colorFore), 2))), colorFore)
       
    tmp.colorDisable = IIf(colorDisable < 0, _
       GetSysColor(Val("&H" & Mid(Hex(colorDisable), 2))), colorDisable)
       
    Set tmp.cmdButton = cmdButton
    tmp.nBevel = nBevel
    m_colButtons.Add tmp, "-" & Str(GetDlgCtrlID(cmdButton.hwnd))

    Dim nStyle As Long
    
    nStyle = GetWindowLong(cmdButton.hwnd, GWL_STYLE)
    nStyle = nStyle Or BS_OWNERDRAW
    SetWindowLong cmdButton.hwnd, GWL_STYLE, nStyle
   
    If Not m_bHooked Then
        Hook frmParent
    End If

End Sub

Public Sub Hook(frm As Form)

    m_bHooked = True

    m_hWndParent = frm.hwnd
    m_lpPrevWndProc = GetWindowLong(m_hWndParent, GWL_WNDPROC)
    SetWindowLong m_hWndParent, GWL_WNDPROC, AddressOf WndProc

End Sub

Public Sub Unhook()

    m_bHooked = False

    SetWindowLong m_hWndParent, GWL_WNDPROC, m_lpPrevWndProc

End Sub

Private Function WndProc(ByVal hwnd As Long, ByVal umessage As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim bPrevProc As Boolean
    bPrevProc = True

    Select Case umessage
        Case WM_DRAWITEM
            On Error Resume Next
            Err.Clear
            m_colButtons.Item "-" & Str(wParam)
            If Err Then
                bPrevProc = True
            Else
                DrawItem lParam, wParam
                bPrevProc = False
            End If
            Err.Clear
        Case WM_NCDESTROY
            Unhook
    End Select
    
    If bPrevProc Then
        WndProc = CallWindowProc(m_lpPrevWndProc, hwnd, _
           umessage, wParam, lParam)
    End If

End Function

Private Sub DrawItem(lpDIS As Long, ID As Long)

    Dim DIS As DRAWITEMSTRUCT
    CopyMemory DIS, ByVal lpDIS, Len(DIS)

    Dim rectFocus As RECT
    Dim rectButton As RECT
    
    CopyRect rectFocus, DIS.rcItem
    CopyRect rectButton, DIS.rcItem
    
    rectFocus.left = rectFocus.left + 4
    rectFocus.right = rectFocus.right - 4
    rectFocus.top = rectFocus.top + 4
    rectFocus.bottom = rectFocus.bottom - 4
    
    DrawFilledRect DIS.hdc, rectButton, _
       m_colButtons.Item("-" & Str(ID)).colorBack
       
    DrawFrame DIS.hdc, rectButton, _
       m_colButtons.Item("-" & Str(ID)).nBevel
    
    DrawButtonText DIS.hdc, rectButton, _
       m_colButtons.Item("-" & Str(ID)).cmdButton.Caption, _
       m_colButtons.Item("-" & Str(ID)).colorFore, DIS.itemState
    
    If DIS.itemState And ODS_FOCUS Then
        DrawFocusRect DIS.hdc, rectFocus
        If DIS.itemState And ODS_SELECTED Then
            DrawFilledRect DIS.hdc, rectButton, _
               m_colButtons.Item("-" & Str(ID)).colorBack
               
            DrawFrame DIS.hdc, rectButton, _
               -m_colButtons.Item("-" & Str(ID)).nBevel
               
            DrawButtonText DIS.hdc, rectButton, _
               m_colButtons.Item("-" & Str(ID)).cmdButton.Caption, _
               m_colButtons.Item("-" & Str(ID)).colorFore, _
               DIS.itemState
               
            DrawFocusRect DIS.hdc, rectFocus
        End If
        
    ElseIf DIS.itemState And ODS_DISABLED Then
    
        DrawButtonText DIS.hdc, rectButton, _
           m_colButtons.Item("-" & Str(ID)).cmdButton.Caption, _
           m_colButtons.Item("-" & Str(ID)).colorDisable, _
           DIS.itemState
        
    End If

End Sub

Private Sub DrawFrame(hdc As Long, r As RECT, Inset As Long)

    Dim colorDark As Long
    Dim colorLight As Long
    Dim colorTL As Long
    Dim colorBR As Long
    
    Dim I As Long
    Dim m As Long
    Dim width As Long
    
    width = Abs(Inset)

    For I = 0 To width - 1
        m = 255 / (i + 2)
        colorDark = PaletteRGB(m, m, m)
        m = 192 + (63 / (i + 1))
        colorLight = PaletteRGB(m, m, m)
        
        If width = 1 Then
            colorLight = RGB(255, 255, 255)
            colorDark = RGB(128, 128, 128)
        End If
        
        If Inset < 0 Then
            colorTL = colorDark
            colorBR = colorLight
        Else
            colorTL = colorLight
            colorBR = colorDark
        End If

        DrawLine2 hdc, r.left, r.top, r.right, r.top, colorTL
        DrawLine2 hdc, r.left, r.top, r.left, r.bottom, colorTL
        
        If (Inset < 0) And (i = width - 1) And (width > 1) Then
            DrawLine2 hdc, r.left + 1, r.bottom - 1, _
               r.right, r.bottom - 1, RGB(1, 1, 1)
            DrawLine2 hdc, r.right - 1, r.top + 1, _
               r.right - 1, r.bottom, 0
        Else
            DrawLine2 hdc, r.left + 1, r.bottom - 1, _
               r.right, r.bottom - 1, colorBR
            DrawLine2 hdc, r.right - 1, r.top + 1, _
               r.right - 1, r.bottom, colorBR
        End If
        
        InflateRect r, -1, -1
        
    Next
    
End Sub

Private Sub DrawFilledRect(hdc As Long, r As RECT, color As Long)
    
    Dim brush As Long
    brush = CreateSolidBrush(color)
    FillRect hdc, r, brush
    DeleteObject brush
    
End Sub

Private Sub DrawLine(hdc As Long, rectEndPoints As RECT, color As Long)
    Dim penNew As Long
    Dim penOld As Long
    
    penNew = CreatePen(PS_SOLID, 1, color)
    penOld = SelectObject(hdc, penNew)
    
    MoveToEx hdc, rectEndPoints.left, rectEndPoints.top, ByVal 0
    LineTo hdc, rectEndPoints.right, rectEndPoints.bottom
    
    SelectObject hdc, penOld
    
    DeleteObject penNew
    
End Sub

Private Sub DrawLine2(hdc As Long, left As Long, top As Long, _
   right As Long, bottom As Long, color As Long)
   
    Dim penNew As Long
    Dim penOld As Long
    
    penNew = CreatePen(PS_SOLID, 1, color)
    penOld = SelectObject(hdc, penNew)
    
    MoveToEx hdc, left, top, ByVal 0
    LineTo hdc, right, bottom
    
    SelectObject hdc, penOld
    
    DeleteObject penNew
    

End Sub

Private Sub DrawButtonText(hdc As Long, r As RECT, _
   strText As String, color As Long, state As Long)

    Dim colorPrev As Long
    colorPrev = SetTextColor(hdc, color)
    
    SetBkMode hdc, TRANSPARENT
    
    Dim rectTemp As RECT
    CopyRect rectTemp, r
    
    If state And ODS_SELECTED Then
        OffsetRect rectTemp, 1, 1
    End If
        
    
    DrawText hdc, strText, Len(strText), rectTemp, _
       DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    
    SetTextColor hdc, colorPrev

End Sub

Private Function PaletteRGB(r As Long, g As Long, b As Long) As Long

    PaletteRGB = RGB(r, g, b) Or &H2000000
 
End Function

' End of Module
'------------------------

'------------------------
' Class: clsButtonInfo

Public colorFore As Long
Public colorBack As Long
Public colorDisable As Long
Public nBevel As Long
Public cmdButton As CommandButton

' End of class
'-------------------------

 
  
Sample Usage: 
  
     DefineColor Me, Me.Command1, vbRed
    DefineColor Me, Me.Command2, vbBlue

 

'---------------------------------------------------------
'---------------------------------------------------------
'---------------------------------------------------------
'---------------------------------------------------------



ListView_AutoSize 

Description: 
 This is the programmatic equivalent of hitting <Ctrl>+<Gray Plus> on a ListView. This also demonstrates how to get around an apparent bug in the ListView control when setting and getting column widths by using API calls to do the same. 
  
Code: 
 Option Explicit

Private Const LVM_FIRST = &H1000
Private Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE = -1
Private Const SM_CXVSCROLL = 2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
   ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, lParam _
   As Any) As Long
   
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As _
   Long) As Long
   
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock _
   As Long) As Long
   
Public Sub ListView_AutoSize( _
   lvToResize As ListView, _
   Optional bLastColumn As Boolean = True, _
   Optional bLockWin As Boolean = True)
    
    Debug.Assert lvToResize.View = lvwReport
    
    
    Dim nColumn As Long
    Dim nNewWidth As Long

    If bLockWin Then
        LockWindowUpdate lvToResize.hWnd
    End If
    
    For nColumn = 0 To lvToResize.ColumnHeaders.Count - 1
        SendMessage lvToResize.hWnd, LVM_SETCOLUMNWIDTH, nColumn, _
                    LVSCW_AUTOSIZE
    Next

    If bLastColumn Then
        For nColumn = 0 To lvToResize.ColumnHeaders.Count - 2
            nNewWidth = nNewWidth + SendMessage(lvToResize.hWnd, _
                        LVM_GETCOLUMNWIDTH, nColumn, ByVal 0)
        Next
        
        nNewWidth = nNewWidth + GetSystemMetrics(SM_CXVSCROLL)
        nNewWidth = ((lvToResize.Width / Screen.TwipsPerPixelX) - _
                    nNewWidth) - 6
        nColumn = lvToResize.ColumnHeaders.Count - 1
        If nNewWidth > 0 Then
            SendMessage lvToResize.hWnd, LVM_SETCOLUMNWIDTH, nColumn, _
                        ByVal nNewWidth
        End If
    End If

    If bLockWin Then
        LockWindowUpdate 0
    End If

End Sub



 
  
Sample Usage: 
  
 ListView_AutoSize ListView1
 
'---------------------------------------------------------------------
'---------------------------------------------------------------------
'---------------------------------------------------------------------
'---------------------------------------------------------------------


LinesIntersect 

Description: 
 This bit of code shows how to determine the point that two lines intersect at. Pretty straightforward code, just a matter of thinking back to all of those geometry classes. See the sample code for usage. 
  
Code: 
 Option Explicit

Private Type PointType
    X As Double
    Y As Double
End Type

Private Type LineType
    Pt1 As PointType
    Pt2 As PointType
End Type
    
Private Sub LinesIntersect(ln1 As LineType, ln2 As LineType, _
   pt As PointType, bParallel As Boolean)

    Dim bVertical(1 To 2) As Boolean
    Dim m(1 To 2) As Double
    Dim b(1 To 2) As Double
    
    'Check to see if either of the lines is vertical
    bVertical(1) = (ln1.Pt1.X = ln1.Pt2.X)
    bVertical(2) = (ln2.Pt1.X = ln2.Pt2.X)
    
    bParallel = False
    If bVertical(1) And bVertical(2) Then
        'Both lines are vertical, so they're parallel
        bParallel = True
    ElseIf bVertical(1) Then
        'Calculate the non-vertical line's slope and offset
        m(2) = (ln2.Pt2.Y - ln2.Pt1.Y) / (ln2.Pt2.X - ln2.Pt1.X)
        b(2) = ln2.Pt1.Y - (m(2) * ln2.Pt1.X)
        'and use it to determine the intersection point
        pt.X = ln1.Pt1.X
        pt.Y = m(2) * pt.X + b(2)
    ElseIf bVertical(2) Then
        'Calculate the non-vertical line's slope and offset
        m(1) = (ln1.Pt2.Y - ln1.Pt1.Y) / (ln1.Pt2.X - ln1.Pt1.X)
        b(1) = ln1.Pt1.Y - (m(1) * ln1.Pt1.X)
        'and use it to determine the intersection point
        pt.X = ln2.Pt1.X
        pt.Y = m(1) * pt.X + b(1)
    Else
        'Calculate the slopes of both lines
        m(1) = (ln1.Pt2.Y - ln1.Pt1.Y) / (ln1.Pt2.X - ln1.Pt1.X)
        m(2) = (ln2.Pt2.Y - ln2.Pt1.Y) / (ln2.Pt2.X - ln2.Pt1.X)
    
        If m(1) = m(2) Then
            'The slopes are equal, so the lines are parallel
            bParallel = True
        Else
            'They differ, so calculate the offset of the lines
            b(1) = ln1.Pt1.Y - (m(1) * ln1.Pt1.X)
            b(2) = ln2.Pt1.Y - (m(2) * ln2.Pt1.X)
            
            'Now, calculate the intersection point
            pt.X = (b(2) - b(1)) / (m(1) - m(2))
            pt.Y = m(1) * pt.X + b(1)
        End If
    End If
       
End Sub

 
  
Sample Usage: 
  
     'Requires two lines (Line1, Line2) and one shape (Shape1)

    Randomize Timer

    Dim bParallel As Boolean
    Dim pt As PointType
    Dim ln1 As LineType
    Dim ln2 As LineType
    
    'Randomly position the lines
    Line1.X1 = Rnd * ScaleWidth
    Line1.X2 = Rnd * ScaleWidth
    Line2.X1 = Rnd * ScaleWidth
    Line2.X2 = Rnd * ScaleWidth
    Line1.Y1 = Rnd * ScaleHeight
    Line1.Y2 = Rnd * ScaleHeight
    Line2.Y1 = Rnd * ScaleHeight
    Line2.Y2 = Rnd * ScaleHeight
    
    'Set our line type to the lines' position
    ln1.Pt1.X = Line1.X1
    ln1.Pt1.Y = Line1.Y1
    ln1.Pt2.X = Line1.X2
    ln1.Pt2.Y = Line1.Y2
    ln2.Pt1.X = Line2.X1
    ln2.Pt1.Y = Line2.Y1
    ln2.Pt2.X = Line2.X2
    ln2.Pt2.Y = Line2.Y2
    
    'Calculate the point of intersection
    LinesIntersect ln1, ln2, pt, bParallel
    
    If bParallel Then
        messagebox "Lines are parallel or the same line."
    Else
        Shape1.Width = 90
        Shape1.Height = 90
        Shape1.Top = pt.Y - Shape1.Height / 2
        Shape1.Left = pt.X - Shape1.Width / 2
    End If

 
'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------


MixPictures 

Description: 
 This is some sample code that mixes the contents of two pictures boxes together. It's not the most effecient, but it gets the job done quicker than any sample I could find online. 
  
Code: 
 Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
       ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
       ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) _
       As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
       (Destination As Any, Source As Any, ByVal Length As Long)

Private Type RGBAComponent
   Red As Byte
   Green As Byte
   Blue As Byte
   Alpha As Byte
End Type

Sub MixPictures(pctTarget As PictureBox, pctSource1 As PictureBox, _
   nSrcX1 As Long, nSrcY1 As Long, nWidth As Long, nHeight As Long, _
   pctSource2 As PictureBox, nSrcX2 As Long, nSrcY2 As Long, _
   nValue1 As Long)

   Dim nValue2 As Long
   nValue2 = 255 - nValue1

   Dim nCurColor1 As Long
   Dim nCurColor2 As Long
   Dim rgba1 As RGBAComponent
   Dim rgba2 As RGBAComponent

   Dim X As Long
   Dim Y As Long

   For Y = nSrcY1 To nSrcY1 + nHeight
       For X = nSrcX1 To nSrcX1 + nWidth
       
           nCurColor1 = GetPixel(pctSource1.hdc, X, Y)
           nCurColor2 = GetPixel(pctSource2.hdc, _
               X + nSrcX2 - nSrcX1, Y + nSrcY2 - nSrcY1)

           CopyMemory rgba1, nCurColor1, 4
           CopyMemory rgba2, nCurColor2, 4

           SetPixel pctTarget.hdc, X - nSrcX1, Y - nSrcY1, _
             RGB((rgba1.Red * nValue1 + rgba2.Red * nValue2) \ 255, _
             (rgba1.Green * nValue1 + rgba2.Green * nValue2) \ 255, _
             (rgba1.Blue * nValue1 + rgba2.Blue * nValue2) \ 255)
       Next
   Next
   
End Sub

 
  
Sample Usage: 
  
     MixPictures Picture3, Picture1, 0, 0, Picture1.Width / _
        Screen.TwipsPerPixelX, Picture2.Height / _
        Screen.TwipsPerPixelY, Picture2, 0, 0, HScroll1.Value

 
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------


PointsToAngle 

Description: 
 This function calculates the angle of rotation that one point is around a given center point. This probably isn't the most efficient formula possible, but I've found it quite useful. 
  
Code: 
 Option Explicit

Public Const PI = 3.14159265358979

Public Function PointsToAngle(ByVal xOne As Double, ByVal yOne As _
       Double, ByVal X As Double, ByVal Y As Double) As Double

    If xOne = X And Y = yOne Then
        Exit Function
    End If

    X = X - xOne
    Y = Y - yOne
    
    Dim nRatio As Double
    Dim nASin As Double
    nRatio = X / Sqr(Y ^ 2 + X ^ 2)
    
    
    If Abs(nRatio) = 1 Then
        nASin = 90 * nRatio
    Else
        nASin = Atn(nRatio / Sqr(-nRatio * nRatio + 1)) * (180 / PI)
    End If

    If Y < 0 Then
        nASin = 90 - nASin + 90
    ElseIf X < 0 Then
        nASin = nASin + 360
    End If

    PointsToAngle = nASin

End Function

 
  
Sample Usage: 
  
 Debug.Print PointsToAngle(1, 1, 9, 9)

 
'----------------------------------------------------------
'----------------------------------------------------------
'----------------------------------------------------------
'----------------------------------------------------------
'----------------------------------------------------------


WebSafeColor 

Description: 
 WebSafeColor converts a given color to the closest match in a Web Safe Color palette. It can handle normal RGB colors, as well as dealing with VB's system color constants. 
  
Code: 
 Option Explicit

Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Function WebSafeColor(ByRef nColor As OLE_COLOR) As OLE_COLOR

    If nColor < 0 Then
        ' If the color is negative, then it's one of the
        '  SystemColorConstants, so look up it's real value
        nColor = GetSysColor(nColor And &H7FFFFFFF)
    End If

    ' Call WebSafeColorSingle for each color
    '  value and add up the results
    WebSafeColor = WebSafeColorSingle(nColor And &HFF&) + _
       (WebSafeColorSingle((nColor And &HFF00&) \ &H100&) * _
       &H100&) + (WebSafeColorSingle((nColor And &HFF0000) \ _
       &H10000) * &H10000)

End Function

Public Function WebSafeColorSingle(nColor As Long) As Long

    ' The color isn't a valid color value
    Debug.Assert nColor >= 0
    Debug.Assert nColor <= 255

    ' Return the web color closest to the input color
    WebSafeColorSingle = ((nColor + 25) \ 51) * 51

End Function

 
  
Sample Usage: 
  
     Debug.Print Hex(WebSafeColor(vbButtonFace))
    Debug.Print Hex(WebSafeColor(RGB(50, 100, 150)))

 
'---------------------------------------------------------
'---------------------------------------------------------
'---------------------------------------------------------
'---------------------------------------------------------


EnumWindows 

Description: 
 This module demonstrates calling the EnumWindows and EnumChildWindows APIs by using these functions to fill a TreeView similar to Spy++'s main screen. To use it, create a form, named frmMain, and place a TreeView on that form and name the TreeView tvMain. Then call the function BeginEnum to populate the TreeView with all the windows on the system. 
  
Code: 
 Option Explicit

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As _
   Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
   
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, _
   ByVal lParam As Long) As Long
   
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
   ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) _
   As Long
   
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
   hwnd As Long, ByVal wmessage As Long, ByVal wParam As Long, lParam As _
   Any) As Long
   
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   
Declare Function GetDesktopWindow Lib "user32" () As Long

Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal _
   hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) _
   As Long
   
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect _
   As RECT) As Long

Public Const WM_GETTEXT = &HD
Public Const GWL_HWNDPARENT = (-8)

Public Sub BeginEnum()

    frmMain.tvMain.Nodes.Clear
    ProcessEnum GetDesktopWindow, -1

    EnumWindows AddressOf EnumProc, 0

End Sub

Private Function EnumProc(ByVal hwnd As Long, ByVal lParam As Long) _
   As Long

    Dim hWndParent As Long
    hWndParent = GetWindowLong(hwnd, GWL_HWNDPARENT)
    
    If hWndParent = lParam Then
        ProcessEnum hwnd, lParam
    End If

    EnumProc = True

End Function

Private Sub ProcessEnum(hwnd As Long, hWndParent As Long)

    Dim sTreeCaption As String
    Dim sWindowClass As String
    Dim sWindowText As String
    Dim nCount As Long
    Dim rt As RECT
    
    sWindowClass = Space(80)
    nCount = GetClassName(hwnd, sWindowClass, 80)
    sWindowClass = Mid(sWindowClass, 1, nCount)
    
    sWindowText = Space(80)
    nCount = GetWindowText(hwnd, sWindowText, 80)
    sWindowText = Mid(sWindowText, 1, nCount)
    
    If sWindowText = "" Then
        sWindowText = Space(80)
        nCount = SendMessage(hwnd, WM_GETTEXT, 80, ByVal sWindowText)
        sWindowText = Mid(sWindowText, 1, nCount)
    End If
        
    GetWindowRect hwnd, rt
    
    sTreeCaption = Chr(34) & sWindowText & Chr(34) & " - " & _
       sWindowClass & ", (" & Str(rt.Left) & "," & Str(rt.Top) & "," & _
       Str(rt.Right) & "," & Str(rt.Bottom) & "), 0x" & _
       Right("00000000" & Hex(hwnd), 8)
    
    If hWndParent = -1 Then

        frmMain.tvMain.Nodes.Add , , "# 0", sTreeCaption
    
    Else

        frmMain.tvMain.Nodes.Add "#" & Str(hWndParent), tvwChild, _
           "#" & Str(hwnd), sTreeCaption
    
        EnumChildWindows hwnd, AddressOf EnumProc, hwnd
        
    End If

End Sub

 
  
Sample Usage: 
  
 n/a
 
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------



LookupIPAddress 

Description: 
 This is a simple VB wrapper function for the Winsock API gethostbyname, but it can be rather useful. And, it's much cleaner than using the VB winsock control just to look up an IP address. 
  
Code: 
 Option Explicit

Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYS_STATUS_LEN = 128

Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(WSADESCRIPTION_LEN) As Byte
    szSystemStatus(WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal _
   szName As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal _
   wVersionRequested As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias _
   "RtlMoveMemory" (Destination As Any, Source As Any, ByVal _
   Length As Long)

Public Function LookupIPAddress(ByVal sHostName As String) As String

    Dim wsa As WSADATA
    Dim nRet As Long
    Dim nTemp As Long
    Dim bTemp(0 To 3) As Byte
    Dim sOut As String
    Dim he As HOSTENT
    
    'Initialize WinSock
    WSAStartup &H10, wsa
        
    'Attempt to lookup the host
    nRet = gethostbyname(sHostName)
    
    'If it failed, just return nothing
    If nRet = 0 Then
        sOut = ""
    Else
        'Take a look at the resulting hostent structure
        CopyMemory he, ByVal nRet, Len(he)
        
        'Are there atlest four bytes, then we have
        ' at least one address
        If he.h_length >= 4 Then
            'Copy the address out,
            CopyMemory nTemp, ByVal he.h_addr_list, 4
            CopyMemory bTemp(0), ByVal nTemp, 4
            ' and format it
            sOut = Format(bTemp(0)) & "." & Format(bTemp(1)) & "." _
               & Format(bTemp(2)) & "." & Format(bTemp(3))
        Else
            sOut = ""
        End If
        
    End If
    
    WSACleanup

    LookupIPAddress = sOut

End Function
 
  
Sample Usage: 
  
 Debug.Print LookupIPAddress("localhost")
 


'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------



IndentCode 

Description: 
 This function re-indents code passed to it, based on where blocks of code should be. I find this very useful when I'm trying to follow someone else's code that doesn't use tabs (or uses them oddly). Note: This code uses the clsStringCollection class, it's available under the Strings->Split on my sample code page. 
  
Code: 
 Option Explicit

Private Const TAB_SIZE = 4

Public Function IndentCode(sCode As String) As String

    Dim sc As clsStringCollection
    Set sc = New clsStringCollection
    sc.SplitInto sCode, vbNewLine
    
    Dim nCurLine As Long
    Dim nCurTab As Long
    Dim sLine As String
    Dim vItem As Variant
    Dim bEscapeEOL As Boolean
    Dim nCount As Long
    Dim nEOLTab As Long
    Dim vLine As Variant
    Dim sOut As String
    
    Dim colIncrease As Collection
    Dim colDecrease As Collection
    Dim colInAll As Collection
    Dim colDeAll As Collection
    Set colIncrease = New Collection
    Set colDecrease = New Collection
    Set colInAll = New Collection
    Set colDeAll = New Collection
    
    colIncrease.Add "If "
    colIncrease.Add "With "
    colIncrease.Add "Select Case "
    colIncrease.Add "Select Case "
    colIncrease.Add "While"
    colIncrease.Add "Do"
    colIncrease.Add "Private Sub "
    colIncrease.Add "Private Function "
    colIncrease.Add "Private Property "
    colIncrease.Add "Public Sub "
    colIncrease.Add "Public Function "
    colIncrease.Add "Public Property "
    colIncrease.Add "Sub "
    colIncrease.Add "Function "
    colIncrease.Add "Property "
    colIncrease.Add "For "
    colIncrease.Add "Type "
    colIncrease.Add "Public Type "
    colIncrease.Add "Private Type "
    
    colDecrease.Add "End If"
    colDecrease.Add "End Select"
    colDecrease.Add "End Select"
    colDecrease.Add "End With"
    colDecrease.Add "Wend"
    colDecrease.Add "Loop"
    colDecrease.Add "End Sub"
    colDecrease.Add "End Function"
    colDecrease.Add "End Property"
    colDecrease.Add "Next"
    colDecrease.Add "End Type"
    
    colDecrease.Add "Else"
    colDecrease.Add "Case"
    colIncrease.Add "Else"
    colIncrease.Add "Case"
    
    nCurTab = 0
    bEscapeEOL = False
    nCount = sc.Count
    
    For Each vLine In sc
        sLine = Trim(vLine)
        
        For Each vItem In colDecrease
            If LCase(Mid(sLine, 1, Len(vItem))) = LCase(vItem) Then
                nCurTab = nCurTab - TAB_SIZE
            End If
        Next
        
        If nCurTab < 0 Then
            nCurTab = 0
            sLine = sLine & " ' TAB COUNT ERROR"
        End If
        
        
        sOut = sOut & Space(nCurTab) & sLine & vbNewLine
        
        If Right(sLine, 2) = " _" Then
            If Not bEscapeEOL Then
                nEOLTab = InStr(1, Trim(sLine), " ")
                If nEOLTab < 3 Then
                    nEOLTab = 3
                End If
                nCurTab = nCurTab + nEOLTab
                bEscapeEOL = True
            End If
        Else
            If bEscapeEOL Then
                nCurTab = nCurTab - nEOLTab
                bEscapeEOL = False
            End If
        End If
        
        For Each vItem In colIncrease
            If LCase(Mid(sLine, 1, Len(vItem))) = LCase(vItem) Then
                nCurTab = nCurTab + TAB_SIZE
            End If
        Next
        
    Next
    
    IndentCode = sOut
    
End Function

 
  
Sample Usage: 
  
     Dim sTemp As String
    sTemp = IndentCode(Clipboard.GetText)
    
    Clipboard.Clear
    Clipboard.SetText sTemp

 
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------



NumToString 

Description: 
 The function NumToString writes out any number, up to about 920 trillion, the limit of the Currency variable type, in English words. It only writes out the integer portion of the number. The DollarToString function does the same, but places the word dollars after the string and also writes out the fractional part of the value as the cents.

1/8/2001: Added DateToString and NumToStringTh (ie, "23" -> "twenty-third") helper functions.

7/20/2002: Corrected error in DollarToString that limited the range of values. 
  
Code: 
 Public Function DollarToString(ByVal nAmount As Currency) As String
'Example:
'  DollarToString(5.99) = " five dollars and ninety-nine cents"
    Dim nDollar As Currency
    Dim nCent As Currency
    
    nDollar = Int(nAmount)
    nCent = (Abs(nAmount) - Int(Abs(nAmount))) * 100
    
    DollarToString = NumToString(nDollar) & " dollar"
    
    If Abs(nDollar) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
    DollarToString = DollarToString & " and" & _
    NumToString(nCent) & " cent"
    
    If Abs(nCent) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
End Function

Public Function NumToString(ByVal nNumber As Currency) As String
'Example: NumToString(123) = " one hundred twenty-three"
    Dim bNegative As Boolean
    Dim bHundred As Boolean

    If nNumber < 0 Then
        bNegative = True
    End If

    nNumber = Abs(Int(nNumber))
    If nNumber < 1000 Then
        If nNumber \ 100 > 0 Then
            NumToString = NumToString & _
                NumToString(nNumber \ 100) & " hundred"
            bHundred = True
        End If
        nNumber = nNumber - ((nNumber \ 100) * 100)
        Dim bNoFirstDigit As Boolean
        bNoFirstDigit = False
        Select Case nNumber \ 10
            Case 0
                Select Case nNumber Mod 10
                    Case 0
                        If Not bHundred Then
                            NumToString = NumToString & " zero"
                        End If
                    Case 1: NumToString = NumToString & " one"
                    Case 2: NumToString = NumToString & " two"
                    Case 3: NumToString = NumToString & " three"
                    Case 4: NumToString = NumToString & " four"
                    Case 5: NumToString = NumToString & " five"
                    Case 6: NumToString = NumToString & " six"
                    Case 7: NumToString = NumToString & " seven"
                    Case 8: NumToString = NumToString & " eight"
                    Case 9: NumToString = NumToString & " nine"
                End Select
                bNoFirstDigit = True
            Case 1
                Select Case nNumber Mod 10
                    Case 0: NumToString = NumToString & " ten"
                    Case 1: NumToString = NumToString & " eleven"
                    Case 2: NumToString = NumToString & " twelve"
                    Case 3: NumToString = NumToString & " thirteen"
                    Case 4: NumToString = NumToString & " fourteen"
                    Case 5: NumToString = NumToString & " fifteen"
                    Case 6: NumToString = NumToString & " sixteen"
                    Case 7: NumToString = NumToString & " seventeen"
                    Case 8: NumToString = NumToString & " eighteen"
                    Case 9: NumToString = NumToString & " nineteen"
                End Select
                bNoFirstDigit = True
            Case 2: NumToString = NumToString & " twenty"
            Case 3: NumToString = NumToString & " thirty"
            Case 4: NumToString = NumToString & " forty"
            Case 5: NumToString = NumToString & " fifty"
            Case 6: NumToString = NumToString & " sixty"
            Case 7: NumToString = NumToString & " seventy"
            Case 8: NumToString = NumToString & " eighty"
            Case 9: NumToString = NumToString & " ninety"
        End Select
        If Not bNoFirstDigit Then
            If nNumber Mod 10 <> 0 Then
                NumToString = NumToString & "-" & _
                    Mid(NumToString(nNumber Mod 10), 2)
            End If
        End If
    Else
        Dim nTemp As Currency
        nTemp = 10 ^ 12 'trillion
        Do While nTemp >= 1
            If nNumber >= nTemp Then
                NumToString = NumToString & _
                    NumToString(Int(nNumber / nTemp))
                Select Case Int(Log(nTemp) / Log(10) + 0.5)
                    Case 12: NumToString = NumToString & " trillion"
                    Case 9: NumToString = NumToString & " billion"
                    Case 6: NumToString = NumToString & " million"
                    Case 3: NumToString = NumToString & " thousand"
                End Select
                nNumber = nNumber - (Int(nNumber / nTemp) * nTemp)
            End If
            nTemp = nTemp / 1000
        Loop
    End If
    If bNegative Then
        NumToString = " negative " & NumToString
    End If
End Function

Public Function NumToStringTh(ByVal nNumber As Currency) As String
'Example: NumToStringTh(123) = " one hundred twenty-third"

    Dim sNum As String
    Dim sExtra As String
    Dim nSpace As String

    'Convert the number
    sNum = NumToString(nNumber)
    'Find the location of the last space or dash
    nSpace = Len(sNum)
    
    Do Until Mid(sNum, nSpace, 1) = " " Or _
        Mid(sNum, nSpace, 1) = "-"
        nSpace = nSpace - 1
    Loop

    sExtra = Mid(sNum, nSpace + 1)
    sNum = Mid(sNum, 1, nSpace)
    
    'Conver the last word ("one" -> "first", etc)
    Select Case sExtra
        Case "hundred":   NumToStringTh = sNum & "hundredth"
        Case "zero" 'no such thing as 'zeroth'
            NumToStringTh = sNum & "zero"
        Case "one":       NumToStringTh = sNum & "first"
        Case "two":       NumToStringTh = sNum & "second"
        Case "three":     NumToStringTh = sNum & "third"
        Case "four":      NumToStringTh = sNum & "fourth"
        Case "five":      NumToStringTh = sNum & "fifth"
        Case "six":       NumToStringTh = sNum & "sixth"
        Case "seven":     NumToStringTh = sNum & "seventh"
        Case "eight":     NumToStringTh = sNum & "eighth"
        Case "nine":      NumToStringTh = sNum & "ninth"
        Case "ten":       NumToStringTh = sNum & "tenth"
        Case "eleven":    NumToStringTh = sNum & "eleventh"
        Case "twelve":    NumToStringTh = sNum & "twelfth"
        Case "thirteen":  NumToStringTh = sNum & "thirteenth"
        Case "fourteen":  NumToStringTh = sNum & "fourteenth"
        Case "fifteen":   NumToStringTh = sNum & "fifteenth"
        Case "sixteen":   NumToStringTh = sNum & "sixteenth"
        Case "seventeen": NumToStringTh = sNum & "seventeenth"
        Case "eighteen":  NumToStringTh = sNum & "eighteenth"
        Case "nineteen":  NumToStringTh = sNum & "nineteenth"
        Case "twenty":    NumToStringTh = sNum & "twentieth"
        Case "thirty":    NumToStringTh = sNum & "thirtieth"
        Case "forty":     NumToStringTh = sNum & "fortieth"
        Case "fifty":     NumToStringTh = sNum & "fiftieth"
        Case "sixty":     NumToStringTh = sNum & "sixtieth"
        Case "seventy":   NumToStringTh = sNum & "seventieth"
        Case "eighty":    NumToStringTh = sNum & "eightieth"
        Case "ninety":    NumToStringTh = sNum & "ninetieth"
        Case "trillion":  NumToStringTh = sNum & "trillionth"
        Case "billion":   NumToStringTh = sNum & "billionth"
        Case "million":   NumToStringTh = sNum & "millionth"
        Case "thousand":  NumToStringTh = sNum & "thousandth"
        Case Else 'This shouldn't happen, but just in case
            NumToStringTh = NumToString(nNumber)
    End Select
    
End Function

Public Function DateToString(FromDate As String) As String
'Example: ' DateToString(#1/1/2001#) = "January the first, two thousand one"

    DateToString = Format(FromDate, "mmmm") & " the" & _
        NumToStringTh(DatePart("d", FromDate)) & "," & _
        NumToString(DatePart("yyyy", FromDate))
        
End Function

 
  
Sample Usage: 
  
     Debug.Print DollarToString(1234.11)
    Debug.Print NumToString(-54321)

 
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------



StringToNum 

Description: 
 The function StringToNum will take just about any number written out in english, up to about 920 trillion, or the limit of the Currency variable type, and convert it to a Currency number. StringToNum doesn't make any attempt to validate the number being passed, so it will convert some strings that make no sense. 
  
Code: 
 Public Function StringToNum(ByVal sNumber As String) As Currency
'Example: StringToNum("one hundred twenty-three") = 123

    Dim nLen As Long
    Dim nTemp As Currency
    Dim bNegative As Boolean
    
    sNumber = Replace(sNumber, "-", " ")
    sNumber = Replace(sNumber, ".", " ")
    sNumber = Replace(sNumber, ",", " ")
    Do
        nLen = Len(sNumber)
        sNumber = Replace(sNumber, "  ", " ")
    Loop Until Len(sNumber) = nLen
    sNumber = LCase(Trim(sNumber))

    Dim vNumbers As Variant
    Dim vNumber As Variant
    vNumbers = Split(sNumber, " ")
    For Each vNumber In vNumbers
        Select Case vNumber
            Case "zero"
                'nothing
            Case "negative"
                bNegative = Not bNegative
            Case "a": nTemp = nTemp + 1
            Case "one": nTemp = nTemp + 1
            Case "two": nTemp = nTemp + 2
            Case "three": nTemp = nTemp + 3
            Case "four": nTemp = nTemp + 4
            Case "five": nTemp = nTemp + 5
            Case "six": nTemp = nTemp + 6
            Case "seven": nTemp = nTemp + 7
            Case "eight": nTemp = nTemp + 8
            Case "nine": nTemp = nTemp + 9
            Case "ten": nTemp = nTemp + 10
            Case "eleven": nTemp = nTemp + 11
            Case "twelve": nTemp = nTemp + 12
            Case "thirteen": nTemp = nTemp + 13
            Case "fourteen": nTemp = nTemp + 14
            Case "fifteen": nTemp = nTemp + 15
            Case "sixteen": nTemp = nTemp + 16
            Case "seventeen": nTemp = nTemp + 17
            Case "eighteen": nTemp = nTemp + 18
            Case "nineteen": nTemp = nTemp + 19
            Case "twenty": nTemp = nTemp + 20
            Case "thirty": nTemp = nTemp + 30
            Case "forty": nTemp = nTemp + 40
            Case "fifty": nTemp = nTemp + 50
            Case "sixty": nTemp = nTemp + 60
            Case "seventy": nTemp = nTemp + 70
            Case "eighty": nTemp = nTemp + 80
            Case "ninety": nTemp = nTemp + 90
            Case "trillion"
                nTemp = nTemp * 1000000000000#
                StringToNum = StringToNum + nTemp
                nTemp = 0
            Case "billion"
                nTemp = nTemp * 1000000000
                StringToNum = StringToNum + nTemp
                nTemp = 0
            Case "million"
                nTemp = nTemp * 1000000
                StringToNum = StringToNum + nTemp
                nTemp = 0
            Case "thousand"
                nTemp = nTemp * 1000
                StringToNum = StringToNum + nTemp
                nTemp = 0
            Case "hundred"
                nTemp = nTemp * 100
            Case Else
                If IsNumeric(vNumber) Then
                    nTemp = CCur(vNumber)
                Else
                    Debug.Assert False 'Unknown value
                End If
        End Select
    Next
    
    StringToNum = StringToNum + nTemp
    
    If bNegative Then
        StringToNum = -StringToNum
    End If
    
End Function

 
  
Sample Usage: 
  
 Debug.Print StringToNum("Negative Four thousand three hundred twenty-one")
 
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------


WrapString 

Description: 
 This function simply wraps a string at a specified point, looking for spaces and hyphens to wrap the text around. It has some optional arguments that let you control how the wrapping occurs. 
  
Code: 
 Public Function WrapString(ByVal sString As String, _
       Optional nMaxWidth = 72, Optional nMargin = 10, _
       Optional nIndent = 0) As String
       
   'sString = The string to be wrapped
   'nMaxWidth = The maximum width of any line
   'nMargin = The maximum number of characters to be
   '          dropped to the next line
   'nIndent = Additional space to be added to new lines
   
   'Verify that the values make sense
   Debug.Assert nMaxWidth > 1
   Debug.Assert nMargin > 0
   Debug.Assert nMargin < nMaxWidth
   Debug.Assert nIndent >= 0
   Debug.Assert nIndent < nMaxWidth
   
   Dim sBreakChars As String
   sBreakChars = " -" 'These are the characters to word-wrap on
   
   Dim I As Long
   Dim bFound As Boolean
   
   Do Until Len(sString) < nMaxWidth
   
      'Look for a word wrap character
      bFound = False
      For I = nMaxWidth To nMaxWidth - nMargin Step -1
         If InStr(1, sBreakChars, Mid(sString, i, 1)) > 0 Then
            'Found one, so just break out of the loop
            bFound = True
            Exit For
         End If
      Next
      
      'If no word wrap character was found, then just use the
      ' margin size
      If Not bFound Then
         I = nMaxWidth
      End If
      
      'If there's already a line, append a newline character
      If WrapString <> "" Then
         WrapString = WrapString & vbNewLine
      End If
      
      'And now add the next line, and chop it off of our work string
      WrapString = WrapString & Mid(sString, 1, i)
      sString = Mid(sString, I + 1)
      
      'Add the indention spaces if necessary
      If sString <> "" Then
         sString = Space(nIndent) & sString
      End If
   Loop
   
   'If there's still something left over,
   '  then add it to the final output
   If sString <> "" Then
      If WrapString <> "" Then
         WrapString = WrapString & vbNewLine
      End If
      WrapString = WrapString & sString
   End If

End Function

 
  
Sample Usage: 
  
    Debug.Print WrapString("The quick red fox jumped over the lazy " & _
     "brown dog.  Now is the time for all good men to come to the " & _
     "aid of their country")

 
'-----------------------------------------------------------
'-----------------------------------------------------------
'-----------------------------------------------------------
'-----------------------------------------------------------
'-----------------------------------------------------------


WrapString 

Description: 
 This is a simple function that word wraps a string into smaller strings, separating each line with a newline. 
  
Code: 
 Public Function WrapString( _
   sLineToWrap As String, _
   Optional nFirstLine As Long = 5, _
   Optional anotherLines As Long = 0, _
   Optional nLineLength As Long = 72) As String

    Dim bFirst As Boolean
    Dim nPos As Long
    Dim bBreakFound As Boolean
    
    bFirst = True

    Do While Len(sLineToWrap) > nLineLength
        nPos = nLineLength
        bBreakFound = False
        Do Until nPos = 0 Or bBreakFound
            If Mid(sLineToWrap, nPos, 1) = " " Then
                bBreakFound = True
            ElseIf Mid(sLineToWrap, nPos, 1) = "-" Then
                bBreakFound = True
            Else
                nPos = nPos - 1
            End If
        Loop
        If nPos = 0 Then
            nPos = nLineLength
        End If
        WrapString = WrapString & Space(IIf(bFirst, nFirstLine, _
           anotherLines)) & Mid(sLineToWrap, 1, nPos) & vbNewLine
        
        sLineToWrap = Mid(sLineToWrap, nPos + 1)
        bFirst = False
    Loop
    
    WrapString = WrapString & Space(IIf(bFirst, _
       nFirstLine, anotherLines)) & sLineToWrap

End Function

 
  
Sample Usage: 
  
     Dim sLine As String
    sLine = "This is a really long continuous line that I wrote to "
    sLine = sLine & "demonstrate the wrapping capabilities of the "
    sLine = sLine & "WrapString function.  This text serves no "
    sLine = sLine & "other purpose other than as an example for "
    sLine = sLine & "this function, feel free to ignore this text."

    Debug.Print WrapString(sLine)

 
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

**** End Of
---------------------
Count = 160 -- Fri 17-Jul-2009 01:37:21
---------------------
In Internet Explorer -- 
WinTitle = http://www.scottandmichelle.net/scott/code/index.html
URL = Scott's Sample Code
---------------------
Scott's Sample Code - Windows Internet Explorer
---------------------

FINISHED




'----------------------------------------
Try this to find our pursuit of a neat window close
looks same as our last one

'----------------------------------------
'----------------------------------------
'----------------------------------------
'YES EDIT INSERT
'----------------------------------------

PRIVATE Const GW_HWNDNEXT = 2   ' in another
PRIVATE Const WM_QUIT As Long = &H12
PRIVATE Const WM_CLOSE = &H10
Private Const WM_SYSCOMMAND As Long = &H112
PRIVATE Const SC_CLOSE As Long = &HF060&

TargetHwnd = i
TargetHwnd = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
'----------------------------------------
'----------------------------------------
'----------------------------------------


'NOT KILL PROCESS IS NOT CLEAN
Function CloseProgram(Optional ByVal ProcessHandle As Long, Optional ByVal ThreadHandle As Long, Optional CloseThread As Boolean = True) As Boolean
  
  Dim ReturnValue As Long
  Dim ExitCode    As Long

  If CloseThread = True Then
    ReturnValue = GetExitCodeThread(ThreadHandle, ExitCode)
    If ReturnValue <> 0 Then
  '    TerminateProcess ProcessHandle, ExitCode
       TerminateThread ThreadHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
    
  Else
    ReturnValue = GetExitCodeProcess(ProcessHandle, ExitCode)
    If ReturnValue <> 0 Then
   '   TerminateThread ThreadHandle, ExitCode
       TerminateProcess ProcessHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
  End If
  Exit Function
  
End Function











'-----------------            
'Not Good Thought Meant Outlook FireWall Prog
---------------------
Count = 173 -- Mon 27-Jul-2009 13:42:36
---------------------
In Internet Explorer -- 
WinTitle = http://www.freevbcode.com/ShowCode.Asp?ID=6654
URL = Auto Click YES when Outlooks Security Guard opens prompt dialog.
---------------------
Auto Click YES when Outlooks Security Guard opens prompt dialog. - Windows Internet Explorer
---------------------
Option Explicit
'=============================================================================
' ByPassMSOutlook
' Written by Mohamad Rawey Chek Ani, Perlis, Malaysia
' mrawey@yahoo.com.my
' If you use this application then please mention me in your credits
' Thanks and enjoy
' Please also vote for me :-)

'program that sits in the taskbar and clicks the Yes button on
'behalf of you, when Outlook's Security Guard opens prompt dialog
'saying that a program is trying to send an email with Outlook or
'access its address book. You can suspend/resume it by double-clicking
'its taskbar icon. Developers can automate its behavior by sending special
'messages.

'This had been tested for
'   Outlook 2000 SP1+SR1
'   Outlook 2000 SP2
'   Outlook 2002
'   Outlook 2003

Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" (ByVal lpClassName As Any, _
        ByVal lpWindowName As Any) As Long

Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, _
        ByVal wmessage As Long, ByVal wParam As Long, _
        lParam As Any) As Long
Declare Function FindWindowEx Lib "user32" Alias _
        "FindWindowExA" (ByVal hWnd1 As Long, _
        ByVal hWnd2 As Long, ByVal lpsz1 As String, _
        ByVal lpsz2 As String) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" _
        (ByVal hwnd As Long) As Long
Declare Function EnableWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal fEnable As Long) As Long
        
Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, _
        ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function EnumChildWindows Lib "user32" _
        (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, _
        ByVal lParam As Long) As Long
Declare Function GetWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SefocusAPI Lib "user32" Alias "SetFocus" _
        (ByVal hwnd As Long) As Long

        
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_ENTER = &HD
Public Const VK_TAB = &H9

Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const WM_GETTEXT = &HD

---------------------
---------------------
Count = 174 -- Mon 27-Jul-2009 13:42:39
---------------------
In Internet Explorer -- 
WinTitle = http://www.freevbcode.com/ShowCode.Asp?ID=6654
URL = Auto Click YES when Outlooks Security Guard opens prompt dialog.
'---------------------
'Auto Click YES when Outlooks Security Guard opens prompt dialog. - Windows Internet Explorer
'---------------------
Option Explicit
Private Sub SomeProc()
    Dim wnd&, Aimwindow&, wnd_chld As Long
    Dim uClickYes As Long
    Dim Res As Long
    Dim wasenabled As Long
    Dim a As String, j As Long
    
    DoEvents
    ' is there dialog box appear yet !
    Let wnd = FindWindow(vbNullString, "Microsoft Outlook")
    If (wnd = 0) Then Exit Sub
        Let wnd_chld = GetWindow(wnd, GW_CHILD)
            If (wnd_chld = 0) Then Exit Sub
    
        ' is dialog box enable yet
                wasenabled = IsWindowEnabled(wnd_chld)
                    If wasenabled = 0 Then Exit Sub
                            Do While wnd_chld
                                DoEvents
                                a = Space$(128)
                                ' get child box
                                j = SendMessage(wnd_chld, WM_GETTEXT, 100, ByVal a)
                                If (Len(Trim(a)) <> 0) Then
                                    a = Left(a, InStr(a, Chr$(0)) - 1)
                                    If (UCase(a) = "YES") Then
                                        ' is button YES enable yet
                                        wasenabled = IsWindowEnabled(wnd_chld)
                                        If wasenabled <> 0 Then
                                               Beep
                                               SefocusAPI wnd_chld
                                               keybd_event VK_TAB, 0, 0, 0
                                               keybd_event VK_TAB, 0, KEYEVENTF_KEYUP, 0
                        
                                               keybd_event VK_TAB, 0, 0, 0
                                               keybd_event VK_TAB, 0, KEYEVENTF_KEYUP, 0
                        
                                               keybd_event VK_ENTER, 0, 0, 0
                                               keybd_event VK_ENTER, 0, KEYEVENTF_KEYUP, 0
                                        End If
                                    End If
                                End If
                                DoEvents
                                wnd_chld = GetWindow(wnd_chld, GW_HWNDNEXT)
                            Loop
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
DoEvents
SomeProc
Timer1.Enabled = True
End Sub

'---------------------

Public Sub CloseProgramParent(ByVal Caption1 As String, ByVal Caption2 As String)

Dim I As Long
If Caption1 = "" And Caption2 = "" Then Exit Sub

If Caption1 = "" Then Caption1 = vbNullString
If Caption2 = "" Then Caption2 = vbNullString

I = FindWindow(Caption1, Caption2)

If I = 0 Then CLOSEprogram= False: Exit Sub

CLOSEprogram= True

Do While I <> 0
    SendMessage i, WM_CLOSE, 0&, 0&
    I = GetParent(i)
    If I = 0 Then Exit Do
Loop

End Sub




'How was the Program Darling -- It ended badly Ohhh yes it was process terminated
'and
'the Nice way

02:08 29 July 2009

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const WM_CLOSE = &H10

'---------

Public Sub CloseProgram(ByVal Caption As String)

 Handle = FindWindow(vbNullString, Caption)
 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub

Public Sub CloseProgramHwnd(ByVal Handle As Long)

 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub

'------------------------  Another One -- Escaped my EYE had to find for a while Destroy Terminate

If hCapWnd <> 0 Then Call DestroyWindow(hCapWnd)
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long 'C BOOL


'---------------------




Change the Keyboard Speed and Repeat Delay
Author: Waty  
Category: System/API  
Type: Snippets  
Difficulty: Intermediate  


Version Compatibility:  Visual Basic 5   Visual Basic 6  


This code has been viewed 67009 times.

Instructions: Copy the declarations and code below and paste directly into your VB project.


Declarations:



' #VBIDEUtils#*****************************************
' * Programmer Name  : Waty Thierry
' * Web Site   : http://www.geocities.com/ResearchTriangle/6311/
' * E-Mail     : waty.thierry@usa.net
' * Date       : 25/11/1999
' * Time       : 15:59
' *************************************************************
' * Comments   : Change the keyboard speed and the repeat delay
' *
' ********************************************************
Private Declare Function SystemParametersInfo Lib "user32" _
  Alias "SystemParametersInfoA" (ByVal uAction As Long, _
   ByVal uParam As Long, ByVal lpvParam As Any, _
   ByVal fuWinIni As Long) As Long

Private Const SPI_SETKEYBOARDDELAY = 23
Private Const SPI_SETKEYBOARDSPEED = 11
Private Const SPIF_SENDCHANGE = 1

Public Function SetKeyBoardSpeed(NewSpeed As Long) As Boolean

'PASS A NUMBER BETWEEN 0 AND 31
'31 IS THE FASTEST SPEED
'0 IS THE SLOWEST

   Dim Retcode As Long
   Dim dummy As Long

   Retcode = SystemParametersInfo(SPI_SETKEYBOARDSPEED, _
      NewSpeed, dummy, SPIF_SENDCHANGE)
 
  SetKeyBoardSpeed = (Retcode > 0)
End Function

Public Function SetKeyBoardDelay(NewDelay As Long) As Boolean
'PASS A VALUE BETEEN 0 AND 3.  0 for the SHORTEST DELAY
'(ABOUT 250 ms) 3 for the longset (ABOUT 1 sec)

   Dim Retcode As Long
   Dim dummy As Long
   dummy = 0
    
     Retcode = SystemParametersInfo(SPI_SETKEYBOARDDELAY, _
         NewDelay, dummy, SPIIF_SENDCHANGE)

      SetKeyBoardDelay = (Retcode > 0)
End Function








'-----------------------------------------------------------------------------------------------------------

Option Explicit

Private Type SHFILEOPSTRUCT
hWnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FOF_ALLOWUNDO = &H40&
Private Const FOF_NOCONFIRMATION = &H10&
Private Const FO_COPY = &H2&
Private Const FO_MOVE = &H1&
Private Const FO_DELETE = &H3&
Private Const FO_RENAME = &H4&
Private Const FOF_NOCONFIRMMKDIR = &H200&
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4          ' X - Cannot be set
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10      ' X - CreateDirectory to set
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800    ' X -


Dim WinType_SFO As SHFILEOPSTRUCT
Dim lRet As Long
Dim lflags As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
    
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias _
    "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, _
    ByVal dwFlags As Long) As Long

Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4

Public Function EmptyRecycleBin(ByVal RootPath As String, Optional NoConfirmation As _
    Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
    Dim hWnd As Long, flags As Long

    ' get the handle of the active window, or zero
    On Error Resume Next
    hWnd = Screen.ActiveForm.hWnd
    ' add a colon and backslash, if missing
    If Len(RootPath) > 0 And Mid$(RootPath, 2, 2) <> ":\" Then
        RootPath = Left$(RootPath, 1) & ":\"
    End If
    ' build the flags argument
    flags = (NoConfirmation And SHERB_NOCONFIRMATION) Or (NoProgress And _
        SHERB_NOPROGRESSUI) Or (NoSound And SHERB_NOSOUND)
    SHEmptyRecycleBin hWnd, RootPath, flags
End Function


Public Function ShellFileCopy(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_COPY
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileCopy = (lRet = 0)

End Function

Public Function ShellFileMove(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_MOVE
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileMove = (lRet = 0)

End Function

Public Function ShellFileDelete(src As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_DELETE
.pFrom = src
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileDelete = (lRet = 0)

End Function

Public Function ShellFileRename(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_RENAME
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileRename = (lRet = 0)

End Function

Public Function FileExists(Fname As String) As Boolean

    If Fname = "" Or Right$(Fname, 1) = "\" Then
      FileExists = False: Exit Function
    End If

    FileExists = (Dir(Fname) <> "")
End Function

Public Function DirExists(ByVal DName As String) As Boolean

Dim sDummy As String

If Right(DName, 1) <> "\" Then DName = DName & "\"
sDummy = Dir$(DName & "*.*", vbDirectory)
DirExists = Not (sDummy = "")

End Function

Public Function DeleteOldFiles(DaysOld As Long, FileSpec As _
String, Optional ComparisonDate As Variant) As Boolean

'PURPOSE: DELETES ALL FILES THAT ARE DaysOld Older than
'ComparisonDate, which defaults to now

'RETURNS: True, if succesful
'False otherwise (e.g., user passes non-date argument
'deletion fails because file is in use,
'file doesn't exist, etc.)

'will not delete readonly, hidden or system files

Dim sFileSpec As String
Dim dCompDate As Date
Dim sFileName As String
Dim sFileSplit() As String
Dim iCtr As Integer, iCount As Integer
Dim sDir As String

sFileSpec = FileSpec

If IsMissing(ComparisonDate) Then
    dCompDate = Now
ElseIf Not IsDate(ComparisonDate) Then
    'client passed wrong type
    DeleteOldFiles = False
    Exit Function
Else
    dCompDate = CDate(Format(ComparisonDate, "mm/dd/yyyy"))
End If

sFileName = Dir(sFileSpec)

If sFileName = "" Then
    'returns false is file doesn't exist
    DeleteOldFiles = False
    Exit Function
End If

Do

    If sFileName = "" Then Exit Do

    If InStr(sFileSpec, "\") > 0 Then
        sFileSplit = Split(sFileSpec, "\")
        iCount = UBound(sFileSplit) - 1
        For iCtr = 0 To iCount
            sDir = sDir & sFileSplit(iCtr) & "\"
        Next

        sFileName = sDir & sFileName
    End If

    On Error GoTo errhandler:
    If DateDiff("d", FileDateTime(sFileName), dCompDate) _
       >= DaysOld Then

         Kill sFileName

    End If

    sFileName = Dir
    sDir = ""
Loop

DeleteOldFiles = True

Exit Function

errhandler:
    DeleteOldFiles = False
    Exit Function
End Function









---------------------
Count = 518 -- Tue 04-Aug-2009 00:03:14
---------------------
In Internet Explorer -- 
WinTitle = http://www.freevbcode.com/ShowCode.Asp?ID=8094
URL = A Module of File System Functions
---------------------
A Module of File System Functions - Windows Internet Explorer
---------------------
Option Explicit

Private Type SHFILEOPSTRUCT
hWnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FOF_ALLOWUNDO = &H40&
Private Const FOF_NOCONFIRMATION = &H10&
Private Const FO_COPY = &H2&
Private Const FO_MOVE = &H1&
Private Const FO_DELETE = &H3&
Private Const FO_RENAME = &H4&
Private Const FOF_NOCONFIRMMKDIR = &H200&
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4          ' X - Cannot be set
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10      ' X - CreateDirectory to set
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800    ' X -


Dim WinType_SFO As SHFILEOPSTRUCT
Dim lRet As Long
Dim lflags As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
    
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias _
    "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, _
    ByVal dwFlags As Long) As Long

Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4

Public Function EmptyRecycleBin(ByVal RootPath As String, Optional NoConfirmation As _
    Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
    Dim hWnd As Long, flags As Long

    ' get the handle of the active window, or zero
    On Error Resume Next
    hWnd = Screen.ActiveForm.hWnd
    ' add a colon and backslash, if missing
    If Len(RootPath) > 0 And Mid$(RootPath, 2, 2) <> ":\" Then
        RootPath = Left$(RootPath, 1) & ":\"
    End If
    ' build the flags argument
    flags = (NoConfirmation And SHERB_NOCONFIRMATION) Or (NoProgress And _
        SHERB_NOPROGRESSUI) Or (NoSound And SHERB_NOSOUND)
    SHEmptyRecycleBin hWnd, RootPath, flags
End Function


Public Function ShellFileCopy(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_COPY
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileCopy = (lRet = 0)

End Function

Public Function ShellFileMove(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_MOVE
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileMove = (lRet = 0)

End Function

Public Function ShellFileDelete(src As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_DELETE
.pFrom = src
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileDelete = (lRet = 0)

End Function

Public Function ShellFileRename(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_RENAME
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileRename = (lRet = 0)

End Function

Public Function FileExists(Fname As String) As Boolean

    If Fname = "" Or Right$(Fname, 1) = "\" Then
      FileExists = False: Exit Function
    End If

    FileExists = (Dir(Fname) <> "")
End Function

Public Function DirExists(ByVal DName As String) As Boolean

Dim sDummy As String

If Right(DName, 1) <> "\" Then DName = DName & "\"
sDummy = Dir$(DName & "*.*", vbDirectory)
DirExists = Not (sDummy = "")

End Function

Public Function DeleteOldFiles(DaysOld As Long, FileSpec As _
String, Optional ComparisonDate As Variant) As Boolean

'PURPOSE: DELETES ALL FILES THAT ARE DaysOld Older than
'ComparisonDate, which defaults to now

'RETURNS: True, if succesful
'False otherwise (e.g., user passes non-date argument
'deletion fails because file is in use,
'file doesn't exist, etc.)

'will not delete readonly, hidden or system files

Dim sFileSpec As String
Dim dCompDate As Date
Dim sFileName As String
Dim sFileSplit() As String
Dim iCtr As Integer, iCount As Integer
Dim sDir As String

sFileSpec = FileSpec

If IsMissing(ComparisonDate) Then
    dCompDate = Now
ElseIf Not IsDate(ComparisonDate) Then
    'client passed wrong type
    DeleteOldFiles = False
    Exit Function
Else
    dCompDate = CDate(Format(ComparisonDate, "mm/dd/yyyy"))
End If

sFileName = Dir(sFileSpec)

If sFileName = "" Then
    'returns false is file doesn't exist
    DeleteOldFiles = False
    Exit Function
End If

Do

    If sFileName = "" Then Exit Do

    If InStr(sFileSpec, "\") > 0 Then
        sFileSplit = Split(sFileSpec, "\")
        iCount = UBound(sFileSplit) - 1
        For iCtr = 0 To iCount
            sDir = sDir & sFileSplit(iCtr) & "\"
        Next

        sFileName = sDir & sFileName
    End If

    On Error GoTo errhandler:
    If DateDiff("d", FileDateTime(sFileName), dCompDate) _
       >= DaysOld Then

         Kill sFileName

    End If

    sFileName = Dir
    sDir = ""
Loop

DeleteOldFiles = True

Exit Function

errhandler:
    DeleteOldFiles = False
    Exit Function
End Function





'-----------------------------------------------------------------



'More Code


A Module of File System Functions
Author: Anonymous  
Category: Files and Directories  
Type: Modules  
Difficulty: Intermediate  

Version Compatibility:  Visual Basic 6    

More information: This is a small module that I put together based on snippets and other pieces of code. It uses the shell API and can rename, delete, move, copy and delete files based on number of days old. All this code can be found elsewhere as separate modules, all I did was streamline it and combine it into one simple useful module. It can also empty the recycle bin from any drive and set file attributes for folders or files. Just copy the code to a new module and call the functions as needed.
This code has been viewed 40329 times.
Instructions: Copy the declarations and code below and paste directly into your VB project.


'http://www.freevbcode.com/ShowCode.ASP?SearchString=delete%20to%20recycle&ID=8094


Option Explicit

Private Type SHFILEOPSTRUCT
hWnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FOF_ALLOWUNDO = &H40&
Private Const FOF_NOCONFIRMATION = &H10&
Private Const FO_COPY = &H2&
Private Const FO_MOVE = &H1&
Private Const FO_DELETE = &H3&
Private Const FO_RENAME = &H4&
Private Const FOF_NOCONFIRMMKDIR = &H200&
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4          ' X - Cannot be set
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10      ' X - CreateDirectory to set
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800    ' X -


Dim WinType_SFO As SHFILEOPSTRUCT
Dim lRet As Long
Dim lflags As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
    
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias _
    "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, _
    ByVal dwFlags As Long) As Long

Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4

Public Function EmptyRecycleBin(ByVal RootPath As String, Optional NoConfirmation As _
    Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
    Dim hWnd As Long, flags As Long

    ' get the handle of the active window, or zero
    On Error Resume Next
    hWnd = Screen.ActiveForm.hWnd
    ' add a colon and backslash, if missing
    If Len(RootPath) > 0 And Mid$(RootPath, 2, 2) <> ":\" Then
        RootPath = Left$(RootPath, 1) & ":\"
    End If
    ' build the flags argument
    flags = (NoConfirmation And SHERB_NOCONFIRMATION) Or (NoProgress And _
        SHERB_NOPROGRESSUI) Or (NoSound And SHERB_NOSOUND)
    SHEmptyRecycleBin hWnd, RootPath, flags
End Function


Public Function ShellFileCopy(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_COPY
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileCopy = (lRet = 0)

End Function

Public Function ShellFileMove(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_MOVE
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileMove = (lRet = 0)

End Function

Public Function ShellFileDelete(src As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_DELETE
.pFrom = src
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileDelete = (lRet = 0)

End Function

Public Function ShellFileRename(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_RENAME
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileRename = (lRet = 0)

End Function

Public Function FileExists(Fname As String) As Boolean

    If Fname = "" Or Right$(Fname, 1) = "\" Then
      FileExists = False: Exit Function
    End If

    FileExists = (Dir(Fname) <> "")
End Function

Public Function DirExists(ByVal DName As String) As Boolean

Dim sDummy As String

If Right(DName, 1) <> "\" Then DName = DName & "\"
sDummy = Dir$(DName & "*.*", vbDirectory)
DirExists = Not (sDummy = "")

End Function

Public Function DeleteOldFiles(DaysOld As Long, FileSpec As _
String, Optional ComparisonDate As Variant) As Boolean

'PURPOSE: DELETES ALL FILES THAT ARE DaysOld Older than
'ComparisonDate, which defaults to now

'RETURNS: True, if succesful
'False otherwise (e.g., user passes non-date argument
'deletion fails because file is in use,
'file doesn't exist, etc.)

'will not delete readonly, hidden or system files

Dim sFileSpec As String
Dim dCompDate As Date
Dim sFileName As String
Dim sFileSplit() As String
Dim iCtr As Integer, iCount As Integer
Dim sDir As String

sFileSpec = FileSpec

If IsMissing(ComparisonDate) Then
    dCompDate = Now
ElseIf Not IsDate(ComparisonDate) Then
    'client passed wrong type
    DeleteOldFiles = False
    Exit Function
Else
    dCompDate = CDate(Format(ComparisonDate, "mm/dd/yyyy"))
End If

sFileName = Dir(sFileSpec)

If sFileName = "" Then
    'returns false is file doesn't exist
    DeleteOldFiles = False
    Exit Function
End If

Do

    If sFileName = "" Then Exit Do

    If InStr(sFileSpec, "\") > 0 Then
        sFileSplit = Split(sFileSpec, "\")
        iCount = UBound(sFileSplit) - 1
        For iCtr = 0 To iCount
            sDir = sDir & sFileSplit(iCtr) & "\"
        Next

        sFileName = sDir & sFileName
    End If

    On Error GoTo errhandler:
    If DateDiff("d", FileDateTime(sFileName), dCompDate) _
       >= DaysOld Then

         Kill sFileName

    End If

    sFileName = Dir
    sDir = ""
Loop

DeleteOldFiles = True

Exit Function

errhandler:
    DeleteOldFiles = False
    Exit Function
End Function




'-----------------------------------------------------------------

