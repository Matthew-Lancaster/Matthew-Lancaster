VERSION 5.00
Begin VB.Form FormStart 
   Caption         =   "Form2"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Dim R As Long
Public MNU_TREESIZE_GO
Dim NET_PATH_AND_DRIVE
Dim F1$
Dim D1$

' -------------------------------------
' NOTE FORM VBS COMMON LABEL NAME CLASH
' SO CALL ShowWindow_2
' -------------------------------------
Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True


'-----------------------------------------------------------------
Private Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Function GetWindowsVersion()
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    sngWindowsVersion = CSng((CStr(osinfo.dwMajorVersion) & "." & CStr(osinfo.dwMinorVersion)))
    strWindowsVersion = CStr(osinfo.dwMajorVersion) & "." & CStr(osinfo.dwMinorVersion) & "." & CStr(osinfo.dwBuildNumber)
    GetWindowsVersion = osinfo.dwMajorVersion
    GetWindowsVersion = CSng((CStr(osinfo.dwMajorVersion) & "." & CStr(osinfo.dwMinorVersion)))
    
End Function
'-----------------------------------------------------------------


Public Sub FormStartLoader()

'Set FS = CreateObject("Scripting.FileSystemObject")
Call SET_UP_PULIC_FSO

FontSizez = 12

'Dir1.Path = "C:\Program Files\00 WinAmp's"
'Dir1.Path = "E:\01 VB Shell Folders\00 Shell Notepad Launcher"
'File1.Path = "E:\01 VB Shell Folders\00 Shell Notepad Launcher"


On Error Resume Next
For R = 3 To 25
    
    Err.Clear
    Set z = FSO.GetDrive(FSO.GetDriveName(FSO.GetAbsolutePathName(Chr$(R + 64) + ":")))
     
    Select Case z.DriveType
        Case 0: T = "Unknown"
        Case 1: T = "Removable"
        Case 2: T = "Fixed"
        Case 3: T = "Network"
        Case 4: T = "CD-ROM"
        Case 5: T = "RAM Disk"
    End Select
    
    
    If Err.Number = 0 Then
        tt$ = z.DriveLetter + ":" + "- Type - " + T + " --- Serial - " + Hex$(z.SerialNumber) + " - Vol.Name - " + z.VolumeName + vbCrLf
        'RD$(tg) = tt$
        tg = tg + 1
        Y1$ = Y1$ + tt$
        Filename = z.DriveLetter + ":\ __ " + z.VolumeName
        Path = "--Drive"
        
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Filename)
            LV.SubItems(1) = Path
        End With
    
    End If

Next



'Filename_VAR(1) = "\\1-ASUS-X5DIJ"
'Filename_VAR(2) = "\\2-ASUS-EEE"
'Filename_VAR(3) = "\\3-LINDA-PC"
'Filename_VAR(4) = "\\4-ASUS-GL522VW"
'Filename_VAR(5) = "\\5-ASUS-P2520LA"
'Filename_VAR(6) = "\\7-ASUS-GL522VW"
'Filename_VAR(7) = "\\8-MSI-GP62M-7RD"
'Filename_VAR(8) = "\\NAS-QNAP-ML"

' ----------------------------------
' ADD HEADER TITLE FOR NETWORK COMPUTER NAME
' ----------------------------------
With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\NETWORK COMPUTER NAME\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With

Dim Filename_VAR(40)

FR_1 = FreeFile
NET_FILENAME = "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\NETWORK_COMPUTER_NAME.txt"
Open NET_FILENAME For Input As #FR_1
R_L = 0
Do
    Line Input #FR_1, LINE_STINGER
    
    R_L = R_L + 1
    Filename_VAR(R_L) = "\\" + LINE_STINGER
    
    HERE_GO = False
    If (Val(Mid(LINE_STINGER, 1, 1)) > 0 And Mid(LINE_STINGER, 2, 1) = "-") = True Then
        HERE_GO = True
    End If
    
    If HERE_GO = True Then
        Path = "--DriveRemote"
        Path = "--Drive"
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Filename_VAR(R_L))
            LV.SubItems(1) = Path
        End With
    ' ADD THE HTTPS:\\" TO "\BTHUB"
    End If
    
    If HERE_GO = False And LINE_STINGER = "BTHUB" Then
        Filename = "HTTPS:" + Filename_VAR(R_L) '+ "\BTHUB"
        'Path = "--Drive"
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Filename)
            LV.SubItems(1) = Path
        End With
        HERE_GO = True
    End If
    If HERE_GO = False Then
        'Path = "--Drive"
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Filename_VAR(R_L))
            LV.SubItems(1) = Path
        End With
    End If

Loop Until EOF(FR_1)
Close #FR_1

' ----------------------------------
' ADD HEADER TITLE FOR NETWORK DRIVE DISK
' ----------------------------------
With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\NETWORK DRIVE DISK\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With

R_L_X = 0
FR_1 = FreeFile
NET_FILENAME = "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\NETWORK_COMPUTER_NAME.txt"
Open NET_FILENAME For Input As #FR_1
Path = "--DriveRemote"
Do

    HERE_GO = False
    Line Input #FR_1, LINE_STINGER
    If (Val(Mid(LINE_STINGER, 1, 1)) > 0 And Mid(LINE_STINGER, 2, 1) = "-") = True Then
        R_L = R_L + 1
        R_L_X = R_L_X + 1
        Filename_VAR(R_L) = "\\" + LINE_STINGER
        HERE_GO = True
    End If

    If HERE_GO = True Then
        With ScanPath.ListView1
            NET_PATH_AND_DRIVE = Filename_VAR(R_L) + "__C"
            Set LV = .ListItems.Add(, , NET_PATH_AND_DRIVE)
            LV.SubItems(1) = Path
            Call GET_COMPUTR_NETWORK_NAME
            NET_C_PATH = NET_C_PATH + NET_PATH_AND_DRIVE + vbCrLf
            NET_PATH_ALL_2 = NET_PATH_ALL_2 + NET_PATH_AND_DRIVE + vbCrLf
        End With
        NET_PATH_ALL = NET_PATH_ALL + vbCrLf
        With ScanPath.ListView1
            NET_PATH_AND_DRIVE = Filename_VAR(R_L) + "__D"
            Set LV = .ListItems.Add(, , NET_PATH_AND_DRIVE)
            LV.SubItems(1) = Path
            Call GET_COMPUTR_NETWORK_NAME
            NET_D_PATH = NET_D_PATH + NET_PATH_AND_DRIVE + vbCrLf
            NET_PATH_ALL_2 = NET_PATH_ALL_2 + NET_PATH_AND_DRIVE + vbCrLf
            
        End With
        NET_PATH_ALL = NET_PATH_ALL + vbCrLf
        If R_L_X > 0 Then
            With ScanPath.ListView1
                NET_PATH_AND_DRIVE = Filename_VAR(R_L) + "__E"
                Set LV = .ListItems.Add(, , NET_PATH_AND_DRIVE)
                LV.SubItems(1) = Path
                Call GET_COMPUTR_NETWORK_NAME
                NET_E_PATH = NET_E_PATH + NET_PATH_AND_DRIVE + vbCrLf
                NET_PATH_ALL_2 = NET_PATH_ALL_2 + NET_PATH_AND_DRIVE + vbCrLf
            End With
        End If
        NET_PATH_ALL = NET_C_PATH + vbCrLf + NET_D_PATH + vbCrLf + NET_E_PATH + vbCrLf
    End If
Loop Until EOF(FR_1)
Close #FR_1



' ----------------------------------
' ADD HEADER TITLE FOR NETWORK USER FOLDER
' ----------------------------------
With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\NETWORK USER FOLDER\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With

NET_PATH_ALL_R = Split(NET_PATH_ALL, vbCrLf)

Dim M()
ReDim M(UBound(NET_PATH_ALL_R) + 10)
Dim R3
For R3 = 0 To UBound(NET_PATH_ALL_R)
    CK1 = UCase(NET_PATH_ALL_R(R3))
    If InStr(CK1, "1_ASUS") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "2_ASUS") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "D_DRIVE") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "03_FAT32") > 0 Then NET_PATH_ALL_R(R3) = ""
Next
For R3 = 0 To UBound(NET_PATH_ALL_R)
    If NET_PATH_ALL_R(R3) <> "" Then
        
        For R5 = 1 To 5
        
        CK1 = NET_PATH_ALL_R(R3)
        CK3 = CK1 + "\USERS\MATT " + Format(R5, "00")
        If 1 = 2 Then
        If Dir(CK1 + "\USERS\MATT " + Format(R5, "00") + "\Desktop", vbDirectory) <> "" Then
            
            Path = "--DriveRemote_USER"
            CK2 = Mid(CK1, 1, InStr(4, CK1, "\") - 1)
            If InStr(CK1, "3_LINDA") > 0 Then STRV = 8
            If InStr(CK1, "4_ASUS") > 0 Then STRV = 2
            If InStr(CK1, "5_ASUS") > 0 Then STRV = 3
            If InStr(CK1, "7_ASUS") > 0 Then STRV = 2
            If InStr(CK1, "8_MSI") > 0 Then STRV = 2
            STRV = STRV - 1
            CK2 = CK2 + " " + String(STRV, "_") + " C:\USER\MATT " + Format(R5, "00")
            SET_HERE_2 = False
            If InStr(CK2, "3-L") > 0 And InStr(CK2, "MATT 01") > 0 Then SET_HERE_2 = True
            If InStr(CK2, "4-A") > 0 And InStr(CK2, "MATT 01") > 0 Then SET_HERE_2 = True
            If InStr(CK2, "5-A") > 0 And InStr(CK2, "MATT 01") > 0 Then SET_HERE_2 = True
            If InStr(CK2, "7-A") > 0 And InStr(CK2, "MATT 04") > 0 Then SET_HERE_2 = True
            If InStr(CK2, "8-M") > 0 And InStr(CK2, "MATT 01") > 0 Then SET_HERE_2 = True
            If SET_HERE_2 = True Then
            CK2 = CK2 + " #"
            End If
            With ScanPath.ListView1
                Set LV = .ListItems.Add(, , CK2)
                LV.SubItems(1) = Path
                LV.SubItems(2) = CK3
            End With
        End If
        End If
        Next
    End If
Next


'-------------------------------
'-------------------------------
'-------------------------------
'-------------------------------



'For R_L = 1 To Frm_M_P_S.ListView1.ListItems.Count
'    If UCase(Frm_M_P_S.ListView1.ListItems(R_L).SubItems(2)) = "BTHUB" Then
'        Filename = "HTTPS:\\" + UCase(Frm_M_P_S.ListView1.ListItems(R_L)) + "\BTHUB"
'        Path = "--DriveRemote"
'        'Path = "--Drive"
'        With ScanPath.ListView1
'            Set LV = .ListItems.Add(, , Filename)
'            LV.SubItems(1) = Path
'        End With
'    End If
'Next


' ----------------------------------
' ADD HEADER TITLE FOR SPECIAL FOLDER
' ----------------------------------
With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\SPECIAL FOLDER\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With



Dim R_X As Long

For R_X = 0 To 120
    If Mid(UCase(GetSpecialfolder(R_X)), 2) = UCase(":\USERS\" + GetUserName) Then
        GetSpecialfolder_VAR = GetSpecialfolder(R_X)
        Exit For
    End If
Next

X_COUNT = 0
For R = 67 To 90
    If Dir(Chr(R) + ":\TEMP", vbDirectory) <> "" Then
        X_COUNT = X_COUNT + 1
        Path = "--SPECIAL"
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Chr(R) + ":\TEMP")
            LV.SubItems(1) = Path
        End With
    Else
        Exit For
    End If
    If X_COUNT = 3 Then Exit For
Next


GetSpecialfolder_VAR = Mid(GetSpecialfolder_VAR, 1, Len(GetSpecialfolder_VAR) - 1)
For R_L = 1 To 9
    GET_USER_NAME_VAR_NAME = GetSpecialfolder_VAR + Format(R_L, "0")
    Filename = GET_USER_NAME_VAR_NAME
    If Dir(Filename, vbDirectory) <> "" Then
    
        Path = "--Drive"
        
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Filename)
            LV.SubItems(1) = Path
        End With
        
    End If
Next

'On Error Resume Next
'For R = 0 To 255
'
'    If GetSpecialfolder(R) <> "" Then
'        Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
'    End If
'
'Next
'Stop




On Error Resume Next
For R = 0 To 255
    q = 0
'    If R = 0 Then q = 1
'    If R = 2 Then q = 1
'    If R = 5 Then q = 1
'    If R = 6 Then q = 1
'    If R = 9 Then q = 1
'    If R = 11 Then q = 1
'    If R = 16 Then q = 1
'    If R = 20 Then q = 1
'    If R = 22 Then q = 1
'    If R = 25 Then q = 1
'    If R = 26 Then q = 1
'    If R = 31 Then q = 1
'    If R = 35 Then q = 1
'    If R = 36 Then q = 1
'    If R = 37 Then q = 1
'    If R = 38 Then q = 1
'    If R = 39 Then q = 1
'    If R = 40 Then q = 1
'    If R = 41 Then q = 1
'    If R = 42 Then q = 1
'    If R = 53 Then q = 1
'    If R = 54 Then q = 1
'    If R = 55 Then q = 1
'    If R = 56 Then q = 1
    
    If R = 7 Then q = 1
    If R = 29 Then q = 1
    If R = 19 Then q = 1
    If R = 40 Then q = 1
    
    If GetSpecialfolder(R) <> "" Then
'        Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
    End If
    
    If GetSpecialfolder(R) <> "" And q = 0 Then
        Filename = GetSpecialfolder(R)
        'Filename = GetSpecialfolder(R)
        Path = "--SPECIAL"
        
        SET_GO = True
        F1 = UCase(Filename)
        If InStr(F1, "DOCUMENTS AND") > 0 And InStr(F1, "ADMINISTRATIVE") > 0 Then
            SET_GO = False
        End If
        If InStr(F1, "DOCUMENTS AND") > 0 And InStr(F1, "MICROSOFT\CD BURNING") > 0 Then
            SET_GO = False
        End If
        
        If InStr(F1, UCase("AppData\Roaming\Microsoft\Windows\Printer Shortcuts")) > 0 Then
            SET_GO = False
        End If
        'C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Printer Shortcuts
        '3-LINDA-PC
        
        'DUPE CHECKER
        
        
        
        If InStr(DUPE_CHECK, "__" + Filename + "__") > 0 Then
            SET_GO = False
        End If
        DUPE_CHECK = DUPE_CHECK + "__" + Filename + "__"
        
        
        
        If SET_GO = True Then
            With ScanPath.ListView2
                XF_1 = Filename
                XF_2 = Filename
                If InStr(UCase(XF_1), "TOOLS") > 0 Then
                    XF_1 = Replace(XF_1, "Tools", "Tool")
                End If
                Set LV = .ListItems.Add(, , Format(R, "00 ") + XF_1)
                LV.SubItems(1) = Path
                If InStr(UCase(XF_2), "TOOLS") > 0 Then
                    LV.SubItems(2) = XF_2
                End If
            End With
        End If
        
        If InStr(Filename + "--", "Program Files (x86)" + "--") > 0 Then
        If InStr(DUPE_CHECK, "__" + Filename + "__") > 0 Then
            SET_GO = False
        End If
        DUPE_CHECK = DUPE_CHECK + "__" + Filename + "__"
            With ScanPath.ListView2
                Filename = Replace(Filename, " (x86)", "")
                Set LV = .ListItems.Add(, , Format(R, "00 ") + Filename)
                LV.SubItems(1) = Path
            End With
        End If
        
        
    End If

Next
'End


' FAKE THE NAME 26 ADD IT IN WITH OTHER ORDER
Path = "--SPECIAL"
With ScanPath.ListView2
    Set LV = .ListItems.Add(, , "26 C:\Users\" + GetUserName + "\AppData\Roaming\Microsoft\Windows")
    LV.SubItems(1) = Path
End With

Path = "--SPECIAL"
With ScanPath.ListView2
    Set LV = .ListItems.Add(, , "26 C:\Users\" + GetUserName + "\AppData\Local\Microsoft")
    LV.SubItems(1) = Path
End With

ScanPath.ListView2.SortOrder = lvwAscending
ScanPath.ListView2.SortKey = 0
ScanPath.ListView2.Sorted = True
ScanPath.ListView2.Sorted = False

For R = 1 To ScanPath.ListView2.ListItems.Count
    Path = "--SPECIAL"
    With ScanPath.ListView1
        Set LV = .ListItems.Add(, , ScanPath.ListView2.ListItems.Item(R))
        LV.SubItems(1) = ScanPath.ListView2.ListItems.Item(R).SubItems(1)
        LV.SubItems(2) = ScanPath.ListView2.ListItems.Item(R).SubItems(2)
    End With
    
Next




'----
'Shell Commands List for Windows 10 | Tutorials
'https://www.tenforums.com/tutorials/3109-shell-commands-list-windows-10-a.html
'----
'shell:3D Objects    %UserProfile%\3D Objects
'shell:AccountPictures   %AppData%\Microsoft\Windows\AccountPictures
'shell:AddNewProgramsFolder  Control Panel\All Control Panel Items\Get Programs
'shell:Administrative Tools  %AppData%\Microsoft\Windows\Start Menu\Programs\Administrative Tools
'shell:AppData   %AppData%
'shell:Application Shortcuts %LocalAppData%\Microsoft\Windows\Application Shortcuts
'shell: AppsFolder Applications
'shell:AppUpdatesFolder  Installed Updates
'shell:Cache %LocalAppData%\Microsoft\Windows\INetCache
'shell:Camera Roll   %UserProfile%\Pictures\Camera Roll
'shell:CD Burning    %LocalAppData%\Microsoft\Windows\Burn\Burn
'shell:ChangeRemoveProgramsFolder    Control Panel\All Control Panel Items\Programs and Features
'shell:Common Administrative Tools   %ProgramData%\Microsoft\Windows\Start Menu\Programs\Administrative Tools
'shell:Common AppData    %ProgramData%
'shell:Common Desktop    %Public%\Desktop
'shell:Common Documents  %Public%\Documents
'shell:CommonDownloads   %Public%\Downloads
'shell:CommonMusic   %Public%\Music
'shell:CommonPictures    %Public%\Pictures
'shell:Common Programs   %ProgramData%\Microsoft\Windows\Start Menu\Programs
'shell:CommonRingtones   %ProgramData%\Microsoft\Windows\Ringtones
'shell:Common Start Menu %ProgramData%\Microsoft\Windows\Start Menu
'shell:Common Startup    %ProgramData%\Microsoft\Windows\Start Menu\Programs\Startup
'shell:Common Templates  %ProgramData%\Microsoft\Windows\Templates
'??shell:CommonVideo %Public%\Videos
'shell:ConflictFolder    Control Panel\All Control Panel Items\Sync Center\Conflicts
'shell:ConnectionsFolder Control Panel\All Control Panel Items\Network Connections
'shell:Contacts  %UserProfile%\Contacts
'shell:ControlPanelFolder    Control Panel\All Control Panel Items
'shell:Cookies   %LocalAppData%\Microsoft\Windows\INetCookies
'shell:Cookies\Low   %LocalAppData%\Microsoft\Windows\INetCookies\Low
'shell:CredentialManager %AppData%\Microsoft\Credentials
'shell:CryptoKeys    %AppData%\Microsoft\Crypto
'shell: Desktop Desktop
'shell:device Metadata Store %ProgramData%\Microsoft\Windows\DeviceMetadataStore
'shell:documentsLibrary??    Libraries\Documents
'shell:downloads %UserProfile%\Downloads
'shell:dpapiKeys %AppData%\Microsoft\Protect
'shell:Favorites %UserProfile%\Favorites
'shell:Fonts %WinDir%\Fonts
'shell:Games (removed in version 1803)   Games
'shell:GameTasks %LocalAppData%\Microsoft\Windows\GameExplorer
'??shell:History %LocalAppData%\Microsoft\Windows\History
'shell: HomeGroupCurrentUserFolder Homegroup \ (User - Name)
'shell: HomeGroupFolder Homegroup
'shell:ImplicitAppShortcuts  %AppData%\Microsoft\Internet Explorer\Quick Launch\User Pinned\ImplicitAppShortcuts
'shell:InternetFolder    Internet Explorer
'shell: Libraries Libraries
'shell:Links %UserProfile%\Links
'shell:Local AppData %LocalAppData%
'shell:LocalAppDataLow   %UserProfile%\AppData\LocalLow
'??shell:MusicLibrary    Libraries\Music
'shell:MyComputerFolder  This PC
'shell:My Music  %UserProfile%\Music
'shell:My Pictures   %UserProfile%\Pictures
'shell:My Video  %UserProfile%\Videos
'shell:NetHood   %AppData%\Microsoft\Windows\Network Shortcuts
'shell: NetworkPlacesFolder Network
'shell: OneDrive OneDrive
'shell:OneDriveCameraRoll    %UserProfile%\OneDrive\Pictures\Camera Roll
'shell:OneDriveDocuments %UserProfile%\OneDrive\Documents
'shell:OneDriveMusic %UserProfile%\OneDrive\Music
'shell:OneDrivePictures  %UserProfile%\OneDrive\Pictures
'shell:Personal  %UserProfile%\Documents
'shell: PicturesLibrary Libraries \ Pictures
'shell:PrintersFolder    All Control Panel Items\Printers
'??shell:PrintHood   %AppData%\Microsoft\Windows\Printer Shortcuts
'shell:Profile   %UserProfile%
'??shell:ProgramFiles    %ProgramFiles%
'shell:ProgramFilesCommon    %ProgramFiles%\Common Files
'shell:ProgramFilesCommonX64 %ProgramFiles%\Common Files (64-bit Windows only)
'shell:ProgramFilesCommonX86 %ProgramFiles(x86)%\Common Files (64-bit Windows only)
'shell:ProgramFilesX64   %ProgramFiles% (64-bit Windows only)
'shell:ProgramFilesX86   %ProgramFiles(x86)% (64-bit Windows only)
'shell:Programs  %AppData%\Microsoft\Windows\Start Menu\Programs
'shell:Public    %Public%
'shell:PublicAccountPictures %Public%\AccountPictures
'shell:PublicGameTasks   %ProgramData%\Microsoft\Windows\GameExplorer
'shell:PublicLibraries   %Public%\Libraries
'shell:Quick Launch  %AppData%\Microsoft\Internet Explorer\Quick Launch
'shell:Recent    %AppData%\Microsoft\Windows\Recent
'shell:RecordedTVLibrary Libraries\Recorded TV
'??shell:RecycleBinFolder    Recycle Bin
'shell:ResourceDir   %WinDir%\Resources
'shell:Ringtones %ProgramData%\Microsoft\Windows\Ringtones
'shell:Roamed Tile Images    %LocalAppData%\Microsoft\Windows\RoamedTileImages
'shell:Roaming Tiles %AppData%\Microsoft\Windows\RoamingTiles
'shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}  Run dialog box
'shell:SavedGames    %UserProfile%\Saved Games
'shell:Screenshots   %UserProfile%\Pictures\Screenshots
'shell:Searches  %UserProfile%\Searches
'shell:SearchHistoryFolder   %LocalAppData%\Microsoft\Windows\ConnectedSearch\History
'shell: SearchHomeFolder search - ms:
'shell:SearchTemplatesFolder %LocalAppData%\Microsoft\Windows\ConnectedSearch\Templates
'shell:SendTo    %AppData%\Microsoft\Windows\SendTo
'shell:Start Menu    %AppData%\Microsoft\Windows\Start Menu
'shell: StartMenuAllPrograms StartMenuAllPrograms
'shell:Startup   %AppData%\Microsoft\Windows\Start Menu\Programs\Startup
'shell:SyncCenterFolder  Control Panel\All Control Panel Items\Sync Center
'shell:SyncResultsFolder Control Panel\All Control Panel Items\Sync Center\Sync Results
'shell:SyncSetupFolder   Control Panel\All Control Panel Items\Sync Center\Sync Setup
'shell:System    %WinDir%\System32
'shell:SystemCertificates    %AppData%\Microsoft\SystemCertificates
'shell:SystemX86 %WinDir%\SysWOW64
'shell:Templates %AppData%\Microsoft\Windows\Templates
'shell: ThisPCDesktopFolder Desktop
'??shell:UsersFilesFolder    %UserProfile%
'shell:User Pinned   %AppData%\Microsoft\Internet Explorer\Quick Launch\User Pinned
'shell:UserProfiles  %HomeDrive%\Users
'shell:UserProgramFiles  %LocalAppData%\Programs
'shell:UserProgramFilesCommon    %LocalAppData%\Programs\Common
'shell: UsersLibrariesFolder Libraries
'??shell:VideosLibrary   Libraries\Videos
'shell:Windows   %WinDir%
'


' ----------------------------------
' ADD HEADER TITLE FOR EXPLORER SHELL FOLDER
' ----------------------------------
With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\EXPLORER SHELL FOLDER\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With
Dim ARRAY_V()
Dim ARRAY_V_2()
ReDim ARRAY_V(100)
ReDim ARRAY_V_2(100)
AVI = 0
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
'AVI = AVI + 1: ARRAY_V(AVI) = "shell:AddNewProgramsFolder"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:Administrative Tools"
'AVI = AVI + 1: ARRAY_V(AVI) = "shell:Common Administrative Tools"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:Contacts"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:ControlPanelFolder"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:Cookies"
'AVI = AVI + 1: ARRAY_V(AVI) = "shell:Cookies\Low"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:desktop"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:HomeGroupFolder"
'AVI = AVI + 1: ARRAY_V(AVI) = "shell:ImplicitAppShortcuts"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:MyComputerFolder"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:NetHood"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:NetworkPlacesFolder"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:OneDrive"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:Quick Launch"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:Recent"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:RecycleBinFolder"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:SendTo"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:Startup"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:System"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:SystemX86"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:ThisPCDesktopFolder"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:User Pinned"
AVI = AVI + 1: ARRAY_V(AVI) = "shell:Windows"
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""

AVI = 0
For R_AVI = 1 To UBound(ARRAY_V)
    If ARRAY_V(R_AVI) <> "" Then
        AVI = AVI + 1
        ARRAY_V_2(AVI) = ARRAY_V(R_AVI)
    End If
Next
ReDim Preserve ARRAY_V(AVI)
ReDim Preserve ARRAY_V_2(AVI)
For AVI = 1 To UBound(ARRAY_V_2)
    ARRAY_V(AVI) = ARRAY_V_2(AVI)
Next

For AVI = 1 To UBound(ARRAY_V)
    With ScanPath.ListView1
        Set LV = .ListItems.Add(, , ARRAY_V(AVI))
        LV.SubItems(1) = "EXPLORER_SHELL"
    End With
Next



'----
'How to Access Device Manager From the Command Prompt
'https://www.lifewire.com/how-to-access-device-manager-from-the-command-prompt-2626360
'----
'----
'How to run Control Panel tools by typing a command
'https://support.microsoft.com/en-gb/help/192806/how-to-run-control-panel-tools-by-typing-a-command
'----

'   Control panel tool             Command
'   -----------------------------------------------------------------
'   Accessibility Options          control access.cpl
'   Add New Hardware               control sysdm.cpl add new hardware
'   Add/Remove Programs            control appwiz.cpl
'   Date/Time Properties           control timedate.cpl
'   Display Properties             control desk.cpl
'   FindFast                       control findfast.cpl
'   Fonts Folder                   control fonts
'   Internet Properties            control inetcpl.cpl
'   Joystick Properties            control joy.cpl
'   Keyboard Properties            control main.cpl keyboard
'   Microsoft Exchange             control mlcfg32.cpl
'      (or Windows Messaging)
'   Microsoft Mail Post Office     control wgpocpl.cpl
'   Modem Properties               control modem.cpl
'   Mouse Properties               control main.cpl
'   Multimedia Properties          control mmsys.cpl
'   Network Properties             control netcpl.cpl
'                                  NOTE: In Windows NT 4.0, Network
'                                  properties is Ncpa.cpl, not Netcpl.cpl
'   Password Properties            control password.cpl
'   PC Card                        control main.cpl pc card (PCMCIA)
'   Power Management (Windows 95)  control main.cpl power
'   Power Management (Windows 98)  control powercfg.cpl
'   Printers Folder                control printers
'   Regional Settings              control intl.cpl
'   Scanners and Cameras           control sticpl.cpl
'   Sound Properties               control mmsys.cpl sounds
'   System Properties              control sysdm.cpl



With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\CONTROL PANEL COMMAND LINE TOOL\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With

ReDim ARRAY_V(100)
ReDim ARRAY_V_2(100)

DIFA = 14
AVI = AVI + 1: ARRAY_V(AVI) = "MMC DEVMGMT.MSC" + Space(24 - DIFA) + " ---- DEVICE MANAGER"
AVI = AVI + 1: ARRAY_V(AVI) = "CONTROL HDWWIZ.CPL" + Space(19 - DIFA) + " ---- DEVICE MANAGER"
AVI = AVI + 1: ARRAY_V(AVI) = "CONTROL /NAME MICROSOFT.DEVICEMANAGER"
AVI = AVI + 1: ARRAY_V(AVI) = "control appwiz.cpl" + Space(20 - DIFA) + " ---- Add / Remove Programs"
AVI = AVI + 1: ARRAY_V(AVI) = "control main.cpl" + Space(25 - DIFA) + " ---- Mouse"
AVI = AVI + 1: ARRAY_V(AVI) = "control netcpl.cpl" + Space(21 - DIFA) + " ---- Network"
'control mmsys.cpl sounds
AVI = AVI + 1: ARRAY_V(AVI) = "control mmsys.cpl" + Space(22 - DIFA) + " ---- Sound Properties"
AVI = AVI + 1: ARRAY_V(AVI) = "control sysdm.cpl" + Space(23 - DIFA) + " ---- System"

AVI = 0
For R_AVI = 1 To UBound(ARRAY_V)
    If ARRAY_V(R_AVI) <> "" Then
        AVI = AVI + 1
        ARRAY_V_2(AVI) = UCase(ARRAY_V(R_AVI))
    End If
Next
ReDim Preserve ARRAY_V(AVI)
ReDim Preserve ARRAY_V_2(AVI)
For AVI = 1 To UBound(ARRAY_V_2)
    ARRAY_V(AVI) = ARRAY_V_2(AVI)
Next

For AVI = 1 To UBound(ARRAY_V)
    With ScanPath.ListView1
        Set LV = .ListItems.Add(, , ARRAY_V(AVI))
        LV.SubItems(1) = "CONTROL_PANEL_SHELL"
    End With
Next



'----
'Control Panel Command Line – Pahoehoe
'https://www.pahoehoe.net/control-panel-command-line/
'----

With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\WINDOWS GOD MODE FOLDER\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With

ReDim ARRAY_V(100)
ReDim ARRAY_V_2(100)
AVI = AVI + 1: ARRAY_V(AVI) = "Explorer.exe Shell:::{ED7BA470-8E54-465E-825C-99712043E01C}"




' -----------------------------------------------
AVI = 0
For R_AVI = 1 To UBound(ARRAY_V)
    If ARRAY_V(R_AVI) <> "" Then
        AVI = AVI + 1
        ARRAY_V_2(AVI) = ARRAY_V(R_AVI)
    End If
Next
ReDim Preserve ARRAY_V(AVI)
ReDim Preserve ARRAY_V_2(AVI)
For AVI = 1 To UBound(ARRAY_V_2)
    ARRAY_V(AVI) = ARRAY_V_2(AVI)
Next

For AVI = 1 To UBound(ARRAY_V)
    With ScanPath.ListView1
        Set LV = .ListItems.Add(, , ARRAY_V(AVI))
        LV.SubItems(1) = "CONTROL_PANEL_SHELL"
    End With
Next



'----
'Control Panel Command Line – Pahoehoe
'https://www.pahoehoe.net/control-panel-command-line/
'----
'
'Other Useful Tools  Command
'Windows 10 Settings start ms-settings:
'(Powershell Start-Process "ms-settings:"
'Computer Management compmgmt.msc
'Disk Management diskmgmt.msc
'Event Viewer    eventvwr.msc
'Device Manager  devmgmt.msc
'Services & Applications services.msc
'Local Group Policy Editor   gpedit.msc
'MMC (Management Console)    mmc
'Component Services  dcomcnfg
'Disk Cleanup    C:\windows\SYSTEM32\cleanmgr.exe /d{DRIVELETTER}
'System Configuration    msconfig
'System Information  msinfo32
'Task Scheduler  taskschd.msc
'Windows Firewall    wf.msc
'


With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\OTHER\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With

ReDim ARRAY_V(100)
ReDim ARRAY_V_2(100)
' AVI = AVI + 1: ARRAY_V(AVI) = "ms-settings"
AVI = AVI + 1: ARRAY_V(AVI) = "start ms-settings:"
AVI = AVI + 1: ARRAY_V(AVI) = "MMC DEVMGMT.MSC" + Space(24 - DIFA) + " ---- DEVICE MANAGER"
AVI = AVI + 1: ARRAY_V(AVI) = "MMC diskmgmt.msc"
AVI = AVI + 1: ARRAY_V(AVI) = "MMC eventvwr.msc"
AVI = AVI + 1: ARRAY_V(AVI) = "MMC services.msc"
AVI = AVI + 1: ARRAY_V(AVI) = ""
AVI = AVI + 1: ARRAY_V(AVI) = ""

AVI = 0
For R_AVI = 1 To UBound(ARRAY_V)
    If ARRAY_V(R_AVI) <> "" Then
        AVI = AVI + 1
        ARRAY_V_2(AVI) = UCase(ARRAY_V(R_AVI))
    End If
Next
ReDim Preserve ARRAY_V(AVI)
ReDim Preserve ARRAY_V_2(AVI)
For AVI = 1 To UBound(ARRAY_V_2)
    ARRAY_V(AVI) = ARRAY_V_2(AVI)
Next

For AVI = 1 To UBound(ARRAY_V)
    With ScanPath.ListView1
        Set LV = .ListItems.Add(, , ARRAY_V(AVI))
        LV.SubItems(1) = "CONTROL_PANEL_SHELL"
    End With
Next


With ScanPath.ListView1
    Set LV = .ListItems.Add(, , "\GOODSYNC PROFILE FOLDERING\")
    LV.SubItems(1) = "TITLE_BLOCK"
End With


NET_PATH_ALL_R = Split(NET_PATH_ALL_2, vbCrLf)
' Dim M()
ReDim M(UBound(NET_PATH_ALL_R) + 10)
' Dim R3
For R3 = 0 To UBound(NET_PATH_ALL_R)
    CK1 = UCase(NET_PATH_ALL_R(R3))
    ' ------------------------------------------------------
    ' YOU MIGHT THINK MY CODE IS CRAPPY
    ' BUT PUT >0 AFTER EACH INSTR AS WHEN COMPARE 2
    ' AND WITH AN AND STATEMENT BETWEEN
    ' IT CHANGE THE VALUE OF LOGIC
    ' SO NOT ABLE LEAVE >0 OUT
    ' INSTR RETURN RESULT HOW FAR IN THE STRING IS POSITION
    ' ------------------------------------------------------
    If InStr(CK1, "1_ASUS") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "2_ASUS") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "3_LINDA") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "E_DRIVE") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "2_ASUS") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "5_ASUS") > 0 And InStr(CK1, "D_DRIVE") > 0 Then
        NET_PATH_ALL_R(R3) = ""
    End If
    ' If InStr(CK1, "7_ASUS") > 0 And InStr(CK1, "D_DRIVE") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "8_MSI") > 0 And InStr(CK1, "D_DRIVE") > 0 Then NET_PATH_ALL_R(R3) = ""
Next
i = -1
For R3 = 0 To UBound(NET_PATH_ALL_R)
    CK1 = UCase(NET_PATH_ALL_R(R3))
    CK2 = NET_PATH_ALL_R(R3)
    If CK1 <> "" Then
        SET_GO = False
        If InStr(CK1, "4_ASUS") > 0 Then SET_GO = True
        If InStr(CK1, "7_ASUS") > 0 Then SET_GO = True
        If InStr(CK1, "8_MSI") > 0 Then SET_GO = True
        If SET_GO = True Then
        If InStr(CK1, "C_DRIVE") > 0 Then
            i = i + 1: M(i) = CK2 + "\GoodSync\Profile\jobs-groups-options.tic"
            ' -----------------------------------------------------------------
            ' GOODSYNC -- CHANGED THEIR PATH OTHER DAY FROM
            ' \ROAMING TO C:\Users\MATT 00\AppData\Local\GoodSync
            ' -----------------------------------------------------------------
            For R5 = 1 To 5
                GS_1 = CK2 + "\Users\MATT " + Format(R5, "00") + "\AppData\Local\GoodSync\jobs-groups-options.tic"
                If Dir(GS_1) <> "" Then
                    i = i + 1: M(i) = GS_1
                End If
            Next
            ' -----------------------------------------------------------------
        End If
        End If
        If InStr(CK1, "D_DRIVE") > 0 Then
            i = i + 1: M(i) = CK2 + "\GoodSync\Profile\jobs-groups-options.tic"
        End If
    End If
Next

ReDim Preserve M(i)

For R3 = 0 To UBound(M)
    If Dir(M(R3)) <> "" Then
        Path = "--DriveRemote_GS"
        
        With ScanPath.ListView1
            TT1 = M(R3)
            HDD = ""
            If InStr(TT1, "D_DRIVE") > 0 Then HDD = "D"
            If InStr(TT1, "C_DRIVE") > 0 Then HDD = "C"
            X5 = InStr(TT1, "_DRIVE")
            TT2 = Mid(TT1, X5 + 7)
            TT3 = Mid(TT2, 1, InStrRev(TT2, "\") - 1)
            TT1 = Mid(TT1, 1, InStr(4, TT1, "\") - 1) + " " + HDD + ":\" + TT3
            TT1 = Mid(TT1, 1, 8) + "  " + Mid(TT1, InStr(4, TT1, " "))
            TT1 = Replace(TT1, "MSI- ", "MSI    ")
            Set LV = .ListItems.Add(, , TT1)
            LV.SubItems(1) = Path
            TT3 = Mid(M(R3), 1, InStrRev(M(R3), "\") - 1)
            LV.SubItems(2) = TT3
        End With
    End If
Next







'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

ad = Dir("E:\01 VB Shell Folders\00 Shell *", vbDirectory)
Do
    If ad <> "" And FSO.FolderExists("E:\01 VB Shell Folders\" + ad) = True Then
            With ScanPath.ListView1
                Set LV = .ListItems.Add(, , ad)
                LV.SubItems(1) = "E:\01 VB Shell Folders\"
            End With
    End If
    ad = Dir
Loop Until ad = ""



EE = ""
For R = ScanPath.ListView1.ListItems.Count To 0 Step -1
    If InStr(EE, ScanPath.ListView1.ListItems.Item(R) + "**") > 0 Then
        ScanPath.ListView1.ListItems.Remove (R)
    End If
    EE = EE + ScanPath.ListView1.ListItems.Item(R) + "**"
Next
EE = ""

For R = ScanPath.ListView1.ListItems.Count To 0 Step -1
    If InStr(ScanPath.ListView1.ListItems.Item(R) + "--", "Program Files (x86)--") > 0 Then
        If FSO.FolderExists("C:\Program Files (x86)") = False Then
            ScanPath.ListView1.ListItems.Remove (R)
        End If
    End If
Next



End Sub


Sub GET_COMPUTR_NETWORK_NAME()

            B1 = NET_PATH_AND_DRIVE
            COMPUTER_NAME_VAR = Mid(B1, 3, InStr(B1, "__") - 3)
            COMPUTER_NAME_UNDERSCORE_VAR = Replace(COMPUTER_NAME_VAR, "-", "_")
            COMPUTER_NAME_DRIVE_LETTER_VAR = Mid(B1, InStr(B1, "__") + 2, 1)
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "C" Then DRIVE_VAR = "01"
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "D" Then DRIVE_VAR = "02"
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "E" Then DRIVE_VAR = "03"
            COMPUTER_NAME_PUT_01 = "\\" + COMPUTER_NAME_VAR + "\" + COMPUTER_NAME_UNDERSCORE_VAR + "_"
            COMPUTER_NAME_PUT_02 = DRIVE_VAR + "_" + COMPUTER_NAME_DRIVE_LETTER_VAR + "_DRIVE"
            If InStr(COMPUTER_NAME_PUT_02, "E_DRIVE") > 0 Then
                'fat32_4gb
                COMPUTER_NAME_PUT_02 = Replace(COMPUTER_NAME_PUT_02, "E_DRIVE", UCase("fat32_4gb"))
            End If
            COMPUTER_NAME_PUT = COMPUTER_NAME_PUT_01 + COMPUTER_NAME_PUT_02
            ' \\8-msi-gp62m-7rd\8_msi_gp62m_7rd_02_d_drive
            NET_PATH_AND_DRIVE = COMPUTER_NAME_PUT

End Sub


Sub LabelClick(Index)

If SetTrueToLoadLast = False Then
    H1$ = ScanPath.ListView1.ListItems.Item(Index).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(Index).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(Index)
    C1$ = Form1.Label1.Item(Index)
End If

TARGET_PATH_ALREADY_GOT = False

If A1$ = "--SPECIAL" Then
    A1$ = ""
    ' COME AFTER THE SPACE AND INDEX NUMBER
    ' -------------------------------------
    B1$ = Mid(B1$, InStr(B1$, " ") + 1)
    If H1$ <> "" Then B1$ = H1$
    
    TARGET_PATH_ALREADY_GOT = True
End If

'Call Form1.LoadLoggs

If A1$ = "CONTROL_PANEL_SHELL" Then
    SET_GO = 0
    If InStr(B1$, " ---- ") > 0 Then
    B1$ = Trim(Mid(B1$, 1, InStr(B1$, " ---- ") - 1))
    End If
    If UCase(B1$) = UCase("start ms-settings:") Then
        SET_GO = 1
    End If
    If SET_GO = 0 Then Shell B1$, vbMaximizedFocus
    If SET_GO = 1 Then
    
        
        Call EXPLORER_WITH_SHELL("/Select,", PATH_WANTER)
        Call SHELL_CMD(B1$, "CMD /C """ + B1$ + """")
    End If
End If


If A1$ = "EXPLORER_SHELL" Then
    If CLIPBOARDOR_PATH_NAME = True Then
        Call Form1.CLIP_PATH_NAME(B1$)
    End If
    
    Call EXPLORER_WITH_SHELL("", B1$)

End If


If SetTrueToLoadLast = True Then
            
    PATH_WANT = Trim(B1$)
    If IsNumeric(Mid(PATH_WANT, 1, 2)) Then
        PATH_WANT = Trim(Mid(PATH_WANT, 3))
    End If
    
    ' ------------------------------------------------------------
    If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
        ' SELECT THE NETWORK FOLDER
        ' -------------------------------
        If NETWORK_2_STEP_JUMPER > 0 Then
            If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = "" Then
                NETWORK_2_STEP_JUMPER = NETWORK_2_STEP_JUMPER - 1
                Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True
                Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Caption = "NETWORK DRIVE SELECT - " + Format(2 - NETWORK_2_STEP_JUMPER, "00") + " OF 02"
            End If
            COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = COMPUTER_NAME_PUT
            Beep
            
            
            If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = "" Then
                Exit Sub
            End If
            If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = "" Then
                Exit Sub
            End If
            
            Call LOAD_NETWORK_PATH_TO_EXPLORER
            
        End If
        
        PATH_WANT = D1$
        
        If CLIPBOARDOR_PATH_NAME = True Then
            Call Form1.CLIP_PATH_NAME(PATH_WANT)
        End If
        ' ------------------------------------------------------------
    End If
    
    Call EXPLORER_WITH_SHELL("/Select,", PATH_WANT)
End If


que = 0
If Mid(A1$, 1, 2) = "--" Then
    If A1$ = "--Drive" And Mid(B1, 2, 1) = ":" Then
        B1 = Mid(B1, 1, 2)
        'Shell "Explorer.exe /e," + B1$, vbNormalFocus
        'Shell "Explorer.exe /select," + B1$, vbNormalFocus
        If CLIPBOARDOR_PATH_NAME = True Then
            Call Form1.CLIP_PATH_NAME(B1$)
        End If
        
        Call EXPLORER_WITH_SHELL("", B1$)
        Exit Sub
    End If
    
    If A1$ = "--DriveRemote_USER" And Mid(B1, 1, 2) = "\\" Then
    
    
            COMPUTER_NAME_PUT = H1$
            If CLIPBOARDOR_PATH_NAME = True Then
                Call Form1.CLIP_PATH_NAME(COMPUTER_NAME_PUT)
            End If
            
'            If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
'                COMPUTER_NAME_PUT = D1$
'            End If
            
            Call EXPLORER_WITH_SHELL("", COMPUTER_NAME_PUT)

    End If
    
    If A1$ = "--DriveRemote_GS" Then 'And Mid(B1, 1, 2) = "\\" Then
    
    
            COMPUTER_NAME_PUT = H1$
            If CLIPBOARDOR_PATH_NAME = True Then
                Call Form1.CLIP_PATH_NAME(COMPUTER_NAME_PUT)
            End If
            
'            If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
'                COMPUTER_NAME_PUT = D1$
'            End If
            
            Call EXPLORER_WITH_SHELL("", COMPUTER_NAME_PUT)
    End If
    
    If A1$ = "--DriveRemote" And Mid(B1, 1, 2) = "\\" Then
        If InStr(B1, "__") > 0 Then
'                        NET_C_PATH = NET_C_PATH + Filename_VAR(R_L) + "__C" + vbCrLf
'            NET_D_PATH = NET_D_PATH + Filename_VAR(R_L) + "__D" + vbCrLf
'            NET_E_PATH = NET_E_PATH + Filename_VAR(R_L) + "__C" + vbCrLf

            COMPUTER_NAME_VAR = Mid(B1, 3, InStr(B1, "__") - 3)
            COMPUTER_NAME_UNDERSCORE_VAR = Replace(COMPUTER_NAME_VAR, "-", "_")
            COMPUTER_NAME_DRIVE_LETTER_VAR = Mid(B1, InStr(B1, "__") + 2, 1)
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "C" Then DRIVE_VAR = "01"
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "D" Then DRIVE_VAR = "02"
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "E" Then DRIVE_VAR = "03"
            COMPUTER_NAME_PUT_01 = "\\" + COMPUTER_NAME_VAR + "\" + COMPUTER_NAME_UNDERSCORE_VAR + "_"
            COMPUTER_NAME_PUT_02 = DRIVE_VAR + "_" + COMPUTER_NAME_DRIVE_LETTER_VAR + "_DRIVE"
            If InStr(COMPUTER_NAME_PUT_02, "E_DRIVE") > 0 Then
                'fat32_4gb
                COMPUTER_NAME_PUT_02 = Replace(COMPUTER_NAME_PUT_02, "E_DRIVE", UCase("fat32_4gb"))
            End If

            COMPUTER_NAME_PUT = COMPUTER_NAME_PUT_01 + COMPUTER_NAME_PUT_02
            ' \\8-msi-gp62m-7rd\8_msi_gp62m_7rd_02_d_drive
            
            
            ' SELECT THE NETWORK FOLDER
            ' -------------------------------
            If NETWORK_2_STEP_JUMPER > 0 Then
                If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = "" Then
                    NETWORK_2_STEP_JUMPER = NETWORK_2_STEP_JUMPER - 1
                    Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True
                    Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Caption = "NETWORK DRIVE SELECT - " + Format(2 - NETWORK_2_STEP_JUMPER, "00") + " OF 02"
                End If
                COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = COMPUTER_NAME_PUT
                Beep
                
                
                If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = "" Then
                    Exit Sub
                End If
                If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = "" Then
                    Exit Sub
                End If
                
                Call LOAD_NETWORK_PATH_TO_EXPLORER
                
            End If
            
            If CLIPBOARDOR_PATH_NAME = True Then
                Call Form1.CLIP_PATH_NAME(COMPUTER_NAME_PUT)
            End If
            
            If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
                COMPUTER_NAME_PUT = D1$
            End If
            
            Call EXPLORER_WITH_SHELL("", COMPUTER_NAME_PUT)
            
        Else
            If CLIPBOARDOR_PATH_NAME = True Then
                Call Form1.CLIP_PATH_NAME(B1$)
            End If
            
            Call EXPLORER_WITH_SHELL("", B1$)
        End If
    End If
    
    If A1$ = "--Drive" Then
        'Shell "Explorer.exe /e," + B1$, vbNormalFocus
        'Shell "Explorer.exe /select," + B1$, vbNormalFocus
            If CLIPBOARDOR_PATH_NAME = True Then
                Call Form1.CLIP_PATH_NAME(B1$)
            End If
        
        Call EXPLORER_WITH_SHELL("", B1$)
    End If
End If

If TARGET_PATH_ALREADY_GOT = True Then
    PATH_WANTER = B1$
    
    ' MsgBox PATH_WANTER
    If CLIPBOARDOR_PATH_NAME = True Then
        Call Form1.CLIP_PATH_NAME(PATH_WANTER)
    End If
    
    Call EXPLORER_WITH_SHELL("/Select,", PATH_WANTER)

End If

If SetTrueToLoadLast = False Then
    COUNT_EXIT = 0
    Do
        Call GETSHORTLINK(A1$ + B1$)
        D1$ = txtTargetPath
        ' -----------------------------------------------
        ' DON'T KNOW SEEM RUN TWICE FOR CORRECT RESULT
        ' OR SPEED
        ' GETSHORTLINK IS WANT RUN FEW TIME TO RETURN D1$
        ' -----------------------------------------------
        F1$ = GetLongName(D1$)
        DoEvents
        If F1$ <> ":" Then Exit Do
        Sleep 100
        COUNT_EXIT = COUNT_EXIT + 1
    Loop Until F1$ <> ":" Or COUNT_EXIT = 500
    
    If D1$ = "" Then
        MsgBox "SHELL EXPLORER" + vbCrLf + "RETRY 500" + vbCrLf + "Call GETSHORTLINK(A1$ + B1$) RESULT NOTHING IN D1$"
        Unload Me
        Exit Sub
    End If
    
    If Len(F1$) < 2 Then
        MsgBox "SHELL EXPLORER" + vbCrLf + "RETRY 500" + vbCrLf + "F1$ = GetLongName(D1$) F1$ RESULT NOTHING"
        Unload Me
        Exit Sub
    End If
    
    If Trim(D1$) = "" Then
        If CLIPBOARDOR_PATH_NAME = True Then
            Call Form1.CLIP_PATH_NAME(A1$ + B1$)
        End If
        vLaunch A1$ + B1$
        End
    End If
End If

If SetTrueToLoadLast = True Then

    Form1.Dir1.Path = txtTargetPath
    Form1.File1.Path = txtTargetPath
    If Form1.Dir1.ListCount > 0 Then
        PATH_WANTER = Form1.Dir1.List(0)
    Else
        PATH_WANTER = Form1.File1.Path + "\" + Form1.File1.List(0)
    End If

    ' -------------------------------------------------
    ' IF ASK FOR THIS HERE CLIPBOARDOR_PATH_NAME
    ' AND THEN GET DOUBLE OFFERING
    ' AND CLIPBOARD TWICE
    ' AND THE 1ST TIME AROUND
    ' WHEN DOUBLE STEP BUILD PATH UP NETOWRKER
    ' AND 2ND TIME
    ' WHEN SECOND TIME MIGHT BE TWO HITTER QUICKER DONE
    ' UNLESS CODE IN - MORE OFTEN
    ' -------------------------------------------------
    ' I HAVE TO FIND OUT BY TRACE DEBUGGER
    ' WHAT THE 2 STEP VALUE CHANGE IS LIKE
    ' -------------------------------------------------
    ' HA HA
    ' -------------------------------------------------

    If CLIPBOARDOR_PATH_NAME = True Then
        Call Form1.CLIP_PATH_NAME(PATH_WANTER)
    End If
        
    If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
        If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = "" Then
            NETWORK_2_STEP_JUMPER = NETWORK_2_STEP_JUMPER - 1
        End If
        Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Caption = "NETWORK DRIVE SELECT - " + Format(2 - NETWORK_2_STEP_JUMPER, "00") + " OF 02"
        COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = PATH_WANTER
        Beep
    End If
    
    SET_GO = True
    If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = "" Then SET_GO = False
    If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = "" Then SET_GO = False
    
    If SET_GO = False Then
        If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
            Exit Sub
        End If
    End If
    If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = False Then
        SET_GO = True
    End If
    
    If SET_GO = True Then
        Call LOAD_NETWORK_PATH_TO_EXPLORER
    End If

    
    If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
        PATH_WANTER = D1$
    End If
    
    Call EXPLORER_WITH_SHELL("/Select,", PATH_WANTER)
End If


If D1$ <> "" Then

    ' COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 -- NETWORK PATH
    ' COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 -- LOCAL   PATH
    ' IF NETWORK PATH DONE THEN READY TO GO
    ' OR GET LOCAL PATH AND WAIT NETWORK PATH
    '
    ' IF NOTHING AT NETWORK THEN LOCAL GET
    '
    ' IF NETWORK 2 STEP MODE
    ' WAIT UNTIL BOTH PARAM GOT
    '
    ' I AM HERE BUT MORE INTERESTED IN COLLECT A LOCAL PATH
    '
    If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
        If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = "" Then
            NETWORK_2_STEP_JUMPER = NETWORK_2_STEP_JUMPER - 1
        End If
        Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Caption = "NETWORK DRIVE SELECT - " + Format(2 - NETWORK_2_STEP_JUMPER, "00") + " OF 02"
        COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = F1$
        Beep
    End If
    
    SET_GO = True
    If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = "" Then SET_GO = False
    If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = "" Then SET_GO = False
    
    If SET_GO = False Then
        If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
            
            Exit Sub
        End If
    End If
    If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = False Then
        ' WORK UNDERWAY UNFINSIH PART HAD ERROR WANT BEHAVIOUR
        ' INVERT STATE SOUND REASON
        SET_GO = False
    End If
    
    If SET_GO = True Then
        Call LOAD_NETWORK_PATH_TO_EXPLORER
    End If

    Call EXPLORER_WITH_SHELL("", D1$)
        
End If

End Sub

Public Sub SHELL_CMD(FOLDER_NAME, FULL_COMMAND)

    If MNU_TREESIZE_GO = True Then
        Exit Sub
    End If
    If MNU_TREESIZE_GO = True Then
        TREESIZE_EXE = "C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe"
        Shell TREESIZE_EXE + " " + FOLDER_NAME, vbMaximizedFocus
        Stop
        End
    End If

    If CLIPBOARDOR_PATH_NAME = True Then
        Call Form1.CLIP_PATH_NAME(FOLDER_NAME)
    End If

    Shell FULL_COMMAND, vbMaximizedFocus

'    vLaunch "CMD /C ""start ms-settings:""", vbMaximizedFocus
'    Dim objShell
'    Set objShell = CreateObject("Wscript.Shell")
'    objShell.Run "CMD /K " + B1$, 0, True
'    Set objShell = Nothing

End Sub

Private Sub Form_Load()

' HERE ROUTINE GET CALL FOR FORM_LOAD
' Public Sub FormStartLoader()
' NOT HERE FORM_LOAD
    
Call SET_UP_PULIC_FSO

End Sub


Public Sub EXPLORER_WITH_SHELL(SELECT_OPTION, FOLDER_NAME)

    If MNU_TREESIZE_GO = True Then
        If Mid$(FOLDER_NAME, Len(FOLDER_NAME), 1) = ":" Then
            FOLDER_NAME = FOLDER_NAME + "\"
        End If
    End If
    
    If FormStart.MNU_TREESIZE_GO = True Then
        TREESIZE_EXE = "C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe"
        SCRIPT_PATH_LOADER = "D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer Loader\VBS 40-RUN EXE EXPLORER LANCHER.VBS"
        ' ----------------------------------------------------
        ' ENCODE -- WELL DONE IT AT LONG LAST NEW IDEA IN HEAD
        ' [ Wednesday 21:09:50 Pm_18 September 2019 ]
        ' ----------------------------------------------------
        ' ----
        ' Percent -encoding - Wikipedia
        ' https://en.wikipedia.org/wiki/Percent-encoding
        ' ----
        ' ----------------------------------------------------
        If Dir(SCRIPT_PATH_LOADER) = "" Then MsgBox "FILE NOT FIND" + vbCrLf + SCRIPT_PATH_LOADER: Stop
        
        TREESIZE_EXE = Replace(TREESIZE_EXE, " ", "*")
        ' TREESIZE_EXE = "MM"
        FOLDER_NAME = Replace(FOLDER_NAME, " ", "*")
    
        Dim WSHShell
        Set WSHShell = CreateObject("WScript.Shell")
            l = """" + SCRIPT_PATH_LOADER + """" + " " + TREESIZE_EXE + " " + FOLDER_NAME
            Debug.Print l
            WSHShell.Run l, ShowWindow_2, DontWaitUntilFinished
        Set WSHShell = Nothing
        
        ' SHELL WONT RUN VBS
        ' Shell SCRIPT_PATH_LOADER
        End
        
    End If
    
    If CLIPBOARDOR_PATH_NAME = True Then
        Call Form1.CLIP_PATH_NAME(FOLDER_NAME)
    End If
    SELECT_OPTION = Trim(SELECT_OPTION) + " "
    If Trim(SELECT_OPTION) = "" Then SELECT_OPTION = ""
    Shell "Explorer.exe " + SELECT_OPTION + FOLDER_NAME, vbMaximizedFocus
    End

End Sub

Sub LOAD_NETWORK_PATH_TO_EXPLORER()
    
    SET_GO = False
    If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 <> "" Then SET_GO = True
    If COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 <> "" Then SET_GO = True
    If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True Then
        If SET_GO = True Then
            F1$ = COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02
        End If
    End If
    If Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = False Then
        SET_GO = True
    End If
    
    PATH_USE_X = COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 + Mid(F1$, 3)
    ' -----------------------------------------------------------
    ' HERE ABOVE IT WILL HAVE THE PATH FROM SHORTCUT
    ' AND SHORTCUT WILL ALWAYS BE SHORTNAME
    ' F1$ IS THE LONG NAME -- ONLY IS PATH EXIST ORGINAL LOCATION
    ' FROM THE SHORTCUT -- OR NOT POSIABLE SHORT TO LONG
    ' FOR RELOCATION
    ' -----------------------------------------------------------
    ' X1 = F1$ ' X1 = GetLongName(PATH_USE_X)
    ' If X1 = "" Then X1 = PATH_USE_X
    X1 = PATH_USE_X
    DR1 = Mid(X1, InStr(X1, "_0") + 1, 2)
    ' GOT NUMBER NETWORK PATH -- 01 02 OR 03
    DR2 = Format(Val(DR1) + 64, "00")
    DR2 = DR1
    ' -----------------------------------------------------------
    ' DR1 GET DRIVE LETTER OF PATH REQUEST ON LOCAL
    ' DR2 IT WILL CONVERT DRIVE NUMERIC
    ' LETTER TO NUMBER 01 OR 02 OR 03 FOR NETWORK PATH C D OR E
    ' -----------------------------------------------------------

'    If Asc(DR1) > 64 And Asc(DR1) < 64 + 27 Then
'        DR2 = Format(Asc(DR1) - 50 - 10 - 4 - 2, "00")
'    End If
    
    X2 = X1
    ' IT DOES EXIST THEN NOT HAPPEN HERE
    ' ----------------------------------
    If FSO.FolderExists(X1) = False Then
    ' If Dir(X2, vbDirectory) = "" Then
        If InStr(X2, "_01_C") > 0 Then
            X2 = Replace(X2, "_01_C", "_" + DR2 + "_" + DR1)
        End If
        If InStr(X2, "_02_D") > 0 Then
            X2 = Replace(X2, "_02_D", "_" + DR2 + "_" + DR1)
        End If
        If InStr(X2, "_03_FAT32_4GB") > 0 Then
            X2 = Replace(X2, "_03_FAT32_4GB", "_" + DR2 + "_" + DR1)
        End If
    End If
    D1$ = X2
    ' COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_03 = D1$

    If FSO.FolderExists(D1$) = False Then
        MsgBox "FOLDER NOT EXIST" + vbCrLf + D1$ + vbCrLf + "TRY AND PLAY ANOTHER" + vbCrLf + "NOT ABLE RESELECT NETWORK FOLDER AGAIN YET"
        End
    End If
    '\\4-asus-gl522vw\4_asus_gl522vw_02_d_drive\0 00 ART LOGGERS
End Sub



