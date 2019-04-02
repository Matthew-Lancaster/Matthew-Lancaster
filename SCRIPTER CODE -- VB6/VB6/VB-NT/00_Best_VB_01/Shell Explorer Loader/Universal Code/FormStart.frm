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

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'
Dim R As Long

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
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    
    
    If Err.Number = 0 Then
        tt$ = z.DriveLetter + ":" + "- Type - " + t + " --- Serial - " + Hex$(z.SerialNumber) + " - Vol.Name - " + z.VolumeName + vbCrLf
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
            Set LV = .ListItems.Add(, , Filename_VAR(R_L) + "__C")
            LV.SubItems(1) = Path
        End With
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Filename_VAR(R_L) + "__D")
            LV.SubItems(1) = Path
        End With
        If R_L_X > 0 Then
            With ScanPath.ListView1
                Set LV = .ListItems.Add(, , Filename_VAR(R_L) + "__E")
                LV.SubItems(1) = Path
            End With
        End If
    End If
Loop Until EOF(FR_1)
Close #FR_1




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
        f1 = UCase(Filename)
        If InStr(f1, "DOCUMENTS AND") > 0 And InStr(f1, "ADMINISTRATIVE") > 0 Then
            SET_GO = False
        End If
        If InStr(f1, "DOCUMENTS AND") > 0 And InStr(f1, "MICROSOFT\CD BURNING") > 0 Then
            SET_GO = False
        End If
        
        If InStr(f1, UCase("AppData\Roaming\Microsoft\Windows\Printer Shortcuts")) > 0 Then
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
                Set LV = .ListItems.Add(, , Format(R, "00 ") + Filename)
                LV.SubItems(1) = Path
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


'ScanPath.ListView1.SortOrder = lvwAscending
'ScanPath.ListView1.SortKey = 0
'ScanPath.ListView1.Sorted = True
'ScanPath.ListView1.Sorted = False
ScanPath.ListView2.SortOrder = lvwAscending
ScanPath.ListView2.SortKey = 0
ScanPath.ListView2.Sorted = True
ScanPath.ListView2.Sorted = False

For R = 1 To ScanPath.ListView2.ListItems.Count
    Path = "--SPECIAL"
    With ScanPath.ListView1
        Set LV = .ListItems.Add(, , ScanPath.ListView2.ListItems.Item(R))
        LV.SubItems(1) = ScanPath.ListView2.ListItems.Item(R).SubItems(1)
    End With
Next


'----
'Shell Commands List for Windows 10 | Tutorials
'https://www.tenforums.com/tutorials/3109-shell-commands-list-windows-10-a.html
'----
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

Sub LabelClick(Index)

If SetTrueToLoadLast = False Then
    A1$ = ScanPath.ListView1.ListItems.Item(Index).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(Index)
    C1$ = Form1.Label1.Item(Index)
End If

'Call Form1.LoadLoggs

If A1$ = "EXPLORER_SHELL" Then
    Shell "Explorer.exe " + B1$, vbMaximizedFocus ', vbNormalFocus
    Beep
    End
End If




que = 0
If Mid(A1$, 1, 2) = "--" Then
    If A1$ = "--Drive" And Mid(B1, 2, 1) = ":" Then
        B1 = Mid(B1, 1, 2)
        'Shell "Explorer.exe /e," + B1$, vbNormalFocus
        'Shell "Explorer.exe /select," + B1$, vbNormalFocus
        Shell "Explorer.exe " + B1$, vbMaximizedFocus ', vbNormalFocus
        End
    End If
    If A1$ = "--DriveRemote" And Mid(B1, 1, 2) = "\\" Then
        If InStr(B1, "__") > 0 Then
            
            COMPUTER_NAME_VAR = Mid(B1, 3, InStr(B1, "__") - 3)
            COMPUTER_NAME_UNDERSCORE_VAR = Replace(COMPUTER_NAME_VAR, "-", "_")
            COMPUTER_NAME_DRIVE_LETTER_VAR = Mid(B1, InStr(B1, "__") + 2, 1)
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "C" Then DRIVE_VAR = "01"
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "D" Then DRIVE_VAR = "02"
            If COMPUTER_NAME_DRIVE_LETTER_VAR = "E" Then DRIVE_VAR = "03"
            COMPUTER_NAME_PUT_01 = "\\" + COMPUTER_NAME_VAR + "\" + COMPUTER_NAME_UNDERSCORE_VAR + "_"
            COMPUTER_NAME_PUT_02 = DRIVE_VAR + "_" + COMPUTER_NAME_DRIVE_LETTER_VAR + "_DRIVE"
            COMPUTER_NAME_PUT = COMPUTER_NAME_PUT_01 + COMPUTER_NAME_PUT_02
            ' \\8-msi-gp62m-7rd\8_msi_gp62m_7rd_02_d_drive
            Shell "Explorer.exe " + COMPUTER_NAME_PUT, vbNormalFocus
            End
            
        Else
            Shell "Explorer.exe " + B1$, vbMaximizedFocus ', vbNormalFocus
            End
        End If
    End If
    
    If A1$ = "--Drive" Then
        'Shell "Explorer.exe /e," + B1$, vbNormalFocus
        'Shell "Explorer.exe /select," + B1$, vbNormalFocus
        Shell "Explorer.exe " + B1$, vbMaximizedFocus ', vbNormalFocus
        End
    End If
End If


If SetTrueToLoadLast = False Then
    Call GETSHORTLINK(A1$ + B1$)
    
    D1$ = txtTargetPath
    If Trim(D1$) = "" Then
        vLaunch A1$ + B1$
    Else
    
        Form1.Dir1.Path = txtTargetPath
        Form1.File1.Path = txtTargetPath
        If Form1.Dir1.ListCount > 0 Then
            PATH_WANTING = Form1.Dir1.List(0)
        Else
            PATH_WANTING = Form1.File1.Path + "\" + Form1.File1.List(0)
        End If
    
        ' MsgBox PATH_WANTING
        Shell "Explorer.exe /Select," + PATH_WANTING, vbNormalFocus
    End If
End If

If SetTrueToLoadLast = True Then
            
    PATH_WANT = Trim(B1$)
    If IsNumeric(Mid(PATH_WANT, 1, 2)) Then
        PATH_WANT = Trim(Mid(PATH_WANT, 3))
    End If
    
    Shell "Explorer.exe /Select," + PATH_WANT, vbMaximizedFocus
End If

End

End Sub

Private Sub Form_Load()

Call SET_UP_PULIC_FSO

End Sub
