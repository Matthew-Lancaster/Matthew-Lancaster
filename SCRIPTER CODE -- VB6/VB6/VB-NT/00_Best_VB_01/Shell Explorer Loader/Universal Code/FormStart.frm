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


Dim Filename_VAR(40)
'Filename_VAR(1) = "\\1-ASUS-X5DIJ"
'Filename_VAR(2) = "\\2-ASUS-EEE"
'Filename_VAR(3) = "\\3-LINDA-PC"
'Filename_VAR(4) = "\\4-ASUS-GL522VW"
'Filename_VAR(5) = "\\5-ASUS-P2520LA"
'Filename_VAR(6) = "\\7-ASUS-GL522VW"
'Filename_VAR(7) = "\\8-MSI-GP62M-7RD"
'Filename_VAR(8) = "\\NAS-QNAP-ML"

'Load Frm_M_P_S
'For R_L = 1 To Frm_M_P_S.ListView1.ListItems.Count
'    If UCase(Frm_M_P_S.ListView1.ListItems(R_L).SubItems(2)) <> "BTHUB" Then
'        Filename_VAR(R_L) = "\\" + UCase(Frm_M_P_S.ListView1.ListItems(R_L).SubItems(2))
'        Filename_VAR(R_L) = Replace(Filename_VAR(R_L), ".HOME", "")
'    End If
'Next

FR_1 = FreeFile
NET_FILENAME = "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\NETWORK_COMPUTER_NAME.txt"
Open NET_FILENAME For Input As #FR_1
R_L = 0
Path = "--DriveRemote"
Do
    Line Input #FR_1, LINE_STINGER
    If LINE_STINGER <> "BTHUB" Then
        R_L = R_L + 1
        Filename_VAR(R_L) = "\\" + LINE_STINGER
    End If
    
    'Path = "--Drive"
    With ScanPath.ListView1
        Set LV = .ListItems.Add(, , Filename_VAR(R_L))
        LV.SubItems(1) = Path
    End With
    
Loop Until EOF(FR_1)
Close #FR_1


'COULD DO IN ONE BUT MAKE SURE ADD TO END
'----------------------------------------
FR_1 = FreeFile
NET_FILENAME = "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\NETWORK_COMPUTER_NAME.txt"
Open NET_FILENAME For Input As #FR_1
Path = "--DriveRemote"
Do
    Line Input #FR_1, LINE_STINGER
    
    If LINE_STINGER = "BTHUB" Then
        Filename = "HTTPS:\\" + LINE_STINGER + "\BTHUB"
        'Path = "--Drive"
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Filename)
            LV.SubItems(1) = Path
        End With
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


Dim R_X As Long

For R_X = 0 To 120
    If Mid(UCase(GetSpecialfolder(R_X)), 2) = UCase(":\USERS\" + GetUserName) Then
        GetSpecialfolder_VAR = GetSpecialfolder(R_X)
        Exit For
    End If
Next

For R = 67 To 90
    If Dir(Chr(R) + ":\TEMP", vbDirectory) <> "" Then
        Path = "--SPECIAL"
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , Chr(R) + ":\TEMP")
            LV.SubItems(1) = Path
        End With
    Else
        Exit For
    End If
Next


GetSpecialfolder_VAR = Mid(GetSpecialfolder_VAR, 1, Len(GetSpecialfolder_VAR) - 1)
For R_L = 1 To 9
    GET_USER_NAME_VAR_NAME = GetSpecialfolder_VAR + Format(R_L, "0")
    Filename = GET_USER_NAME_VAR_NAME
    If Dir(Filename, vbDirectory) <> "" Then
    
        Path = "--Drive"
        
        With ScanPath.ListView2
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
    If R = 0 Then q = 1
    If R = 2 Then q = 1
    If R = 5 Then q = 1
    If R = 6 Then q = 1
    If R = 9 Then q = 1
    If R = 11 Then q = 1
    If R = 16 Then q = 1
    If R = 20 Then q = 1
    If R = 22 Then q = 1
    If R = 25 Then q = 1
    If R = 26 Then q = 1
    If R = 31 Then q = 1
    If R = 35 Then q = 1
    If R = 36 Then q = 1
    If R = 37 Then q = 1
    If R = 38 Then q = 1
    If R = 39 Then q = 1
    If R = 40 Then q = 1
    If R = 41 Then q = 1
    If R = 42 Then q = 1
    If R = 53 Then q = 1
    If R = 54 Then q = 1
    If R = 55 Then q = 1
    If R = 56 Then q = 1
    
    If GetSpecialfolder(R) <> "" Then
        Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
    End If
    
    If GetSpecialfolder(R) <> "" And q = q Then
        Filename = GetSpecialfolder(R)
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
                Set LV = .ListItems.Add(, , Filename)
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
                Set LV = .ListItems.Add(, , Filename)
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

For R = 0 To ScanPath.ListView2.ListItems.Count
    Path = "--SPECIAL"
    With ScanPath.ListView1
        Set LV = .ListItems.Add(, , ScanPath.ListView2.ListItems.Item(R))
        LV.SubItems(1) = ScanPath.ListView2.ListItems.Item(R).SubItems(1)
    End With
Next

Set FSO = CreateObject("Scripting.FileSystemObject")

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

que = 0
If Mid(A1$, 1, 2) = "--" Then
    If A1$ = "--Drive" Then B1 = Mid(B1, 1, 2)
    'Shell "Explorer.exe /e," + B1$, vbNormalFocus
'    Shell "Explorer.exe /select," + B1$, vbNormalFocus
'  End
      Shell "Explorer.exe " + B1$, vbNormalFocus
  End
End If

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

End

End Sub

Private Sub Form_Load()

Call SET_UP_PULIC_FSO

End Sub
