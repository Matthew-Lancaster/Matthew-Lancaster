VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Drive File Lister"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   19080
   Icon            =   "MP3 Tags.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   19080
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   5520
      Left            =   13455
      TabIndex        =   19
      Top             =   0
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   9737
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4860
      Top             =   3300
   End
   Begin MSComctlLib.ProgressBar ProgBar1 
      Height          =   405
      Left            =   -15
      TabIndex        =   17
      Top             =   810
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5175
      Top             =   2280
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   3075
      TabIndex        =   9
      Top             =   1200
      Width           =   10365
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7800
      Top             =   4125
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5520
      Left            =   15375
      TabIndex        =   20
      Top             =   0
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   9737
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4304
      EndProperty
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   5520
      Left            =   17130
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   9737
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   420
      Width           =   13440
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "FreeSpace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   4995
      Width           =   3075
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   4695
      Width           =   3075
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Used"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5295
      Width           =   3075
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   13
      Top             =   4440
      Width           =   3075
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Total Meg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   4140
      Width           =   3075
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   11
      Top             =   3885
      Width           =   3075
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Meg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   3600
      Width           =   3075
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Total Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   3060
      Width           =   3075
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   3345
      Width           =   3075
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   3075
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Scanning"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   2460
      Width           =   3075
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   3075
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   1530
      Width           =   3075
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13440
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Is To Do"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   1830
      Width           =   3075
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "To Do"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   1215
      Width           =   3075
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Begin VB.Menu Mnu_Log 
         Caption         =   "Show Log"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'STARTS ON Private Sub Timer2_Timer()


'**************************************
'Windows API/Global Declarations for :Dr
'     ive Type Finder
'**************************************
'Const ape$ = "M:\01_Drive_Lists"
Const ape$ = "M:\00 Loggers Text\01_Drive_Lists"
Dim Pool
Dim TT As Long, FileSpec


Dim Dri$(20)
Public TFV
Public TS1, TFS1, TSFS1 As Currency
Public TaMeg As Currency
Public ToMeg As Currency
Public TfMeg As Currency


Private Sub Form_Activate()

If IsIDE = False Then Form1.WindowState = 1

End Sub

Private Sub Form_Load()

Call Form_Resize

Form1.Refresh



'Good Code to Launch VBP IN autorun IDE for Debug Intermitant Faults afetr you have started froming from exe
'Try to Make sure your project filename is same as EXe Easyier then
'Extra Process low Later if you want

'FileSpec = App.Path + "\" + App.EXEName + ".vbp"
'If IsIDE = False And Dir$(FileSpec) <> "" Then
'    TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec + """", vbMinimizedNoFocus)
'    'Call KingMod.Process_Low(TT)
'    End
'End If

On Error Resume Next
Kill "M:\00 Loggers Text\01_Drive_Lists\01 Drive List --Drive--*.TXT"



End Sub

Private Sub Form_Resize()
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True Then
        If Control.Width + Control.Left > xx Then xx = Control.Width + Control.Left
        If Control.Height + Control.Top > yy Then yy = Control.Height + Control.Top
    End If
Next
'On Error GoTo 0

Form1.Width = xx + 100
Form1.Height = yy + 740

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Mnu_Log_Click()

Shell "Notepad2.exe " + ape$ + "\01 Drive List -- Total List All Log.txt", vbNormalFocus



End Sub

Private Sub Timer1_Timer()
End
End Sub

Private Sub Timer2_Timer()

Timer2.Enabled = False

'Load Mp3_Tag
'ScanPath.Show
List1.AddItem "Drive List Top File Count Each Day"
List1.AddItem ""

On Local Error Resume Next
freef1 = FreeFile
Open ape$ + "\01 Drive List -- Total List All Log.txt" For Input As #freef1
Do
    Line Input #1, a1$
    List1.AddItem a1$
Loop Until EOF(1)
Close #1
On Local Error GoTo 0


List1.AddItem ""

List1.ListIndex = List1.ListCount - 1
List1.Refresh
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
Dim fs2
Set fs2 = CreateObject("Scripting.FileSystemObject")

Dim Drivetxt(50, 50)
Dim Drivesize(50)


Dim strDrive As String
Dim strMessage As String
Dim intCnt As Long
Load ScanPath

freef2 = FreeFile
Open ape$ + "\01 Drive Lists Total All\01 Drive List --" + Format$(Now, "YYYY-MM-DD HH-MM-SS ") + "Drive--" + strDrive + "-- Total All.txt" For Output As #freef2


List1.AddItem "---------------"
List1.AddItem "Drives To Do --"
List1.AddItem "---------------"



lastdrive = 86

For intCnt = 67 To lastdrive '67=C W=87  90=z '67 90
Do
    
    strDrive = Chr(intCnt)
    
    tp = fs.DriveExists(strDrive)
    
    If tp = True Then
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(strDrive + ":")))
    Err.Clear
    On Local Error Resume Next
    s = "Drive " & d.DriveLetter & ": - " & d.volumename
    If Err.Number > 0 Then tp = False
    End If
    
    Select Case GetDriveType(strDrive + ":\")
        Case DRIVE_REMOVABLE
        rtn = "Floppy Drive"
        Case DRIVE_FIXED
        rtn = "Hard Drive"
        Case DRIVE_REMOTE
        rtn = "Network Drive"
        Case DRIVE_CDROM
        rtn = "CD-ROM Drive"
        Case DRIVE_RAMDISK
        rtn = "RAM Disk"
        Case Else
        rtn = ""
    End Select
    
    If rtn = "Network Drive" Then tp = False
    
    xp$ = d.volumename
    If tp = True And InStr(Pool, xp$) = 0 Then
        List1.AddItem s
    End If
    Pool = Pool + xp$
    
    If tp = False Then intCnt = intCnt + 1
If intCnt > lastdrive Then Exit Do
Loop Until tp = True
If intCnt > lastdrive Then Exit For
Next
List1.AddItem "---------------"

List1.ListIndex = List1.ListCount - 1
Pool = ""
lastdrive = 86
For intCnt = 67 To lastdrive '67=C W=87  90=z '67 90
'For intCnt = 68 To 70 '67=C W=87  90=z '67 90

Do
    
    strDrive = Chr(intCnt)
    
    tp = fs.DriveExists(strDrive)
    
    Select Case GetDriveType(strDrive + ":\")
        Case DRIVE_REMOVABLE
        rtn = "Floppy Drive"
        Case DRIVE_FIXED
        rtn = "Hard Drive"
        Case DRIVE_REMOTE
        rtn = "Network Drive"
        Case DRIVE_CDROM
        rtn = "CD-ROM Drive"
        Case DRIVE_RAMDISK
        rtn = "RAM Disk"
        Case Else
        rtn = ""
    End Select
    
    If rtn = "Network Drive" Then tp = False
    
    xp$ = d.volumename
    
    'Make sure subst drives are not included you can do this chk vol name keep all subst drive after your drive
    'after you drive you subst on
    If tp = True And InStr(Pool, xp$) = 0 Then
    Pool = Pool + xp$
    
    
    On Error Resume Next
    'Set G = fs.GetDrive(fs.GetDriveName(strDrive))
    Err.Clear
    Set g = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(strDrive + ":")))
    'g.volumename
    ease = g.volumename
    If ease = "<Disk not ready>" Or ease = "" Then tp = False
    If Err.Number = 0 And tp = True Then
        TS1 = g.totalsize
        TFS1 = g.freespace
        TSFS1 = TS1 - TFS1
        On Local Error Resume Next
        ProgBar1.Max = TSFS1
        On Local Error GoTo 0
    
        TaMeg = TaMeg + TS1
        ToMeg = ToMeg + TSFS1
        TfMeg = TfMeg + TFS1
    
        Label16 = Format$(TS1 / (1048576), "##,###,###,### ") + "MB Size"
        Label17 = Format$(TFS1 / (1048576), "##,###,### ") + "MB Free"
        Label14 = Format$(TSFS1 / (1048576), "##,###,###,### ") + "MB Used"
    
        'LblSizetxt.Caption = Format$(HogX1 / 1048576, "####,###,000")
    
        Form1.Label11.Caption = Format$(TSFS1 / 1048576, "####,###,000")
        Form1.Label13.Caption = Format$(ToMeg / 1048576, "####,###,000")
    End If
    
    End If
    
    If tp = True Then
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(strDrive + ":")))
    Err.Clear
    On Local Error Resume Next
    s = "Drive " & d.DriveLetter & ": - " & d.volumename
    If Err.Number > 0 Then tp = False
    
    
    On Local Error GoTo 0
    End If
    
    'xp$ = fs.Getvolumename(strDrive)
    
    
    
If tp = False Then intCnt = intCnt + 1
If intCnt > lastdrive Then Exit Do
Loop Until tp = True

If intCnt > lastdrive Then Exit For

Label4.Caption = s
uy = uy + 1
Drivetxt(uy, 1) = s

Label5.Caption = "Doing.."
Label6.Caption = "Doing.."


ScanPath.txtPath = strDrive + ":\"
ScanPath.cboMask.Text = "*.*"

Call ScanPath.cmdScan_Click

TotalAll = TotalAll + ScanPath.lblCount.Caption
freef1 = FreeFile

Open ape$ + "\02 Drive Lists Individule\01 Drive List --" + "Drive--" + strDrive + "--.txt" For Output Lock Write As #freef1
freef5 = FreeFile
Open ape$ + "\03 Drive Lists Dates\01 Drive List Date Cre --" + "Drive--" + strDrive + "--.txt" For Output Lock Write As #freef5
freef7 = FreeFile
Open ape$ + "\03 Drive Lists Dates\01 Drive List Date Mod --" + "Drive--" + strDrive + "--.txt" For Output Lock Write As #freef7
freef8 = FreeFile
Open ape$ + "\03 Drive Lists Dates\01 Drive List Date Acc --" + "Drive--" + strDrive + "--.txt" For Output Lock Write As #freef8


stf2 = s + vbCrLf
stf2 = stf2 + "---------------------------------" + vbCrLf
stf2 = stf2 + "Date: " + Format$(Now, "DDD DD-MM-YYYY HH-MM-SS") + vbCrLf
stf2 = stf2 + "Total Files -- " + ScanPath.lblCount.Caption + vbCrLf
stf2 = stf2 + "Total Meg -- " + Label11.Caption + vbCrLf
stf2 = stf2 + "---------------------------------" + vbCrLf
stf2 = stf2 + "Total Files So Far -- " + Trim(Str(TotalAll)) + vbCrLf
stf2 = stf2 + "Total Meg So Far -- " + Label13.Caption + vbCrLf
stf2 = stf2 + "---------------------------------" + vbCrLf


Print #freef1, stf2
Print #freef2, stf2
Print #freef5, stf2
Print #freef7, stf2
Print #freef8, stf2

Drivetxt(uy, 2) = Str(ScanPath.lblCount.Caption)

tg2$ = Format$(Val(ScanPath.lblCount.Caption), "00,000,000 ") + "Files -- "
tg2$ = tg2$ + Format$((TS1 / 1048576), "0,000,000") + " Total Megs -- "
tg2$ = tg2$ + Format$((TSFS1 / 1048576), "0,000,000") + " Used Megs -- "
tg2$ = tg2$ + Format$((TFS1 / 1048576), "0,000,000") + " Free Meg -- "
tg2$ = tg2$ + "On " + s

TFV = TFV + 1
Dri$(TFV) = tg2$
List1.AddItem tg2$
List1.ListIndex = List1.ListCount - 1
List1.Refresh

Label6.Caption = ScanPath.lblCount.Caption

ListView1.ListItems.Clear
ListView2.ListItems.Clear
ListView3.ListItems.Clear

On Error Resume Next
For we = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ScanPath.ListView1.ListItems.Item(we)
    Label5.Caption = Trim(Str(we))
    
    Err.Clear
    Set f = fs.GetFile(a1$ + b1$)
    If Err.Number = 0 Then
        ttp = Format(f.datecreated, "YYYY-MM-DD - HH-MM-SS") + " -- " + a1$ + b1$
    Else
        ttp = "YYYY-MM-DD - HH-MM-SS" + " -- " + a1$ + b1$
    End If
    
    Err.Clear
    If Err.Number = 0 Then
        ttp2 = Format(f.datelastmodified, "YYYY-MM-DD - HH-MM-SS") + " -- " + a1$ + b1$
    Else
        ttp2 = "YYYY-MM-DD - HH-MM-SS" + " -- " + a1$ + b1$
    End If
    
    Err.Clear
    If Err.Number = 0 Then
        ttp3 = Format(f.datelastaccessed, "YYYY-MM-DD - HH-MM-SS") + " -- " + a1$ + b1$
    Else
        ttp3 = "YYYY-MM-DD - HH-MM-SS" + " -- " + a1$ + b1$
    End If
 
    With ListView1
        Set LV = .ListItems.Add(, , ttp)
    End With
    With ListView2
        Set LV = .ListItems.Add(, , ttp2)
    End With
    'With ListView3
    '    Set LV = .ListItems.Add(, , ttp3)
    'End With

    Print #freef1, "Cre - " + ttp + " - Mod - " + ttp1 + " - " + a1$; b1$
    Print #freef2, "Cre - " + ttp + " - Mod - " + ttp1 + " - "; a1$; b1$
Next
On Error GoTo 0

ListView1.SortOrder = lvwAscending
ListView1.SortKey = 0
ListView1.Sorted = True
ListView1.Sorted = False

ListView2.SortOrder = lvwAscending
ListView2.SortKey = 0
ListView2.Sorted = True
ListView2.Sorted = False

ListView3.SortOrder = lvwAscending
ListView3.SortKey = 0
ListView3.Sorted = True
ListView3.Sorted = False


For we = 1 To ListView1.ListItems.Count
    a1$ = ListView1.ListItems.Item(we)
    Print #freef5, a1$
Next
For we = 1 To ListView2.ListItems.Count
    a1$ = ListView2.ListItems.Item(we)
    Print #freef7, a1$
Next
For we = 1 To ListView3.ListItems.Count
    a1$ = ListView3.ListItems.Item(we)
    Print #freef8, a1$
Next

Print #freef2, ""
Print #freef2, ""
Print #freef2, ""

Close #freef1, #free5, #free7, #free8
Next

Print #freef2, "---------------------------------"
Print #freef2, "Date: " + Format$(Now, "DDD DD-MM-YYYY HH-MM-SS")
Print #freef2, "Total Files So Far:"; TotalAll
Print #freef2, "Total Meg So Far:"; Label13.Caption
Print #freef2, "---------------------------------"

Close #freef2


'The Short Logg
freef1 = FreeFile
Open ape$ + "\01 Drive List -- Total List All Log.txt" For Append As #freef1

tt1$ = "Date: " + Format$(Now, "DDD DD-MM-YYYY HH-MM-SS") + " -- " + Format$(TotalAll, "00,000,000") + " - Total Files"
tt2$ = "Date: " + Format$(Now, "DDD DD-MM-YYYY HH-MM-SS") + " -- " + Format$(TaMeg / 1048576, "00,000,000") + " - Total Size Meg"
tt3$ = "Date: " + Format$(Now, "DDD DD-MM-YYYY HH-MM-SS") + " -- " + Format$(ToMeg / 1048576, "00,000,000") + " - Total Used Meg"
tt4$ = "Date: " + Format$(Now, "DDD DD-MM-YYYY HH-MM-SS") + " -- " + Format$(TfMeg / 1048576, "00,000,000") + " - Total Free Meg"

List1.AddItem ""
List1.AddItem tt1$
List1.AddItem tt2$
List1.AddItem tt3$
List1.AddItem tt4$

For r = List1.ListCount - 1 To 1 Step -1
    If List1.List(r) = "Drives To Do --" Then Exit For
Next

Print #freef1, ""

For r2 = r - 1 To List1.ListCount - 1
    Print #freef1, List1.List(r2)
Next

List1.AddItem "..........."
List1.AddItem "Done...."
List1.ListIndex = List1.ListCount - 1
List1.Refresh

Close #1

Timer3.Enabled = True

End Sub

Private Sub Timer3_Timer()
Shell "D:\VB6\VB-NT\00_Best_VB_01\WinRar Archive Drive Lists\WinRar Archive Drive Lists.exe", vbMinimizedNoFocus
End
End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************

