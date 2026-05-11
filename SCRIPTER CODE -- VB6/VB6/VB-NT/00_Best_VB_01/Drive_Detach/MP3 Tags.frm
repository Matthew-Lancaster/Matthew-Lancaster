VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Drive Detach"
   ClientHeight    =   2925
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8550
   Icon            =   "MP3 Tags.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4410
      Top             =   300
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3975
      Top             =   300
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   3540
      Top             =   300
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   8565
   End
   Begin VB.Menu Mnu_Exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Begin VB.Menu Mnu_OpenLogg_Folder 
         Caption         =   "Open Logg Folder"
      End
      Begin VB.Menu Mnu_Log 
         Caption         =   "Show Log Start Up"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Log_Change 
         Caption         =   "Show Log change"
      End
   End
   Begin VB.Menu Mnu_Minimise 
      Caption         =   "Minimise"
   End
   Begin VB.Menu Mnu_ListDrives 
      Caption         =   "List Drives"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
'Windows API/Global Declarations for :Dr
'     ive Type Finder
'**************************************
Public OldDrives$, VVV, OldTG$
Dim RD$(100), R2

Public Ff, Ffx, Rf, Z1$, StartUp, OldFF, OldFF2, LK, Xgag, HossPig


Private Sub Command1_Click()
R2 = Not R2
If R2 = 0 Then
    Shell "subst P: /d", vbMinimizedNoFocus
Else
    Shell "subst P: M:\00-C-Drive-Week", vbMinimizedNoFocus
End If

End Sub

Private Sub Form_Activate()

'Form1.WindowState = 1

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
End
Call Form_Resize
Me.WindowState = 1

StartUp = 1
OldTG$ = "Prog StartUp"
OldDrives$ = "Prog StartUp"

'---Optional
Open App.Path + "\Drive Detatch Change Logg.txt" For Input As #1
tt = " "
Do
    Line Input #1, II
    HH = Mid$(II, 5)
    If Len(HH) = 18 Then
    HH1 = HH
    'HH1 = DateValue(HH) + TimeValue(HH)
    'HH2=FORMAT(HH
    End If
    rr = InStr(II, "Name")
    If rr > 0 Then
        uu = Mid$(II, rr)
        If InStr(tt, uu) = 0 Then
            cc = cc + 1
            OO = OO + Format(cc, "000 ") + " - " + HH1 + " - " + uu + vbCrLf
            tt = tt + uu
    
        End If
    
    End If

Loop Until EOF(1)
Close #1


Open App.Path + "\## LIST OF DIFF DRIVES.txt" For Output As #1
Print #1, OO;
Close #1


'Debug.Print OO
'End
'RUN ONE END
Call Timer2_Timer
End
End Sub

Private Sub Form_Resize()

On Error Resume Next
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2


End Sub

Private Sub Form_Unload(Cancel As Integer)

If Me.WindowState <> 1 Then
    Me.WindowState = 1
    Cancel = True
End If

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Mnu_Exit_Click()
End
End Sub

Private Sub Mnu_ListDrives_Click()

Shell "Notepad.exe " + App.Path + "\## LIST OF DIFF DRIVES.txt", vbNormalFocus

End Sub

Private Sub Mnu_Log_Click()

Shell "Notepad.exe " + App.Path + "\Drive Detatch Start Up Logg.txt", vbNormalFocus
Me.WindowState = 1

End Sub

Private Sub Mnu_Log_Change_Click()

Shell "Notepad.exe " + App.Path + "\Drive Detatch Change Logg.txt", vbNormalFocus



End Sub

Private Sub Mnu_Minimise_Click()

Form1.WindowState = 1

End Sub

Private Sub Mnu_OpenLogg_Folder_Click()
Shell "explorer /e, " + App.Path + "\", vbNormalFocus
End Sub

Private Sub Timer1_Timer()
'End
End Sub

Private Sub Timer2_Timer()


On Error Resume Next


'Dim fs2
Set fs = CreateObject("Scripting.FileSystemObject")

Y2$ = Format$(Now, "DDD DD-MMM-YY HH:MM:SS") + vbCrLf
Y1$ = ""
tg = 0
tz$ = ""

Err.Clear
'Set z = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName("X:")))
'pp = z.DriveType

'rr$ = Dir$("Y:\*.*")
'If Err.Number > 0 And VVV = 0 Or rr$ = "" Then
'    VVV = Now + TimeSerial(0, 0, 20)
'    Form2.Visible = True
'End If

'If VVV > 0 And VVV < Now Then
'    Timer4.Enabled = True
'End If

'1=A to 26=Z
Drives$ = ""
gg$ = ""
For r = 3 To 25
    Err.Clear
    Set z = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(Chr$(r + 64) + ":")))
     
    Select Case z.DriveType
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    'err.description
    If Err.Number = 0 Then
        
        HossPig = 0
        If InStr("RS", z.DriveLetter) Then HossPig = 1
        
        gg$ = gg$ + z.DriveLetter
        tt$ = z.DriveLetter + ":" + "- Type - " + t + " --- Serial - " + Hex$(z.serialnumber) + " - Vol.Name - " + z.VolumeName + vbCrLf
        If tt$ <> tz$ And HossPig = 0 Then
            Drives$ = Drives$ + z.DriveLetter
            tz$ = tt$
            RD$(tg) = tt$
            tg = tg + 1
            Y1$ = Y1$ + tt$
        End If
    End If
Next

'If StartUp = 1 Then Open App.Path + "\Drive Detatch Start Up Logg.txt" For Append As #1

On Error GoTo Errortrap


If Y1$ <> Z1$ Then
    If Len(Drives$) > Len(OldDrives$) Then drin = True Else drin = False
    
    Open App.Path + "\Drive Detatch Change Logg.txt" For Append As #1

    If List1.ListCount > 0 Then
        List1.AddItem ""
    End If
    
    Print #1,
    If StartUp = 1 Then
        List1.AddItem "#Program StartUp ---------------------------"
        Print #1, List1.List(List1.ListCount - 1)
    End If
    
    List1.AddItem Y2$

    Print #1, List1.List(List1.ListCount - 1)
    List1.AddItem Trim(Str(tg - 1)) + " Drives Connected"
    Print #1, List1.List(List1.ListCount - 1)

    OldFF = ""
    For r = 0 To tg - 1
        List1.AddItem RD$(r)
        Print #1, List1.List(List1.ListCount - 1);
        OldFF = OldFF + RD$(r)
    Next
    Print #1,

    'If OldTG < 0 Then OldTG = 0

    If drin = False Then
        List1.AddItem "Less Drives Connected = " + Trim(Str(tg - 1)) + " --- Before = " + OldTG$
        Print #1, List1.List(List1.ListCount - 1)
    Else
        List1.AddItem "More Drives Connected = " + Trim(Str(tg - 1)) + " --- Before = " + OldTG$
        Print #1, List1.List(List1.ListCount - 1)
    End If

    OldTG$ = Trim(Str(tg - 1))
    

    List1.AddItem "Old Drives Connected   = " + OldDrives$
    Print #1, List1.List(List1.ListCount - 1)
    List1.AddItem "New Drives Connected = " + Drives$
    Print #1, List1.List(List1.ListCount - 1)


xag = 0
If tg - 1 > LK Then
If OldFF2 = OldFF Then xag = 1
OldFF2 = OldFF
End If
LK = tg - 1
List1.ListIndex = List1.ListCount - 1

'Xgag = 0

'    If fs.Driveexists("Y") Then
'        Set z = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName("Y:")))
'        On Error Resume Next
'        Err.Clear
'        If z.VolumeName = "<Disk not ready>" Then
'            Xgag = 1
'        End If
'        If Err.Description = "Disk not ready" Then Xgag = 1: MsgBox "Y Drive Problems"
'        On Error GoTo Errortrap
'    End If

'If fs.Driveexists("Y") = False Then Xgag = 1

If OldDrives$ = "Prog StartUp" Then OldDrives$ = ""
For r = 1 To Len(OldDrives$)
    If InStr(Drives$, Mid$(OldDrives$, r, 1)) = 0 Then mountun = mountun + Mid(OldDrives$, r, 1)
Next
For r = 1 To Len(Drives$)
    If InStr(OldDrives$, Mid$(Drives$, r, 1)) = 0 Then mountdo = mountdo + Mid(Drives$, r, 1)
Next
        
        If drin = True Then
            tthat$ = "New Drives Mounted = " + mountdo
        Else
            tthat$ = "New Drives UnMounted = " + mountun
        End If
    
        List1.AddItem tthat$
        Print #1, List1.List(List1.ListCount - 1)
        List1.ListIndex = List1.ListCount - 1

Close #1


OldDrives$ = Drives$


'dont pop up for removeables networks
xg = 0
For r = 1 To Len(Drive$)
    Set z = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(Mid(Drive$, r, 1) + ":")))
    If z.DriveType <> "Removable" Then xg = 1
Next

If xg = 0 Then xag = 1



If StartUp = 0 And xag = 0 Then
    Form1.WindowState = 0
End If
StartUp = 0
End If

Z1$ = Y1$

Errortrap:
If Err.Number = 0 Then Exit Sub
MsgBox "Error #" + Str(Err.Number) + vbCrLf + " Describe #" + vbCrLf + Err.Description
'Stop
Resume Next
End Sub


Private Sub Timer4_Timer()
'    Shell "D:\VB6\VB-NT\00_Best_VB_01\VPN_AutoDial\VPN_Auto-Dialer.exe", vbNormalNoFocus
'    Form2.Visible = False
'    Timer4.Enabled = False
'    VVV = 0
End Sub
