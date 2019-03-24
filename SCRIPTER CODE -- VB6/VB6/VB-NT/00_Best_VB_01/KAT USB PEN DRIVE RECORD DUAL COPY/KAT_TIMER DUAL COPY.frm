VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   ClientHeight    =   564
   ClientLeft      =   132
   ClientTop       =   840
   ClientWidth     =   11004
   LinkTopic       =   "Form1"
   ScaleHeight     =   564
   ScaleWidth      =   11004
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3300
      Top             =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "KAT RECORDER TO USB PEN DRIVE RUNNER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
   Begin VB.Menu MNU_MENU 
      Caption         =   "MOVE RECORDS"
      Begin VB.Menu MNU_RECORDS 
         Caption         =   "YES MOVE RECORDS OR NO COPY RECORDS"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ODUCKIE, TF, OTT

Public cProcesses As New clsCnProc
Dim TT, XHWND

Private Sub Form_Load()
If App.PrevInstance = True Then End
' = cProcesses.Convert(CurProcHwnd, Otss, cnFromhWnd Or cnToProcessID)
' = cProcesses.Convert(Otss, totss2, cnTohWnd Or cnFromProcessID)
'RF = cProcesses.GetEXEID(-1, "C:\Program Files\KatMouse\KatMouse.exe")

Label1 = "KAT RECORDER TO USB PEN DRIVE AND HDD DUAL RECORDER"
Label1.Left = 0
Label1.Top = 0
Me.Width = Label1.Width + 200
Me.Caption = "KAT RECORDER"
Me.Height = Label1.Height + 800
Set FS = CreateObject("Scripting.FileSystemObject")



End Sub

Private Sub Label1_Click()

If IsIDE = True Then Stop: End
    Me.Visible = False
    Me.Enabled = False
    Me.WindowState = 1
    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End

End Sub

Private Sub MNU_RECORDS_Click()
MNU_RECORDS.Checked = Not MNU_RECORDS.Checked
End Sub

Private Sub Timer1_Timer()

'Set F = FS.GETDRIVE("G:")
'D = F.FREESPACE
'If ODUCKIE = 0 Then ODUCKIE = D
'If ODUCKIE = D Then Exit Sub
'ODUCKIE = D

If Second(Now) Mod 5 = 0 Then
    Rf = cProcesses.GetEXEID(-1, "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.exe")
    If Rf = 0 Then
        Shell "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.exe", vbMinimizedNoFocus
    End If
End If

On Error Resume Next
DIRT = "D:\Kat MP3 Recorder\"
Set F = FS.GETFOLDER(DIRT)

If Err.Number > 0 Then
    Err.Clear
    MkDir "D:\Kat MP3 Recorder\"
    If Err.Number > 0 Then Exit Sub

End If

Set F1 = F.Files
TT = F1.Count

If OTT <> TT And TT > 1 Then
OTT = TT

For Each Files In F1
    If FS.FileExists("M:\05 Media\KAT RECORD\" + Files.Name) = False Then
        If FileInUse(DIRT + Files.Name) = False Then
            If IsIDE = True Then Debug.Print Time$ + " -- FILE COPIED " + DIRT + Files.Name
            On Error Resume Next
            If Files.Size < 3 * 1024 Then
                Kill Files.Name
                XJ = 1
                Else
                XJ = 0
            End If
            If XJ = 0 Then Files.Copy "M:\05 Media\KAT RECORD\" + Files.Name
            On Error GoTo 0
                    
        End If
    End If
Next

End If

If TT = 1 Then
    If IsIDE = True Then Debug.Print Time$ + " -- FILE COPIED " + DIRT + XTF
    For Each Files In F1
        On Error Resume Next
            Files.Copy "M:\05 Media\KAT RECORD\" + Files.Name
        On Error GoTo 0
        Exit For
    Next
End If

If MNU_RECORDS.Checked = True Then
    'CHKED = WE WANT TO MOVE
    'SO INSTEAD MOVE WE JUST DEL WATS LEFT AFTER COPY

    For Each Files In F1
        XTF = Files.Name
        If FS.FileExists("M:\05 Media\KAT RECORD\" + Files.Name) = True Then
            If FileInUse(DIRT + Files.Name) = False Then
                FS.DeleteFile DIRT + Files.Name
            End If
        End If
    Next
End If



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

