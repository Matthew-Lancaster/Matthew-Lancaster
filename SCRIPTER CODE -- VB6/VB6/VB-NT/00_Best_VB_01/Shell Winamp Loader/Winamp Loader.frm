VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Winamp Loader"
   ClientHeight    =   5928
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   5268
   Icon            =   "Winamp Loader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5928
   ScaleWidth      =   5268
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   432
      Left            =   195
      TabIndex        =   20
      Top             =   4890
      Width           =   4965
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Load Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2355
      TabIndex        =   18
      Top             =   90
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   2685
      TabIndex        =   8
      Top             =   1005
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Lbl3 
      Alignment       =   2  'Center
      Caption         =   "*-- - Last Winamp Loads and Playing - --*"
      Height          =   225
      Left            =   1050
      TabIndex        =   21
      Top             =   4395
      Width           =   3210
   End
   Begin VB.Label Lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Winamp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1935
      TabIndex        =   19
      Top             =   540
      Width           =   2115
   End
   Begin VB.Label TitleLbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "*** -- Select One -- ***"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   17
      Top             =   75
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   360
      TabIndex        =   16
      Top             =   2505
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   360
      TabIndex        =   15
      Top             =   2715
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   360
      TabIndex        =   14
      Top             =   2925
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   360
      TabIndex        =   13
      Top             =   3135
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   360
      TabIndex        =   12
      Top             =   3345
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   360
      TabIndex        =   11
      Top             =   3555
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   10
      Top             =   3765
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   9
      Top             =   3975
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   375
      TabIndex        =   7
      Top             =   2295
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   375
      TabIndex        =   6
      Top             =   2085
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   375
      TabIndex        =   5
      Top             =   1875
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   375
      TabIndex        =   4
      Top             =   1665
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   375
      TabIndex        =   3
      Top             =   1455
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   375
      TabIndex        =   2
      Top             =   1245
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   1035
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   0
      Top             =   825
      Width           =   1545
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_EXE_PATH_CLIP 
      Caption         =   "EXE PATH CLIP"
   End
   Begin VB.Menu MNU_EXE_FOLDER 
      Caption         =   "EXE FOLDER"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Dim EXEPATH


Private Sub Form_Load()
If App.PrevInstance Then End

'gcw2 = GetForegroundWindow
Winamp22Handle2 = FindWindow("Winamp v1.x", vbNullString)
Winamp22Handle3 = FindWindow("Winamp PE", "Winamp Playlist Editor")
Winamp22Handle4 = FindWindow("Winamp Video", vbNullString)
If Winamp22Handle2 > 0 Then w2h = Winamp22Handle2
If Winamp22Handle3 > 0 Then w2h = Winamp22Handle3
If Winamp22Handle4 > 0 Then w2h = Winamp22Handle4
         
'If w2h <> gcw2 Then Exit Sub
If gcw$ = "The KMPlayer" Then
    MsgBox "The KMPlayer -- Loaded"
    End
End If

gcw$ = GetWindowTitle(w2h)
Debug.Print gcw$
         
         
Dim Tex$
Dim Pid As Long
         
Tex$ = GetFileFromHwnd(w2h)
'If InStr(Tex$, "Video X") Then Call MuteVol
        
Lbl2.Caption = Tex$
If Tex$ <> "" Then
gg$ = Mid$(Tex$, InStrRev(Tex$, "\", Len(Tex$) - 11) + 1)
gg$ = Mid$(gg$, 1, InStrRev(gg$, "\") - 1)
Else
gg$ = "No Winamp Loaded"
Call LOAD_WINAMP_AT_START
'Me.WindowState = vbMinimized
End If
Lbl2.Caption = gg$


Form1.Top = 480
Form1.Left = 0

'Dir1.Path = "C:\Program Files\00 WinAmp's"
Dir1.Path = "C:\Program Files ROOT\00 WinAmp's"

Lbl2.Top = 0
Lbl2.Left = 0
Lbl2.Width = Form1.Width


x = Lbl2.Height + 15

TitleLbl.Top = x
TitleLbl.Left = 0
TitleLbl.Width = Form1.Width / 2
Check1.Top = x
Check1.Left = (Form1.Width / 2) + 10
Check1.Width = Form1.Width / 2

x = TitleLbl.Top + TitleLbl.Height + 15



w = 0

rd = 0


rg = 0
For Each Control In Controls
If InStr(Control.Name, "Lab") Then rg = rg + 1
Next

For R = 0 To rg - 1
Label1(R).Top = x
Label1(R).Left = w
Label1(R).Height = 360
Label1(R).Width = Form1.Width
If Dir1.List(rd) = "" Then
    Label1(R).Visible = False
Else
    tt = Mid$(Dir1.List(rd), InStrRev(Dir1.List(rd), "\") + 1)
    If tt = "Winamp Gold 07 Video X" Then tt = "Winamp Gold 07 Video One"
    Label1(R).Caption = Format$(R, "00") + ". " + tt
    x = x + Label1(R).Height + 15
    fheight = Label1(R).Top + Label1(R).Height + 420
    rd = rd + 1
End If
Next

Form1.Height = fheight
Lbl3.Top = fheight - 410
Lbl3.Width = Form1.Width
Lbl3.Left = 0

List1.Top = fheight - 410 + Lbl3.Height
List1.Height = 1500
List1.Width = Form1.Width - 120
List1.Left = 0

Form1.Height = List1.Top + List1.Height + 410

If Dir$(App.Path + "\Winloads.txt") = "" Then Exit Sub
Open App.Path + "\Winloads.txt" For Input As #1
On Local Error Resume Next
Seek #1, LOF(1) - 4000
Do
Line Input #1, ed$
List1.AddItem ed$
Loop Until EOF(1)
Close #1

List1.AddItem Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Lbl2.Caption
Open App.Path + "\Winloads.txt" For Append As #1
Print #1, Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Lbl2.Caption
Close #1

List1.ListIndex = List1.ListCount - 1
End Sub

Sub LOAD_WINAMP_AT_START()
'Exit Sub

'Open App.Path + "\Winloads.txt" For Append As #1
'Print #1, Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Label1(Index)
'Close #1
'ChDrive Dir1.List(Index)
'ChDir Dir1.List(Index)


'Shell "C:\Program Files\00 WinAmp's\Winamp Gold 05 My Music 01\winamp.exe", vbMinimizedNoFocus
'Exit Sub

For R = 1 To Dir1.ListCount - 1
    xw = 0
    If InStr(Dir1.List(R), "MIDI") > 0 Then xw = 1
    If InStr(Dir1.List(R), "Music 01") > 0 Then xw = 1
    If xw = 1 Then
        Shell Dir1.List(R) + "\winamp.exe", vbMinimizedNoFocus
    End If
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

If Me.WindowState <> 1 Then
    Me.WindowState = 1
    'test may have to put back form need reseting
    'Unload FrmJoy
    'Cancel = True
    Exit Sub
End If
End Sub

Private Sub Label1_Click(Index As Integer)
'Form1.Visible = False

If EXEPATH = True Then
Clipboard.Clear
Clipboard.SetText Dir1.List(Index) + "\winamp.exe"
Exit Sub
End
End If

If EXEFOLDER = True Then
Shell "EXPLORER " + Dir1.List(Index) + "\winamp.exe", vbNormalFocus
Exit Sub
End
End If


If Check1.Value = vbUnchecked Then
Open App.Path + "\Winloads.txt" For Append As #1
Print #1, Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Label1(Index)
Close #1
ChDrive Dir1.List(Index)
ChDir Dir1.List(Index)
Shell Dir1.List(Index) + "\winamp.exe", vbNormalFocus
Me.WindowState = 1
Exit Sub

End

Else
Shell "notepad " + Dir1.List(Index) + "\Current Play To Text File Append.txt", vbNormalFocus
Exit Sub

End

End If




End Sub

Private Sub MNU_EXE_PATH_CLIP_Click()
EXEPATH = True
End Sub

Private Sub MNU_EXIT_Click()
    End
End Sub


Private Sub MNU_VB_ME_Click()
VBPATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VBPATH) = "" Then
    VBPATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
End If

Shell VBPATH + " """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    
    Beep
    CODER_EXE_FILE_NAME_1 = App.Path + "\" + Dir(App.Path + "\*.VBP")
    Shell "EXPLORER /SELECT, """ + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    
    'End

End Sub
