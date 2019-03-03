VERSION 5.00
Begin VB.Form AlarmForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm Clock"
   ClientHeight    =   3135
   ClientLeft      =   5010
   ClientTop       =   4245
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ALARM.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "To Quarter Past"
      Height          =   450
      Left            =   2430
      TabIndex        =   10
      Top             =   1440
      Width           =   1530
   End
   Begin VB.Timer Timer2 
      Left            =   765
      Top             =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "40mins Laundry"
      Height          =   450
      Left            =   2430
      TabIndex        =   9
      Top             =   480
      Width           =   1530
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   450
      Left            =   3180
      TabIndex        =   8
      Top             =   0
      Width           =   780
   End
   Begin VB.CommandButton Command2 
      Caption         =   "30mins"
      Height          =   450
      Left            =   2430
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hour"
      Height          =   450
      Left            =   2430
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   765
      Top             =   165
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -15
      TabIndex        =   7
      Top             =   1935
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -15
      TabIndex        =   6
      Top             =   390
      Width           =   2415
   End
   Begin VB.Label Lbl3 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -15
      TabIndex        =   5
      Top             =   1545
      Width           =   2415
   End
   Begin VB.Label Lbl2 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -15
      TabIndex        =   4
      Top             =   780
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -15
      TabIndex        =   1
      Top             =   1170
      Width           =   2415
   End
   Begin VB.Label lblTime 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "AlarmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim AlarmTime
Dim AlarmSounded
Dim StartTime
Const conMinimized = 1


Private Sub Command1_Click()
    
    'AlarmTime = Now + TimeSerial(1, 0, 0) ' so store it as a date/time value in AlarmTime.
    'hour
    StartTime = Now
    AlarmTime = Now + TimeSerial(1, 0, 0)

End Sub

Private Sub Command2_Click()
    
    'AlarmTime = Now + TimeSerial(1, 0, 0) ' so store it as a date/time value in AlarmTime.
    'hour
    StartTime = Now
    AlarmTime = Now + TimeSerial(0, 30, 0)

End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
End Sub

Private Sub Command4_Click()

    StartTime = Now
    AlarmTime = Now + TimeSerial(0, 40, 0)

End Sub

Private Sub Command5_Click()
    StartTime = Now
    
    AlarmTime = Int(Now) + TimeSerial(Hour(Now), 15, 0)
    If AlarmTime < Now Then AlarmTime = Int(Now) + TimeSerial(Hour(Now) + 1, 15, 0)
    'AlarmTime = Now + TimeSerial(0, 0, 2)

End Sub

Private Sub Form_Click()
    AlarmTime = InputBox("Enter alarm time", "VB Alarm", AlarmTime)
    If AlarmTime = "" Then Exit Sub
    If Not IsDate(AlarmTime) Then
        MsgBox "The time you entered was not valid."
    Else                                    ' String returned from InputBox is a valid time,
        AlarmTime = CDate(AlarmTime)        ' so store it as a date/time value in AlarmTime.
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    For Each Control In Controls
    If Control.Width + Control.Left > xx Then xx = Control.Width + Control.Left
    If Control.Height + Control.Top > yy Then yy = Control.Height + Control.Top
    Next
    On Error GoTo 0
    AlarmForm.Height = yy + 450
    AlarmForm.Width = xx + 110
    
    AlarmTime = ""
    'times
    StartTime = Int(Now) + TimeSerial(18, 5, 0)  ' so store it as a date/time value in AlarmTime.
    AlarmTime = Int(Now) + TimeSerial(19, 5, 0) ' so store it as a date/time value in AlarmTime.
            
    'hour
    StartTime = Now
    AlarmTime = Now + TimeSerial(1, 0, 0)
    
    Label2.Caption = TimeValue(StartTime)
    rf = DateDiff("s", StartTime, AlarmTime)
    mm = Int(rf / 60)
    ss = rf Mod 60
    Label3.Caption = Trim(Str(mm)) + ":" + Trim(Str(ss))
    
    End Sub

Private Sub Form_Resize()
    If WindowState = conMinimized Then      ' If form is minimized, display the time in a caption.
        SetCaptionTime
    Else
        Caption = "Alarm Clock"
    End If
End Sub

Private Sub SetCaptionTime()
    Caption = Format(Time, "Medium Time")   ' Display time using medium time format.
End Sub

Private Sub Timer1_Timer()

Static AlarmSounded As Integer
    
    If lblTime.Caption <> CStr(Time) Then
        ' It's now a different second than the one displayed.
        rf = DateDiff("s", Now, AlarmTime)
        
        If rf < 0 And Not AlarmSounded Then
            Beep
            MsgBox "Alarm at " & Time
            AlarmSounded = True
        ElseIf rf > 0 Then
            AlarmSounded = False
        End If
        If WindowState = conMinimized Then
            ' If minimized, then update the form's Caption every minute.
            If Minute(CDate(Caption)) <> Minute(Time) Then SetCaptionTime
        Else
            ' Otherwise, update the label Caption in the form every second.
            lblTime.Caption = Time
            Lbl2.Caption = TimeValue(AlarmTime)
            If AlarmTime <> "" Then
                rf = DateDiff("s", Now, AlarmTime)
                mm = Int(rf / 60)
                ss = rf Mod 60
                Label1.Caption = Trim(Str(mm)) + ":" + Trim(Str(ss))
            
                rf = DateDiff("s", StartTime, Now)
                mm = Int(rf / 60)
                ss = rf Mod 60
                Lbl3.Caption = Trim(Str(mm)) + ":" + Trim(Str(ss))
            
            
            End If
        End If
    End If
End Sub

