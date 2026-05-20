VERSION 5.00
Begin VB.Form frmAskDateTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert date & time"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKo 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1860
      TabIndex        =   17
      Top             =   1620
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Accept"
      Height          =   435
      Left            =   210
      TabIndex        =   16
      Top             =   1620
      Width           =   1035
   End
   Begin VB.TextBox txtMillisecond 
      Height          =   315
      Left            =   1965
      MaxLength       =   3
      TabIndex        =   11
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox txtSecond 
      Height          =   315
      Left            =   1530
      MaxLength       =   2
      TabIndex        =   10
      Top             =   840
      Width           =   315
   End
   Begin VB.TextBox txtMinute 
      Height          =   315
      Left            =   1095
      MaxLength       =   2
      TabIndex        =   9
      Top             =   840
      Width           =   315
   End
   Begin VB.TextBox txtHour 
      Height          =   315
      Left            =   660
      MaxLength       =   2
      TabIndex        =   8
      Top             =   840
      Width           =   315
   End
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   2460
      MaxLength       =   4
      TabIndex        =   6
      Top             =   450
      Width           =   495
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      ItemData        =   "AskDateTime.frx":0000
      Left            =   1065
      List            =   "AskDateTime.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   450
      Width           =   1305
   End
   Begin VB.TextBox txtDay 
      Height          =   315
      Left            =   660
      MaxLength       =   2
      TabIndex        =   4
      Top             =   450
      Width           =   315
   End
   Begin VB.Label lblMilliseconds 
      AutoSize        =   -1  'True
      Caption         =   "ms"
      Height          =   195
      Left            =   1965
      TabIndex        =   15
      Top             =   1185
      Width           =   195
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "s"
      Height          =   195
      Left            =   1530
      TabIndex        =   14
      Top             =   1185
      Width           =   75
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   195
      Left            =   1095
      TabIndex        =   13
      Top             =   1185
      Width           =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "h"
      Height          =   195
      Left            =   660
      TabIndex        =   12
      Top             =   1185
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Year"
      Height          =   195
      Left            =   2460
      TabIndex        =   2
      Top             =   210
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Month"
      Height          =   195
      Left            =   1065
      TabIndex        =   1
      Top             =   210
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Day"
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   210
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Time"
      Height          =   195
      Left            =   195
      TabIndex        =   7
      Top             =   930
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Left            =   195
      TabIndex        =   3
      Top             =   510
      Width           =   345
   End
End
Attribute VB_Name = "frmAskDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public inUseMilliseconds As Boolean
Public inDay As Integer, inMonth As Integer, inYear As Integer
Public inHour As Integer, inMinute As Integer, inSecond As Integer, inMillisecond As Integer

Public outOk As Boolean
Public outDay As Integer, outMonth As Integer, outYear As Integer
Public outHour As Integer, outMinute As Integer, outSecond As Integer, outMillisecond As Integer

Private Sub Form_Load()
   Me.lblMilliseconds.Visible = inUseMilliseconds
   Me.txtMillisecond.Visible = inUseMilliseconds
   
   Me.txtDay = Format(inDay, "00")
   Me.cmbMonth.ListIndex = inMonth - 1
   Me.txtYear = inYear
   
   Me.txtHour = Format(inHour, "00")
   Me.txtMinute = Format(inMinute, "00")
   Me.txtSecond = Format(inSecond, "00")
   Me.txtMillisecond = Format(inMillisecond, "000")
End Sub

Private Sub cmdOk_Click()
   Me.outOk = True
   Me.Hide
End Sub

Private Sub cmdKo_Click()
   Me.outOk = False
   Me.Hide
End Sub

Private Sub txtDay_LostFocus()
   SetTextNum Me.txtDay, 1, 31, 2
End Sub
Private Sub txtYear_LostFocus()
   SetTextNum Me.txtYear, 0, 9999
End Sub
Private Sub txtHour_LostFocus()
   SetTextNum Me.txtHour, 0, 23, 2
End Sub
Private Sub txtMinute_LostFocus()
   SetTextNum Me.txtMinute, 0, 59, 2
End Sub
Private Sub txtSecond_LostFocus()
   SetTextNum Me.txtSecond, 0, 59, 2
End Sub
Private Sub txtMillisecond_LostFocus()
   SetTextNum Me.txtMillisecond, 0, 999, 3
End Sub
Private Sub SetTextNum(t As TextBox, MinVal As Long, MaxVal As Long, Optional NumDigits As Integer = 0)
   Dim v As Long
   
   v = 0
   On Error Resume Next: v = CLng(t.Text): Err.Clear: On Error GoTo 0
   If v < MinVal Then v = MinVal
   If v > MaxVal Then v = MaxVal
   
   If NumDigits > 0 Then
      t.Text = Format(v, String(NumDigits, "0"))
   Else
      t.Text = v
   End If
End Sub
