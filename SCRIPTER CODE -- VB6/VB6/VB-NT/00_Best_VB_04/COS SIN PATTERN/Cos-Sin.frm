VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   3495
   ClientTop       =   1920
   ClientWidth     =   7710
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Ef 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   333
      ScaleMode       =   0  'User
      ScaleWidth      =   333
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   11
      Left            =   1920
      Top             =   3960
   End
   Begin VB.Frame Frame3 
      Caption         =   "Prop"
      Height          =   4215
      Left            =   5160
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      Begin MSComctlLib.Slider Slider5 
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Exit"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Stop"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Full Screen"
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   2760
         Width           =   615
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   1
         Max             =   1500
         SelStart        =   1
         TickFrequency   =   100
         Value           =   1499
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   1
         Max             =   1500
         SelStart        =   1
         TickFrequency   =   100
         Value           =   1499
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   975
         Left            =   1320
         TabIndex        =   13
         Top             =   1200
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   1
         SelStart        =   9
         Value           =   8
      End
      Begin MSComctlLib.Slider Slider4 
         Height          =   975
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   1
         SelStart        =   9
         Value           =   8
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "DrawWidth"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "0,8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "0,8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "1499"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "1499"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MidX As Long, MidY As Long
Const PI = 3.14159
Dim fa As Boolean
Dim s, X, Y, j, OldX, OldY, f, t, r As Long



Private Sub Combo1_Change()

End Sub

Private Sub Command2_Click()
Start
End Sub

Private Sub Command3_Click()
Stopp
End Sub

Private Sub Command4_Click()
Ef.Cls
End Sub

Private Sub Command5_Click()
FullScreen
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Ef_Click()
Exit_FullScreen
End Sub

Private Sub Form_Load()
    Timer1.Interval = 1
    s = Ef.ScaleWidth / 2
    j = 1
    MidX = Ef.ScaleWidth \ 2
    MidY = Ef.ScaleHeight \ 2
    OldX = s * Cos(PI * Label2 * (j)) + MidX
    OldY = s * Sin(PI * Label3 * (j)) + MidY
    fa = True
End Sub

Private Sub Slider1_Change()
Label3 = Slider1.Value
End Sub

Private Sub Slider2_Change()
Label2 = Slider2.Value
End Sub

Private Sub Slider3_Change()
Label4 = Slider3.Value / 10
End Sub

Private Sub Slider4_Change()
Label5 = Slider4.Value / 10
End Sub

Private Sub Slider5_Click()
Ef.DrawWidth = Slider5.Value
End Sub

Private Sub Timer1_Timer()
    j = j + 1
    Label1 = j
    If f = 255 Then
    fa = False
    r = Int(Rnd + 255) + 1
    ElseIf f = 0 Then
    fa = True
    r = Int(Rnd + 255) + 1
    End If
    If fa = True Then
    f = f + 1
    ElseIf fa = False Then
    f = f - 1
    End If
    X = s * Cos((PI * Label2) * Label4 * (j)) + MidX
    Y = s * Sin((PI * Label3) * Label5 * (j)) + MidY
    Ef.Line (X, Y)-(OldX, OldY), RGB(f, f, f)
    Ef.Line (Y, X)-(OldY, OldX), RGB(f, f, f)

    OldX = X
    OldY = Y
End Sub
Sub Start()
Timer1.Enabled = True
Command3.Enabled = True
Slider1.Enabled = False
Slider2.Enabled = False
Command2.Enabled = False
OldX = s * Cos(PI * Label2 * Label4 * (j)) + MidX
OldY = s * Sin(PI * Label3 * Label5 * (j)) + MidY
End Sub

Sub Stopp()
Timer1.Enabled = False
Command2.Enabled = True
Slider1.Enabled = True
Slider2.Enabled = True
Command3.Enabled = False
Ef.Cls
End Sub
Sub FullScreen()
Ef.Cls
Timer1.Enabled = False
Form1.WindowState = 2
Form1.BorderStyle = 0
Ef.Height = Form1.ScaleHeight
Ef.Width = Form1.ScaleWidth
Ef.ScaleHeight = Form1.ScaleHeight
Ef.ScaleWidth = Form1.ScaleHeight
    s = Ef.ScaleWidth / 2

    MidX = Ef.ScaleWidth \ 2
    MidY = Ef.ScaleHeight \ 2
    OldX = s * Cos(PI * Label2 * Label4 * (j)) + MidX
    OldY = s * Sin(PI * Label3 * Label5 * (j)) + MidY
Start
End Sub
Sub Exit_FullScreen()
Stopp
Timer1.Enabled = False
Form1.WindowState = 0
Form1.BorderStyle = 2

Ef.Height = 281
Ef.Width = 337
Ef.ScaleHeight = 333
Ef.ScaleWidth = 333
Form1.Height = 4200
Form1.Width = 7710
    s = Ef.ScaleWidth / 2
    MidX = Ef.ScaleWidth \ 2
    MidY = Ef.ScaleHeight \ 2
    OldX = s * Cos(PI * Label2 * Label4 * (j)) + MidX
    OldY = s * Sin(PI * Label3 * Label5 * (j)) + MidY

End Sub
