VERSION 5.00
Begin VB.Form Pulse 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   Caption         =   "Blood Pressure & Pulse Rate History"
   ClientHeight    =   7530
   ClientLeft      =   570
   ClientTop       =   600
   ClientWidth     =   9900
   Icon            =   "Pulse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   0
      TabIndex        =   28
      Top             =   6435
      Visible         =   0   'False
      Width           =   9510
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1980
      TabIndex        =   10
      Top             =   1530
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   990
      TabIndex        =   9
      Top             =   1530
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   1530
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Edit With Notepad"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   3510
      Width           =   2520
   End
   Begin VB.Label Label25 
      BackColor       =   &H8000000E&
      Caption         =   "Edit With Notepad"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   34
      Top             =   3015
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5490
      TabIndex        =   33
      Top             =   6030
      Width           =   90
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "List Box"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8055
      TabIndex        =   32
      Top             =   450
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6390
      TabIndex        =   31
      Top             =   450
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Norm"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8010
      TabIndex        =   30
      Top             =   45
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Aver"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8865
      TabIndex        =   29
      Top             =   45
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Aver"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7200
      TabIndex        =   27
      Top             =   45
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Average Taken From Last 5 Readings"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   45
      TabIndex        =   26
      Top             =   45
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Norm"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6345
      TabIndex        =   25
      Top             =   45
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7965
      TabIndex        =   24
      Top             =   6030
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   675
      TabIndex        =   23
      Top             =   6030
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   3825
      TabIndex        =   22
      Top             =   1215
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   5
      Left            =   3825
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   3825
      TabIndex        =   20
      Top             =   2385
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   3825
      TabIndex        =   19
      Top             =   2970
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   3825
      TabIndex        =   18
      Top             =   4005
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   3825
      TabIndex        =   17
      Top             =   4590
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   3825
      TabIndex        =   16
      Top             =   5175
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Dia"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   4410
      TabIndex        =   15
      Top             =   45
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Sys"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   3645
      TabIndex        =   14
      Top             =   45
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Pulse"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   5175
      TabIndex        =   13
      Top             =   45
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "X Time"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4185
      TabIndex        =   12
      Top             =   6030
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   180
      TabIndex        =   11
      Top             =   2835
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "Enter Readings"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   2025
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Skip This Part"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Pulse"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1980
      TabIndex        =   5
      Top             =   1035
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Dia"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   990
      TabIndex        =   4
      Top             =   1035
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "Sys"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   1035
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Date Time"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Mum"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1035
      TabIndex        =   1
      Top             =   45
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Matt"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   915
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Displaying Last 10 Results"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   45
      TabIndex        =   36
      Top             =   45
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Menu Info 
      Caption         =   "&Info"
   End
End
Attribute VB_Name = "Pulse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefDbl A-Z
Public commvar$
Public listboxvar1 As Integer
Public xjag As Integer
Public lastfew As Integer
Public display2 As Integer
Public display3 As Integer
Public xscale As Integer
Public yscale As Integer
Public xs As Double
Public ys As Double
Public fs As Double
Dim Ttpulse5 As Integer 'Double






Private Sub Form_Activate()

'Call grid

Call Form_Resize

End Sub

Private Sub Form_Load()
Pulse.Top = 1400
Pulse.Left = 480
Pulse.Width = 9100
'Pulse.ScaleWidth = 400
If App.PrevInstance = True Then Pulse.Left = 9120 + 480


xjag = 15

commvar$ = Command$


Label26.Caption = "Software matthew.lancaster@btopenworld.com Version:" + Str$(App.Major) + "." + Str$(App.Minor) + "." + Str$(App.Revision)


Label18.Caption = "Average Taken From Last " + vbCr + Trim$(Str$(xjag)) + " Readings"

End Sub








Private Sub proc1()

Label1.Visible = False
Label2.Visible = False
Label3.Visible = True

Text1.Visible = True
Text2.Visible = True
Text3.Visible = True

Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True

Label25.Visible = True


da$ = Mid$(Date$, 4, 2) + "-" + Mid$(Date$, 1, 2) + "-" + Mid$(Date$, 7, 4)
ti$ = Mid$(Time$, 1, 5)

If commvar$ = "1" Then

Label3.Caption = da$ + "-" + ti$

End If

If commvar$ = "4" Then
'Open "D:\VB6\vb-nt\pulse\pulsem.txt" For Append As #1
'Print #1, da$; "-"; ti$; "-     -      -"
'Close #1
'Shell "Notepad D:\VB6\vb-nt\pulse\pulsem.txt"
Label3.Caption = da$ + "-" + ti$

End If




End Sub





Private Sub proc2()

Label18.Visible = False


Label12.Visible = True
Label13.Visible = True
Label6.Visible = True








Label15.Visible = True
Label16.Visible = True
Label4.Visible = True
Label5.Visible = True


Label17.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True



Label22.Visible = True
Label23.Visible = True


Label25.Visible = False
Label26.Visible = False

wq = 0


If commvar$ = "3" Or commvar$ = "2" Or commvar$ = "" Then
If commvar$ = "2" Or commvar$ = "" Then fe$ = App.Path + "\pulse.txt": fe2$ = App.Path + "\Pulseavr.txt": fe3$ = "D:\VB6\vb-nt\pulse\Pulseav2.txt"
If commvar$ = "3" Then fe$ = App.Path + "\pulsem.txt": fe2$ = App.Path + "\Pulsemvr.txt": fe3$ = "D:\VB6\vb-nt\pulse\Pulsemv2.txt"
Dim op(3)
Dim qp(3)
'Screen 12


Call grid


Open fe$ For Input As #1
Do
Line Input #1, a$
count2 = count2 + 1
ads = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
If ads > DateSerial(2004, 1, 1) And wq = 0 Then wq = count2
Loop Until EOF(1)
Close #1


Label27.Caption = "Displaying This Years Last" + vbCr + Str$((count2 - wq) + 1) + " Readings"

If lastfew = 1 Then Label27.Visible = True Else Label27.Visible = False



Open fe$ For Input As #1
Line Input #1, a$
If lastfew = 1 Then
For rtyy = 1 To wq - 2
Line Input #1, a$
Next
End If


B = 0
Line Input #1, a$
ads = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
ads = ads + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
B = B + 1
If Not EOF(1) Then
Do
Line Input #1, a$
B = B + 1
Loop Until EOF(1)
c = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
c = c + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
End If
 
Label15.Caption = Format$(ads, "DD/MM/YY")
Label16.Caption = Format$(c, "DD/MM/YY")


Close #1


If c = 0 Then
Ttpulse5 = Val(Mid$(a$, 36, 4))
dia = Val(Mid$(a$, 27, 4))
sys = Val(Mid$(a$, 18, 4))


'ERROR CODE COMMA
Line (80 / ys, 400 / ys)-(600 / ys, (400 - (Ttpulse5 * 2)) / ys), RGB(255, 0, 0)
Line (80 / ys, 400 / ys)-(600 / ys, (400 - (dia * 2)) / ys), RGB(0, 255, 0)
Line (80 / ys, 400 / ys)-(600 / ys, (400 - (sys * 2)) / ys), RGB(0, 0, 255)
End If
If c > 0 Then
qp(1) = 400
qp(2) = 400
qp(3) = 400
Open fe$ For Input As #1

If lastfew = 1 Then
For rtyy = 1 To wq - 2
Line Input #1, a$
Next
End If


Line Input #1, a$
l = (c - ads) / 520
q = 80
po = 1
For r = 1 To B
Line Input #1, a$
e = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
e = e + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
t = 0
For o = 0 To (c - ads) Step l
t = t + 1
If (c - e) < o Then Exit For
Next

'PRINT t


Ttpulse5 = Val(Mid$(a$, 36, 4))
dia = Val(Mid$(a$, 27, 4))
sys = Val(Mid$(a$, 18, 4))

p = 80 + (520 - t)

op(1) = 400 - (Ttpulse5 * 2)
op(2) = 400 - (dia * 2)
op(3) = 400 - (sys * 2)
If po = 0 Then
Line (q / xs, qp(1) / ys)-(p / xs, op(1) / ys), RGB(255, 0, 0)
Line (q / xs, qp(2) / ys)-(p / xs, op(2) / ys), RGB(0, 255, 0)
Line (q / xs, qp(3) / ys)-(p / xs, op(3) / ys), RGB(0, 0, 255)
End If
If po = 1 Then
po = 0
Line (p / xs, op(1) / ys)-(p / xs, op(1) / ys), RGB(255, 0, 0)
Line (p / xs, op(2) / ys)-(p / xs, op(2) / ys), RGB(0, 255, 0)
Line (p / xs, op(3) / ys)-(p / xs, op(3) / ys), RGB(0, 0, 255)
End If
qp(1) = op(1)
qp(2) = op(2)
qp(3) = op(3)
q = p

Next
Close #1
End If

a$ = " "
'PAINT (3, 3), 7


End If

Label17.Visible = True

'normal
listboxvar1 = 0
Call proc4




End Sub

Private Sub proc3()




Label18.Visible = True










If commvar$ = "3" Or commvar$ = "2" Or commvar$ = "" Then
If commvar$ = "2" Or commvar$ = "" Then fe$ = App.Path + "\pulse.txt": fe2$ = App.Path + "\Pulseavr.txt": fe3$ = "D:\VB6\vb-nt\pulse\Pulseav2.txt"
If commvar$ = "3" Then fe$ = App.Path + "\pulsem.txt": fe2$ = App.Path + "\Pulsemvr.txt": fe3$ = "D:\VB6\vb-nt\pulse\Pulsemv2.txt"
Dim op(3)
Dim qp(3)


Dim averrun1(600) As Double
Dim averrun2(600) As Double
Dim averrun3(600) As Double

aver2 = 0
averp1 = 0
averp2 = 0
averd1 = 0
averd2 = 0
avers1 = 0
avers2 = 0
avera1 = 0
avera2 = 0





Call grid




Open fe$ For Input As #1
Line Input #1, a$
B = 0
Line Input #1, a$
Adq = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
Adq = Adq + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
B = B + 1
If Not EOF(1) Then
Do
Line Input #1, a$
B = B + 1
Loop Until EOF(1)
c = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
c = c + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
End If
Close #1

Label15.Caption = Format$(Adq, "DD/MM/YY")
Label16.Caption = Format$(c, "DD/MM/YY")













If c = 0 Then
Ttpulse5 = Val(Mid$(a$, 36, 4))
dia = Val(Mid$(a$, 27, 4))
sys = Val(Mid$(a$, 18, 4))
Line (80 / xs, 400 / ys)-(600 / xs, (400 - (Ttpulse5 * 2)) / ys), RGB(255, 0, 0)
Line (80 / xs, 400 / ys)-(600 / xs, (400 - (dia * 2)) / ys), RGB(0, 255, 0)
Line (80 / xs, 400 / ys)-(600 / xs, (400 - (sys * 2)) / ys), RGB(0, 0, 255)
End If
If c > 0 Then
'DIM op(3)
'DIM qp(3)
qp(1) = 400
qp(2) = 400
qp(3) = 400
Open fe2$ For Output As #2
Print #2, "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"
Open fe$ For Input As #1
Line Input #1, a$
l = (c - Adq) / 520
q = 80
po = 1
For r = 1 To B
Line Input #1, a$
e = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
e = e + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
t = 0
For o = 0 To (c - Adq) Step l
t = t + 1
If (c - e) < o Then Exit For
Next

Ttpulse5 = Val(Mid$(a$, 36, 4))
dia = Val(Mid$(a$, 27, 4))
sys = Val(Mid$(a$, 18, 4))

aver2 = aver2 + 1
averp1 = averp1 + Ttpulse5
averp2 = averp1 / aver2
averd1 = averd1 + dia
averd2 = averd1 / aver2
avers1 = avers1 + sys
avers2 = avers1 / aver2
avera1 = avera1 + ((sys + dia + Ttpulse5) / 3)
avera2 = avera1 / aver2

averrun1(aver2) = Ttpulse5
averrun2(aver2) = dia
averrun3(aver2) = sys

jelly1 = xjag
If aver2 < xjag Then jelly1 = aver2
jelly2 = aver2 - xjag


If jelly2 <= 0 Then jelly2 = 1

avertp1 = 0
avertd1 = 0
averts1 = 0
For rty2 = jelly2 To aver2
avertp1 = avertp1 + averrun1(rty2)
avertp2 = avertp1 / jelly1
avertd1 = avertd1 + averrun2(rty2)
avertd2 = avertd1 / jelly1
averts1 = averts1 + averrun3(rty2)
averts2 = averts1 / jelly1

Next




Print #2, Mid$(a$, 1, 17); Format$(avers2, "#00.000000"); "-"; Format$(averd2, "#00.000000"); "-"; Format$(averp2, "#00.000000"); "-"; Format$(avera2, "#00.000000"); "-"; Mid$(a$, 50)

p = 80 + (520 - t)

op(1) = 400 - (avertp2 * 2)
op(2) = 400 - (avertd2 * 2)
op(3) = 400 - (averts2 * 2)

If po = 0 Then
Line (q / xs, qp(1) / ys)-(p / xs, op(1) / ys), RGB(255, 0, 0)
Line (q / xs, qp(2) / ys)-(p / xs, op(2) / ys), RGB(0, 255, 0)
Line (q / xs, qp(3) / ys)-(p / xs, op(3) / ys), RGB(0, 0, 255)
End If
If po = 1 Then
po = 0
Line (p / xs, op(1) / ys)-(p / xs, op(1) / ys), RGB(255, 0, 0)
Line (p / xs, op(2) / ys)-(p / xs, op(2) / ys), RGB(0, 255, 0)
Line (p / xs, op(3) / ys)-(p / xs, op(3) / ys), RGB(0, 0, 255)
End If
qp(1) = op(1)
qp(2) = op(2)
qp(3) = op(3)
q = p

Next
Close #1
Close #2
End If

'Label19.Visible = True
listboxvar1 = 1

Call proc4

End If


End Sub


Private Sub proc4()

If display3 = 0 Then

For r = List1.ListCount - 1 To 0 Step -1
List1.RemoveItem (r)
Next

If commvar$ = "3" Or commvar$ = "2" Or commvar$ = "" Then
If commvar$ = "2" Or commvar$ = "" Then fe$ = App.Path + "\pulse.txt": fe2$ = App.Path + "\Pulseavr.txt": fe3$ = App.Path + "\Pulseav2.txt"
If commvar$ = "3" Then fe$ = App.Path + "\pulsem.txt": fe2$ = App.Path + "\Pulsemvr.txt": fe3$ = "D:\VB6\vb-nt\pulse\Pulsemv2.txt"



Open fe2$ For Input As #1
Do
Line Input #1, a$
'Print A$
Loop Until EOF(1)
Close #1

'PRINT "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"

Open fe2$ For Input As #1
lines = 0
Do
Line Input #1, a$
lines = lines + 1
Loop Until EOF(1)
Close #1
'ytr = lines - 51''lines to produce 60 in pulseav2.txt
ytr = lines
Open fe2$ For Input As #1
lines = 0
Do
Line Input #1, a$
lines = lines + 1
Loop Until lines = ytr Or EOF(1)
If EOF(1) Then Close #1: Open fe$ For Input As #1
Open fe3$ For Output As #2
Print #2, "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"

Do
Line Input #1, a$
Print #2, a$
Loop Until EOF(1)
Close #1, #2

'Shell "c:\utils\u " + fe2$ + " /#:" + LTrim$(Str$(ytr - 21))


''SHELL "t " + fe3$

fe33 = FreeFile

fe22$ = fe$
If listboxvar1 = 1 Then fe22$ = fe2$
Open fe22$ For Input As fe33
Do
Line Input #fe33, dutch$
List1.AddItem dutch$
Loop Until EOF(fe33)
Close fe33
List1.ListIndex = List1.ListCount - 1


End If
Label24.Caption = Trim$(Str$(List1.ListCount - 1)) + " Readings."

End If

List1.Visible = True
List1.Height = 94 / ys
List1.Width = 639 / xs
List1.Top = 430 / ys
List1.Left = 1
List1.FontSize = 10 / fs
List1.ListIndex = List1.ListCount - 2
List1.ListIndex = List1.ListCount - 1
List1.Refresh



End Sub


Private Sub Form_Resize()
If Pulse.ScaleHeight > 0 And Pulse.ScaleWidth > 0 Then

xscale = Pulse.ScaleWidth
yscale = Pulse.ScaleHeight
xs = 640 / xscale
ys = 524 / yscale
fs = (640 + 524) / (xscale + yscale)

End If
display3 = 1
If display2 = 1 Then Call proc2
If display2 = 2 Then Call proc3
If display2 = 3 Then Call proc4
If display2 = 4 Then Call proc4
display3 = 0
Pulse.Refresh


End Sub

Private Sub Info_Click()

MsgBox ("Values are usually given in millimetres of mercury (mmHg). Normal ranges for blood pressure in adult humans are:" + vbCr + "Systolic between 90 and 135 mmHg (12 to 18 kPa)" + vbCr + "Diastolic between 50 and 90 mmHg (7 to 12 kPa)")

End Sub

Private Sub Label1_Click()
'matt
commvar$ = "1"
Call proc1

End Sub

Private Sub Label10_Click()
'skip
Label3.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False

Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False

If commvar$ = "1" Then commvar$ = "2"
If commvar$ = "4" Then commvar$ = "3"


display2 = 1
Call proc2

End Sub

Private Sub Label11_Click()
'enter readings

If Val(Text1.Text) > 0 Then


Label3.Visible = False
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False

Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False


za$ = "-        -        -   "
Mid$(za$, 2) = Text1.Text

Mid$(za$, 11) = Text2.Text
Mid$(za$, 20) = Text3.Text


If commvar$ = "1" Then

Open "D:\VB6\vb-nt\pulse\pulse.txt" For Append As #1
Print #1, Label3.Caption + za$
Close #1

commvar$ = "2"

End If

If commvar$ = "4" Then

Open "D:\VB6\vb-nt\pulse\pulsem.txt" For Append As #1
Print #1, Label3.Caption + za$
Close #1

commvar$ = "3"
End If

display2 = 1
Call proc2

End If

End Sub

Private Sub Label17_Click()
'norm graph

display2 = 1

If lastfew = 0 Then lastfew = 1: GoTo jumpkl
If lastfew = 1 Then lastfew = 0: GoTo jumpkl

jumpkl:

If lastfew = 1 Then Label27.Visible = True Else Label27.Visible = False

Call proc2
End Sub

Private Sub Label19_Click()
'average
display2 = 2

Label27.Visible = False
Call proc3

End Sub
Private Sub Label21_Click()

'normal

display2 = 3

listboxvar1 = 0

Call proc4

End Sub
Private Sub Label20_Click()
'Label19.Visible = False

'average

display2 = 4

listboxvar1 = 1

Call proc4

End Sub


Private Sub Label2_Click()
'mum
commvar$ = "4"
Call proc1

End Sub

Private Sub grid()

ForeColor = RGB(0, 0, 0)

DrawWidth = 3 ' Use wider brush.


tugy = 338
For r = 0 To 6
Label14(r).Visible = True
Label14(r).Caption = Trim$(Str$((r + 1) * 25))
Label14(r).Top = tugy / ys

Label14(r).Left = 30 / xs
Label14(r).FontSize = 14 / fs

tugy = tugy - 50
Next

Label4.Top = 160 / ys
Label4.Left = 10 / xs
Label4.FontSize = 18 / (fs / 0.7)

Label5.Top = 402 / ys
Label5.Left = 282 / xs
Label5.FontSize = 18 / (fs / 0.7)

Label24.Top = 402 / ys
Label24.Left = 380 / xs
Label24.FontSize = 14 / (fs / 0.8)
'Label24.Width = 110 / fs

Label27.Top = 0 / ys
Label27.Left = 3 / xs
Label27.FontSize = 12 / (fs / 0.8)

If 238 / fs < 230 Then Label27.Width = 238 / fs

Label18.Top = 0 / ys
Label18.Left = 3 / xs
Label18.FontSize = 12 / (fs / 0.8)
If 238 / fs < 230 Then Label18.Width = 238 / fs

'A = 0
'Color 14

'Pulse.Refresh


ast = 0

'For r = 400 To 50 Step -50
'LOCATE (r / 17) + 1, 4


'Next




Label12.Top = 3 / ys
Label12.Left = 243 / xs
Label12.FontSize = 18 / (fs / 0.8)
'Label12.Width = 46 / xs

Label13.Top = 3 / ys
Label13.Left = 294 / xs
Label13.FontSize = 18 / (fs / 0.8)
'Label13.Width = 46 / xs

Label6.Top = 3 / ys
Label6.Left = 345 / xs
Label6.FontSize = 18 / (fs / 0.8)
'Label6.Width = 70 / xs



Label17.Top = 3 / ys
Label17.Left = 423 / xs
Label17.FontSize = 14 / (fs / 0.8)
'Label17.Width = 52 / xs

Label19.Top = 3 / ys
Label19.Left = 480 / xs
Label19.FontSize = 14 / (fs / 0.8)
'Label19.Width = 52 / xs

Label20.Top = 3 / ys
Label20.Left = 591 / xs
Label20.FontSize = 14 / (fs / 0.8)
'Label20.Width = 52 / xs

Label21.Top = 3 / ys
Label21.Left = 534 / xs
Label21.FontSize = 14 / (fs / 0.8)
'Label21.Width = 52 / xs

Label22.Top = 24 / ys
Label22.Left = 423 / xs
Label22.FontSize = 8 / (fs)
'Label22.Width = 46 / xs

Label23.Top = 24 / ys
Label23.Left = 534 / xs
Label23.FontSize = 8 / (fs)
'Label23.Width = 46 / xs


Label15.Top = 402 / ys
Label15.Left = 42 / xs
Label15.FontSize = 14 / fs
'Label15.Width = 190 / xs

Label16.Top = 402 / ys
Label16.Width = 100 / xs
Label16.Left = Pulse.ScaleWidth - Label16.Width
'abel16.FontSize = 14 / fs

Cls



Line (0, 0)-(639 / xs, 430 / ys), , B
For r = 400 To 50 Step -50
Line (75 / xs, r / ys)-(80 / xs, r / ys)
Next
For r = 400 To 50 Step -50
Line (80 / xs, r / ys)-(600 / xs, r / ys)
Next
Line (80 / xs, 400 / ys)-(600 / xs, 400 / ys) 'x Along
Line (80 / xs, 400 / ys)-(80 / xs, 50 / ys) 'y Up

DrawWidth = 1.5 ' Use wider brush.


End Sub

Private Sub Label25_Click()

If commvar$ = "1" Then Shell "notepad D:\VB6\vb-nt\pulse\pulse.txt", vbNormalFocus

If commvar$ = "4" Then Shell "notepad D:\VB6\vb-nt\pulse\pulsem.txt", vbNormalFocus



End Sub
