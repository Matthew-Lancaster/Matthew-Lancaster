VERSION 5.00
Begin VB.Form Form2_WORK_SWAPPER 
   Caption         =   "Form2"
   ClientHeight    =   7464
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   12744
   LinkTopic       =   "Form2"
   ScaleHeight     =   7464
   ScaleWidth      =   12744
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAPageFile 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   6696
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   6852
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtPageFile 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6696
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtVirtualMemory 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6696
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6108
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtAVirtualMemory 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6696
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   6348
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtPhysicalMemory 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6696
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5592
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtAPhysicalMemory 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6696
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5844
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.Label Label35 
      Caption         =   "Avaible Page File:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   3960
      TabIndex        =   34
      Top             =   6852
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label36 
      Caption         =   "Total Page File _ K:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3960
      TabIndex        =   33
      Top             =   6600
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label37 
      Caption         =   "Total Virtual Memory _ K:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3960
      TabIndex        =   32
      Top             =   6108
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label38 
      Caption         =   "Avaible Virtual Memory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3960
      TabIndex        =   31
      Top             =   6348
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label39 
      Caption         =   "Total Physical Memory _ K:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3960
      TabIndex        =   30
      Top             =   5592
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label40 
      Caption         =   "Avaible Physical Memory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3960
      TabIndex        =   29
      Top             =   5856
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label4 
      Caption         =   "Total Physical Memory My Computer MAXIMUM 32G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   3960
      TabIndex        =   28
      Top             =   5340
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Label txtMemoryUsage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "txtMemoryUser"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   228
      Left            =   7308
      TabIndex        =   27
      Top             =   5088
      Visible         =   0   'False
      Width           =   2076
   End
   Begin VB.Label Label41 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Memory Usage 1 Second Ticker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   228
      Left            =   3960
      TabIndex        =   26
      Top             =   5088
      Visible         =   0   'False
      Width           =   3312
   End
   Begin VB.Label Label62 
      Caption         =   "TASKKILLER /F /IM *"
      Height          =   240
      Left            =   3948
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   1776
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   3960
      X2              =   9300
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   3960
      X2              =   9300
      Y1              =   4872
      Y2              =   4872
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   3960
      X2              =   9300
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      Caption         =   "TASKKILLER PID_NUMERIC OPTION_INSTEAD"
      Height          =   240
      Left            =   5748
      TabIndex        =   24
      Top             =   3828
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   3960
      X2              =   9300
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label43 
      Caption         =   "TASKKILLER BY HWND HANDLE _ PROCESS KILL FORCE"
      Height          =   240
      Left            =   3948
      TabIndex        =   23
      Top             =   4608
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Label Label34 
      Caption         =   "TASKKILLER BY HWND HANDLE _ POST MESSENGER"
      Height          =   240
      Left            =   3948
      TabIndex        =   22
      Top             =   4356
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "WILD CARD ****"
      Height          =   240
      Left            =   3948
      TabIndex        =   21
      Top             =   2376
      Visible         =   0   'False
      Width           =   1776
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Without WildCard Butter"
      Height          =   240
      Left            =   5736
      TabIndex        =   20
      Top             =   2376
      Visible         =   0   'False
      Width           =   1848
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "KILLER With /F FORCE"
      Height          =   240
      Left            =   7608
      TabIndex        =   19
      Top             =   2376
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   7608
      TabIndex        =   18
      Top             =   2688
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TASKKILLER /F /IM /T"
      Height          =   240
      Left            =   7620
      TabIndex        =   17
      Top             =   2976
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TASKKILLER /F /IM"
      Height          =   240
      Left            =   7620
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TASKKILLER /F /IM"
      Height          =   240
      Left            =   5736
      TabIndex        =   15
      Top             =   3516
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label25 
      Caption         =   "TASKLIST GO"
      Height          =   240
      Left            =   3948
      TabIndex        =   14
      Top             =   4092
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Label Label21 
      Caption         =   "TASKKILLER /IM"
      Height          =   240
      Left            =   5736
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label20 
      Caption         =   "TASKKILLER /IM /T"
      Height          =   240
      Left            =   5736
      TabIndex        =   12
      Top             =   2976
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TASKKILLER /F /IM /T"
      Height          =   240
      Left            =   5736
      TabIndex        =   11
      Top             =   2688
      Visible         =   0   'False
      Width           =   1848
   End
   Begin VB.Label Label18 
      Caption         =   "TASKKILLER /IM *"
      Height          =   240
      Left            =   3948
      TabIndex        =   10
      Top             =   3516
      Visible         =   0   'False
      Width           =   1776
   End
   Begin VB.Label Label17 
      Caption         =   "TASKKILLER /IM * /T"
      Height          =   240
      Left            =   3948
      TabIndex        =   9
      Top             =   2976
      Visible         =   0   'False
      Width           =   1776
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   3948
      TabIndex        =   8
      Top             =   2688
      Visible         =   0   'False
      Width           =   1776
   End
   Begin VB.Label Label57 
      BackColor       =   &H00DFFFFF&
      Caption         =   "TASKKILLER COMMAND LINE EXECUTE STATUS"
      Height          =   240
      Left            =   4968
      TabIndex        =   1
      Top             =   1104
      Visible         =   0   'False
      Width           =   5724
   End
   Begin VB.Label Label23 
      Caption         =   "HITT TO KILL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   4968
      TabIndex        =   0
      Top             =   1368
      Visible         =   0   'False
      Width           =   5724
   End
   Begin VB.Menu MNU_ITEM 
      Caption         =   "FORM2_WORK_SWAPPER_TRANSFER_FORM_ITEM_WITH_FORM1"
   End
End
Attribute VB_Name = "Form2_WORK_SWAPPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize_OLD()
    
'    Dim Me_Top, hWndTask, hWndResult
'
'    If SIZE_SET = True Then Exit Sub
'    SIZE_SET = True
'    'Taskbar at the Bottom
'    '---------------------
'    Dim RectTaskbar As RECT
'    hWndTask = FindWindow("Shell_TrayWnd", vbNullString)
'    hWndResult = GetWindowRect(hWndTask, RectTaskbar)
'    ' CENTER ON SCREEN COMPENSATE TASKBAR AT BOTTOM
'    '----------------------------------------------
'    If Screen.height - RectTaskbar.Top * Screen.TwipsPerPixelY < 2500 Then
'        Me.Left = Screen.width / 2 - Me.width / 2
'        Me_Top = (Screen.height - ((RectTaskbar.Bottom - RectTaskbar.Top) * Screen.TwipsPerPixelY)) / 2 - Me.height / 2
'    Else
'        Me.Left = Screen.width / 2 - Me.width / 2
'        Me_Top = Screen.height / 2 - Me.height / 2
'    End If
'    If Me_Top < 0 Then Me_Top = 0
'    Me.Top = Me_Top
'
'    Me.width = LSTPROCESS_3_SORTER_LISTVIEW.Left + LSTPROCESS_3_SORTER_LISTVIEW.width + 300

End Sub

