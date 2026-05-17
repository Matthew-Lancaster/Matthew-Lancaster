VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_ClipTest 
   BackColor       =   &H00404040&
   Caption         =   "Clipboard Viewer"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   2148
   ClientWidth     =   15672
   Icon            =   "frmClipTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   15672
   Visible         =   0   'False
   Begin VB.Timer TIMER_CLIPBOARD_PROCESS_ADBLOCK_FILE 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5580
      Top             =   2748
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   5004
      Sorted          =   -1  'True
      TabIndex        =   53
      Top             =   3144
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Timer Timer_SECOND 
      Interval        =   1000
      Left            =   4632
      Top             =   3072
   End
   Begin MCI.MMControl MMControl4 
      Height          =   264
      Left            =   7968
      TabIndex        =   49
      Top             =   2832
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   6244
      _ExtentY        =   572
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer TIMER_CLIPBOARD_TIMER_RETRY 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4284
      Top             =   3072
   End
   Begin VB.DirListBox Dir1 
      Height          =   288
      Left            =   5400
      TabIndex        =   38
      Top             =   3000
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.FileListBox File2 
      Height          =   264
      Left            =   6468
      TabIndex        =   37
      Top             =   2724
      Visible         =   0   'False
      Width           =   1308
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   336
      Left            =   6864
      TabIndex        =   36
      Top             =   2904
      Visible         =   0   'False
      Width           =   1044
      _ExtentX        =   1842
      _ExtentY        =   572
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   3864
      Top             =   3060
   End
   Begin VB.Timer Timer_GIVEN_FILENAME_AND_CMD_LINE_STRING_FOR_COMPARE 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3744
      Top             =   3060
   End
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   14808
      TabIndex        =   35
      Top             =   8484
      Visible         =   0   'False
      Width           =   1248
   End
   Begin VB.Timer TIMER_Form_Resize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3408
      Top             =   3060
   End
   Begin VB.Timer Timer_RESET_API_CLIPPER 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   3060
      Top             =   3060
   End
   Begin VB.Timer Timer_CHECK_PROJECT_DATE_IN_PROCESS 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2712
      Top             =   3060
   End
   Begin VB.Timer Timer_CLIPBOARD_API 
      Interval        =   4000
      Left            =   2364
      Top             =   3060
   End
   Begin VB.Timer TIMER_TO_RESIZE 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2004
      Top             =   3060
   End
   Begin VB.Timer Timer_EXIT 
      Interval        =   10
      Left            =   1644
      Top             =   3060
   End
   Begin VB.Timer Timer_SHOW_THE_TIME 
      Interval        =   1000
      Left            =   1284
      Top             =   3060
   End
   Begin MCI.MMControl MMControl1 
      Height          =   348
      Left            =   10332
      TabIndex        =   20
      Top             =   8748
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   6244
      _ExtentY        =   614
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin RichTextLib.RichTextBox Text3 
      Height          =   3276
      Left            =   10356
      TabIndex        =   13
      Top             =   4620
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   5779
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmClipTest.frx":1272
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   3276
      Left            =   912
      TabIndex        =   12
      Top             =   4620
      Width           =   9408
      _ExtentX        =   16595
      _ExtentY        =   5779
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmClipTest.frx":12F4
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   252
      Left            =   14820
      TabIndex        =   11
      Top             =   7908
      Visible         =   0   'False
      Width           =   1236
      _ExtentX        =   2180
      _ExtentY        =   445
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmClipTest.frx":1376
   End
   Begin VB.Timer Timer_Test_Logic 
      Interval        =   1000
      Left            =   924
      Top             =   3060
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Text"
      Height          =   276
      Left            =   13284
      TabIndex        =   2
      Top             =   8484
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Picture 2"
      Height          =   276
      Index           =   1
      Left            =   13284
      TabIndex        =   1
      Top             =   8196
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Picture 1"
      Height          =   276
      Index           =   0
      Left            =   13284
      TabIndex        =   0
      Top             =   7908
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MCI.MMControl MMControl8 
      Height          =   264
      Left            =   7968
      TabIndex        =   50
      Top             =   3132
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   6244
      _ExtentY        =   572
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "COMPARE VIEWER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   7
      Left            =   5472
      TabIndex        =   52
      Top             =   4224
      Width           =   4440
   End
   Begin VB.Label LABEL_COMPARE_CASE_SENSITIVE 
      Alignment       =   2  'Center
      Caption         =   "Compare Not Sensitive Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   912
      TabIndex        =   51
      Top             =   2736
      Width           =   4644
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   9
      Left            =   19068
      TabIndex        =   48
      Top             =   2388
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   8
      Left            =   19068
      TabIndex        =   47
      Top             =   3504
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   7
      Left            =   19056
      TabIndex        =   46
      Top             =   2760
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   6
      Left            =   19068
      TabIndex        =   45
      Top             =   3132
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   5
      Left            =   19056
      TabIndex        =   44
      Top             =   1632
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   4
      Left            =   19056
      TabIndex        =   43
      Top             =   2016
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   3
      Left            =   19056
      TabIndex        =   42
      Top             =   1260
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   2
      Left            =   19056
      TabIndex        =   41
      Top             =   876
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   2  'Center
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   0
      Left            =   19056
      TabIndex        =   40
      Top             =   108
      Width           =   2220
   End
   Begin VB.Label MMM_SHOW_THE_TIME_ARR 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE TIME 03 BUTTON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   1
      Left            =   19056
      TabIndex        =   39
      Top             =   492
      Width           =   2220
   End
   Begin VB.Label Label_FindWindowPart_VB_CLIPVIEWER_AUTO_BUTTON_COMPARE 
      Alignment       =   2  'Center
      Caption         =   "Make All Instance of Clipboard Viewer Do Compare With The Loader File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   912
      TabIndex        =   34
      Top             =   2376
      Width           =   9408
   End
   Begin VB.Label MNU_SHOW_THE_TIME_03 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE ME TIME 04"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   10368
      TabIndex        =   33
      Top             =   480
      Width           =   8592
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Picker for Clipboard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   924
      TabIndex        =   32
      Top             =   2004
      Width           =   9408
   End
   Begin VB.Label Label_CHECK_PROJECT_DATE_IN_PROCESS 
      Alignment       =   2  'Center
      Caption         =   "CheckForm Project Date in Process"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   912
      TabIndex        =   31
      Top             =   1632
      Width           =   9408
   End
   Begin VB.Label Label1 
      Caption         =   "COMPARE SAME TRUE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   6
      Left            =   5676
      TabIndex        =   30
      Top             =   96
      Width           =   4656
   End
   Begin VB.Label Label_COLOR_RED 
      BackColor       =   &H0005A9A5&
      Caption         =   "Label4"
      Height          =   300
      Left            =   10356
      TabIndex        =   29
      Top             =   7992
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label MNU_SHOW_THE_TIME_03 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE ME TIME 03"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   0
      Left            =   10380
      TabIndex        =   28
      Top             =   108
      Width           =   8592
   End
   Begin VB.Label LabelCRC5 
      Alignment       =   2  'Center
      Caption         =   "DON'T MINIMIZE CLIPPER -- -- TIME && COMBINE --"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   15204
      TabIndex        =   27
      Top             =   2028
      Width           =   3780
   End
   Begin VB.Label Label5 
      Caption         =   " UPPER CASE ALL URL CLIPPER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   10380
      TabIndex        =   26
      Top             =   3108
      Width           =   8616
   End
   Begin VB.Label LabelCRC4 
      Caption         =   "CRC4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   10380
      TabIndex        =   25
      Top             =   2400
      Width           =   4800
   End
   Begin VB.Label Label3 
      Caption         =   " LOWER CASE ALL URL CLIPPER FOR YOUTUBE DESCRIPTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   10380
      TabIndex        =   24
      Top             =   2784
      Width           =   8616
   End
   Begin VB.Label Lab_TIME_CLIPBOARD_CHANGE 
      Alignment       =   2  'Center
      Caption         =   "Last Clipboard Time API Ticker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   912
      TabIndex        =   23
      Top             =   1260
      Width           =   9408
   End
   Begin VB.Label MNU_SHOW_THE_TIME_03 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE ME TIME 02"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   2
      Left            =   10380
      TabIndex        =   22
      Top             =   888
      Width           =   8592
   End
   Begin VB.Label MNU_SHOW_THE_TIME_03 
      Alignment       =   1  'Right Justify
      Caption         =   "GIVE ME TIME 01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   3
      Left            =   10380
      TabIndex        =   21
      Top             =   1272
      Width           =   8604
   End
   Begin VB.Label Label2 
      Caption         =   "MESSENGER CRC32"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   10380
      TabIndex        =   19
      Top             =   1644
      Width           =   8604
   End
   Begin VB.Label LabelCRC3 
      Caption         =   "CRC3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   10380
      TabIndex        =   18
      Top             =   2028
      Width           =   4800
   End
   Begin VB.Label LabelFILE2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FILE2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   924
      TabIndex        =   17
      Top             =   3828
      Width           =   18060
   End
   Begin VB.Label LabelFILE1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FILE1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   924
      TabIndex        =   16
      Top             =   3432
      Width           =   18060
   End
   Begin VB.Label LabelCRC1 
      Caption         =   "CRC1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   912
      TabIndex        =   15
      Top             =   876
      Width           =   4656
   End
   Begin VB.Label LabelCRC2 
      Caption         =   "CRC2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5676
      TabIndex        =   14
      Top             =   876
      Width           =   4656
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Compare"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   5
      Left            =   11616
      TabIndex        =   10
      Top             =   8328
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Compare"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Index           =   4
      Left            =   10356
      TabIndex        =   9
      Top             =   8316
      Visible         =   0   'False
      Width           =   1224
   End
   Begin VB.Label Label1 
      Caption         =   "Compare when not the Same"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   912
      TabIndex        =   8
      Top             =   96
      Width           =   4656
   End
   Begin VB.Label Label_SIZE_2 
      Caption         =   "SIZE 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5676
      TabIndex        =   7
      Top             =   492
      Width           =   4656
   End
   Begin VB.Label Label_SIZE_1 
      Caption         =   "SIZE 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   924
      TabIndex        =   6
      Top             =   492
      Width           =   4644
   End
   Begin VB.Label Label1 
      Caption         =   "Text in Clipboard -- PREVIOUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   10392
      TabIndex        =   5
      Top             =   4224
      Width           =   5064
   End
   Begin VB.Label Label1 
      Caption         =   "Text in Clipboard -- CURRENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   936
      TabIndex        =   4
      Top             =   4212
      Width           =   4452
   End
   Begin VB.Label Label1 
      Caption         =   "Picture in Clipboard"
      Height          =   300
      Index           =   0
      Left            =   13284
      TabIndex        =   3
      Top             =   8772
      Visible         =   0   'False
      Width           =   1476
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   14820
      Stretch         =   -1  'True
      Top             =   8196
      Visible         =   0   'False
      Width           =   1236
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_ALWAYS_ON_TOP 
      Caption         =   "ON TOP"
   End
   Begin VB.Menu MNU_FOLDER_DIGITAL_STILL_CAMERA 
      Caption         =   "FOLDER DIGITAL STILL CAMERA"
   End
   Begin VB.Menu MNU_CAMERA_VIDEO_FOLDER_4G 
      Caption         =   "VIDEO FOLDER 4G"
   End
   Begin VB.Menu MNU_VERSION 
      Caption         =   "VERSION"
   End
   Begin VB.Menu MNU_API_RESET 
      Caption         =   "API"
   End
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMIN"
   End
   Begin VB.Menu MNU_CAPS_TXT_LOW 
      Caption         =   "CAPS LOW"
   End
   Begin VB.Menu MNU_CAP_TEXT 
      Caption         =   "CAPS UP"
   End
   Begin VB.Menu MNU_CHAR_CODE 
      Caption         =   "SINGLE CHARACTOR CODE -- 0"
   End
   Begin VB.Menu MNU_FIND_DIFFER 
      Caption         =   "FIND DIFFER"
   End
   Begin VB.Menu MNU_NUMBER_TO_STRING 
      Caption         =   "NUMBER TO STRING"
   End
   Begin VB.Menu MNU_NUMBER_TO_STRING_AND_DATE 
      Caption         =   "NUMBER TO STRING && DATE"
   End
   Begin VB.Menu MNU_CLIP_TIME 
      Caption         =   "GIVE ME TIME -->"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_GIVE_TIME_GITHUB 
      Caption         =   "GIVE TIME GITHUB"
   End
   Begin VB.Menu MNU_GOODSYN_CCONVERT_SMBD_PATH_TO_DOS 
      Caption         =   "GOODSYNC CONVERT smbd:// TO NET_PATH //"
   End
   Begin VB.Menu MNU_MESSENGER_GIVE_CRC32 
      Caption         =   "MESSENGER GIVE CRC32"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_GIVE_ME_TIME_WITHER_UTC 
      Caption         =   "GIVE ME TIME AND UNI_ UTC Time Toggle = YES"
   End
   Begin VB.Menu MNU_GIVE_ME_68_DASH_LINE_REMARK 
      Caption         =   "GET 70 DASH LINE REMARK"
   End
   Begin VB.Menu MNU_GIVE_ME_68_DASH_LINE_2 
      Caption         =   "GET 70 DASHES"
   End
   Begin VB.Menu MNU_GIVE_ME_40_DASH_LINE 
      Caption         =   "GET 40 DASHES"
   End
   Begin VB.Menu MNU_GIVE_ME_10_DASH_LINE 
      Caption         =   "GET 10 DASHER"
   End
   Begin VB.Menu MNU_COMPARE 
      Caption         =   "COMPARE 1 2"
   End
   Begin VB.Menu MNU_COMPARE_IN_HEXVIEWER 
      Caption         =   "COMPARE HEXVIEWER"
   End
   Begin VB.Menu MNU_CLIPBOARDER_REPLACE_ER_AND 
      Caption         =   "CLIPBOARD REPLACE ""AND"""
   End
   Begin VB.Menu MNU_CLIP_TIME_TESTER 
      Caption         =   "CLIPPER TIME TESTER"
   End
   Begin VB.Menu MNU_STRIP_MEDIA_INFO 
      Caption         =   "STRIP MEDIA INFO"
   End
   Begin VB.Menu MNU_TRIM_MENU 
      Caption         =   "TRIM MENU - 3 ITEM"
      Begin VB.Menu MNU_TRIM_END_OF_LINE_SPACE 
         Caption         =   "TRIM END OF LINE SPACE"
      End
      Begin VB.Menu MNU_TRIM_SPACE_ON_EACH_LINE_AND_WITHIN 
         Caption         =   "TRIM SPACE ON EACH LINE AND WITHIN"
      End
      Begin VB.Menu MNU_TRIM_SPACE_ON_EACH_LINE_AND_WITHIN_UPTO_FIRST_COLON 
         Caption         =   "TRIM SPACE ON EACH LINE AND WITHIN UPTO FIRST COLON"
      End
   End
   Begin VB.Menu MNU_SPACE_DASH_TO_UNDERSCORE 
      Caption         =   "SPACE && DASH TO UNDERSCORE"
   End
   Begin VB.Menu MNU_KILL_OTHER_CLIPVIEWER 
      Caption         =   "KILL OTHER CLIP VIEWER"
   End
   Begin VB.Menu MNU_KILL_ALL_CLIPVIEWER 
      Caption         =   "KILL ALL CLIP VIEWER"
   End
   Begin VB.Menu MNU_TIME_DIFFERENCE 
      Caption         =   "TIME DIFFERENCE"
   End
   Begin VB.Menu MNU_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE 
      Caption         =   "AUTO CONVERT -- FALSE"
   End
   Begin VB.Menu MNU_CONVERT_GSD_PATH_TO_NETWORK_FRIENDLY_ONE 
      Caption         =   "CONVERT GOODSYNC GSD DIRECT PATH TO NETWORK FRIENDLY PATH"
   End
   Begin VB.Menu MNU_GOODSYNC_VOLUME_LABLE_DRIVE 
      Caption         =   "GOODSYNC VOLUME LABEL DRIVE CONVERT"
   End
   Begin VB.Menu MNU_DIRECTORY_PICKER 
      Caption         =   "DIRECTORY PICKER"
   End
   Begin VB.Menu MNU_POUNDS_ON_NEWLINE_AND_CUT_ABOVE_AMOUNT 
      Caption         =   "POUNDS ON NEWLINE AND CUT ABOVE £200"
   End
   Begin VB.Menu MNU_TIME_BACKWARD_800_DOOR_BELL 
      Caption         =   "TIME BACKWARD 800 DOOR BELL"
   End
   Begin VB.Menu MNU_RUN_GOODSYNC_CAMERA 
      Caption         =   "RUN GOODSYNC CAMERA D-DRIVE PORTABLE"
   End
   Begin VB.Menu MNU_CALC_COOP_20 
      Caption         =   "CALC COOP £20"
   End
   Begin VB.Menu MNU_CALC_COOP_40 
      Caption         =   "CALC COOP £40"
   End
   Begin VB.Menu MNU_RS232_LOGGER_OPEN 
      Caption         =   "RS232 LOGGER OPEN"
   End
   Begin VB.Menu MNU_RS232_LOGGER_CLOSE 
      Caption         =   "RS232 LOGGER CLOSE"
   End
   Begin VB.Menu MNU_CLIPBOARD_LOGGER 
      Caption         =   "CLIPBOARD LOGGER"
   End
   Begin VB.Menu MNU_TIMEZONE_CLOCK 
      Caption         =   "TIMEZONE_CLOCK"
   End
   Begin VB.Menu MNU_VIDEODOWNLOADERULTIMATE_EXE 
      Caption         =   "Video Downloader Ultimate.exe"
   End
   Begin VB.Menu MNU_MENU_ITEM_COUNT 
      Caption         =   "MENU ITEM COUNT"
   End
   Begin VB.Menu MNU_SetPoint_exe 
      Caption         =   "SET POINT MOUSE CONTROL"
   End
   Begin VB.Menu MNU_UNIFY_AR_LOGITECH 
      Caption         =   "UNIFY AR LOGITECH"
   End
   Begin VB.Menu MNU_PROCESS_ADBLOCK_FILE 
      Caption         =   "PROCESS ADBLOCK FILE"
   End
   Begin VB.Menu MNU_ADBLOCK_FOLDER 
      Caption         =   "ADBLOCK FOLDER"
   End
   Begin VB.Menu MNU_CLIPBOARD_ADBLOCK 
      Caption         =   "CLIPBOARD ADBLOCK"
   End
   Begin VB.Menu MNU_NOTEPAD_CLIPPPER 
      Caption         =   "NOTEPAD CLIPPER"
   End
End
Attribute VB_Name = "FRM_ClipTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Option Explicit


Dim HWND_2_LATCH

Dim ADBLOCK_INFO_MAIN_CHUNK

Dim HWND_2_OLD
Dim HWND_2_TEXT_OLD

Dim COOP_CALC_SETTER

Dim MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER

Dim MMCONTROL4_FIRST_LOAD_SOUND_OFF
Dim HY, WX

Dim CLIPBOARD_SET_TEXT_FOR_BANK_VAR

Dim A_NOW
Dim FOLDER_NAME
Dim FILE_NAME_4


' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 001 AND HALF HOUR OF CODE WENT IN
' Sat 27-Jul-2019 21:59:07
' Sat 27-Jul-2019 11:28:00
' ANOTHER UP TO HERE
' Sun 28-Jul-2019 02:34:00 END OF ALL SESSION
' ----------------------------------------------------------------------
' 01.
' GET GIVE_ME_TIME GOES WITH CRC32
' AND IF TIME GOT AFTER CRC32 IT COMBINE TIME WITH CRC32
' 02.
' API IS CHECK FOR RESET BY WAY OF SECOND SINCE LAST CLIPBOARD
' AMATEUR FOR NOW
' 03.
' MENU HAD MENUBAR WORK AND CONTROL COUNTER
' 04.
' MNU_GIVE_ME_68_DASH_LINE_REMARK FOR CODE REMMER REMARK
' LIKE THESE -- MAKE LINE LENGTH 70 INCLUDE REM
' ----------------------------------------------------------------------
' 05.
' PUT THE SCROLL BAR ON BOTH TEXT BOX
' 07.
' PUTTER LOWER CASE OPTION FOR YOUTUBE URL -- FUSSY ABOUT THAT ONE
' HIGHLIGHT THEM HTML TO CLICK -- MUST BE LOWER CASE
' GOOGLE YOUTUBE HELP PAGE GONE IDIOT OVER THIS ONE
' WORK IT OUT YOUR OWN
' 08. ADD ANOTHER TIME FORMAT TO GRAB LIKE CLIPBOARD LOGGER ONE
' NOW MAKE 3
' 09. LABEL SET BETTER
' 10. MENUBAR NOW WORK WITH FORM SIZE SETTING
' 11. API BUTTON IS ON LABEL FORM AS WELL AS MENU
' 12. MORE COLOUR CHANGER CRC32 NOT MATCH
' ----------------------------------------------------------------------
' Sat 27-Jul-2019 21:59:07
' Sat 27-Jul-2019 11:28:00
' 07.
' Sun 28-Jul-2019 01:39:00 UP TO COMPLETE SESSION
' 08. 09. 10. 11.
' Sun 28-Jul-2019 02:34:00 END OF ALL SESSION
' 12.
' Sun 28-Jul-2019 03:10:00 END OF ALL SESSION
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 002 -- SESSION 002 SINCE LOGGER
' ----------------------------------------------------------------------
' LOT WORK TODAY
' IMPORT CODE THAT RUN OTHER VB APP I MAKE
' AND HAS SYSTEM OF UPDATE AROUND NETWORK
' WHICH UNLOAD ITSELF TO DO
' NEW CODE FOR THAT OPERATION
' MAKE EASIER TO IMPLIMENT INTO OTHER FOM IN FURTURE
' WITH LEAST REQUIRE DIFFERCULTY
' AND SIMPLE CODE ROUTINE ADD OTHER FORM
' AND STOP ERROR WHERE OTHER FORM NOT UPDATED YET
' THAT USER CODE
' NEW CODE WILL
' CHECK IF A FORM HAS A CONTROL
' CHECK THE FORM NAME EXIST WHEN RUN START
' WHEN FIND CONTROL IT WILL CALL THAT ROUTINE TO ADD WHATEVER WANT DO
' IT USE A METHOD TO CHECK FORM EXIST WITH MAKE ERROR TRAP EASIER
' NORMAL REQUEST TO SEARCHER A FORM WILL FLAG BAD ERROR IF NOT EXIST
' THAT BEYOUND EERROR TRAPPER
' SO WE FIND NEW WAY TO FIND WITHOUT THAT ERROR
' AND WITH DONTROL FIND NEAR SMAE NETHOD
' WE WEAPONISE THE NEW COMMAND WHICH IS CALL
' CallByName Form, SUBNAME, VbMethod
' AND THEN ROUTINE HERE
' Sub ControlCall_Find(ControlName As String)
' ----------------------------------------------------------------------
' FROM -- Sat 31-Aug-2019 07:53:44
' TO   -- Sat 31-Aug-2019 12:02:00 -- LONG TIME SPENT ON THIS ONE 4 HOUR
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 003 -- AND THEN TWO SESSION ONE DAY
' ----------------------------------------------------------------------
' ADD NEW EASY PICKER FOR STRING I OFTEN WANT CLIPBOARDER
' AND LATER ADD
' BUG FIX THAT CRC32 IS DISPLAY WITH FILE COUNT
' MUST OF BEEN RUSH TO ENDER BEFORE
' AND ALSO
' NOW GOT CLOSE OTHER INSTANCE OF THIS APP WHEN COMPARE WAS SUPPOSE TO SEE FEW GONE
' AND FOR TEST THE COMMAND LINE AS SAVED
' SO WHEN IN IDE EASY LINKER FOR COMPLETE
' THE NEW PICKER FORM
' HAD OPTION TO MINIMIZE WHEN CALLED UP
' MORE LIKE AFTER A PICK IT MINIMIZE AND BOTH FORM ARE DUAL AND UP DOWN SAME
' BUT HARD WORK THINK IT THROUGH
' WHEN PICK SOMETHING I GUESS IT THAT PICK MODE IS ON SO WHEN RESTORE BRING BACK PICKER AGIN
' IT START IN MAXIMISED
' AND HARD SET FOCUS WITHOUT JITTER FROM OTHER FORM
' WHAT HAPPEN AT END
' THE FORM BOTH HAVE SAME TITLE AND ICON
' AND WHEN ONE IS HIDE IT REMOVE FROM TAKSBAR
' THAT WHY FIRST CALLER
' TWO FORM OF SAME SHOW - FOR STARTING
' GET HANG IT THERE
' ALSO RATHER THAN MINIMIZE IT ADDITIONALLY HAS STAY ON TOP MODE
' ----------------------------------------------------------------------
' FROM -- Sat 31-Aug-2019 07:53:44 %%%%
' TO   -- Sat 31-Aug-2019 21:30:00 -- LONG TIME SPENT ON THIS ONE 4 HOUR
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 003 -- AND THEN THREE SESSION ONE DAY
' ----------------------------------------------------------------------
' SESSION 003 IS MOST WHAT IN SESSION 002
' EXCEPT ALL THE PICKER FORM WAS DONE
' ----------------------------------------------------------------------
' FROM -- Sat 31-Aug-2019 22:30:00
' TO   -- Sat 31-Aug-2019 23:59:59 -- LONG TIME SPENT ON THIS ONE 4 HOUR
' ----------------------------------------------------------------------

' LATEST WORK
' Timer_GIVEN_FILENAME_AND_CMD_LINE_STRING_FOR_COMPARE_Timer

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 004 -- NUMBER TO STRING WORK PLUS OTHER
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' FROM -- Mon 25-Nov-2019 06:25:29
' TO   -- Mon 25-Nov-2019 07:11:00 -- 1 HOUR 46 MINUTE
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 005 --  WORK -- MNU_COMPARE_IN_HEXVIEWER
' --------------------------------
' Sat 04-Jan-2020 09:47:56
' Sat 04-Jan-2020 12:11:09 -- 2 HOUR 23 MINUTE
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION SIXER -- WORK ADD FOURTH COLUMN -- AND ARRAY SOME THAT AREA
' WORK BETTER
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' FROM -- Sat 11-Jan-2020 21:21:17
' TO   -- Sun 12-Jan-2020 00:22:00 -- 3 HOUR 1 MINUTE -- MIDNIGHT
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 007 -- CORRECT CRC32 COUNT FILESIZE
' ----------------------------------------------------------------------
' Wed 22-Jan-2020 00:06:34
' Wed 22-Jan-2020 04:48:03 - 18 HOUR 35 MINUTE
' ----------------------------------------------------------------------
' 22-Jan-2020 00:06:34
' 22-Jan-2020 04:48:03 -- 5 HOUR 41 MINUTE
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 008 -- MORE TIME DATE DIFFERNCE MENU -- MIX OTHER CODE WITHER
' INCLUDE RESULT TIMER
' ----------------------------------------------------------------------
' MNU_TIME_DIFFERENCE_CODE_Click
' ----------------------------------------------------------------------
' NOW CLICK GRAB DATE IN FORM STYLE
' MUCH BETTER DEBUG IT DATEVALUE
' REMOVE CHECK OF OVER IMDNIGHT DAY BEFORE
' MAKE -- DATE ALLOW TWO
' TRIM AND REMOVE LINE FEED CARRIAGE RETURN 1ST LINE ONLY
' BEFORE IF NONE SECOND DATE USER NOW
' IF DATE IS TIME ONLY
' REQUIRE ADDTIONAL DATE FROM NOW TODAY TO WORK DATEDIFF()
' ----------------------------------------------------------------------
' Thu 23-Jan-2020 13:51:38
' Thu 23-Jan-2020 17:54:29 -- 4 HOUR 3 MINUTE
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' SESSION 009 -- ADD LOAD BIT AN BOBS
' ----------------------------------------------------------------------
' ADD THE CALL
' Public Sub TIMER_CLIPBOARD_TIMER_RETRY_Timer()
' NOW ABLE TO USE TIMER WITH CLIPBOARD
' AND NOT LOOP EVER
' IF CLIPBOARD FAIL HAS COOMAND TO EXIT PROCEDURE AND THEN
' TIMER CALL BACK
' CallByName Form, VAR_TIMER_CLIPBOARD_TIMER_RETRY, VbMethod
' ----------------------------------------------------------------------
' AND TIMEDIFF CODE -- UPDATER
' ----------------------------------------------------------------------
' DONE CODE HERE -- STILL GOT TO DO
' Project_Check_Date
' ----------------------------------------------------------------------
' Thu 23-Jan-2020 13:51:38
' Thu 24-Jan-2020 05:40:00 -- 16 HOUR 48 MINUTE
' ----------------------------------------------------------------------


' ----------------------------------------------------------------------
' SESSION 010
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' WORK HERE
' MNU_POUNDS_ON_NEWLINE_AND_CUT_ABOVE_AMOUNT_Click()
' ----------------------------------------------------------------------
' Fri 07-Feb-2020 01:58:27
' Fri 07-Feb-2020 05:12:00 -- 3 HOUR 14 MINUTE
' ----------------------------------------------------------------------


' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------


' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------

Dim WAIT_FORM_BOOT_BEFORE_GO

Dim VAR_TIMER_CLIPBOARD_TIMER_RETRY

Dim VARTIME_DIFFERENCE

Dim FRM_CLIPTEST_02_ENABLE

Dim HH_2

Dim FILENAME_AND_COMMAND_STRING_GIVE

Dim SHOW_WINDOW_ALWAYS_ON_TOP_AFTER_FORM_LOAD

Dim FORM_LOAD_GO_TEXT_MULTI_INSTANCE_MINIMIZED

Dim RUN_AT_FORM_LOAD

Dim AT_END_EXIT_CLOSE

Dim DO_FORM_RESIZE_ONCE

Dim FILE_NAME_SIZE_1
Dim FILE_NAME_SIZE_2

Dim RESIZE_WINDOWSTATE_CHANGE

Dim CRC_1
Dim CRC_2

Public DONT_UPDATE_CLIPBOARD_THIS_ONE

Public EXIT_TRUE
Public CHECK_PROJECT_DATE_IN_PROCESS

Dim RUN_ONCE_01

Dim ARCHIVE_Menu_Height

Dim GET_CRC32_IN_LOWER_CASE_FOR_YOUTUBE

Dim DONT_MIN


Dim TO_SETTER

Dim A_DATE_TIME_PM
'Dim TIME_CLIPBOARD_CHANGE
Dim COMMAND_STRING
Dim TEXT_CLIPPER
Dim COMV1
Dim FILENAME_1 As String
Dim FILENAME_2 As String

Public FSO

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Const DontWaitUntilFinished = False, WaitUntilFinished = True, ShowWindow_2 = 1, DontShowWindow = 0


'I = MsgBox("ClipBoard API Has Stopped and Gone Missing" + vbCrLf + "Use the Menu Option *INFO* to Diagnostic and Reload It" + vbCrLf + "This Can Happen If ChkDsk Unlocked All Handles to The Hard Drive and the ClipboardViewer.ocx Driver Couldn't Get Access", vbMsgBoxSetForeground)
'ClipboardViewer.ocx Driver
'ClipboardViewer.oca Driver -- TRUE NAME EXT
'THE VERSION USE IS
'ClipboardViewer.oca
'C:\Program Files\Microsoft Visual Studio\VB98\ClipboardViewer.oca   0x488   29/08/2015 22:56:01 03/04/2016 21:48:47 A   10,240  *           *           0x00120089  1,696   7244    VB6.EXE C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE   oca 16.56%
'MUST BE THE REGISTER VERSION
'AS ONE IN ROOT OF APP EXE DON'T LOCATE TO USE

Dim i, MSGQ, IR
Dim CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2
Dim FileSpec, FILE_SPEC_GO

Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long

' Public constants for set window on top
Private Const hWnd_TOPMOST = -1
Private Const hWnd_NOTOPMOST = -2

' Public constants for SetWindowPos API declaration
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long



Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer



' -----------------------------------------------------------
' -----------------------------------------------------------
' Public Function SetFileDateTime(ByVal Filename As String, ByVal TheDate As String) As Boolean
' -----------------------------------------------------------
' -----------------------------------------------------------
Private Type FILETIME
    dwLowDate  As Long
    dwHighDate As Long
End Type

Private Type SYSTEMTIME
    wYear      As Integer
    wMonth     As Integer
    wDayOfWeek As Integer
    wDay       As Integer
    wHour      As Integer
    wMinute    As Integer
    wSecond    As Integer
    wMillisecs As Integer
End Type

Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
  
Private Declare Function CreateFile Lib "kernel32" Alias _
   "CreateFileA" (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) _
   As Long

Private Declare Function LocalFileTimeToFileTime Lib _
     "kernel32" (lpLocalFileTime As FILETIME, _
      lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
   (ByVal hFile As Long, ByVal MullP As Long, _
    ByVal NullP2 As Long, lpLastWriteTime _
    As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
    "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
    As FILETIME) As Long
    
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long
' -----------------------------------------------------------
' -----------------------------------------------------------
' Public Function SetFileDateTime(ByVal Filename As String, ByVal TheDate As String) As Boolean
' -----------------------------------------------------------
' -----------------------------------------------------------
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long


' --------------------------------------------------------------------
' --------------------------------------------------------------------
' SET OF DECLARE WORTH HAVE ALL CODE PROJECT BEGGINER
' --------------------------------------------------------------------
' --------------------------------------------------------------------
' BAS MODULE DECLARE
' Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
' --------------------------------------------------------------------
'Private Type Rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
'Private Type MENUBARINFO
'  cbSize As Long
'  rcBar As Rect
'  hMenu As Long
'  hWndMenu As Long
'  fBarFocused As Boolean
'  fFocused As Boolean
'End Type
'Private MenuInfo As MENUBARINFO
'Private Const OBJID_MENU As Long = &HFFFFFFFD
'Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
'Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
'Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'----------------------------------------------------------------------------------
'Private Type FILETIME
'   LowDateTime          As Long
'   HighDateTime         As Long
'End Type
Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type
' --------------------------------------------------------------------
' --------------------------------------------------------------------
'Const HWND_TOPMOST = -1
'Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
'Const SWP_NOSIZE = &H1
'Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
'Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
' --------------------------------------------------------------------
Const GWL_WNDPROC = -4
Const GWL_HINSTANCE = -6
Const GWL_HWNDPARENT = -8
Const GWL_STYLE = -16
Const GWL_EXSTYLE = -20
Const GWL_USERDATA = -21
Const GWL_ID = -12
' --------------------------------------------------------------------
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function FindWindow2 Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cchar As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' --------------------------------------------------------------------
'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Const hWnd_TOPMOST = -1
'Const hWnd_NOTOPMOST = -2
'Const MF_BYPOSITION = &H400&
'Const SWP_NOSIZE = &H1
'Const SWP_NOMOVE = &H2
'Const SPI_SCREENSAVERRUNNING = 97
'Const SWP_NOACTIVATE = &H10
'Const SWP_SHOWWINDOW = &H40
' Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim flags
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function



' --------------------------------------------------------------------
' FUNCTION SET HERE
' --------------------------------------------------------------------
Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
' --------------------------------------------------------------------
Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
' --------------------------------------------------------------------
' --------------------------------------------------------------------
Public Function GetSpecialFolder_Show_Script_Debug(CSIDL As Long) As String
Dim r As Long
On Error Resume Next
For r = 0 To 120
    If Trim(GetSpecialfolder(r)) <> "" Then
        'Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
        'AAX = GetSpecialfolder(R)
    End If
Next
End
End Function
Public Function GetSpecialfolder(CSIDL As Long) As String
    '##############################################################################################
    'Returns the Path to a "Special" Folder (i.e. Internet History)
    '##############################################################################################
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function
Private Function GetTitle(ByVal hWnd As Long) As String
    Dim sTitle                As String
    sTitle = String(MAX_LENGTH, Chr$(0))
    GetWindowText hWnd, sTitle, MAX_LENGTH
    i = InStr(1, sTitle, Chr$(0))
    GetTitle = Mid$(sTitle, 1, i - 1)
End Function
Private Function GetClass(ByVal hWnd As Long)
    Dim sClass                As String
    sClass = String(MAX_LENGTH, Chr$(0))
    GetClassName hWnd, sClass, MAX_LENGTH
    i = InStr(1, sClass, Chr$(0))
    GetClass = Mid$(sClass, 1, i - 1)
End Function
Private Function WindowClass(ByVal hWnd As Long) As String
    Const MAX_LEN As Byte = 255
    Dim strBuff As String, intLen As Integer
    strBuff = String(MAX_LEN, vbNullChar)
    intLen = GetClassName(hWnd, strBuff, MAX_LEN)
    WindowClass = Left(strBuff, intLen)
End Function
Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
End Function
Public Function StripNulls_2(OriginalStr) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls_2 = OriginalStr
End Function
Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hWnd)
   S = Space(L + 1)
   GetWindowText hWnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function
Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
    If IsIconic(lngHwnd) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lngHwnd) <> 0 Then
        GetWindowState = vbMaximized
    End If
    End If
End Function
Public Sub vLaunch(vFile As String)
    Dim Tri
    Tri = 1
    If Mid$(vFile, 1, 5) = "Http:" Then Tri = 1
    If Mid$(vFile, 2, 2) = ":\" Then Tri = 2
    If Right$(vFile, 1) = "\" Then Tri = 3
    Select Case Tri
    Case 1
        ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    Case 2
        vFile = "file:\\" + vFile
        ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
        'Shell "C:\Program Files\Internet Explorer\iexplore.exe " + vFile, vbNormalFocus
    Case 3
        Shell "Explorer.exe /e," + vFile, vbNormalFocus
    End Select
End Sub






Public Function SetFileDateTime(ByVal Filename As String, ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
If Dir(Filename) = "" Then Exit Function
If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

lFileHnd = CreateFile(Filename, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)
    
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
SetFileDateTime = lRet > 0

End Function


Sub GO_VIRTUAL_GIRL()
    Exit Sub
    If IsIDE = False Then Exit Sub
    
    Me.WindowState = vbMaximized

    With LISTVIEW1
        .ColumnHeaders.Add , "PID", "PID", 700 - 50, lvwColumnLeft
        .ColumnHeaders.Add , "EXE", "EXE", 9000, lvwColumnLeft
        .View = lvwReport
        .height = 7000
        .FullRowSelect = True
    End With

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim STRING_1
    Dim STRING_2
    Dim STRING_4
   
    Dim TD()
    ReDim TD(200)
   
    Dir1.Path = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\X VIRTUAL FEMALE\Extra Girl 0"
    i = 0
    For r = 0 To Dir1.ListCount - 1
        i = i + 1: TD(i) = Dir1.List(r)
    Next
    Dir1.Path = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\X VIRTUAL FEMALE\Extra Girl 1"
    For r = 0 To Dir1.ListCount - 1
        i = i + 1: TD(i) = Dir1.List(r)
    Next
    Dir1.Path = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\X VIRTUAL FEMALE\Extra Girl 2"
    For r = 0 To Dir1.ListCount - 1
        i = i + 1: TD(i) = Dir1.List(r)
    Next
   
    
    ReDim Preserve TD(i)
    
    For R4 = 1 To UBound(TD)
        XVCOUNT1 = 0
        XVCOUNT2 = 0
        XVCOUNT3 = 0
        STRING_4 = ""
        STRING_5 = ""
        
        DT4 = DateValue("2008-04-02")
        ' ADD ONE DAY
        DT4 = DT4 + (R4 - 1)
        ' ADJUST UP FOR TIMEZONE AND NOT GET CAUGHT UNDER A DAY
        DT4 = DT4 + TimeSerial(10, 0, 0)
        
    
        File1.Path = TD(R4)
        File1.Pattern = "*.MPG"
        File1.Refresh
    
        ' DON'T RUN MAIN CODE EXTRACTOR IF DONE ONCE SUCCESS
        ' --------------------------------------------------
        If 1 = 2 Then
            
            TT = TD(R4) + "\brain.bin"
            If Dir(TD(R4)) = "" Then
                TT = TD(R4) + "\# brain.bin"
            End If
        
            FR2 = FreeFile
            Open TT For Binary As #FR2
                TTCOM = Input(LOF(FR2), FR2)
            Close #FR2
            
            STRING_1 = ""
            
            TTCOM2 = TTCOM
            TTCOM1 = UCase(TTCOM)
        
            Dim AR
            FF = Replace(TTCOM2, Chr(0), "##;")
            AR = Split(FF, "##;")
            For R2 = 0 To UBound(AR)
                AR(R2) = Replace(AR(R2), Chr(0), "")
                AR(R2) = Replace(AR(R2), Chr(5), "")
                If Len(AR(R2)) = 1 Then AR(R2) = ""
            Next
            
            STRING1 = "EXRACT TALK" + vbCrLf
            STRING1 = STRING1 + "--------" + vbCrLf
            
            For R2 = 0 To UBound(AR)
            If AR(R2) <> "" Then
        '        Debug.Print AR(R2)
                STRING_1 = STRING_1 + AR(R2) + vbCrLf
            End If
            Next
            
            ' MsgBox STRING_1
            
            TH = TD(R4) + "\# WORDIN EXTRACTOR.TXT"
            
            FR2 = FreeFile
            Open TH For Output As #FR2
                Print #FR2, STRING_1
            Close #FR2
            
            TG_1 = TD(R4) + "\brain.bin"
            TG_2 = TD(R4) + "\# brain.bin"
            If Dir(TG_1) <> "" Then
                FSO.CopyFile TG_1, TG_2
                Kill TG_1
            End If
            
            ' ---------------------------
            ' NEXT PART
            ' ---------------------------
            ' DO THE FILENAME
            ' ---------------------------
            
            RXE_2 = 0
            XT = 1
            XR = XT
            Do
                XR = InStr(XR, TTCOM1, ".MPG")
                If XR > 0 Then
                    XZ = InStrRev(TTCOM1, Chr(0), XR - 1)
                    If XZ > 0 Then
                        YY = Mid(TTCOM2, XZ + 1, XR - XZ + 3)
                        TRUE_FOUND = "FALSE"
                        FILE_NAME_CAPTURE_2 = YY
                        
                        HAD_IT = False
                        For RL = 0 To File1.ListCount - 1
                            If InStr(UCase(File1.List(RL)), UCase(YY)) Then
                                FILE_NAME_CAPTURE_1 = File1.Path + "\" + File1.List(RL)
                                FILE_NAME_CAPTURE_2 = YY
                                HAD_IT = True
                                Exit For
                            End If
                        Next
                        
                        If HAD_IT = True Then
                            XVCOUNT1 = XVCOUNT1 + 1
                            XVCOUNT3 = XVCOUNT3 + 1
        '                    Debug.Print Str(XVCOUNT3) + Str(XVCOUNT1) + " -- " + YY
                            STRING2 = STRING2 + Str(XVCOUNT3) + Str(XVCOUNT1) + " -- " + YY + vbCrLf
                            STRING4 = STRING4 + YY + vbCrLf
                            TRUE_FOUND = "TRUE"
                        End If
                        If TRUE_FOUND = "FALSE" Then
                            XVCOUNT2 = XVCOUNT2 + 1
                            XVCOUNT3 = XVCOUNT3 + 1
                            ' Debug.Print Str(XVCOUNT3) + " FILE NOT FOUND " + Str(XVCOUNT2) + " -- " + YY
                            STRING2 = STRING2 + Str(XVCOUNT3) + " FILE NOT FOUND " + Str(XVCOUNT2) + " -- " + YY + vbCrLf
                        End If
                    End If
                End If
                If XR = 0 Then Exit Do
                XR = XR + 1
        
            'MsgBox STRING4
        
            'AR = Split(STRING4, vbCrLf)
            Dim ARX(1)
            
            R2 = 1
            STRING_4 = "RENAME" + vbCrLf
            STRING_4 = STRING_4 + "--------" + vbCrLf
            
            'If TRUE_FOUND <> "TRUE" Then Stop
            If TRUE_FOUND = "TRUE" Then RXE_2 = RXE_2 + 1
            
            ' Debug.Print Str(XVCOUNT2) + " -- " + Str(XVCOUNT3)
            
            ' Debug.Print FILE_NAME_CAPTURE_2
            ' ---------------------------------------
            ARX(R2) = FILE_NAME_CAPTURE_2
            ARX(R2) = UCase(Mid(ARX(R2), 1, 1)) + Mid(ARX(R2), 2)
            ARX(R2) = Replace(ARX(R2), ".mpg", ".MPG")
            ' NOT SURE REQUIRE
            ' ---------------------------------------
            If Mid(ARX(R2), 1, 1) = "0" Then
            ' TAKE OFF TH NUMBER BEFORE WHEN DO RERUN
            ' ---------------------------------------
            If InStrRev(ARX(ARX_2), "__") > 0 Then
                ARX(R2) = Mid(ARX(R2), InStrRev(ARX(R2), "__"))
            End If
            End If
            
            NEW_RENAME = Format(RXE_2 + 1, "0000") + "__" + ARX(R2)
            FILE_RENAME_2 = TD(R4) + "\" + NEW_RENAME
            
            If FILE_NAME_CAPTURE_1 <> "" Then
            If Dir(FILE_RENAME_2) = "" Then
            If Dir(FILE_NAME_CAPTURE_1) <> "" Then
                Name FILE_NAME_CAPTURE_1 As FILE_RENAME_2
            End If
            End If
            End If
            
            STRING_4 = STRING_4 + FILE_RENAME_1 + vbCrLf
            STRING_5 = STRING_5 + FILE_RENAME_2 + vbCrLf
                
            If XR = 0 Then Exit Do
            XR = XR + 1
            Loop Until XR = 0
    
            STRING_4 = STRING_4 + "---------------------------------------" + vbCrLf + STRING_5 + vbCrLf
        
            TH = TD(R4) + "\# FILE OUTPUT.TXT"
            FR2 = FreeFile
            Open TH For Output As #FR2
                Print #FR2, STRING_4
            Close #FR2
    
        End If ' SKIP IF ONLY WANT SET DATE
        ' ---------------------------------
        
        File1.Refresh
    
        'ListView1.ListItems.Clear
    
        LOW_DATE = DT4
        For RL = 0 To File1.ListCount - 1
            Set F = FSO.GetFile(File1.Path + "\" + File1.List(RL))
            DT1 = F.DateLastModified
            If DT1 < LOW_DATE Then LOW_DATE = DT1
        Next
        DT1 = LOW_DATE
        DT1 = DT4
    
        For RL = 0 To File1.ListCount - 1
            FF2 = File1.List(RL)
            TTR = SetFileDateTime(File1.Path + "\" + File1.List(RL), DT1)
            HAS_A_PRE_DIGIT_NUMBER = False
            INX = ""
            If Mid(File1.List(RL), 1, 1) = "0" Then
                HAS_A_PRE_DIGIT_NUMBER = True
                INX = Mid(File1.List(RL), 1, 4)
                FF2 = Mid(File1.List(RL), 7)
            End If
            With LISTVIEW1
                Set LV1 = .ListItems.Add(, , INX)
                LV1.SubItems(1) = FF2
                '------------------------
                'PAIN TO GET THE FORMULAR
                'EVEN THEN FAR OUT AGAIN
                '------------------------
            End With
            
            If HAS_A_PRE_DIGIT_NUMBER = False Then
                RENAME_V = Format(400, "0000") + "__" + File1.List(RL)
                PATH_2 = TD(R4)
                PATH_2 = File1.Path
                FILE_RENAME_1 = File1.Path + "\" + File1.List(RL)
                FILE_RENAME_2 = File1.Path + "\" + Trim(RENAME_V)
                
                If Dir(FILE_RENAME_1) <> "" Then
                    ' Name FILE_RENAME_1 As FILE_RENAME_2
                    ' -----------------------------------
                    ' NOT RENAME NOW DONE ONCE SUCCESSFUL
                    ' -----------------------------------
                End If
            End If
            
        Next

        LISTVIEW1.SortOrder = lvwAscending
        LISTVIEW1.SortKey = 1
        LISTVIEW1.Sorted = True
    Next
End
    
Exit Sub
End

End Sub



'Function GetUserName() As String
'   Dim UserName As String * 255
'   Call GetUserNameA(UserName, 255)
'   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function
'
'Function GetComputerName() As String
'   Dim UserName As String * 255
'   Call GetComputerNameA(UserName, 255)
'   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function


Private Sub Command1_Click()

Clipboard.SetText Text1.Text

End Sub


Private Sub Form_Activate()


    
' -----------------------------------
' For i = 0 To Screen.FontCount - 1
'     A = A + Screen.Fonts(i) + vbCrLf
' Next i
' -----------------------------------
' Clipboard.Clear
' Clipboard.SetText A
' -----------------------------------
' Fri 25-Sep-2020 19:51:41
' -----------------------------------
' ----
' Change Fontstyle Using VB - VB6 | Dream.In.Code
' https://www.dreamincode.net/forums/topic/69063-change-fontstyle-using-vb/
' ----
Dim CONT As Control
On Error Resume Next
For Each CONT In Me.Controls
    
    FONT_SIZE = 12
    Select Case UCase(CONT.Name)
        Case "MMM_SHOW_THE_TIME_ARR": FONT_SIZE = 9
    End Select
    If InStr(UCase(CONT.Name), "TEXT") > 0 Then
        FONT_SIZE = 12
    End If
    If InStr(UCase(CONT.Name), "MMM_SHOW_THE_TIME_ARR") > 0 Then
    If InStr(UCase(CONT.Caption), "000") > 0 Then FONT_SIZE = 13
    If InStr(UCase(CONT.Caption), "TIME SLICE") > 0 Then FONT_SIZE = 8.5
    If InStr(UCase(CONT.Caption), "0008") > 0 Then
        FONT_SIZE = 6.5
        CONT.Caption = "CRC" + vbCr + "SHORTER"
    End If
    If InStr(UCase(CONT.Caption), "0009") > 0 Then
        FONT_SIZE = 6.5
        CONT.Caption = "VERY" + vbCr + "SHORTER"
    End If
    If CONT.Index = 5 Then FONT_SIZE = 9
    If CONT.Index = 6 Then FONT_SIZE = 9.4
    If CONT.Index = 7 Then FONT_SIZE = 9.4
    End If
    FONT_NAME = "SOURCE CODE PRO SEMIBOLD"
    FONT_NAME = "SOURCE CODE PRO BLACK"
    FONT_NAME = "ARIAL"
    FONT_NAME = "ARIAL BLACK"
    CONT.Font.Name = FONT_NAME
    CONT.FontBold = True
    CONT.FontSize = FONT_SIZE
    CONT.Font.Size = FONT_SIZE
    ' TEXT BOX REQUIRE HERE STYLE __ CONT.FONT.Size = FONT_SIZE
Next

    


End Sub

Private Sub Form_Click()
    If IsIDE = True Then
        End
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If IsIDE = True Then
        If KeyCode = 27 Then End
    End If
End Sub

Sub GO_WORD_EDITOR()

Exit Sub

Dim D1, D2

FR1 = FreeFile
Open "C:\DOWNLOADS\WordNavigator-com - Search for -o-.mhtml" For Input As #FR1
FR2 = FreeFile
Open "C:\DOWNLOADS\WordNavigator-com - 2ND LETTER O -- GOOGLE GOGGLE.TXT" For Output As #FR2
Do
    Line Input #FR1, NN
    If InStr(NN, ".com/word/") > 0 Then
        F1 = ".com/word/"
        D1 = InStr(NN, ".com/word/")
        D2 = InStr(D1 + Len(F1) + 1, NN, "/") - Len(F1)
        If D2 <= 0 Then
            D2 = InStr(D1, NN, "=") - Len(F1)
        End If
        D3 = Mid(NN, D1 + Len(F1), D2 - D1)
        If Mid(D3, 2, 1) = "o" Then
            If Mid(D3, Len(D3), 1) <> "s" Then
                If Len(D3) <= 4 Then
                    D3 = UCase(Mid(D3, 1, 1)) + Mid(D3, 2)
                    If Mid(D3, 2, 2) = "oo" Then
                        Print #FR2, D3 + " #"
                    Else
                        D3 = Mid(D3, 1, 1) + Mid(D3, 2, 1) + "o" + Mid(D3, 3)
                        Print #FR2, D3
                    End If
                End If
            End If
        End If
    End If
Loop Until EOF(FR1)
Close #FR1, #FR2

End

End Sub


Private Sub Form_Load()


'Dim AR(5)
'AR(1) = 12
'AR(2) = 24
'AR(3) = 31
'AR(4) = 52
'AR(5) = 91
'
'For R4 = 1 To 5
'For R2 = 1 To 10000 Step 1
'    If MODX Mod 1000 = 0 Then DoEvents
'    MODX = MODX + 1
'    For R1 = 1 To AR(R4)
'        ' A = A + Format(R1 * R2, "0.00000") + " --"
'        A = A + Format(R1 * R2, "0") + " --"
'    Next
'    If InStr(A, "6") > 0 Then A = ""
'    If A <> "" Then Exit For
'    If A <> "" Then A = A + vbCr
'Next
'If A <> "" Then
'    Debug.Print AR(R4) & " -##- " & A
'End If
'Next
'Debug.Print "ENDER"
'Stop
'End
'



'    If App.PrevInstance = True Then End

'    If Command$ <> "" Then
'        MsgBox Command$
'    End If

'    Call GO_VIRTUAL_GIRL
'    Call GO_WORD_EDITOR

    Dim VAR, PID As Long

    Dim START_MINIMIZED

    RUN_ONCE_01 = True
    
    ' ONLY RUN MINIMIAL BY A LAUNCHER NOT TASKBAR
    ' -------------------------------------------
    COMMAND_STRING = Command$
    
'    COMMAND_STRING = """C:\SCRIPTER\NOTEPAD TALK\ZZ TEXT 2018-08-15 __ EMMA D & B _ AOT.txt"" ""\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE\SCRIPTER\NOTEPAD TALK\ZZ TEXT 2018-08-15 __ EMMA D & B _ AOT.txt"""

    ' --------------------------------------
    ' TEST DEBUGGER
    ' --------------------------------------
    ' COMMAND_STRING = "MINIMAL____START_22"
    ' --------------------------------------
    ' HERE IS NOT THE USUAL FO COMMAND LINE THAT HAS PARAM FILE SET
    ' AND WANT MINIMIZED
    ' -------------------------------------------------------------
    If InStr(COMMAND_STRING, "MINIMAL____START_22") > 0 Then
        START_MINIMIZED = True
        COMMAND_STRING = Replace(COMMAND_STRING, "MINIMAL____START_22", "")
    End If
    
    ' DO COME IN WITH PATH FILE STRING COMMONER
    ' BY COMMAND LINE
    ' -----------------------------------------
    If COMMAND_STRING = "" Or InStr(COMMAND_STRING, "MINIMAL____START_22") > 0 Then
        If App.PrevInstance = True Then End
    End If
    
    'KILL ITSELF IN __.EXE KILL SOFTLY
    'WHILE ISIDE LEARN
    '---------------------------------
    If IsIDE = True Then
        A = cProcesses.GetEXEID_KILL_ALL_INSTANCE(App.Path + "\" + App.EXEName + ".exe")
        If A = "FOUND SOMETHING" Then
            Beep
            End
        End If
    End If
    
    
'    MNU_GIVE_TIME_DIFF.Visible = False
    
    ' -------------------------------------------------
    ' IF WANT RENAME FORM FOR DIFFERENT PROJECT
    ' & ALSO KEEP UNIVERSAL FOR SYNC PURPOSE
    ' UNABLE TO RENAME FORM IN ANYWAY AND SYNC
    ' ANSWER CHANGE NAME WITHOUT FRM_ PRECEDE
    ' SO LAND AT TOP OF FORM LISTER NOT IN THE WAY
    ' SOME PROJECT HAVE FRM_ OR FORM_ SORT ORDER
    ' -------------------------------------------------
    ' 03 OF 03
    ' -------------------------------------------------
    ' DETECT PRESENCE OF FORM
    ' REQUEST ANSWER WILL LOAD THE FORM ALSO
    ' -------------------------------------------------
    On Error Resume Next
    Dim frmX As Form, FrmxName_T As String
    Dim FrmxName() As String
    ReDim Preserve FrmxName(10)
    i = 0
    i = i + 1: FrmxName(i) = "Form2_Check_Project_Date"  ' PREVIOUS  NAME USER
    i = i + 1: FrmxName(i) = "Frm_Project_Check_Date"    ' ATTEMPTED NAME USER
    i = i + 1: FrmxName(i) = "Project_Check_Date"        ' STANDARDISE
    ReDim Preserve FrmxName(i)
    For i = 1 To UBound(FrmxName)
        Err.Clear
        Set frmX = Forms.Add(FrmxName(i))
        If Err.Number = 0 Then
            FrmxName_T = FrmxName(i)
            Exit For
        End If
    Next
    ' -------------------------------------------------

    Dim F1, F2
    Dim SAME_VALUE_TRUE_FALSE
    
    EXIT_TRUE = False
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'HOOK_CLIPBOARD_API_lOADED = True
    
    '-------------------------------TEST2017
    'Form1.Mnu_API_Reload_Status.Caption = "The API Form Clipper Logger Sub Call Loaded @ " + Format(Now, "DDD DD-MM-YYYY HH:MM:SS")
    
    '-------------------------------TEST2017
    'If Form1.Mnu_API_UNLoad_Status.Caption = "The API Form Clipper Logger Sub Call UN-Loaded @ " Then
    '    Form1.Mnu_API_UNLoad_Status.Caption = "The API Form Clipper Logger Sub Call UN-Loaded @ NOT HAPPEN YET NOT A TIME SET AT MOMENT"
    'End If
    
    'Me.Hide
    'Start to monitor the clipboard
    
    
    ' ---------------------------------------------------------------
    RUN_AT_FORM_LOAD = True
    Call Label_FindWindowPart_VB_CLIPVIEWER_AUTO_BUTTON_COMPARE_Click
    ' ---------------------------------------------------------------
    
    FRM_CLIPTEST_02_ENABLE = True
    
    COMV1 = COMMAND_STRING
    FILENAME_AND_COMMAND_STRING_GIVE = False
    If COMV1 <> "" Then
        FILENAME_1 = Mid(COMV1, 1, InStr(COMV1, """ """))
        FILENAME_1 = Replace(FILENAME_1, """", "")
        FILENAME_2 = Mid(COMV1, InStr(COMV1, """ """) + 2)
        FILENAME_2 = Replace(FILENAME_2, """", "")
        FILENAME_2 = Replace(FILENAME_2, "file:\\\", "")
        FILENAME_1 = Replace(FILENAME_1, "file:\\\", "")
'        If FSO.FileExists(FILENAME_1) = True Then
'            Me.Show
'            DoEvents
'            MsgBox FILENAME_1
'            Else
'            MsgBox "DONT COMPARE" & FILENAME_1
'        End If
'
        
        If FSO.FileExists(FILENAME_1) = True Or FSO.FileExists(FILENAME_2) = True Then
            ' -------------------------------------------------
            ' DON'T RUN CLIPBOARD API IF TEXT MODE AND FILENAME
            ' OPTION FOR USE LATER BY BUTTON OR MENU
            ' HELP STOP ERROR CHANGE RESULT IF CLIPBOARD COME IN
            ' USE WHEN MULTI APP COMPARE GO
            ' -------------------------------------------------
            FRM_CLIPTEST_02_ENABLE = False
            SHOW_WINDOW_ALWAYS_ON_TOP_AFTER_FORM_LOAD = True
            SHOW_WINDOW_ALWAYS_ON_TOP_AFTER_FORM_LOAD = False
            AT_END_EXIT_CLOSE = True
            FILENAME_AND_COMMAND_STRING_GIVE = True
            Timer_GIVEN_FILENAME_AND_CMD_LINE_STRING_FOR_COMPARE.Enabled = True
            Label_FindWindowPart_VB_CLIPVIEWER_AUTO_BUTTON_COMPARE.BackColor = Label1(4).BackColor
            ' HEX(Label1(4).BackColor) ' &H0000FF00& -- BRIGHT GREEN
            ' HEX(Label1(5).BackColor) ' &HC0FFFF&   -- PALE YELLOW
        End If
    End If
    If FILENAME_AND_COMMAND_STRING_GIVE = False Then
        FILENAME_1 = ""
        FILENAME_2 = ""
        Label_FindWindowPart_VB_CLIPVIEWER_AUTO_BUTTON_COMPARE.BackColor = Label1(5).BackColor
        ' HEX(Label1(4).BackColor) ' &H0000FF00& -- BRIGHT GREEN
        ' HEX(Label1(5).BackColor) ' &HC0FFFF&   -- PALE YELLOW
    End If
    
    Dim FR1
    Dim FILENAME_SAVE
    COMV1 = COMMAND_STRING
    ' ONLY FOR ISIDE = TRUE
    ' ---------------------
    ' SO FILE NAME PICKUP FROM COMMAND LINE AND SAVED
    ' AND THEN NEXT LOAD WILL SHOW FILENAME USER TIME BEFORE
    ' IT NOT REALLY A MULTI TASK ISSUE OF NEXT
    FILENAME_SAVE = App.Path + "\# DATA\" + GetComputerName + "\Data\#COMPARE _" + GetComputerName + "_ 0 COMMAND_LINE.Txt"
    If COMV1 <> "" Then
        If Dir(App.Path + "\# DATA\" + GetComputerName + "\Data", vbDirectory) <> "" Then
            FR1 = FreeFile
            Open FILENAME_SAVE For Output As #FR1
                Print #FR1, COMV1
            Close #FR1
        End If
    End If
    
    ' -----------------------------------------
    ' WHEN IDE DEBUG FOR LASTY FILE SET PAIR
    ' CHECK IS MORE REQUIRE IF FILE DON'T EXIST
    ' -----------------------------------------
    Dim COMV2
    If IsIDE = True And 1 = 2 Then
        If Dir(App.Path + "\# DATA\" + GetComputerName + "\Data", vbDirectory) <> "" Then
            FR1 = FreeFile
            Open FILENAME_SAVE For Input As #FR1
                Line Input #FR1, COMV2
            Close #FR1
        End If
        If COMV2 <> "" Then
            FILENAME_1 = Mid(COMV2, 1, InStr(COMV2, """ """))
            FILENAME_1 = Replace(FILENAME_1, """", "")
            FILENAME_2 = Mid(COMV2, InStr(COMV2, """ """) + 2)
            FILENAME_2 = Replace(FILENAME_2, """", "")
            If FSO.FileExists(FILENAME_1) = False Or FSO.FileExists(FILENAME_2) = False Then
                If COMV1 = "" Then
                    COMV2 = ""
                    FILENAME_1 = ""
                    FILENAME_2 = ""
                End If
            End If
        End If
    End If
    ' END IF -- AMEN
    ' END IF -- AMEN
    ' TEST DEBUG LOAD LAST COMMAND LINE
    ' COMV1 = """C:\TEMP\-7-ASUS-GL522VW\gst_58292_666.tmp"" ""M:\0 00 ART LOGGERS\# APP AND SCREEN AUTO\5-ASUS-P2520LA_AUTO\AUTO_Form_Shot\FormCapture_2016-07-14-Thu_FORM_SWAP\FormCapture- 2016-07-14 23-45-18 Thu.jpg"""
    ' COMV1 = """C:\SCRIPTER\NOTEPAD TALK\TEXT 2018-08-15 __ EMMA D & B _ AOT.txt"" ""C:\TEMP\-7-ASUS-GL522VW\gst_22968_362.tmp"""
    
    
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    
    If FRM_CLIPTEST_02_ENABLE = True Then
        ' ----------------------------------------------------
        ' DONT USER CLIPBOARD WHEN LOAD FILENAME
        ' AS MAKE RESULT COMPARE CHANGE IF GET NEW CLIPBOARDER
        ' AND THAT NOT ACCURATE PRESENTATION OF RESULT DISPLAY
        ' ----------------------------------------------------
        ' ----------------------------------------------------
        Load FRM_ClipTest_02
        FRM_ClipTest_02.ctlClipboard1.StartClipboardViewer
    End If
    
    '-----------------------------
    O_VAL_MINUTE_API = Now
    'Call Form1.MNU_API_RESET_SUB
    '-----------------------------
    
    If Clipboard.GetFormat(vbCFText) = True Then
        Label1(1).BackColor = Label1(5).BackColor
        Label1(7).BackColor = Label1(5).BackColor
    End If
    
    Dim FILENAME2
    FILENAME2 = App.Path + "\Sound_Wav's\Camera1a_2_8kHz.wav"
    If Dir$(FILENAME2) <> "" Then
        ' Set properties needed by MCI to open.
        MMControl1.Notify = True
        MMControl1.Wait = True
        MMControl1.Shareable = False
        MMControl1.DeviceType = "WaveAudio"
        MMControl1.Filename = FILENAME2
        ' Open the MCI WaveAudio device.
        MMControl1.Command = "Open"
    '    MMControl1.Command = "Play"
    End If
    
'    Dim FILENAME4
'    FILENAME4 = App.Path + "\Sound_Wav's--2\Harp run #2.WAV"
'    If Dir$(FILENAME2) <> "" Then
'        ' Set properties needed by MCI to open.
'        MMControl4.Notify = True
'        MMControl4.Wait = True
'        MMControl4.Shareable = False
'        MMControl4.DeviceType = "WaveAudio"
'        ' MMControl4.Filename = FILENAME4
'        ' Open the MCI WaveAudio device.
'        MMControl4.Command = "Open"
'        ' If IsIDE = True Then MMControl4.Command = "Play"
'    End If
    
    
    FILENAME8 = "D:\VB6\VB-NT\00_Best_VB_01\CLIPBOARD_VIEWER\Sound_Wav's--2\AKKORD.WAV"
    FILENAME8 = ""
    If Dir$(FILENAME8) <> "" Then
        ' Set properties needed by MCI to open.
        MMControl8.Notify = True
        MMControl8.Wait = True
        MMControl8.Shareable = False
        MMControl8.DeviceType = "WaveAudio"
        MMControl8.Filename = FILENAME8
        ' Open the MCI WaveAudio device.
        MMControl8.Command = "Open"
    '    MMControl4.Command = "Play"
    End If
    
    
    Call Text2_Change
    
    Call MNU_ADMINISTRATOR_Click
    
    ' D:\VB6\VB-NT\00_Best_VB_01\CLIPBOARD_VIEWER\ClipBoard Viewer.exe
    ' C:/Program Files (x86)/WinMerge/WinMergeU.exe
    
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    
    If FILENAME_1 <> "" Then
        Set F1 = FSO.GetFile(FILENAME_1)
        Set F2 = FSO.GetFile(FILENAME_2)
        Label_SIZE_1.Caption = "SIZE_1     =" + Str(F1.Size)
        Label_SIZE_2.Caption = "SIZE_2     =" + Str(F2.Size)
        FILE_NAME_SIZE_1 = F1.Size
        FILE_NAME_SIZE_2 = F2.Size
    
        CRC_1 = Hex(m_CRC.CalculateFile(FILENAME_1))
        CRC_2 = Hex(m_CRC.CalculateFile(FILENAME_2))
        
        LabelCRC1.Caption = "CRC32_1 = " + CRC_1
        LabelCRC2.Caption = "CRC32_2 = " + CRC_2
        LabelFILE1.Caption = "F1 = " + FILENAME_1
        LabelFILE2.Caption = "F2 = " + FILENAME_2
        ' LabelFILE1.FontSize = 12
        ' LabelFILE2.FontSize = LabelFILE1.FontSize
        
        If CRC_1 = CRC_2 And F1.Size = F2.Size Then
            SAME_VALUE_TRUE_FALSE = "COMPARE IS TRUE"
            Label1(3).BackColor = Label1(5).BackColor
            Label_SIZE_1.BackColor = Label1(5).BackColor
            Label_SIZE_2.BackColor = Label1(5).BackColor
            LabelCRC1.BackColor = Label1(4).BackColor
            LabelCRC2.BackColor = Label1(4).BackColor
        Else
            SAME_VALUE_TRUE_FALSE = "COMPARE IS FALSE"
            Label1(3).BackColor = Label1(4).BackColor
            Label_SIZE_1.BackColor = Label_COLOR_RED.BackColor
            Label_SIZE_2.BackColor = Label_COLOR_RED.BackColor
            LabelCRC1.BackColor = Label_COLOR_RED.BackColor
            LabelCRC2.BackColor = Label_COLOR_RED.BackColor
        End If
        
        Set F1 = Nothing
        Set F2 = Nothing
    End If
    
    
    Set F = FSO.GetFile(App.Path + "\" + App.EXEName + ".exe")
    ADATE1 = F.DateLastModified
    
    MOTION_VERSION = Format(ADATE1, "YYYY_MMM_DD_DDD__HH:MM:00")
    
    MNU_VERSION.Caption = "Ver_2020_" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) + "  " + MOTION_VERSION ' + " _ Matt.Lan@btinternet.com"
    
    If FILENAME_1 = "" Then
        Label_SIZE_1.Caption = "SIZE =" + Str(Len(Text2.Text)) + " -- " + SAME_VALUE_TRUE_FALSE
        Label_SIZE_2.Caption = "SIZE =" + Str(Len(Text3.Text)) + " -- " + SAME_VALUE_TRUE_FALSE
    End If
    
    If FILENAME_1 <> "" Then
        Label_SIZE_1.Caption = "SIZE =" + Str(FILE_NAME_SIZE_1) + " -- " + SAME_VALUE_TRUE_FALSE
        Label_SIZE_2.Caption = "SIZE =" + Str(FILE_NAME_SIZE_2) + " -- " + SAME_VALUE_TRUE_FALSE
    End If
    
    If FILENAME_1 <> "" Then
        Call SET_TEXT_2_AND_TEXT_COLOUR_AND_TEXT_IF_SAME_WHEN_FILE
    Else
        Call SET_TEXT_2_AND_TEXT_COLOUR_AND_TEXT_IF_SAME
    End If
    
    LabelCRC5.Caption = "DON'T MINIMIZE CLIPPER" + vbCrLf + "TIME && COMBINE"
    
    Call ME_TOP_LEFT_CENTER
    
    If START_MINIMIZED = True Then
        If Me.WindowState <> vbMinimized Then
            Me.WindowState = vbMinimized
        End If
    End If
    
    If FILENAME_1 <> "" Then
        ' START_MINIMIZED = True ONLY IF MORE THAN 2 INSTANCE
        ' SEE AS MOST TIME ONE IS ALWAYS RUN
        ' ---------------------------------------------------
        VAR = cProcesses.GetEXEID_COUNT(App.Path + "\" + App.EXEName + ".exe")
    End If
    
    If FindWindowPart_VB_CLIPVIEWER_AUTO_BUTTON_COMPARE > 2 Then
    End If
    
    If FILENAME_1 <> "" Then
        If VAR > 2 Then
            Me.WindowState = vbMinimized
        End If
    End If
    
    ' Call MNU_TIME_DIFFERENCE_CODE_Click
    
    WAIT_FORM_BOOT_BEFORE_GO = True
    
    MMM_SHOW_THE_TIME_ARR(4).BackColor = Label1(4).BackColor
    ' MMM_SHOW_THE_TIME_ARR(4).FontSize = 10
    MMM_SHOW_THE_TIME_ARR(4).Caption = "TIME SLICE"
    MMCONTROL4_FIRST_LOAD_SOUND_OFF = True
    ' MMM_SHOW_THE_TIME_ARR(5).FontSize = 10
    ' MMM_SHOW_THE_TIME_ARR(6).FontSize = 10
    MMM_SHOW_THE_TIME_ARR(6).Caption = "COOP £40"
    ' MMM_SHOW_THE_TIME_ARR(7).FontSize = 10
    MMM_SHOW_THE_TIME_ARR(7).Caption = "COOP £20"
    
    LINE_PICKER_COMMON.LINE_PICKER_COMMON_FROM_START = True
    Load LINE_PICKER_COMMON
    
    
'    If IsIDE = False Then
'        MsgBox "H"
'    End If
    Call frmListMenu.SET_MENU_PADD_WORK
    
    
    
End Sub

Sub FONT_INFO()

'System
'8514 oem
'Fixedsys
'Terminal
'Modern
'Roman
'Script
'Courier
'MS Serif
'MS Sans Serif
'Small Fonts
'TeamViewer15
'Marlett
'Arial
'Arabic TRANSPARENT
'Arial Baltic
'Arial CE
'Arial CYR
'Arial Greek
'Arial TUR
'Arial Black
'Bahnschrift Light
'Bahnschrift Semilight
'Bahnschrift
'Bahnschrift Semibold
'Calibri
'Calibri Light
'Cambria
'Cambria Math
'Candara
'Comic Sans MS
'Consolas
'Constantia
'Corbel
'Courier New
'Courier New Baltic
'Courier New CE
'Courier New CYR
'Courier New Greek
'Courier New TUR
'Ebrima
'Franklin Gothic Medium
'Gabriola
'Gadugi
'Georgia
'Impact
'Javanese Text
'Leelawadee UI
'Leelawadee UI Semilight
'Lucida Console
'Lucida Sans Unicode
'Malgun Gothic
'@Malgun Gothic
'Malgun Gothic Semilight
'@Malgun Gothic Semilight
'Microsoft Himalaya
'Microsoft JhengHei
'@Microsoft JhengHei
'Microsoft JhengHei UI
'@Microsoft JhengHei UI
'Microsoft JhengHei Light
'@Microsoft JhengHei Light
'Microsoft JhengHei UI Light
'@Microsoft JhengHei UI Light
'Microsoft New Tai Lue
'Microsoft PhagsPa
'Microsoft Sans Serif
'Microsoft Tai Le
'Microsoft YaHei
'@Microsoft YaHei
'Microsoft YaHei UI
'@Microsoft YaHei UI
'Microsoft YaHei Light
'@Microsoft YaHei Light
'Microsoft YaHei UI Light
'@Microsoft YaHei UI Light
'Microsoft Yi Baiti
'MingLiU -ExtB
'@MingLiU-ExtB
'PMingLiU -ExtB
'@PMingLiU-ExtB
'MingLiU_HKSCS -ExtB
'@MingLiU_HKSCS-ExtB
'Mongolian Baiti
'MS Gothic
'@MS Gothic
'MS UI Gothic
'@MS UI Gothic
'MS PGothic
'@MS PGothic
'MV Boli
'Myanmar Text
'Nirmala UI
'Nirmala UI Semilight
'Palatino Linotype
'Segoe MDL2 Assets
'Segoe Print
'Segoe Script
'Segoe UI
'Segoe UI Black
'Segoe UI Emoji
'Segoe UI Historic
'Segoe UI Light
'Segoe UI Semibold
'Segoe UI Semilight
'Segoe UI Symbol
'SimSun
'@SimSun
'NSimSun
'@NSimSun
'SimSun -ExtB
'@SimSun-ExtB
'Sitka Small
'Sitka Text
'Sitka Subheading
'Sitka Heading
'Sitka Display
'Sitka Banner
'Sylfaen
'Symbol
'Tahoma
'Times New Roman
'Times New Roman Baltic
'Times New Roman CE
'Times New Roman CYR
'Times New Roman Greek
'Times New Roman TUR
'Trebuchet MS
'Verdana
'Webdings
'Wingdings
'Yu Gothic
'@Yu Gothic
'Yu Gothic UI
'@Yu Gothic UI
'Yu Gothic UI Semibold
'@Yu Gothic UI Semibold
'Yu Gothic Light
'@Yu Gothic Light
'Yu Gothic UI Light
'@Yu Gothic UI Light
'Yu Gothic Medium
'@Yu Gothic Medium
'Yu Gothic UI Semilight
'@Yu Gothic UI Semilight
'HoloLens MDL2 Assets
'Andale Mono IPA
'Proxima Nova Rg
'Proxima Nova Lt
'MT Extra
'Nina
'Segoe Condensed
'Haettenschweiler
'MS Outlook
'Book Antiqua
'Century Gothic
'Bookshelf Symbol 7
'MS Reference Sans Serif
'MS Reference Specialty
'Bradley Hand ITC
'Freestyle Script
'French Script MT
'Juice ITC
'Kristen ITC
'Lucida Handwriting
'Mistral
'Papyrus
'Pristina
'Tempus Sans ITC
'Arial Narrow
'Garamond
'Monotype Corsiva
'Agency FB
'Arial Rounded MT Bold
'Blackadder ITC
'Bodoni MT
'Bodoni MT Black
'Bodoni MT Condensed
'Bookman Old Style
'Calisto MT
'Castellar
'Century Schoolbook
'Copperplate Gothic Bold
'Copperplate Gothic Light
'Curlz MT
'Edwardian Script ITC
'Elephant
'Engravers MT
'Eras Bold ITC
'Eras Demi ITC
'Eras Light ITC
'Eras Medium ITC
'Felix Titling
'Forte
'Franklin Gothic Book
'Franklin Gothic Demi
'Franklin Gothic Demi Cond
'Franklin Gothic Heavy
'Franklin Gothic Medium Cond
'Gigi
'Gill Sans MT
'Gill Sans MT Condensed
'Gill Sans Ultra Bold
'Gill Sans Ultra Bold Condensed
'Gill Sans MT Ext Condensed Bold
'Gloucester MT Extra Condensed
'Goudy Old Style
'Goudy Stout
'Imprint MT Shadow
'Lucida Sans
'Lucida Sans Typewriter
'Maiandra GD
'OCR A Extended
'Palace Script MT
'Perpetua
'Perpetua Titling MT
'Rage Italic
'Rockwell
'Rockwell Condensed
'Rockwell Extra Bold
'Script MT Bold
'Tw Cen MT
'Tw Cen MT Condensed
'Tw Cen MT Condensed Extra Bold
'Algerian
'Baskerville Old Face
'Bauhaus 93
'Bell MT
'Berlin Sans FB
'Berlin Sans FB Demi
'Bernard MT Condensed
'Bodoni MT Poster Compressed
'Britannic Bold
'Broadway
'Brush Script MT
'Californian FB
'Centaur
'Chiller
'Colonna MT
'Cooper Black
'Footlight MT Light
'Harlow Solid Italic
'Harrington
'High Tower Text
'Jokerman
'Kunstler Script
'Lucida Bright
'Lucida Calligraphy
'Lucida Fax
'Magneto
'Matura MT Script Capitals
'Modern No. 20
'Niagara Engraved
'Niagara Solid
'Old English Text MT
'Onyx
'Parchment
'Playbill
'Poor Richard
'Ravie
'Informal Roman
'Showcard Gothic
'Snap ITC
'Stencil
'Viner Hand ITC
'Vivaldi
'Vladimir Script
'Wide Latin
'Century
'Wingdings 2
'Wingdings 3
'Arial Unicode MS
'@Arial Unicode MS
'MS Mincho
'@MS Mincho
'Source Code Pro
'Source Code Pro Medium
'Source Code Pro Semibold
'Source Code Pro Black

End Sub



'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'UnloadMode = 3
'End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Form As Form

On Error Resume Next
For Each Form In Forms
    Err.Clear
    If Form.EXIT_TRUE = True Then
        If Err.Number = 0 Then
            EXIT_TRUE = True
            Exit For
        End If
    End If
Next Form
For Each Form In Forms
    Form.EXIT_TRUE = EXIT_TRUE
Next Form


If AT_END_EXIT_CLOSE = True Then Me.EXIT_TRUE = True
' CHECK RIPER DON'T WANT TO EXIT
' ------------------------------
If AT_END_EXIT_CLOSE = False Then
If IsIDE = False Then
    If Me.WindowState <> vbMinimized And Me.EXIT_TRUE = False Then
        Me.WindowState = vbMinimized
        End
        Cancel = True
        Exit Sub
    End If
End If
End If

Dim i, Control

'SET ALL TIMERS IN ALL FORMS ENABLED TO =FALSE
On Error Resume Next
    For i = 0 To Forms.Count - 1
        For Each Control In Forms(i).Controls
            If InStr(UCase(Control.Name), "TIMER") > 0 Then
                'Debug.Print Control.Name
                Control.Enabled = False
            End If
        Next
    Next i
On Error GoTo 0

If AT_END_EXIT_CLOSE = True Then Me.EXIT_TRUE = True
If Me.EXIT_TRUE = True Then Cancel = False

' DELETE THE TEMP FILE FOR MULTI COMPARE IS ONLY 1 INSTANCE
' ---------------------------------------------------------
Call Label_FindWindowPart_VB_CLIPVIEWER_AUTO_COMP_2

For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form



End


End Sub


Private Sub Lab_TIME_CLIPBOARD_CHANGE_Click()
' -------------------------------------------------------
' ALLOW API CLIPBOARDER IF REQUEST MADE BY BUTTON OR MENU
' WHEN NOT ENABLE BEFORE AS TEXT MODE WILL BE ON
' -------------------------------------------------------
' A BUTTON ON THE FORM LABEL
' -------------------------------------------------------
FRM_CLIPTEST_02_ENABLE = True

' Call MNU_API_RESET_Click

FRM_ClipTest_02.ctlClipboard1.EndClipboardViewer
'Unload FRM_ClipTest_02
DoEvents
'Load FRM_ClipTest_02
FRM_ClipTest_02.ctlClipboard1.StartClipboardViewer

End Sub


Private Sub Label1_Click(Index As Integer)
If Index = 1 Then
    If Clipboard.GetFormat(vbCFText) = True Then
        Text3.Text = Text2.Text
        On Error Resume Next
        Do
            Err.Clear
            Text2.Text = Clipboard.GetText
            If Err.Number > 0 Then Sleep 200
        Loop Until Err.Number = 0
        On Error GoTo 0
        
    End If
    Exit Sub
End If

If Index = 7 Then
    Call MNU_FIND_DIFFER_Click
    Exit Sub
End If

If Index <> 3 Then Exit Sub

'Label1
'Call MNU_COMPARE_Click

    Dim FILENAME_COMPARE_1
    Dim FILENAME_COMPARE_2
    Dim COMPARE_EXE_1, COMPARE_EXE_2
    Dim FR1
    Dim FN_VAR_1
    Dim FN_VAR_2
    
    Beep
    
    If COMV1 <> "" Then
        Me.WindowState = vbMinimized
    End If
    
    If COMV1 = "" Then
    
        FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\"
        If Dir(FILENAME_IN_USE_CHECK, vbDirectory) = "" Then
            CreateFolderTree FILENAME_IN_USE_CHECK
        End If
        
        
        FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\#COMPARE _" + GetComputerName + "_ 1 DUPE CHECKER.Txt"
        'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
        FILENAME_COMPARE_1 = FILENAME_IN_USE_CHECK
        FR1 = FreeFile
        On Error Resume Next
        Open FILENAME_IN_USE_CHECK For Output As #FR1
            Print #FR1, Text2.Text;
        Close #FR1
    
        FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\#COMPARE _" + GetComputerName + "_ 2 DUPE CHECKER.Txt"
        'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
        FILENAME_COMPARE_2 = FILENAME_IN_USE_CHECK
        FR1 = FreeFile
        On Error Resume Next
        Open FILENAME_IN_USE_CHECK For Output As #FR1
            Print #FR1, Text3.Text;
        Close #FR1

    End If

    If COMV1 <> "" Then
        FILENAME_COMPARE_1 = FILENAME_1
        FILENAME_COMPARE_2 = FILENAME_2
    End If

    Beep
'    Shell "C:\Program Files (x86)\WinMerge\WinMergeU.exe" + " """ + FILENAME_COMPARE_1 + """ """ + FILENAME_COMPARE_2 + """"
    ' Beep
    
'   Const DontWaitUntilFinished = False, WaitUntilFinished = True, ShowWindow_2 = 1, DontShowWindow = 0
        
    FN_VAR_1 = "D:\VB6\VB-NT\00_Best_VB_01\CLIPBOARD_VIEWER\ClipBoard Viewer.exe"
    FN_VAR_1 = "C:\Program Files (x86)\WinMerge\WinMergeU.exe"
    
    FN_VAR_2 = " """ + FILENAME_COMPARE_1 + """ """ + FILENAME_COMPARE_2 + """"
    Me.WindowState = vbMinimized

    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.RUN """" + FN_VAR_1 + """" + " " + FN_VAR_2, ShowWindow_2, WaitUntilFinished
    Set WSHShell = Nothing

'    FN_VAR_1 = FN_VAR_1 + FN_VAR_2

'    ExecCmdWAIT (FN_VAR_1)
    
    Me.WindowState = vbMinimized

    Sleep 2000


    Me.WindowState = vbMinimized
    
    If COMV1 <> "" Then
        Unload Me
'        End
        A1_VarsMod.Terminate_Process

    End If


End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
    CODER_EXE_FILE_NAME_1 = i
    If InStr(i, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 64 BIT.vbp", ".vbp")
    If InStr(i, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 32 BIT.vbp", ".vbp")
    
    CODER_EXE_FILE_NAME_1 = Replace(CODER_EXE_FILE_NAME_1, ".vbp", ".exe")
    
    FileSpec = CODER_VBP_FILE_NAME_2
    If Dir(FileSpec) <> "" Then
        FILE_SPEC_GO = FileSpec
    Else
        If LCase(GetSpecialfolder(&H29)) = LCase("C:\Windows\SysWOW64") Then
            FileSpec = Replace(FileSpec, ".vbp", " - 64 BIT.vbp")
        Else
            FileSpec = Replace(FileSpec, ".vbp", " - 32 BIT.vbp")
        End If
    End If
    
    If Dir(FileSpec) <> "" Then FILE_SPEC_GO = FileSpec
    If Dir(FileSpec) = "" Then FILE_SPEC_GO = CODER_VBP_FILE_NAME_2
    
    If Dir(FILE_SPEC_GO) = "" Then MsgBox "ERROR NOT EXISTOR _ " + FILE_SPEC_GO
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub

Private Sub Label2_Click()
Call LabelCRC3_Click
End Sub

Private Sub Label3_Click()
GET_CRC32_IN_LOWER_CASE_FOR_YOUTUBE = True
End Sub

Private Sub Label4_Click()
    Load LINE_PICKER_COMMON
    RESIZE_DONE_ONCE = False
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
    RESIZE_WINDOWSTATE_CHANGE = -200
    FORM_SET_WIDTH_ONE = False
    LINE_PICKER_COMMON.Visible = True
    ' LINE_PICKER_COMMON.WindowState = vbMaximized
    LINE_PICKER_COMMON.WindowState = Normal
    FRM_ClipTest.WindowState = vbMinimized
    
    LINE_PICKER_COMMON.SetFocus
    
    ' ----------------------------------------------
    ' LINE_PICKER_COMMON_LIST_SET
    ' ----------------------------------------------
    
    ' ----------------------------------------------
    ' HERE REQUIRE SOMETIME CERTAIN STUFF
    ' WHEN HERE IS AND THE WINDOW NOT DROP OF PICKER
    ' Thu 24-Sep-2020 13:38:00
    ' EVEN WHEN SET TO GO THAT
    ' WHEN FORM LOST FOCUS ERROR WINDOW NOT THERE
    ' USER CLOSE AND FINE QUICKER -- NOT UNLOAD
    ' ----------------------------------------------
    ' LINE_PICKER_COMMON.MNU_ALWAYS_ON_TOP_Click
    ' ----------------------------------------------
    
End Sub


Private Sub MNU_ADBLOCK_FOLDER_Click()
    FOLDER_NAME_____ = "C:\SCRIPTER\SCRIPT 08 IN_NET_ ADBLOCK"
    Shell "EXPLORER """ + FOLDER_NAME_____ + """", vbMaximizedFocus
    Me.WindowState = vbMinimized
End Sub

Private Sub MNU_ALWAYS_ON_TOP_Click()
    
    ' MENU FLIPPER
    If InStr(MNU_ALWAYS_ON_TOP.Caption, "ON TOP __ CORRECT") > 0 Then
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __]"
        AlwaysOnTop_MODE = False
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    Else
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __ CORRECT __]"
        AlwaysOnTop_MODE = True
        RESULT_API = AlwaysOnTop(Me.hWnd)
    End If
    Me.SetFocus

End Sub

Private Sub MNU_ALWAYS_ON_TOP_SET()
    
    ' SET CORRECT AND RESULT REQUIRE THAT
    '
    ' DO WHAT MENU GOT
    '
    If InStr(MNU_ALWAYS_ON_TOP.Caption, "ON TOP __ CORRECT") > 0 Then
        AlwaysOnTop_MODE = True
        RESULT_API = AlwaysOnTop(Me.hWnd)
    Else
        AlwaysOnTop_MODE = False
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    End If

End Sub

Private Sub MNU_ALWAYS_ON_TOP_SET_TRUE()
    
    ' FORCE ON
    MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __ CORRECT __]"
    AlwaysOnTop_MODE = True
    RESULT_API = AlwaysOnTop(Me.hWnd)

End Sub



Private Sub MNU_CLIPBOARD_ADBLOCK_Click()

If ADBLOCK_INFO_MAIN_CHUNK = "" Then
    Call MNU_PROCESS_ADBLOCK_FILE_Click
Else
    Call CLIPBOARD_PROCESS_ADBLOCK_FILE
End If

Me.WindowState = vbMinimized

End Sub

Private Sub MNU_GIVE_TIME_GITHUB_Click()

Dim XI
Dim A_DATE_TIME_UTC
Dim A_DATE_TIME As String

'2026-Apr-29 -- 12:37:10 - 12:37:10-PM - Wed
A_DATE_TIME = Format(Now, "YYYY-MMM-DD") + " -- " + Format(Now, "HH:MM:SS") + " - " + Format(Now, "HH:MM:SS AM/PM") + " - " + Format(Now, "DDD")
A_DATE_TIME = Replace(A_DATE_TIME, " PM", "-PM")
A_DATE_TIME = Replace(A_DATE_TIME, " AM", "-AM")

'A_DATE_TIME = Replace(A_DATE_TIME, "PM", "Pm")
'A_DATE_TIME = Replace(A_DATE_TIME, "AM", "Am")

Clipboard.Clear
Clipboard.SetText A_DATE_TIME

Beep
If IsIDE = False Then Me.WindowState = vbMinimized

End Sub

Private Sub MNU_GOODSYN_CCONVERT_SMBD_PATH_TO_DOS_Click()

' smbd://4-ASUS-GL522VW/4_ASUS_GL522VW_40_4_SAMSUNG_5TB/DSC_4G_1TB/2015+SONY_MP4/2020 CyberShot HX60V __ MP4smbd://4-ASUS-GL522VW/4_ASUS_GL522VW_40_4_SAMSUNG_5TB/DSC_4G_1TB/2015+SONY_MP4/2020 CyberShot HX60V __ MP4smbd://4-ASUS-GL522VW/4_ASUS_GL522VW_40_4_SAMSUNG_5TB/DSC_4G_1TB/2015+SONY_MP4/2020 CyberShot HX60V __ MP4smbd://4-ASUS-GL522VW/4_ASUS_GL522VW_40_4_SAMSUNG_5TB/DSC_4G_1TB/2015+SONY_MP4/2020 CyberShot HX60V __ MP4
' TO EXAMPLE
' \\4-asus-gl522vw\4_asus_gl522vw_02_d_drive
On Error Resume Next

Dim VAT_ST_7
Dim VAT_ST_8

VAR_ST_7 = Clipboard.GetText
VAR_ST_8 = VAR_ST_7
VAR_ST_7 = Replace(VAR_ST_7, "smbd://", "\\")
VAR_ST_7 = Replace(VAR_ST_7, "/", "\")
'A_DATE_TIME = Replace(A_DATE_TIME, "PM", "Pm")
'A_DATE_TIME = Replace(A_DATE_TIME, "AM", "Am")

Clipboard.Clear
Clipboard.SetText VAR_ST_7

If Err.Number > 0 Then MsgBox "ERROR SOURCE PATH = " & vbCrLf & VAT_ST_8 & vbCrLf & vbCrLf & "RESULT PATH = " & vbCrLf & VAT_ST_7

Beep
If IsIDE = False Then Me.WindowState = vbMinimized

End Sub

Private Sub MNU_NOTEPAD_CLIPPPER_Click()

    FILENAME_LOAD_01 = App.Path + "\" + "NOTPAD__CLIPPER_#NFS_EX.TXT"
    FR1 = FreeFile
    Open FILENAME_LOAD_01 For Output As #FR1
        Print #FR1, Text2.Text;
    Close #FR1
    APP_NAME = "C:\PROGRAM FILES (X86)\NOTEPAD++\NOTEPAD++.EXE"
    If FSO.FileExists(APP_NAME) = True Then
        Shell APP_NAME + " """ + FILENAME_LOAD_01 + """", vbMaximizedFocus
    End If

End Sub

Private Sub MNU_PROCESS_ADBLOCK_FILE_Click()

Dim FILENAME_LOAD_01 As String
Dim FILENAME_LOAD_02 As String

' ---------------------------------------------------------------------
' LOAD FILE INTO ARRAY 02 OF 02 BEGIN
' ---------------------------------------------------------------------
FILENAME_NAME_01 = "TEXT_ADBLOCK_TEXT__01__MAIN.TXT"
FILENAME_NAME_01 = "TEXT_ADBLOCK_TEXT__01__MAIN.TXT"
FILENAME_NAME_01 = "TEXT_ADBLOCK_01_MAIN.TXT"
FILENAME_NAME_02 = "TEXT_ADBLOCK_TEXT__02__MERGE_COMBINE.TXT"
FILENAME_NAME_02 = "TEXT_ADBLOCK_TEXT__02__MERGE_COMBINE.TXT"
FILENAME_NAME_02 = "TEXT_ADBLOCK_02_MERGE_COMBINE.TXT"
FILENAME_LOAD_01 = "C:\SCRIPTER\SCRIPT 08 IN_NET_ ADBLOCK\" + FILENAME_NAME_01
FILENAME_LOAD_02 = "C:\SCRIPTER\SCRIPT 08 IN_NET_ ADBLOCK\" + FILENAME_NAME_02

' LATER HERE ARCHIVE VAR IS EDIT AGAIN
' FORMAT_TIME = Format(Now, "YYYY-MM-DD  HH-MM-SS")
' FILENAME_LOAD_04 = "C:\SCRIPT

FOLDER_NAME_____ = "C:\SCRIPTER\SCRIPT 08 IN_NET_ ADBLOCK"
If Dir(FILENAME_LOAD_01) = "" Then
    MsgBox "ERROR NOT FILE NAME" + vbCrLf + vbCrLf + FILENAME_LOAD_01, vbMsgBoxSetForeground
    Exit Sub
End If
If Dir(FILENAME_LOAD_02) = "" Then
    MsgBox "ERROR NOT FILE NAME" + vbCrLf + vbCrLf + FILENAME_LOAD_02, vbMsgBoxSetForeground
    Exit Sub
End If

Dim STRARRAY1
Dim STRARRAY2


' ------------------------------------------
' GET A CRC32 TO COMPARE WHEN DONE
' ------------------------------------------
' 1ST THING WRITE COPY FILE TO ARCHIVE BACKUP
' ------------------------------------------------
' LINE NUMBER 0002
CRC_ADBLOCK_01 = Hex(m_CRC.CalculateFile(FILENAME_LOAD_01))
FORMAT_TIME = Format(Now, "YYYY-MM-DD  HH-MM-SS")
FORMAT_TIME = Format(Now, "YYYY-MM-DD_HH-MM")
FILENAME_LOAD_04 = "C:\SCRIPTER\SCRIPT 08 IN_NET_ ADBLOCK\TEXT_ADBLOCK_TEXT__ARCHIVE__" + FORMAT_TIME + "__CRC32_" + CRC_ADBLOCK_01 + ".TXT"

FILENAME_LOAD_04 = Replace(FILENAME_LOAD_04, "TEXT_ADBLOCK_TEXT__ARCHIVE__", "TEXT_ADBLOCK_ARC_")
FILENAME_LOAD_04 = Replace(FILENAME_LOAD_04, "__", "_")
FILENAME_LOAD_04 = Replace(FILENAME_LOAD_04, "CRC32", "CRC")
FSO.CopyFile FILENAME_LOAD_01, FILENAME_LOAD_04, True
' ------------------------------------------------



' GET CHUNK __ FILENAME_LOAD_01 __ PUT ARRAY
' ------------------------------------------
Dim VALUE_CHUNK As String
FR1 = FreeFile
Open FILENAME_LOAD_01 For Binary As #FR1
    VALUE_CHUNK = Space$(LOF(FR1))
    Get #FR1, 1, VALUE_CHUNK
Close #FR1
VALUE_CHUNK = VALUE_CHUNK + vbCr
STRARRAY1 = Split(VALUE_CHUNK, vbCr)

' GET CHUNK __ FILENAME_LOAD_02 __ PUT ARRAY
FR1 = FreeFile
Open FILENAME_LOAD_02 For Binary As #FR1
    VALUE_CHUNK = Space$(LOF(FR1))
    Get #FR1, 1, VALUE_CHUNK
Close #FR1
VALUE_CHUNK = VALUE_CHUNK + vbCr
STRARRAY2 = Split(VALUE_CHUNK, vbCr)

' STRIP VBCR AND VBLF OFF
For R_COUNT = 0 To UBound(STRARRAY1)
    STRARRAY1(R_COUNT) = Replace(STRARRAY1(R_COUNT), vbCr, "")
    STRARRAY1(R_COUNT) = Replace(STRARRAY1(R_COUNT), vbLf, "")
Next
For R_COUNT = 0 To UBound(STRARRAY2)
    STRARRAY2(R_COUNT) = Replace(STRARRAY2(R_COUNT), vbCr, "")
    STRARRAY2(R_COUNT) = Replace(STRARRAY2(R_COUNT), vbLf, "")
Next

' ------------------------------------------
' WORK STRARRAY1 FOR MOMENT
' ------------------------------------------
' WORK WAY FROM END LIKE CROP HEADER FOOTER
' ------------------------------------------

Dim M_ARRAY()
ReDim M_ARRAY(2000)
' REBUILD STRING ARRAY STRARRAY1 REMOVE NOTHING END AND BEGINNER -- EG HEADER FOOTER KEEP
COUNT_ARRAY = 0
For R_COUNT = 0 To UBound(STRARRAY1)
    If Trim(STRARRAY1(R_COUNT)) <> "" Then
        COUNT_ARRAY = COUNT_ARRAY + 1
        M_ARRAY(COUNT_ARRAY) = STRARRAY1(R_COUNT)
    End If
Next
ReDim Preserve M_ARRAY(COUNT_ARRAY)


' ----------------------------------------------
' GET HEADER AND FOOTER
' COUNT BEGIN FIND REM FROM START __ 01 OF 02
BLOCK_REMLINE = 0
For R_COUNT = 0 To UBound(M_ARRAY)
    If M_ARRAY(R_COUNT) <> "" Then
        If Mid(M_ARRAY(R_COUNT), 1, 1) <> "!" Then Exit For
        POS_START_REMMER = R_COUNT
    End If
Next
' COUNT BEGIN FIND REM FROM START __ 02 OF 02
For R_COUNT = 0 To POS_START_REMMER
    If M_ARRAY(R_COUNT) <> "" Then
        STARTER_BLOCK = STARTER_BLOCK + M_ARRAY(R_COUNT) + vbCrLf
    End If
Next
' COUNT BACK FIND REM FROM ENDER  __ 01 OF 02
For R_COUNT = UBound(M_ARRAY) To 0 Step -1
    If M_ARRAY(R_COUNT) <> "" Then
        If Mid(M_ARRAY(R_COUNT), 1, 1) <> "!" Then Exit For
        POS_ENDER_REMMER = R_COUNT
    End If
Next
' COUNT BACK FIND REM FROM ENDER  __ 02 OF 02
For R_COUNT = POS_ENDER_REMMER To UBound(M_ARRAY)
    If Trim(M_ARRAY(R_COUNT)) <> "" Then
        ENDER_BLOCK = ENDER_BLOCK + M_ARRAY(R_COUNT) + vbCrLf
    End If
Next
' ----------------------------------------------

' ----------------------------------------------
' REMOVE DOUBLE LINE FEED AND TRIM
' ----------------------------------------------
ENDER_BLOCK = Replace(ENDER_BLOCK, vbCrLf + vbCrLf, vbCrLf)
STARTER_BLOCK = Replace(STARTER_BLOCK, vbCrLf + vbCrLf, vbCrLf)
ENDER_BLOCK = Trim(ENDER_BLOCK)
STARTER_BLOCK = Trim(STARTER_BLOCK)
' ----------------------------------------------
'MsgBox STARTER_BLOCK, vbMsgBoxSetForeground
'MsgBox ENDER_BLOCK, vbMsgBoxSetForeground
' ----------------------------------------------


' ------------------------------------------------
' LOAD ARRAY LIST BOX FOR SORT -- NOT THE REM LINE
' 01 OF 02 ---------------------------------------
' M_ARRAY() -- HAVE -- STRARRAY1()
List1.Clear
i = 0
For R_COUNT = 0 To UBound(M_ARRAY)
    If Trim(M_ARRAY(R_COUNT)) <> "" Then
    If Mid(M_ARRAY(R_COUNT), 1, 1) <> "!" Then
        List1.AddItem M_ARRAY(R_COUNT)
        i = i + 1
    End If
    End If
Next
TOTAL_COUNT_LINE_AD_BLOCKER = i
' ------------------------------------------------
' ------------------------------------------------
' LOAD ARRAY LIST BOX FOR SORT -- NOT THE REM LINE
' 02 OF 02 ---------------------------------------
' STRARRAY2()
' ------------------------------------------------
i = 0
For R_COUNT = 0 To UBound(STRARRAY2)
    GET_INFO = Trim(STRARRAY2(R_COUNT))
    If GET_INFO <> "" Then
    If Mid(GET_INFO, 1, 1) <> "!" Then
        List1.AddItem GET_INFO
        i = i + 1
    End If
    End If
Next
TOTAL_COUNT_LINE_MERGE_AD_BLOCKER = i

' ------------------------------------------------
' SORT DO AND REMOVE DUPLICATE
' ------------------------------------------------
For r = List1.ListCount - 1 To 1 Step -1
    If List1.List(r) = List1.List(r + 1) Then
        List1.RemoveItem (r)
    End If
Next
For r = List1.ListCount - 1 To 1 Step -1
    If Trim(List1.List(r)) = "" Then
        List1.RemoveItem (r)
    End If
Next
TOTAL_COUNT_LINE_AFTER_MERGE_IF_ANY_AD_BLOCKER = List1.ListCount - 1

' ------------------------------------------------
' DO SOME TALLY INFO
' -------------------------------------------------
' -------------------------------------------------
IM1 = ""
'FILENAME_NAME_01 = "TEXT_ADBLOCK_TEXT__01__MAIN.TXT"
'FILENAME_NAME_02 = "TEXT_ADBLOCK_TEXT__02__MERGE_COMBINE.TXT"
IM1 = IM1 + "________________________________________"
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "MERGE BLEND____"
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "\" + FILENAME_NAME_01
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "\" + FILENAME_NAME_02
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "________________________________________"
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "TOTAL_COUNT_LINE_AD_BLOCKER____ORIGINAL"
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + Trim(Str(TOTAL_COUNT_LINE_AD_BLOCKER))
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "TOTAL_COUNT_LINE____MERGE_AD_BLOCKER"
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + Trim(Str(TOTAL_COUNT_LINE_MERGE_AD_BLOCKER))
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + "TOTAL_COUNT_LINE_AFTER_MERGE_IF_ANY____AD_BLOCKER"
IM1 = IM1 + "" + vbCrLf
IM1 = IM1 + Trim(Str(TOTAL_COUNT_LINE_AFTER_MERGE_IF_ANY_AD_BLOCKER))
' -------------------------------------------------
' -------------------------------------------------

' ------------------------------------------------
' 01 OF 03 ---------------------------------------
' ------------------------------------------------
STRING_WRITE_FILE = STARTER_BLOCK
' ------------------------------------------------

' ------------------------------------------------
' 02 OF 03 ---------------------------------------
' 02 OF 03 -- LIST_BOX_WITH_INFO -- ADD TO STRING
' MAIN DETAIL MIDDLE STRING -- ADBLOCK INO
' ------------------------------------------------
ADBLOCK_INFO_MAIN_CHUNK = ""
For r = 0 To List1.ListCount - 1
    LIST_BOX_WITH_INFO = List1.List(r)
    If LIST_BOX_WITH_INFO <> "" Then
        STRING_WRITE_FILE = STRING_WRITE_FILE + LIST_BOX_WITH_INFO + vbCrLf
        ADBLOCK_INFO_MAIN_CHUNK = ADBLOCK_INFO_MAIN_CHUNK + LIST_BOX_WITH_INFO + vbCrLf
    End If
Next

' ------------------------------------------------
' 03 OF 03 ---------------------------------------
' ------------------------------------------------
STRING_WRITE_FILE = STRING_WRITE_FILE + ENDER_BLOCK
' ------------------------------------------------

' ------------------------------------------------
' WRITE TO FILE TEMP
' ------------------------------------------------
FR1 = FreeFile
Open FILENAME_LOAD_01 + ".TXT2" For Output As #FR1
    Print #FR1, STRING_WRITE_FILE;
Close #FR1
' ---------------------------------------------------

' ---------------------------------------------------
' ' TEMP FILE COMPARE GOOD TO MEMORY -- FILENAME_LOAD_01 + ".TXT2"
' COPY TEMP TO FINAL DESTINATION
' FSO.CopyFile FILENAME_LOAD_01 + ".TXT2", FILENAME_LOAD_01, True
' OKAY NEXT BIT

'If Dir(FILENAME_LOAD_01) = "" Then
'    Shell "EXPLORER """ + FOLDER_NAME_____ + """", vbNormalFocus
'    MsgBox "FILE COPY CREATE FAIL -- " + vbCr + vbCr + FILENAME_LOAD_01 + vbCr + vbCr + "CHECK HERE TEMP BACKUP FILE WILL STILL EXIST" + vbCr + vbCr + FILENAME_LOAD_01 + ".TXT2" + vbCr + vbCr + IM1, vbMsgBoxSetForeground
'    Exit Sub
'End If

'' VERIFY DESTINATION FILE
'FR1 = FreeFile
'Open FILENAME_LOAD_01 For Binary As #FR1
'    VALUE_CHUNK = Space$(LOF(FR1))
'    Get #FR1, 1, VALUE_CHUNK
'Close #FR1

' ---------------------------------------------------------
' ---------------------------------------------------------
' COMPLETE
' ---------------------------------------------------------
' ---------------------------------------------------------



' ---------------------------------------------------------
' TRIM THE ARCHIVE
' AT LEAST 10 NOT MORE THAN 40 DAY
' GET THE DIR FILE INFO INTO A LISTBOX
' ------------------------------------
List1.Clear
' T_FILE = Dir(FOLDER_NAME_____ + "\TEXT_ADBLOCK_TEXT__ARCHIVE__*.TXT")
T_FILE = Dir(FOLDER_NAME_____ + "\TEXT_ADBLOCK_ARC*.TXT")
Do
    If T_FILE = "" Then Exit Do
    List1.AddItem FOLDER_NAME_____ + "\" + T_FILE
    T_FILE = Dir
Loop Until T_FILE = ""

' START WITH NEW -- TOP DOWN -- FAR NAUGHTER
' GOT TO GO BACK WAY WHEN REMOVE SAME TIME
' REMOVE SOME AND LEAVE 10 THERE ONLY IF DATE RIGHT 40 DAY
' AT LEAST 10 NOT MORE THAN 40 DAY
OLD_LIST_COUNT = -2
Do
    LIST_COUNT = List1.ListCount
    Debug.Print List1.List(1)
    If List1.ListCount <= 10 Then Exit Do
    
    Set F = FSO.GetFile(List1.List(1))
    ADATE1 = F.DateLastModified
    
    TIME_SPAN = Now - 10                     ' 10 DAY
    ' TIME_SPAN = Now - TimeSerial(1, 0, 0)    ' A HOUR
    If ADATE1 < TIME_SPAN Then
        Kill List1.List(1)
        List1.RemoveItem (1)
    End If
    
    ' INSTRUCTION RETURN NONE DO AND THEN EXIT -- OLD_LIST_COUNT = LIST_COUNT
    If OLD_LIST_COUNT = LIST_COUNT Then Exit Do
    OLD_LIST_COUNT = LIST_COUNT
Loop Until True = False
' ------------------
' GOT 10 COUNTER
' 10 OR MORE
' NOT 40 DAY FOR LONG TIME -- NICE CODIN
' DEBUG TEST ON ONE HOUR
' 1 ONE -- 1 3 -- 1 TO 3 -- 1 2 3 -- 01 HOUR -- 001 HOUR --
' ------------------
' ------------------



' LINE NUMBER 0004
CRC_ADBLOCK_02 = Hex(m_CRC.CalculateFile(FILENAME_LOAD_01 + ".TXT2"))
If CRC_ADBLOCK_02 <> CRC_ADBLOCK_01 Then
    ' -------------------------------------------------------------
    ' THE CREATOR FILE TEMP IS NOT SAME BEFORE AND WRITE SOME
    ' COPY TEMP REPLACE OVER BEFORE
    ' -------------------------------------------------------------
    FSO.CopyFile FILENAME_LOAD_01 + ".TXT2", FILENAME_LOAD_01, True
    ' ---------------------------------
    ' CALL LOAD CLIPBOARD ON
    ' ---------------------------------
    Call CLIPBOARD_PROCESS_ADBLOCK_FILE
    ' ----------------------------------------------------------------
    ' LOAD NOTEPAD++
    ' ----------------------------------------------------------------
    
    APP_NAME = "C:\PROGRAM FILES (X86)\NOTEPAD++\NOTEPAD++.EXE"
    If FSO.FileExists(APP_NAME) = True Then
        Shell APP_NAME + " """ + FILENAME_LOAD_01 + """", vbMaximizedFocus
        DO_DONE = True
    End If
    ' NOT LOAD WHEN 32 BIT COMPUTER XP RANGE
    ' ----------------------------------------------------
'    APP_NAME = "C:\PROGRAM FILES\NOTEPAD++\NOTEPAD++.EXE"
'    If DO_DONE = False And FSO.FileExists(APP_NAME) = True Then
'        Shell APP_NAME + " """ + FILENAME_LOAD_01 + """", vbMaximizedFocus
'        DO_DONE = True
'    End If
Else
    ' -------------------------------------------------------
    ' THE CREATOR FILE TEMP IS SAME BEFORE AND NOT WRITE OVER
    ' -------------------------------------------------------
    ' ---------------------------------
    ' CALL LOAD CLIPBOARD ON
    ' ---------------------------------
    Call CLIPBOARD_PROCESS_ADBLOCK_FILE
    ' -------------------------------------------------------
End If

CRC_ADBLOCK_01 = Hex(m_CRC.CalculateFile(FILENAME_LOAD_01))
CRC_ADBLOCK_02 = Hex(m_CRC.CalculateFile(FILENAME_LOAD_01 + ".TXT2"))
If CRC_ADBLOCK_02 = CRC_ADBLOCK_01 Then
    ' -------------------------------------------
    ' TEMP COPY PERFECT DESTINATION
    ' TEMP IS GONNA
    ' -------------------------------------------
    If Dir(FILENAME_LOAD_01 + ".TXT2") <> "" Then
        Kill FILENAME_LOAD_01 + ".TXT2"
    End If
End If
    
Me.WindowState = vbMinimized

' -------------------------------------------
' MSGBOX WITH NUMBER INFO
' -------------------------------------------
IM2 = ""
IM2 = IM2 + "ADBLOCK INFO AND MERGER" + vbCr
IM2 = IM2 + "TRIGGER NOTPAD++ \" + FILENAME_NAME_01
MsgBox IM2 + vbCr + IM1, vbMsgBoxSetForeground

FOLDER_NAME_____ = "C:\SCRIPTER\SCRIPT 08 IN_NET_ ADBLOCK"
'Shell "EXPLORER """ + FOLDER_NAME_____ + """", vbNormalFocus

End Sub

Public Sub CLIPBOARD_PROCESS_ADBLOCK_FILE()
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText ADBLOCK_INFO_MAIN_CHUNK
    If Err.Number > 0 Then
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "TIMER_CLIPBOARD_PROCESS_ADBLOCK_FILE_Timer"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
    End If
End Sub



Private Sub Timer_SECOND_Timer()

HWND_2 = FindWindow("Chrome_WidgetWin_1", "Halifax - Enter Memorable Information - Google Chrome")
HWND_2_TEXT = ""
If HWND_2 > 0 Then HWND_2_TEXT = GetWindowTitle(HWND_2)
If HWND_2_TEXT_OLD <> HWND_2_TEXT Then
    Call LINE_PICKER_COMMON.MNU_UNLOAD_FORM_Click
    HWND_2_OLD = 0
End If
HWND_2_TEXT_OLD = HWND_2_TEXT

If GetForegroundWindow = HWND_2 Then
    If HWND_2 <> HWND_2_OLD Then
        HWND_2_OLD = HWND_2
        ' Call Label4_Click ' ---- LINE_PICKER_COMMON
        Load LINE_PICKER_COMMON
        RESIZE_DONE_ONCE = False
        RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
        RESIZE_WINDOWSTATE_CHANGE = -200
        FORM_SET_WIDTH_ONE = False
        LINE_PICKER_COMMON.Visible = True
        LINE_PICKER_COMMON.WindowState = Normal
        FRM_ClipTest.WindowState = vbMinimized
        
        LINE_PICKER_COMMON.MNU_ALWAYS_ON_TOP.Caption = "ALWAYS ON TOP __"
        Call LINE_PICKER_COMMON.MNU_ALWAYS_ON_TOP_SET_TRUE
        Call LINE_PICKER_COMMON.ME_POSITION_HALIFAX
    End If
End If

' HWND_2_OLD = HWND_2
'If GetForegroundWindow <> HWND_2 Then
'If GetForegroundWindow <> Me.hWnd Then
'If GetForegroundWindow <> LINE_PICKER_COMMON.hWnd Then
'    HWND_2_OLD = 0
'End If
'End If
'End If
End Sub


Private Sub MMControl4_Done(NotifyCode As Integer)
'MMControl4.
End Sub

Private Sub MNU_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE_Click()
If InStr(MNU_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE.Caption, "FALSE") Then
    MNU_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE.Caption = "AUTO CONVERT -- TRUE"
Else
    MNU_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE.Caption = "AUTO CONVERT -- FALSE"
End If
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_CALC_COOP_20_Click()

'    If MNU_VAR_CLIPPER = "" Then
'        Do
'            On Error Resume Next
'            Err.Clear
'            MNU_VAR_CLIPPER = Clipboard.GetText
'            On Error GoTo 0
'            If Err.Number > 0 Then Sleep 500
'        Loop Until Err.Number = 0
'    End If
    
    If IsNumeric(MNU_VAR_CLIPPER) = False Or Trim(MNU_VAR_CLIPPER) = "" Then
        MNU_VAR_CLIPPER = Text2.Text
        MNU_VAR_CLIPPER = Replace(MNU_VAR_CLIPPER, vbCr, "")
        MNU_VAR_CLIPPER = Replace(MNU_VAR_CLIPPER, vbLf, "")
        MNU_VAR_CLIPPER = Trim(MNU_VAR_CLIPPER)
    End If
    If IsNumeric(MNU_VAR_CLIPPER) = False Then
        MsgBox "NOT NUMERIC NUMBER GIVE BY CLIPBOARD OR IN THE BOX THERE"
    End If
    i = MNU_VAR_CLIPPER
    
    If COOP_CALC_SETTER = 0 Then COOP_CALC_SETTER = 20
        
    i = COOP_CALC_SETTER - Val(MNU_VAR_CLIPPER)
    
    MNU_VAR_CLIPPER = i
    
    Me.WindowState = vbMinimized
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    'Clipboard.SetText "CALC FOR COOP MAKE TO £20 IS __ " + Str(MNU_VAR_CLIPPER)
    Clipboard.SetText Format(MNU_VAR_CLIPPER, "###00.00") '+ " "
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_CALC_COOP_20_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If

    MsgBox "CALC FOR COOP MAKE TO £" + Trim(Str(COOP_CALC_SETTER)) + " IS __ £" + Format(MNU_VAR_CLIPPER, "###0.00") + " ", vbMsgBoxSetForeground

    MNU_VAR_CLIPPER = ""

End Sub

Private Sub MNU_CALC_COOP_40_Click()

COOP_CALC_SETTER = 40
Call MNU_CALC_COOP_20_Click
COOP_CALC_SETTER = 0

End Sub

Private Sub MNU_CLIPBOARD_LOGGER_Click()

'D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe

End Sub

Private Sub MNU_COMPARE_IN_HEXVIEWER_Click()
    ' --------------------------------
    ' WORK -- MNU_COMPARE_IN_HEXVIEWER
    ' --------------------------------
    ' Sat 04-Jan-2020 09:47:56
    ' Sat 04-Jan-2020 12:11:09
    ' --------------------------------

    Dim FILENAME_COMPARE_1
    Dim FILENAME_COMPARE_2
    Dim COMPARE_EXE_1, COMPARE_EXE_2
    Dim FR1
    Dim FN_VAR_1
    Dim FN_VAR_2
    
    Beep
    Me.WindowState = vbMinimized

    FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\"
    If Dir(FILENAME_IN_USE_CHECK, vbDirectory) = "" Then
        CreateFolderTree FILENAME_IN_USE_CHECK
    End If
    
    
    FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\#COMPARE _" + GetComputerName + "_ 1 DUPE CHECKER.Txt"
    'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_COMPARE_1 = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        Print #FR1, Text2.Text;
    Close #FR1

    FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\#COMPARE _" + GetComputerName + "_ 2 DUPE CHECKER.Txt"
    'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_COMPARE_2 = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        Print #FR1, Text3.Text;
    Close #FR1

    FN_VAR_1 = "D:\VB6\VB-NT\00_Best_VB_01\CLIPBOARD_VIEWER\ClipBoard Viewer.exe"
    FN_VAR_1 = "C:\Program Files\XVI32\XVI32.exe"
    
    FN_VAR_2 = """" + FILENAME_COMPARE_1 + """"
    FN_VAR_4 = """" + FILENAME_COMPARE_2 + """"
    FN_VAR_2 = FILENAME_COMPARE_1
    FN_VAR_4 = FILENAME_COMPARE_2
    
    Shell FN_VAR_1 + " """ + FN_VAR_2 + """", vbNormalFocus
    Shell FN_VAR_1 + " """ + FN_VAR_4 + """", vbNormalFocus

'    Dim WSHShell
'    Set WSHShell = CreateObject("WScript.Shell")
'        WSHShell.Run """" + FN_VAR_1 + """" + " " + FN_VAR_2, ShowWindow_2, DontWaitUntilFinished
'    Set WSHShell = Nothing
'    Set WSHShell = CreateObject("WScript.Shell")
'        WSHShell.Run """" + FN_VAR_1 + """" + " " + FN_VAR_4, ShowWindow_2, DontWaitUntilFinished
'    Set WSHShell = Nothing

End Sub

Public Sub MNU_CONVERT_GSD_PATH_TO_NETWORK_FRIENDLY_ONE_Click()

Dim SS, TTD
Dim PATH_COMPONENT
Dim r
'Dim DOUT
Dim HDD

On Error Resume Next
TTD = Clipboard.GetText

'SS = "\\7-asus-gl522vw\7_ASUS_GL522VW_02_D_DRIVE\VI_ DSC ME 02\2001 MEDIA HARDCORE 01"
'TTD = "gstps://7-asus-gl522vw.matt-lan-btinternet-com.goodsync/D:/VI_ DSC ME 02/2001 MEDIA HARDCORE 01"

TTD = Replace(TTD, LCase("2-ASUS-X"), LCase("1-ASUS-X"))
TTD = Replace(TTD, LCase("1-ASUS-E"), LCase("2-ASUS-E"))

PATH_COMPONENT = Mid(TTD, InStrRev(TTD, ":/") + 1)
PATH_COMPONENT = Replace(PATH_COMPONENT, "/", "\")
Dim AR_1(7)
i = 0
i = i + 1: AR_1(i) = "1-ASUS-X5DIJ"
i = i + 1: AR_1(i) = "2-ASUS-EEE"
i = i + 1: AR_1(i) = "3-LINDA-PC"
i = i + 1: AR_1(i) = "4-ASUS-GL522VW"
i = i + 1: AR_1(i) = "5-ASUS-P2520LA"
i = i + 1: AR_1(i) = "7-ASUS-GL522VW"
i = i + 1: AR_1(i) = "8-MSI-GP62M-7RD"
Dim AR_2(7)
i = 0
i = i + 1: AR_2(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ"
i = i + 1: AR_2(i) = "\\2-ASUS-EEE\2_ASUS_EEE"
i = i + 1: AR_2(i) = "\\3-LINDA-PC\3_LINDA_PC"
i = i + 1: AR_2(i) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW"
i = i + 1: AR_2(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA"
i = i + 1: AR_2(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW"
i = i + 1: AR_2(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD"

For r = 1 To 7
    If InStr(UCase(TTD), AR_1(r)) Then
        DOUT = AR_2(r)
    End If
Next

HDD2 = HDD
If InStr(UCase(TTD), "/C:") Then HDD = UCase("_01_C_DRIVE")
If InStr(UCase(TTD), "/D:") Then HDD = UCase("_02_D_DRIVE")
If InStr(UCase(TTD), "/E:") Then HDD = UCase("_03_FAT32_4GB")

If InStr(TTD, "7-ASUS-GL522VW") > 0 Then
    If InStr(UCase(TTD), "/L:") Then HDD = UCase("_10_1_SAMSUNG_4TB_C")
    If InStr(UCase(TTD), "/M:") Then HDD = UCase("_10_1_SAMSUNG_4TB_D")
    If InStr(UCase(TTD), "/N:") Then HDD = UCase("_10_1_SAMSUNG_4TB_E")
    If InStr(UCase(TTD), "/P:") Then HDD = UCase("_20_2_SAMSUNG_4TB_C")
    If InStr(UCase(TTD), "/Q:") Then HDD = UCase("_20_2_SAMSUNG_4TB_D")
    If InStr(UCase(TTD), "/R:") Then HDD = UCase("_20_2_SAMSUNG_4TB_E")
    If InStr(UCase(TTD), "/O:") Then HDD = UCase("_30_4_SAMSUNG_4TB_HUBIC")
    If InStr(UCase(TTD), "/K:") Then HDD = UCase("_40_1_SAMSUNG_2TB_CLOUD_2TB")
    If InStr(UCase(TTD), "/S:") Then HDD = UCase("_40_2_SAMSUNG_2TB_CLOUD_2TB")
    If InStr(UCase(TTD), "/T:") Then HDD = UCase("_80_3_SAMSUNG_4TB_D")
    If InStr(UCase(TTD), "/U:") Then HDD = UCase("_80_4_SAMSUNG_4TB_D")
End If
' CMD NET SHARE

'If HDD = "" Then
'    i = ""
'    i = i + "SHARE WITH FOLDER NAME -- NOT EXIST ANY -- VAR" + vbCrLf + vbCrLf
'    i = i + "BOTH THE SAME __ REQUIRE MORE PROGRAMMER" + vbCrLf + vbCrLf
'    i = i + "ONLY WORK FROM DRIVE ABOVE E: WITH 7-ASUS-GL522VW" + vbCrLf + vbCrLf
'    i = i + "VAR -- TTD" + vbCrLf + vbCrLf
'    i = i + "----------" + vbCrLf
'    i = i + TTD
'    i = i + "----------"
'    MsgBox i, vbMsgBoxSetForeground
'End If

DOUT = DOUT + HDD + PATH_COMPONENT

'If Dir(DOUT, vbDirectory) = "" Then
'    Beep
'    MsgBox "DIRECTORY NOT EXIST ERR __ CONVERT ANYWAY" + vbCrLf + vbCrLf + DOUT, vbMsgBoxSetForeground
'End If

'If Dir(DOUT, vbDirectory) <> "" Then
    ' Me.WindowState = vbMinimized
    Call MNU_CONVERT_GSD_PATH_TO_NETWORK_FRIENDLY_ONE__0002
'End If

End Sub

Public Sub MNU_CONVERT_GSD_PATH_TO_NETWORK_FRIENDLY_ONE__0002()
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText DOUT
    If Err.Number > 0 Then
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_CONVERT_GSD_PATH_TO_NETWORK_FRIENDLY_ONE__0002"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
    End If
End Sub



Private Sub MNU_DIRECTORY_PICKER_Click()
    Load LINE_PICKER_DIRECTORY

    LINE_PICKER_DIRECTORY.Visible = True
    LINE_PICKER_DIRECTORY.WindowState = vbMaximized
    ' LINE_PICKER_DIRECTORY_LIST_SET
End Sub



Private Sub MNU_FIND_DIFFER_Click()

If Len(Text2.Text) = 0 Or Len(Text3.Text) = 0 Then
    MNU_FIND_DIFFER.Caption = "COMPARE VIEWER"
    Exit Sub
End If

X5 = Label1(7).Caption
NUMBER_01 = 0
If IsNumeric(Mid(X5, 1, 4)) Then
    NUMBER_01 = Val(Mid(X5, 1, 4))
End If
For r = 1 To Len(Text2.Text)
    If Mid(Text2.Text, r, 1) <> Mid(Text3.Text, r, 1) Then
        SET_START_INDEXER = r - 2
        If SET_START_INDEXER < 1 Then SET_START_INDEXER = 1
        If r > NUMBER_01 Then
            Label1(7).Caption = Format(r, "0000 ") + "[_" + Mid(Text2.Text, SET_START_INDEXER, 4) + "_]__[_" + Mid(Text3.Text, SET_START_INDEXER, 4) + "_]"
            NUMBER_01 = r
            Exit For
        End If
    End If
Next
If r - 1 > NUMBER_01 Then
    NUMBER_01 = 0
    For r = 1 To Len(Text2.Text)
        If Mid(Text2.Text, r, 1) <> Mid(Text3.Text, r, 1) Then
            SET_START_INDEXER = r - 2
            If SET_START_INDEXER < 1 Then SET_START_INDEXER = 1
            Label1(7).Caption = Format(r, "0000 ") + "[_" + Mid(Text2.Text, SET_START_INDEXER, 4) + "_]__[_" + Mid(Text3.Text, SET_START_INDEXER, 4) + "_]"
            Exit Sub
        End If
    Next
End If

End Sub


Private Sub MNU_GIVE_ME_10_DASH_LINE_Click()
    Dim STRING_OUT
    
    STRING_OUT = String(10, "-") + vbCrLf
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.SetText STRING_OUT
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If Err.Number > 0 Then Exit Sub
    
    Me.WindowState = vbMinimized
End Sub
Private Sub MNU_GIVE_ME_40_DASH_LINE_Click()
    Dim STRING_OUT
    
    STRING_OUT = String(40, "-") + vbCrLf
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.SetText STRING_OUT
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If Err.Number > 0 Then Exit Sub
    
    Me.WindowState = vbMinimized
End Sub



Public Sub MNU_GOODSYNC_VOLUME_LABLE_DRIVE_Click()

'=3_SAMSUNG_4TB_D:\VI_ DSC 01 V0 01 MM
'\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE\VIDEO\NOT\X 00 NOT ME\00 Vid XXX



Dim SS, TTD
Dim PATH_COMPONENT
Dim r
Dim DOUT
Dim HDD

On Error Resume Next
TTD = Clipboard.GetText


'SS = "\\7-asus-gl522vw\7_ASUS_GL522VW_02_D_DRIVE\V"
'TTD = "gstps://7-asus-gl522vw.matt-lan-btinternet-com.goodsync/D:/V"

If InStr(TTD, "\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE") > 0 Then
    TTD = Replace(TTD, "\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE\", "=D DRIVE:\")
End If

Dim FS, d, S, DRVPATH
DRVPATH = Mid(TTD, 1, 2)
If Trim(DRVPATH) = "" Then Exit Sub

If Mid(DRVPATH, 2, 1) = ":" Then
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set d = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName(DRVPATH)))
    S = "Drive " & d.DriveLetter & ": - " & d.VolumeName
    TTD = "=" + d.VolumeName + ":\" + Mid(TTD, 4)
End If




'
'
'TTD = Replace(TTD, LCase("2-ASUS-X"), LCase("1-ASUS-X"))
'TTD = Replace(TTD, LCase("1-ASUS-E"), LCase("2-ASUS-E"))
'
'PATH_COMPONENT = Mid(TTD, InStrRev(TTD, ":/") + 1)
'PATH_COMPONENT = Replace(PATH_COMPONENT, "/", "\")
'Dim AR_1(7)
'i = 0
'i = i + 1: AR_1(i) = "1-ASUS-X5DIJ"
'i = i + 1: AR_1(i) = "2-ASUS-EEE"
'i = i + 1: AR_1(i) = "3-LINDA-PC"
'i = i + 1: AR_1(i) = "4-ASUS-GL522VW"
'i = i + 1: AR_1(i) = "5-ASUS-P2520LA"
'i = i + 1: AR_1(i) = "7-ASUS-GL522VW"
'i = i + 1: AR_1(i) = "8-MSI-GP62M-7RD"
'Dim AR_2(7)
'i = 0
'i = i + 1: AR_2(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ"
'i = i + 1: AR_2(i) = "\\2-ASUS-EEE\2_ASUS_EEE"
'i = i + 1: AR_2(i) = "\\3-LINDA-PC\3_LINDA_PC"
'i = i + 1: AR_2(i) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW"
'i = i + 1: AR_2(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA"
'i = i + 1: AR_2(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW"
'i = i + 1: AR_2(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD"
'
'For r = 1 To 7
'    If InStr(UCase(TTD), AR_1(r)) Then
'        DOUT = AR_2(r)
'    End If
'Next
'
'If InStr(UCase(TTD), "/C:") Then HDD = UCase("_01_C_DRIVE")
'If InStr(UCase(TTD), "/D:") Then HDD = UCase("_02_D_DRIVE")
'If InStr(UCase(TTD), "/E:") Then HDD = UCase("_03_FAT32_4GB")
'If InStr(UCase(TTD), "/L:") Then HDD = UCase("_10_1_SAMSUNG_4TB_C")
'If InStr(UCase(TTD), "/M:") Then HDD = UCase("_10_1_SAMSUNG_4TB_D")
'If InStr(UCase(TTD), "/N:") Then HDD = UCase("_10_1_SAMSUNG_4TB_E")
'If InStr(UCase(TTD), "/P:") Then HDD = UCase("_20_2_SAMSUNG_4TB_C")
'If InStr(UCase(TTD), "/Q:") Then HDD = UCase("_20_2_SAMSUNG_4TB_D")
'If InStr(UCase(TTD), "/R:") Then HDD = UCase("_20_2_SAMSUNG_4TB_E")
'If InStr(UCase(TTD), "/O:") Then HDD = UCase("_30_4_SAMSUNG_4TB_HUBIC")
'If InStr(UCase(TTD), "/K:") Then HDD = UCase("_40_1_SAMSUNG_2TB_CLOUD_2TB")
'If InStr(UCase(TTD), "/S:") Then HDD = UCase("_40_2_SAMSUNG_2TB_CLOUD_2TB")
'If InStr(UCase(TTD), "/T:") Then HDD = UCase("_80_3_SAMSUNG_4TB_D")
'If InStr(UCase(TTD), "/U:") Then HDD = UCase("_80_4_SAMSUNG_4TB_D")
'' CMD NET SHARE
'
'
'DOUT = DOUT + HDD + PATH_COMPONENT

DOUT = TTD

If Dir(DOUT, vbDirectory) = "" Then Beep
If Dir(DOUT, vbDirectory) <> "" Then
    Me.WindowState = vbMinimized
End If

On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText DOUT
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0




End Sub

Private Sub MNU_KILL_ALL_CLIPVIEWER_Click()
    Call FindWindowPart_VB_CLIPVIEWER_CLOSE_ALL
End Sub

Private Sub MNU_NUMBER_TO_STRING_AND_DATE_Click()

GET_TIME_VAR = Now
A_DATE_TIME_PM = Mid(Format(GET_TIME_VAR, "HH-MM-00 Am/Pm"), 10)
A_DATE_TIME_2 = Format(GET_TIME_VAR, "DDDD HH:MM:00") + " " + A_DATE_TIME_PM + "__" + Format(GET_TIME_VAR, "DD MMMM YYYY")


Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4 As String

Do
    On Error Resume Next
    Err.Clear
    VAR_ST_1 = Clipboard.GetText
    On Error GoTo 0
    If Err.Number > 0 Then Sleep 500
Loop Until Err.Number = 0

VAR_ST_4 = VAR_ST_1

If IsNumeric(VAR_ST_1) Then
    VAR_ST_4 = NumToString(VAR_ST_1)
    VAR_ST_4 = VAR_ST_1 + " __ " + Trim(Mnu_Fix1st_Letters(VAR_ST_4)) + " __ " + A_DATE_TIME_2
End If

If VAR_ST_1 <> VAR_ST_4 Then
    Do
        On Error Resume Next
        Clipboard.Clear
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> "" Then Sleep 500
    Loop Until VAR_ST_2 = ""
    Sleep 100
    Do
        On Error Resume Next
        Clipboard.SetText VAR_ST_4, vbCFText
        Sleep 100
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> VAR_ST_1 Then Sleep 800
    Loop Until VAR_ST_2 = VAR_ST_4

    Me.WindowState = vbMinimized

End If

'Beep

Me.WindowState = vbMinimized




End Sub

Private Sub MNU_NUMBER_TO_STRING_Click()

Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4 As String

Do
    On Error Resume Next
    Err.Clear
    VAR_ST_1 = Clipboard.GetText
    On Error GoTo 0
    If Err.Number > 0 Then Sleep 500
Loop Until Err.Number = 0

VAR_ST_4 = VAR_ST_1

If IsNumeric(VAR_ST_1) Then
    VAR_ST_4 = NumToString(VAR_ST_1)
    VAR_ST_4 = Trim(Mnu_Fix1st_Letters(VAR_ST_4))
End If

If VAR_ST_1 <> VAR_ST_4 Then
    Do
        On Error Resume Next
        Clipboard.Clear
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> "" Then Sleep 500
    Loop Until VAR_ST_2 = ""
    Sleep 100
    Do
        On Error Resume Next
        Clipboard.SetText VAR_ST_4, vbCFText
        Sleep 100
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> VAR_ST_1 Then Sleep 800
    Loop Until VAR_ST_2 = VAR_ST_4

    Me.WindowState = vbMinimized

End If

'Beep
Me.WindowState = vbMinimized



End Sub


Function Mnu_Fix1st_Letters(EE As String)
'HERE IS -- PROCAPS
    
'Dim EE As String

'EE = Text1.Text
'EE = AD$
   
'If EGG = 0 Then EE = LCase(EE)
EE = LCase(EE)

Mid$(EE, 1, 1) = UCase(Mid$(EE, 1, 1))

For r = 65 To 65 + 25
    EE = Replace(EE, " " + LCase(Chr(r)), " " + UCase(Chr(r)))
    EE = Replace(EE, vbLf + LCase(Chr(r)), vbLf + UCase(Chr(r)))
    EE = Replace(EE, vbCr + LCase(Chr(r)), vbCr + UCase(Chr(r)))
    EE = Replace(EE, "-" + LCase(Chr(r)), "-" + UCase(Chr(r)))
    EE = Replace(EE, "(" + LCase(Chr(r)), "(" + UCase(Chr(r)))
    EE = Replace(EE, "." + LCase(Chr(r)), "." + UCase(Chr(r)))
    EE = Replace(EE, "\" + LCase(Chr(r)), "\" + UCase(Chr(r)))
    EE = Replace(EE, "_" + LCase(Chr(r)), "_" + UCase(Chr(r)))
    EE = Replace(EE, """" + LCase(Chr(r)), """" + UCase(Chr(r)))
Next

'HERE IS -- PROCAPS

Mid$(EE, 1, 1) = UCase(Mid$(EE, 1, 1))

AD$ = EE
'
'    EXECUTE_TIMER_ENABLED = False
'        Clipboard.Clear
'    EXECUTE_TIMER_ENABLED = True
'
'    EXECUTE_TIMER_ENABLED = False
'        Clipboard.SetText AD$
'    EXECUTE_TIMER_ENABLED = True
'
'AD_DATE = 0
'
'Me.WindowState = 1

'Text1.Text = EE

Mnu_Fix1st_Letters = AD$

End Function



Private Sub MNU_POUNDS_ON_NEWLINE_AND_CUT_ABOVE_AMOUNT_Click()

' ----------------------------------------------------------------------
' SESSION 010
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' WORK HERE
' MNU_POUNDS_ON_NEWLINE_AND_CUT_ABOVE_AMOUNT_Click()
' ----------------------------------------------------------------------
' Fri 07-Feb-2020 01:58:27
' Fri 07-Feb-2020 05:12:00 -- 3 HOUR 14 MINUTE
' ----------------------------------------------------------------------

Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4 As String

Do
    On Error Resume Next
    Err.Clear
    VAR_ST_1 = Clipboard.GetText
    On Error GoTo 0
    If Err.Number > 0 Then Sleep 500
Loop Until Err.Number = 0

If VAR_ST_1 = "" Then Beep: MsgBox "NONE CLIPBOARD INPUT": Exit Sub

' REMOVE END POUND
StrArray = Split(VAR_ST_1, vbCr)
VAR_ST_1 = ""
For R_COUNT = 0 To UBound(StrArray)
    i = InStrRev(StrArray(R_COUNT), "£")
    StrArray(R_COUNT) = Replace(StrArray(R_COUNT), vbCr, "")
    StrArray(R_COUNT) = Replace(StrArray(R_COUNT), vbLf, "")
    StrArray(R_COUNT) = Replace(StrArray(R_COUNT), "  ", " ")
    If i > 0 Then
        ' REMOVE END POUND
        StrArray(R_COUNT) = Trim(Mid(StrArray(R_COUNT), 1, i - 2))
    Else
        StrArray(R_COUNT) = Trim(StrArray(R_COUNT))
    End If
    VAR_ST_1 = VAR_ST_1 + StrArray(R_COUNT) + vbCrLf
    ' Debug.Print StrArray(R_COUNT)
Next

VAR_ST_1 = Replace(VAR_ST_1, vbLf, "")
VAR_ST_1 = Replace(VAR_ST_1, vbCr, "£")
StrArray = Split(VAR_ST_1, "£")

For R_COUNT = 0 To UBound(StrArray)
    If Len(StrArray(R_COUNT)) < 12 Then
        StrArray(R_COUNT) = Replace(StrArray(R_COUNT), Chr(9) + " ", "") ' TAB + SPACE
        StrArray(R_COUNT) = Replace(StrArray(R_COUNT), Chr(9), "")       ' TAB
    End If
Next

TEXT_BOARD_02 = ""
For R_COUNT = 0 To UBound(StrArray)
    If Trim(StrArray(R_COUNT)) <> "" Then
        If Len(StrArray(R_COUNT)) < 10 Then
            If Val(StrArray(R_COUNT)) > 0 Then
                TOT_SO_FAR = TOT_SO_FAR + Val(StrArray(R_COUNT))
            End If
        End If
    End If
    TEXT_BOARD_DONE = False
    If IsNumeric(StrArray(R_COUNT)) And Val(StrArray(R_COUNT)) > 0 Then
        If TOT_SO_FAR > 0 Then
            If InStr(Str(TOT_SO_FAR), ".") > 0 Then
                TOT_STRING = Format(TOT_SO_FAR, "#.00")
            Else
                TOT_STRING = Format(TOT_SO_FAR, "#")
            End If
            XCOUNT = XCOUNT + 1
            TEXT_BOARD_02 = TEXT_BOARD_02 + "£" + Trim(StrArray(R_COUNT)) + vbCrLf + "TOTAL ACCUMULATOR --" + Str(XCOUNT) + " -- £" + TOT_STRING + vbCrLf
        Else
            TEXT_BOARD_02 = TEXT_BOARD_02 + "£" + Trim(StrArray(R_COUNT)) + vbCrLf
        End If
        TEXT_BOARD_DONE = True
    End If
    ' SHORT LINE
    ' EXAMPLE AT RIGHT
    ' 01 Jan 2020 Transfer to 2 E-SAVINGS £10.00      £659.01
    ' 01 Jan 2020 Transfer to 2 E-SAVINGS
    
    
    StrArray(R_COUNT) = Trim(Replace(StrArray(R_COUNT), Chr(9), " "))               ' TAB
    
    Dim A2(8)
    i = 0
    i = i + 1: A2(i) = "Standing Order    " + Space(8)
    i = i + 1: A2(i) = "Standing order    " + Space(8)
    i = i + 1: A2(i) = "Visa purchase     " + Space(8)
    i = i + 1: A2(i) = "Transfer from     " + Space(8)
    i = i + 1: A2(i) = "Direct debit      " + Space(8)
    i = i + 1: A2(i) = "Visa Credit       " + Space(8)
    i = i + 1: A2(i) = "Transfer to       " + Space(8)
    i = i + 1: A2(i) = "Payment to        " + Space(8)
    For r = 1 To UBound(A2)
        StrArray(R_COUNT) = Trim(Replace(StrArray(R_COUNT), Trim(A2(r)), A2(r)))
    Next

    If TEXT_BOARD_DONE = False Then
        TEXT_BOARD_02 = TEXT_BOARD_02 + Trim(StrArray(R_COUNT)) + vbCrLf
    End If
Next


TEXT_BOARD_02 = Replace(TEXT_BOARD_02, vbCrLf + vbCrLf, vbCrLf)

CLIPBOARD_SET_TEXT_FOR_BANK_VAR = TEXT_BOARD_02
Call CLIPBOARD_SET_TEXT_FOR_BANK

' MsgBox CLIPBOARD_SET_TEXT_FOR_BANK_VAR
Me.WindowState = vbMinimized

'VAR_ST_4 = Trim(Mnu_Fix1st_Letters(VAR_ST_4))
End Sub

Public Sub CLIPBOARD_SET_TEXT_FOR_BANK()
    
    Me.WindowState = vbMinimized
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText CLIPBOARD_SET_TEXT_FOR_BANK_VAR
    If Err.Number > 0 Then
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "CLIPBOARD_SET_TEXT_FOR_BANK"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
    End If

End Sub




Public Sub CLIPBOARD_SET_TEXT_FOR_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE_01()
    
    Exit Sub
    
    If InStr(MNU_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE.Caption, "AUTO CONVERT -- TRUE") = 0 Then
        Exit Sub
    End If
    
    Call MNU_CONVERT_GSD_PATH_TO_NETWORK_FRIENDLY_ONE_Click

    ' Call CLIPBOARD_SET_TEXT_FOR_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE_02

End Sub

Public Sub CLIPBOARD_SET_TEXT_FOR_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE_02()
    
    Exit Sub
    
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText CLIPBOARD_SET_TEXT_FOR_BANK_VAR
    If Err.Number > 0 Then
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "CLIPBOARD_SET_TEXT_FOR_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE_02"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
    End If

End Sub


Private Sub MNU_RS232_LOGGER_OPEN_Click()
    PATH_2 = "VB6\VB-NT\00_Best_VB_01\Tidal_Info"
    NET_PATH = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE"
    If GetComputerName = "4-ASUS-GL522VW" Then
        NET_PATH = "C:"
    End If
    
    FOLDER_NAME = NET_PATH + "\SCRIPTOR DATA\" + PATH_2
    FILE_NAME_4 = "ZZ RS232 FRONT DOOR LOGGER.txt"
    FILE_NAME_4 = FOLDER_NAME + "\" + FILE_NAME_4
    A_NOW = Now
    Call WRITE_LOGGER_OPEN
    
    FILE_NAME_2 = "RS232 FRONT DOOR.txt"
    FILE_NAME_8 = "RS232 FRONT DOOR OPEN.txt"
    FILE_NAME_9 = "RS232 FRONT DOOR CLOSE.txt"
    FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
    FR1 = FreeFile
    Open FILE_NAME For Output As #FR1
    Close #FR1
    FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_8
    FR1 = FreeFile
    Open FILE_NAME For Output As #FR1
    Close #FR1
    
    Me.WindowState = vbMinimized
    APP_NAME = "C:\Program Files (x86)\Notepad++\notepad++.exe"
    Shell APP_NAME + " """ + FILE_NAME_4 + """", vbMaximizedFocus
    
    Shell "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe", vbNormalFocus
    
End Sub

Private Sub MNU_RS232_LOGGER_CLOSE_Click()
    PATH_2 = "VB6\VB-NT\00_Best_VB_01\Tidal_Info"
    NET_PATH = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE"
    If GetComputerName = "4-ASUS-GL522VW" Then
        NET_PATH = "C:"
    End If
    FOLDER_NAME = NET_PATH + "\SCRIPTOR DATA\" + PATH_2
    FILE_NAME_4 = "ZZ RS232 FRONT DOOR LOGGER.txt"
    FILE_NAME_4 = FOLDER_NAME + "\" + FILE_NAME_4
    A_NOW = Now
    Call WRITE_LOGGER_CLOSE
    
    FILE_NAME_2 = "RS232 FRONT DOOR.txt"
    FILE_NAME_8 = "RS232 FRONT DOOR OPEN.txt"
    FILE_NAME_9 = "RS232 FRONT DOOR CLOSE.txt"
    FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
    FR1 = FreeFile
    Open FILE_NAME For Output As #FR1
    Close #FR1
    FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_9
    FR1 = FreeFile
    Open FILE_NAME For Output As #FR1
    Close #FR1
    
    Me.WindowState = vbMinimized
    APP_NAME = "C:\Program Files (x86)\Notepad++\notepad++.exe"
    Shell APP_NAME + " """ + FILE_NAME_4 + """", vbMaximizedFocus
    
    ' Shell "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe", vbNormalFocus
    
    
End Sub
                
Sub WRITE_LOGGER_CLOSE()
    FR1 = FreeFile
    Open FILE_NAME_4 For Append As #FR1
    ' Debug.Print "WRITE FILE_NAME " + FILE_NAME
    XX4 = Format(A_NOW, "YYYY-MM-DD -- HH:MM:00 -- HH:MM:00 AM/PM -- DDD")
    XX4 = Replace(XX4, "6:00 ", "8:00 ")
    XX4 = Replace(XX4, "3:00 ", "4:00 ")
    If InStr(XX4, "9:00 ") > 0 Then
        For r = 1 To 2
            TT = InStr(XX4, "9:00 ")
            If TT > 0 Then
                XX4 = Replace(XX4, "9:00 ", "0:00 ")
                XX5 = Mid(XX4, TT - 1, 5) ' 00:00
                XX4 = Replace(XX4, XX5, XX4)
            End If
        Next
    End If
    Print #FR1, XX4 + " -- DOOR CLOSE"
    Close #FR1
End Sub

Sub WRITE_LOGGER_OPEN()
    FR1 = FreeFile
    Open FILE_NAME_4 For Append As #FR1
    ' Debug.Print "WRITE FILE_NAME " + FILE_NAME
    XX4 = Format(A_NOW, "YYYY-MM-DD -- HH:MM:00 -- HH:MM:00 AM/PM -- DDD")
    XX4 = Replace(XX4, "6:00 ", "8:00 ")
    XX4 = Replace(XX4, "3:00 ", "4:00 ")
    If InStr(XX4, "9:00 ") > 0 Then
        For r = 1 To 2
            TT = InStr(XX4, "9:00 ")
            If TT > 0 Then
                XX4 = Replace(XX4, "9:00 ", "0:00 ")
                XX5 = Mid(XX4, TT - 1, 5) ' 00:00
                XX4 = Replace(XX4, XX5, XX4)
            End If
        Next
    End If
    Print #FR1, XX4 + " -- DOOR OPEN"
    Close #FR1
End Sub

Private Sub MNU_RUN_GOODSYNC_CAMERA_Click()

    ' ------------------------------------------------------------------
    ' BACKUP CAMERA
    ' Wed 22-Jan-2020 04:48:02
    ' ------------------------------------------------------------------

    FN_VAR_1 = "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT"
    
    FN_VAR_2 = "CAMERA_BACKUP_MODE"
    Me.WindowState = vbMinimized

    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.RUN """" + FN_VAR_1 + """" + " " + FN_VAR_2, ShowWindow_2, DontWaitUntilFinished
    Set WSHShell = Nothing
    
    Me.WindowState = vbMinimized

End Sub


Private Sub MNU_TIMEZONE_CLOCK_Click()
    Me.WindowState = vbMinimized
    APP_NAME = "D:\VB6\VB-NT\00_Best_VB_01\TIMEZONE MINI GUI DISPLAY\TIMEZONE MINI GUI DISPLAY.exe"
    Shell APP_NAME, vbMaximizedFocus
End Sub
Private Sub MNU_UNIFY_AR_LOGITECH_Click()
    Me.WindowState = vbMinimized
    APP_NAME = "C:\Program Files\Common Files\Logishrd\Unifying\DJCUHost.exe"
    Shell APP_NAME, vbMaximizedFocus
End Sub
Private Sub MNU_SetPoint_exe_Click()
    Me.WindowState = vbMinimized
    VB_1 = "C:\Program Files\Logitech\SetPointP\SetPoint.exe"
    If Dir(VB_1) = "" Then
        MsgBox "NONE PROGRAM OF HERE NAME__" + vbCrLf + VB_1
        Exit Sub
    End If
    Shell VB_1 + "  /ss", vbMaximizedFocus
'
'    '------------------------------------------------------
'    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
'    '------------------------------------------------------
'    Dim objShell
'    Set objShell = CreateObject("Wscript.Shell")
'    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
'    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
'    Set objShell = Nothing
End Sub

Private Sub MNU_SHOW_THE_TIME_03_Click(Index As Integer)

Dim TEXT_BOARD_02, STRING_OUT As String
Dim XXTX As String


' ----------------------------------------------------------------------
' MNU_SHOW_THE_TIME_03(1)
' IS THE USUAL CRC32 AND TIME
' ----------------------------------------------------------------------
' OTHER ROUTINE SEARCH HERE SOURCE GENERATOR
' ----------------------------------------------------------------------
' A_DATE_TIME_04 = Format(GET_TIME_VAR, "DD-MMM-YYYY HH:MM:SS DDD")
' If UCase(MNU_SHOW_THE_TIME_03(1).Caption) <> UCase(A_DATE_TIME_04) Then
'     MNU_SHOW_THE_TIME_03(1).Caption = UCase(A_DATE_TIME_04)
' End If
' ----------------------------------------------------------------------


If Index = 0 Then
    
    
    ' ---------------------------------------------------------------------
    ' ---------------------------------------------------------------------
    ' EXAMPLE SHOW
    ' VBCRLF BETWEEN THE REM -- AFTER HERE -- October 2020 ]
    ' ENCLOSE OF QUOTE VARIABLE STING_OUT
    ' ---------------------------------------------------------------------
    ' ---------------------------------------------------------------------
    ' GOT AR
    ' MNU_SHOW_THE_TIME_03(0).Caption + vbCrLf + String(69, "-") + vbCrLf
    ' ---------------------------------------------------------------------
    ' STRING_OUT =
    ' "[ Tue 20:58:50 Pm_06 October 2020 ]
    ' ---------------------------------------------------------------------"
    ' ---------------------------------------------------------------------
    ' ---------------------------------------------------------------------
    
    STRING_OUT = MNU_SHOW_THE_TIME_03(0).Caption + vbCrLf + String(69, "-") + vbCrLf
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.SetText UCase(STRING_OUT)
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If Err.Number > 0 Then Exit Sub
    
    Me.WindowState = vbMinimized
End If

If Index = 1 Then
    ' ---------------------------------------------------------------------
    ' LABEL6 WAS HERE
    ' ---------------------------------------------------------------------
    ' DONT_MIN = True
    ' Call LabelCRC3_Click
    ' ---------------------------------------------------------------------
    ' Thu 18-Oct-2018 13:28:00
    ' ---------------------------------------------------------------------
    On Error Resume Next
    Do
        Err.Clear
        If Clipboard.GetFormat(vbCFText) = True Then
            Err.Clear
            TEXT_BOARD_02 = Clipboard.GetText
            If Err.Number > 0 Then Sleep 200
        End If
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    ' ---------------------------------------------------------------------
    ' EXAMPLE -- GOT -- MNU_SHOW_THE_TIME_03(1).Caption
    ' XXTX = "Tue 06-Oct-2020 20:55:08"
    ' ---------------------------------------------------------------------
    ' TEXT_BOARD_02
    ' "Text Size  425 -- CRC32 B34D8B41"
    ' ---------------------------------------------------------------------
    ' NOW UCASE OUTPUT
    ' ---------------------------------------------------------------------
    XXTX = MNU_SHOW_THE_TIME_03(1).Caption
    ' ---------------------------------------------------------------------
    If InStr(TEXT_BOARD_02, "CRC32") Then
        If InStrRev(TEXT_BOARD_02, vbCrLf) > 0 Then
            STRING_OUT = XXTX + vbCrLf + Mid(TEXT_BOARD_02, InStrRev(TEXT_BOARD_02, vbCrLf) + 2)
        Else
            STRING_OUT = XXTX + vbCrLf + TEXT_BOARD_02
        End If
        STRING_OUT = STRING_OUT ' + "*"
    Else
        STRING_OUT = XXTX
    End If
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.SetText UCase(STRING_OUT)
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If Err.Number > 0 Then Exit Sub
    
    Me.WindowState = vbMinimized
End If
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
If Index = 2 Then
    On Error Resume Next
    Do
        Err.Clear
        If Clipboard.GetFormat(vbCFText) = True Then
            Err.Clear
            TEXT_BOARD_02 = Clipboard.GetText
            If Err.Number > 0 Then Sleep 200
        End If
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    XXTX = MNU_SHOW_THE_TIME_03(2)
    If InStr(TEXT_BOARD_02, "CRC32") Then
        If InStrRev(TEXT_BOARD_02, vbCrLf) > 0 Then
            STRING_OUT = XXTX + vbCrLf + Mid(TEXT_BOARD_02, InStrRev(TEXT_BOARD_02, vbCrLf) + 2)
        Else
            STRING_OUT = XXTX + vbCrLf + TEXT_BOARD_02
        End If
        STRING_OUT = STRING_OUT ' + "*"
    Else
        STRING_OUT = XXTX
    End If
    
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.SetText STRING_OUT
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If Err.Number > 0 Then Exit Sub
    
    Me.WindowState = vbMinimized

End If
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
If Index = 3 Then
    On Error Resume Next
    Do
        Err.Clear
        If Clipboard.GetFormat(vbCFText) = True Then
            Err.Clear
            TEXT_BOARD_02 = Clipboard.GetText
            If Err.Number > 0 Then Sleep 200
        End If
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    XXTX = MNU_SHOW_THE_TIME_03(3).Caption
    If InStr(TEXT_BOARD_02, "CRC32") Then
        If InStrRev(TEXT_BOARD_02, vbCrLf) > 0 Then
            STRING_OUT = XXTX + vbCrLf + Mid(TEXT_BOARD_02, InStrRev(TEXT_BOARD_02, vbCrLf) + 2)
        Else
            STRING_OUT = XXTX + vbCrLf + TEXT_BOARD_02
        End If
        STRING_OUT = STRING_OUT ' + "*"
    Else
        STRING_OUT = XXTX
    End If
    

    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.SetText STRING_OUT
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If Err.Number > 0 Then Exit Sub
    Me.WindowState = vbMinimized
End If
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------



If Index = 8 Then
    ' ---------------------------------------------------------------------
    ' LABEL6 WAS HERE
    ' ---------------------------------------------------------------------
    ' DONT_MIN = True
    ' Call LabelCRC3_Click
    ' ---------------------------------------------------------------------
    ' Thu 18-Oct-2018 13:28:00
    ' ---------------------------------------------------------------------
    On Error Resume Next
    Do
        Err.Clear
        If Clipboard.GetFormat(vbCFText) = True Then
            Err.Clear
            TEXT_BOARD_02 = Clipboard.GetText
            If Err.Number > 0 Then Sleep 200
        End If
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    ' ---------------------------------------------------------------------
    ' EXAMPLE -- GOT -- MNU_SHOW_THE_TIME_03(1).Caption
    ' XXTX = "Tue 06-Oct-2020 20:55:08"
    ' ---------------------------------------------------------------------
    ' TEXT_BOARD_02
    ' "Text Size  425 -- CRC32 B34D8B41"
    ' ---------------------------------------------------------------------
    ' NOW UCASE OUTPUT
    ' ---------------------------------------------------------------------
    XXTX = MNU_SHOW_THE_TIME_03(1).Caption
    ' ---------------------------------------------------------------------
    If InStr(TEXT_BOARD_02, "CRC32") Then
        If InStrRev(TEXT_BOARD_02, vbCrLf) > 0 Then
            STRING_OUT = XXTX + vbCrLf + Mid(TEXT_BOARD_02, InStrRev(TEXT_BOARD_02, vbCrLf) + 2)
        Else
            
            STRING_OUT = XXTX + vbCrLf + TEXT_BOARD_02
            STRING_OUT = Mid(STRING_OUT, 1, 17) + Mid(STRING_OUT, 21)
            STRING_OUT = Replace(STRING_OUT, "SIZE ", "")
            STRING_OUT = Replace(STRING_OUT, " -- ", " ")
        End If
        STRING_OUT = STRING_OUT ' + "*"
    Else
        STRING_OUT = XXTX
    End If
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.SetText UCase(STRING_OUT)
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If Err.Number > 0 Then Exit Sub
    
    Me.WindowState = vbMinimized

End If


If Index = 9 Then
    ' ---------------------------------------------------------------------
    ' LABEL6 WAS HERE
    ' ---------------------------------------------------------------------
    ' DONT_MIN = True
    ' Call LabelCRC3_Click
    ' ---------------------------------------------------------------------
    ' Thu 18-Oct-2018 13:28:00
    ' ---------------------------------------------------------------------
    On Error Resume Next
    Do
        Err.Clear
        If Clipboard.GetFormat(vbCFText) = True Then
            Err.Clear
            TEXT_BOARD_02 = Clipboard.GetText
            If Err.Number > 0 Then Sleep 200
        End If
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    ' ---------------------------------------------------------------------
    ' EXAMPLE -- GOT -- MNU_SHOW_THE_TIME_03(1).Caption
    ' XXTX = "Tue 06-Oct-2020 20:55:08"
    ' ---------------------------------------------------------------------
    ' TEXT_BOARD_02
    ' "Text Size  425 -- CRC32 B34D8B41"
    ' ---------------------------------------------------------------------
    ' NOW UCASE OUTPUT
    ' ---------------------------------------------------------------------
    XXTX = MNU_SHOW_THE_TIME_03(1).Caption
    ' ---------------------------------------------------------------------
    If InStr(TEXT_BOARD_02, "CRC32") Then
        If InStrRev(TEXT_BOARD_02, vbCrLf) > 0 Then
            STRING_OUT = XXTX + vbCrLf + Mid(TEXT_BOARD_02, InStrRev(TEXT_BOARD_02, vbCrLf) + 2)
        Else
            STRING_OUT = TEXT_BOARD_02
            STRING_OUT1 = Trim(Str(Val(Mid(STRING_OUT, 6, 5))))
            STRING_OUT2 = Mid(STRING_OUT, Len(STRING_OUT) - 7)
            STRING_OUT = STRING_OUT1 + " " + STRING_OUT2
            
'            STRING_OUT = Replace(STRING_OUT, "SIZE ", "")
'            STRING_OUT = Replace(STRING_OUT, " -- ", " ")
        End If
        STRING_OUT = STRING_OUT ' + "*"
    Else
        STRING_OUT = XXTX
    End If
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.Clear
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    On Error Resume Next
    Do
        Err.Clear
        Clipboard.SetText UCase(STRING_OUT)
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
    If Err.Number > 0 Then Exit Sub
    
    Me.WindowState = vbMinimized

End If





End Sub

Private Sub MMM_SHOW_THE_TIME_ARR_Click(Index As Integer)


If Index = 4 Then
    ' -------------------------------------------------
    ' ADJUST WAY TIME DIFFER DISPLAY LOOK VARIABLE HERE
    ' VARTIME_DIFFERENCE = Format(F1, "0") + " DAY " + Format(F2, "0") + " HOUR " + Format(F3_2, "0") + " MINUTE"
    ' -------------------------------------------------
    MNU_TIME_DIFFERENCE_Click
    
'MMM_SHOW_THE_TIME_ARR(4).FontSize = 13
'If MMM_SHOW_THE_TIME_ARR(4) = "~ ON" Then
'    MMM_SHOW_THE_TIME_ARR(4) = "~ OFF"
'Else
'    MMM_SHOW_THE_TIME_ARR(4) = "~ ON"
'End If
End If

If Index < 4 Then
    Call LabelCRC5_Click
    Call MNU_SHOW_THE_TIME_03_Click(Index)
End If

If Index = 5 Then
    Call MNU_PUT_TIME_CLIPBOARD_END_BUTTON_MENU_005_Click
End If
If Index = 6 Then
    Call MNU_CALC_COOP_40_Click
End If
If Index = 7 Then
    Call MNU_CALC_COOP_20_Click
End If

If Index = 8 Then
    Call LabelCRC5_Click
    Call MNU_SHOW_THE_TIME_03_Click(Index)
End If
If Index = 9 Then
    Call LabelCRC5_Click
    Call MNU_SHOW_THE_TIME_03_Click(Index)
End If


End Sub

Sub TIMER_END_BUTTON_MENU_005_TIMER()
    
    TIME_VAR = Format(Now, "YYYY-MM-DD")
    MMM_SHOW_THE_TIME_ARR(5).Caption = TIME_VAR

End Sub

Public Sub MNU_PUT_TIME_CLIPBOARD_END_BUTTON_MENU_005_Click()
    Me.WindowState = vbMinimized
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    TIME_VAR = Format(Now, "YYYY-MM-DD")
'    MMM_SHOW_THE_TIME_ARR
    Clipboard.SetText TIME_VAR
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_PUT_TIME_CLIPBOARD_END_BUTTON_MENU_005_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If
End Sub

Public Sub MNU_SPACE_DASH_TO_UNDERSCORE_Click()

    If MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER = "" Then
        Do
            On Error Resume Next
            Err.Clear
            MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER = Clipboard.GetText
            On Error GoTo 0
            If Err.Number > 0 Then Sleep 500
        Loop Until Err.Number = 0
    End If
    
    i = MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER
    i = Replace(i, " ", "_")
    i = Replace(i, "-", "_")
    i = Replace(i, "-", "_")
    i = Replace(i, ".", "_")
    i = Replace(i, "?", "_")
    i = Replace(i, "'", "_")
    i = Replace(i, "`", "_")
    i = Replace(i, """", "_")
    ' RENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS__MA_.MP4 -- BATCH
    ' RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH
    MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER = i
    
    Me.WindowState = vbMinimized
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER
    If Err.Number > 0 Then
        ' PUBLIC HERE -- NAME OF SUB WHOLE
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_SPACE_DASH_TO_UNDERSCORE_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
        Exit Sub
    End If

    MNU_SPACE_DASH_TO_UNDERSCORE_VAR_CLIPPER = ""
End Sub

Private Sub MNU_STRIP_MEDIA_INFO_Click()

'----
'How to split a string with linefeed character ( <LF>) as a delimiter in VB6?-VBForums
'http://www.vbforums.com/showthread.php?765013-How-to-split-a-string-with-linefeed-character-(-lt-LF-gt-)-as-a-delimiter-in-VB6
'----

Dim TEXT_BOARD_02, STRING_OUT As String
Dim XXTX As String

TEXT_BOARD_02 = Text2.Text

Dim StrArray
Dim R_COUNT

StrArray = Split(TEXT_BOARD_02, vbCrLf)
TEXT_BOARD_02 = ""
Dim O_LEN_S
For R_COUNT = 0 To UBound(StrArray)
    MODIFY_STRING = RTrim(StrArray(R_COUNT))
    XM_1 = InStr(MODIFY_STRING, ":")
    XM_2 = InStr(MODIFY_STRING, "  ")
    If XM_1 > 0 Then
        i = MODIFY_STRING
        i = Mid(i, 1, XM_2)
        MODIFY_STRING = i
        LEN_S = Len(i)
        If LEN_S > O_LEN_S Then
        O_LEN_S = LEN_S
        End If
    End If
Next



For R_COUNT = 0 To UBound(StrArray)
    MODIFY_STRING = RTrim(StrArray(R_COUNT))
    XM_1 = InStr(MODIFY_STRING, ":")
    XM_2 = InStr(MODIFY_STRING, "  ")
    If XM_1 > 0 Then
        i = MODIFY_STRING
        XI = Mid(i, 1, XM_2)
        PAD = String(O_LEN_S - Len(XI) + 2, "-")
        i = Mid(i, 1, XM_2) + ":" + PAD + Mid(i, XM_1)
        MODIFY_STRING = i
    End If
    TEXT_BOARD_02 = TEXT_BOARD_02 + RTrim(MODIFY_STRING) + vbCrLf
Next


On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText TEXT_BOARD_02
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

Beep

If Err.Number > 0 Then Exit Sub

Me.WindowState = vbMinimized

End Sub

Private Sub MNU_TIME_BACKWARD_800_DOOR_BELL_Click()

Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4 As String

Dim A_DATE_TIME
Dim GET_TIME_VAR As Date

Do
    On Error Resume Next
    Err.Clear
    VAR_ST_1 = Clipboard.GetText
    On Error GoTo 0
    If Err.Number > 0 Then Sleep 500
Loop Until Err.Number = 0

If Not IsNumeric(VAR_ST_1) Then
    VAR_ST_1 = Text2.Text
End If

GET_TIME_VAR = Now - TimeSerial(0, 800 - Val(VAR_ST_1), 0)
A_DATE_TIME = Format(GET_TIME_VAR, "DDD HH:MM:00 Am/Pm")

VAR_ST_4 = A_DATE_TIME

MNU_TIME_BACKWARD_800_DOOR_BELL.Caption = "TIME BACKWARD 800 DOOR BELL _ " + VAR_ST_4

If IsNumeric(VAR_ST_1) Then
    VAR_ST_4 = NumToString(VAR_ST_1)
    VAR_ST_4 = VAR_ST_1 + " __ " + Trim(Mnu_Fix1st_Letters(VAR_ST_4)) + " __ " + A_DATE_TIME_2
End If

If VAR_ST_4 <> A_DATE_TIME Then
    Do
        On Error Resume Next
        Clipboard.Clear
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> "" Then Sleep 500
    Loop Until VAR_ST_2 = ""
    Sleep 100
    Do
        On Error Resume Next
        Clipboard.SetText VAR_ST_4, vbCFText
        Sleep 100
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> VAR_ST_4 Then Sleep 800
    Loop Until VAR_ST_2 = VAR_ST_4

'    Me.WindowState = vbMinimized

End If

'Beep

'Me.WindowState = vbMinimized

End Sub


Public Sub MNU_TIME_DIFFERENCE_Click()
    Me.WindowState = vbMinimized
    On Error Resume Next
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText VARTIME_DIFFERENCE
    If Err.Number > 0 Then
        VAR_TIMER_CLIPBOARD_TIMER_RETRY = "MNU_TIME_DIFFERENCE_Click"
        TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
    End If
End Sub


Private Sub MNU_VideoDownloaderUltimate_exe_Click()

    Me.WindowState = vbMinimized
    Shell "C:\ProgramData\VideoDownloaderUltimateWinApp\VideoDownloaderUltimate.exe"

End Sub

Private Sub Text2_DblClick()
Text2.Text = ""
End Sub

Public Sub TIMER_CLIPBOARD_TIMER_RETRY_Timer()
    ' ---------------------------------------
    ' CLIPBOARD_TIMER
    ' MAKE ALL CALLBACK CODE PUBLIC
    ' ---------------------------------------
    Dim Form As Form
    For Each Form In Forms
       If Form.Name = Me.Name Then
            CallByName Form, VAR_TIMER_CLIPBOARD_TIMER_RETRY, VbMethod
        End If
    Next
    TIMER_CLIPBOARD_TIMER_RETRY.Enabled = False
End Sub

Function REM_POWERSHELL(IN_KEEPER)
    If Mid(IN_KEEPER, 1, 1) = "#" Then
        IN_KEEPER = Replace(IN_KEEPER, "# --", "") ' REM LINE POWERSHELL
    End If
    If Mid(IN_KEEPER, 1, 1) = "#" Then
        IN_KEEPER = Mid(IN_KEEPER, 2)              ' REM LINE POWERSHELL
    End If
    REM_POWERSHELL = IN_KEEPER
End Function
Function DATE_STRING_FROM_TO_WORD(IN_KEEPER)
    IN_KEEPER = Replace(IN_KEEPER, " ; FROM", "")
    IN_KEEPER = Replace(IN_KEEPER, " ; TO", "")
    IN_KEEPER = Replace(IN_KEEPER, " TO", "")
    IN_KEEPER = Replace(IN_KEEPER, " FROM", "")
    IN_KEEPER = Trim(IN_KEEPER)
    DATE_STRING_FROM_TO_WORD = IN_KEEPER
End Function



Private Sub MNU_TIME_DIFFERENCE_CODE_Click()

' -----------------------------
' LAST CODE EDITOR ROUTINE HERE
' MNU_TIME_DIFFERENCE_CODE_Click()
' Thu 23-Jan-2020 17:54:29
' -----------------------------

Dim VAR_ST_1
Dim VAR_ST_2 As Double
Dim VAR_ST_4
Dim VAR_ST_5 As Double

Dim A_DATE_TIME
Dim GET_TIME_VAR As Date

Dim ORIGINAL_DATE

If Len(Text2.Text) > 1000 Then Exit Sub

MNU_TIME_DIFFERENCE.Caption = "[__ TIME DIFFERENCE __]"
MMM_SHOW_THE_TIME_ARR(4).BackColor = Label1(5).BackColor

VAR_ST_1 = Text2.Text

VAR_ST_1 = VAR_ST_1 + vbCr
VAR_SPLITT = Split(VAR_ST_1, vbCr)
For R_COUNT = 0 To UBound(VAR_SPLITT)
    If R_COUNT = 0 Then
        VAR_ST_1 = VAR_SPLITT(R_COUNT)
        VAR_ST_1 = Replace(VAR_ST_1, "'#", "") ' REM LINE FOR VB
        VAR_ST_1 = Replace(VAR_ST_1, "!", "")  ' REM LINE FOR ADBLOCKER
        VAR_ST_1 = Replace(VAR_ST_1, "'", "")
        VAR_ST_1 = Replace(VAR_ST_1, vbLf, "")
        VAR_ST_1 = REM_POWERSHELL(VAR_ST_1)
        VAR_ST_1 = DATE_STRING_FROM_TO_WORD(VAR_ST_1)
        VAR_ST_1 = Trim(VAR_ST_1)
        
    End If
    If R_COUNT = 1 Then
        VAR_ST_4 = VAR_SPLITT(R_COUNT)
        VAR_ST_4 = Replace(VAR_ST_4, "'#", "") ' REM LINE FOR VB
        VAR_ST_4 = Replace(VAR_ST_4, "!", "")  ' REM LINE FOR ADBLOCKER
        VAR_ST_4 = Replace(VAR_ST_4, "'", "")
        VAR_ST_4 = Replace(VAR_ST_4, vbLf, "")
        VAR_ST_4 = REM_POWERSHELL(VAR_ST_4)
        VAR_ST_4 = DATE_STRING_FROM_TO_WORD(VAR_ST_4)
        VAR_ST_4 = Trim(VAR_ST_4)
    End If
Next

VAR_ST_1 = Trim(VAR_ST_1)
If InStr(VAR_ST_1, vbCr) > 0 Then
    VAR_ST_1 = Mid(VAR_ST_1, 1, InStr(VAR_ST_1, vbCr) - 1)
End If
If InStr(VAR_ST_1, vbLf) > 0 Then
    VAR_ST_1 = Mid(VAR_ST_1, 1, InStr(VAR_ST_1, vbLf) - 1)
End If

' ---------------------------------------
' TESTER
' ---------------------------------------
' If IsIDE Then VAR_ST_1 = "08:05:24 10:20:00"
' If IsIDE Then VAR_ST_1 = "23-Jan-2020 11:01:12"
' ---------------------------------------
ORIGINAL_DATE = VAR_ST_1

' ---------------------------------------
' ONLY HOLD ONE DATE OF LINE
' MIGHT IMPROVE ADD TWO LINE GETTER
' ---------------------------------------
VAR_ST_1 = UCase(VAR_ST_1)
VAR_ST_4 = UCase(VAR_ST_4)
For R1 = 1 To 12
    ' Debug.Print UCase(Format(DateSerial(0, R1, 1), "MMM"))
    VAR_ST_1 = Replace(VAR_ST_1, UCase(Format(DateSerial(2000, R1, 1), "MMM")), Format(R1, "00"))
    VAR_ST_4 = Replace(VAR_ST_4, UCase(Format(DateSerial(2000, R1, 1), "MMM")), Format(R1, "00"))
Next
For R1 = 1 To 7
    ' Debug.Print UCase(Format(DateSerial(2018, 1, R1), "DDD")) ' START ON MONDAY
    VAR_ST_1 = Replace(VAR_ST_1, UCase(Format(DateSerial(2018, 1, R1), "DDD")), "")
    VAR_ST_4 = Replace(VAR_ST_4, UCase(Format(DateSerial(2018, 1, R1), "DDD")), "")
Next
FIND_TT_1 = VAR_ST_1
FIND_TT_4 = VAR_ST_4
FIND_TT_1 = Replace(FIND_TT_1, ":", "")
FIND_TT_1 = Replace(FIND_TT_1, ";", "")  ' REM LINE FOR AUTOHOTKEY
FIND_TT_1 = Replace(FIND_TT_1, "'#", "") ' REM LINE FOR VB
FIND_TT_1 = Replace(FIND_TT_1, "'", "")  ' REM LINE FOR VB
FIND_TT_1 = Replace(FIND_TT_1, "!", "")  ' REM LINE FOR ADBLOCKER
FIND_TT_1 = Replace(FIND_TT_1, " ", "")
FIND_TT_1 = Replace(FIND_TT_1, "-", "")
FIND_TT_1 = Replace(FIND_TT_1, "@REM", "") ' FROM LINE -- OF BATCH COMMAND LANGAUGE
FIND_TT_1 = Replace(FIND_TT_1, "REM", "") ' FROM LINE -- OF BATCH COMMAND LANGAUGE
FIND_TT_1 = REM_POWERSHELL(FIND_TT_1)
FIND_TT_1 = DATE_STRING_FROM_TO_WORD(FIND_TT_1)

FIND_TT_4 = Replace(FIND_TT_4, ":", "")
FIND_TT_4 = Replace(FIND_TT_4, ";", "")  ' REM LINE FOR AUTOHOTKEY
FIND_TT_4 = Replace(FIND_TT_4, "'#", "") ' REM LINE FOR VB
FIND_TT_4 = Replace(FIND_TT_4, "'", "")  ' REM LINE FOR VB
FIND_TT_4 = Replace(FIND_TT_4, "!", "")  ' REM LINE FOR ADBLOCKER
FIND_TT_4 = Replace(FIND_TT_4, " ", "")
FIND_TT_4 = Replace(FIND_TT_4, "-", "")
FIND_TT_4 = Replace(FIND_TT_4, "@REM", "")   ' FROM LINE -- OF BATCH COMMAND LANGAUGE
FIND_TT_4 = Replace(FIND_TT_4, "REM", "")    ' FROM LINE -- OF BATCH COMMAND LANGAUGE
FIND_TT_4 = REM_POWERSHELL(FIND_TT_4)
FIND_TT_4 = DATE_STRING_FROM_TO_WORD(FIND_TT_4)

FIND_TT_1 = Replace(FIND_TT_1, "FROM__", "") ' HEADER SCRIPT TIMER
FIND_TT_1 = Replace(FIND_TT_1, "TO__", "")   ' HEADER SCRIPT TIMER
FIND_TT_4 = Replace(FIND_TT_4, "FROM__", "") ' HEADER SCRIPT TIMER
FIND_TT_4 = Replace(FIND_TT_4, "TO__", "")   ' HEADER SCRIPT TIMER

XF1 = InStr(FIND_TT_1, "HOUR")
If XF1 > 0 Then FIND_TT_1 = Mid(FIND_TT_1, 1, XF1 - 1) ' HEADER SCRIPT TIMER
XF1 = InStr(FIND_TT_4, "HOUR")
If XF1 > 0 Then FIND_TT_4 = Mid(FIND_TT_4, 1, XF1 - 1) ' HEADER SCRIPT TIMER

If Len(FIND_TT_1) < 8 Then Exit Sub

FIND_TT_1 = Replace(FIND_TT_1, "  ", " ")
FIND_TT_4 = Replace(FIND_TT_4, "  ", " ")

FIND_TT_1 = Trim(FIND_TT_1)
FIND_TT_1 = Mid(FIND_TT_1, 1, 3)
FIND_TT_4 = Trim(FIND_TT_4)
FIND_TT_4 = Mid(FIND_TT_4, 1, 3)

VAR_ST_1 = Replace(VAR_ST_1, ";", "")  ' REM LINE FOR AUTOHOTKEY
VAR_ST_1 = Replace(VAR_ST_1, "'#", "") ' REM LINE FOR VB
VAR_ST_1 = Replace(VAR_ST_1, "'", "")  ' REM LINE FOR VB
VAR_ST_1 = Replace(VAR_ST_1, "!", "")  ' REM LINE FOR ADBLOCKER
VAR_ST_4 = Replace(VAR_ST_4, ";", "")  ' REM LINE FOR AUTOHOTKEY
VAR_ST_4 = Replace(VAR_ST_4, "'#", "") ' REM LINE FOR VB
VAR_ST_4 = Replace(VAR_ST_4, "'", "")  ' REM LINE FOR VB
VAR_ST_1 = Trim(VAR_ST_1)
VAR_ST_4 = Trim(VAR_ST_4)

VAR_ST_1 = Replace(VAR_ST_1, "--", "") ' FROM LINE -- OF RS232 LOGGER
VAR_ST_4 = Replace(VAR_ST_4, "--", "") ' FROM LINE -- OF RS232 LOGGER
VAR_ST_1 = Replace(VAR_ST_1, "  ", " ")
VAR_ST_4 = Replace(VAR_ST_4, "  ", " ")

VAR_ST_1 = Replace(VAR_ST_1, "@REM", "") ' FROM LINE -- OF BATCH COMMAND LANGAUGE
VAR_ST_4 = Replace(VAR_ST_4, "@REM", "") ' FROM LINE -- OF BATCH COMMAND LANGAUGE
VAR_ST_1 = Replace(VAR_ST_1, "REM", "")  ' FROM LINE -- OF BATCH COMMAND LANGAUGE
VAR_ST_4 = Replace(VAR_ST_4, "REM", "")  ' FROM LINE -- OF BATCH COMMAND LANGAUGE

VAR_ST_1 = Replace(VAR_ST_1, "FROM __ ", "") ' HEADER SCRIPT TIMER
VAR_ST_1 = Replace(VAR_ST_1, "TO  __ ", "")   ' HEADER SCRIPT TIMER
VAR_ST_4 = Replace(VAR_ST_4, "FROM __ ", "") ' HEADER SCRIPT TIMER
VAR_ST_4 = Replace(VAR_ST_4, "TO  __ ", "")   ' HEADER SCRIPT TIMER

XF1 = InStr(VAR_ST_1, " HOUR")
If XF1 > 0 Then VAR_ST_1 = Mid(VAR_ST_1, 1, XF1 - 1) ' HEADER SCRIPT TIMER
XF1 = InStr(VAR_ST_4, " HOUR")
XF2 = InStrRev(VAR_ST_4, " ", XF1 - 1)               ' HOUR HAS A DIGIT BEFORE WITH SPACE EITHER SIDE
If XF1 > 0 And XF2 > 0 Then VAR_ST_4 = Mid(VAR_ST_4, 1, XF2 - 1) ' HEADER SCRIPT TIMER
If XF1 > 0 And XF2 = 0 Then VAR_ST_4 = Mid(VAR_ST_4, 1, 1)  ' HEADER SCRIPT TIMER
' ---------------------------------------------------------------------
' Mon 17-Feb-2020 04:44:23
' WORK
' ---------------------------------------------------------------------


VAR_ST_1 = Replace(VAR_ST_1, "  ", " ")
VAR_ST_4 = Replace(VAR_ST_4, "  ", " ")
VAR_ST_1 = Trim(VAR_ST_1)
VAR_ST_4 = Trim(VAR_ST_4)

TRIG1 = False
For R1 = 1 To Len(VAR_ST_1) - 1
    If Mid(VAR_ST_1, R1, 1) = ":" Then TRIG1 = True
    If Mid(VAR_ST_1, R1 + 1, 1) = ":" And TRIG1 = True Then
        XXIN = InStr(R1, VAR_ST_1, " ")
        If XXIN > 0 Then
            VAR_ST_1 = Mid(VAR_ST_1, 1, XXIN - 1)
        End If
        Exit For
    End If
Next
TRIG1 = False
For R1 = 1 To Len(VAR_ST_4)
    If Mid(VAR_ST_4, R1, 1) = ":" Then TRIG1 = True
    If Mid(VAR_ST_4, R1 + 1, 1) = ":" And TRIG1 = True Then
        XXIN = InStr(R1, VAR_ST_4, " ")
        If XXIN > 0 Then
            VAR_ST_4 = Mid(VAR_ST_4, 1, XXIN - 1)
        End If
        Exit For
    End If
Next

EXIT_SUB_NOTE = 0
If IsNumeric(FIND_TT_1) = True Then EXIT_SUB_NOTE = 1
If IsNumeric(FIND_TT_4) = True Then EXIT_SUB_NOTE = 1
If EXIT_SUB_NOTE = 0 Then
    Exit Sub
End If
VAR_ST_1 = Trim(VAR_ST_1)
VAR_ST_4 = Trim(VAR_ST_4)


' Sat 04-Jan-2020 09:47:56
' Sat 04-Jan-2020 12:11:09
' Sat 11-Jan-2020 21:21:17
' Sun 12-Jan-2020 00:22:00

' TRIM LINE DATE SO EXTRA CHAR AT END GONE
On Error Resume Next
If IsDate(DateValue(VAR_ST_1)) And IsDate(TimeValue(VAR_ST_1)) Then
    If Err.Number > 0 Then
        XF1 = InStr(VAR_ST_1, " ")
        If XF1 > 0 Then
            XF2 = InStr(XF1 + 1, VAR_ST_1, " ")
        End If
        If XF2 > 0 Then XF1 = XF2
        If XF1 > 0 Then
            VAR_ST_1 = Mid(VAR_ST_1, 1, XF1 - 1)
        End If
    End If
End If


' TRIM LINE DATE SO EXTRA CHAR AT END GONE
On Error Resume Next
If IsDate(DateValue(VAR_ST_4)) And IsDate(TimeValue(VAR_ST_4)) Then
    If Err.Number > 0 Then
        XF1 = InStr(VAR_ST_4, " ")
        If XF1 > 0 Then
            XF2 = InStr(XF1 + 1, VAR_ST_4, " ")
        End If
        If XF2 > 0 Then XF1 = XF2
        If XF1 > 0 Then
            VAR_ST_4 = Mid(VAR_ST_4, 1, XF1 - 1)
        End If
    End If
End If


Err.Clear
GOT_DATE = ""
If IsNumeric(FIND_TT_1) = True Then
On Error Resume Next
If IsDate(DateValue(VAR_ST_1)) And IsDate(TimeValue(VAR_ST_1)) Then
    If Err.Number = 0 Then
        GOT_DATE = True
        VAR_ST_1 = DateValue(VAR_ST_1) + TimeValue(VAR_ST_1)
    End If
End If
End If
If Err.Number > 0 Then Exit Sub
On Error GoTo 0



On Error Resume Next
Err.Clear
If IsNumeric(FIND_TT_4) = True Then
If InStr(FIND_TT_4, "£") = 0 Then
If InStr(FIND_TT_4, "$") = 0 Then
IT2 = IsDate(DateValue(VAR_ST_4))
IT2 = IsDate(TimeValue(VAR_ST_4))
If Err.Number = 0 Then
If IsDate(DateValue(VAR_ST_4)) And IsDate(TimeValue(VAR_ST_4)) Then
    If Err.Number = 0 Then
        GOT_DATE = True
        VAR_ST_4 = DateValue(VAR_ST_4) + TimeValue(VAR_ST_4)
        NOW_USER = "2 * DATE"
    End If
End If
End If
End If
End If
End If
If Err.Number > 0 Then Exit Sub
On Error GoTo 0

On Error Resume Next

If IsNumeric(FIND_TT_1) = True Then
If IsNumeric(FIND_TT_4) = False Then

If IsDate(DateValue(VAR_ST_1)) = True Then
XF1 = False: XF2 = False
If VAR_ST_4 <> "" Then
    XF1 = IsDate(DateValue(VAR_ST_4))
    XF2 = IsDate(TimeValue(VAR_ST_4))
End If
If XF1 = False And XF2 = False Then
    GOT_DATE = True
    NOW_VAR = Now
    VAR_ST_4 = DateValue(NOW_VAR) + TimeValue(NOW_VAR)
    NOW_USER = "NOW USER"
End If
End If
End If
End If


' 22-Jan-2020 00:06:34
' 22-Jan-2020 04:48:03


If GOT_DATE = "" Then
If InStr(VAR_ST_1, " ") > 0 Then
    VAR_ST_4 = Mid(VAR_ST_1, InStr(VAR_ST_1, " ") + 1)
    VAR_ST_1 = Mid(VAR_ST_1, 1, InStr(VAR_ST_1, " ") - 1)
End If
End If

' 23-Jan-2020 11:01:12

' ---------------------------------------
' IF NONE DATE OF VARIABLE -- POP ONE ONE OF DAY NOW
' OR DAY BEFORE IF OVER
' ---------------------------------------
If IsDate(VAR_ST_1) = False Then
    ' MsgBox "NOT A DATE VALUE"
    Exit Sub
End If
If NOW_VAR = 0 Then
If IsDate(VAR_ST_4) = False Then
    NOW_VAR = Now
    NOW_USER = "NOW USER"
Else
    NOW_VAR = VAR_ST_4
    NOW_USER = "2 * DATE"
End If
End If
' ---------------------------------------
' IF NONE DATE OF VARIABLE -- POP ONE ONE OF DAY NOW
' OR DAY BEFORE IF OVER
' ---------------------------------------

'' IF GOT TWO TIME AND FIRST HIGHER TO MIDNIGHT DAY BEFORE
'If DateValue(VAR_ST_1) = False And TimeValue(VAR_ST_1) > 0 Then
'    VAR_ST_1 = DateValue(Now) + TimeValue(VAR_ST_1)
'End If
'If VAR_ST_1 > Now Then
'    VAR_ST_1 = VAR_ST_1 - 1
'End If

If DateValue(VAR_ST_1) > 0 And TimeValue(VAR_ST_1) > 0 Then
    VAR_ST_1 = DateValue(VAR_ST_1) + TimeValue(VAR_ST_1)
End If
If DateValue(VAR_ST_1) = False And TimeValue(VAR_ST_1) > 0 Then
    VAR_ST_1 = DateValue(Now) + TimeValue(VAR_ST_1)
End If
If DateValue(VAR_ST_1) = True And TimeValue(VAR_ST_1) = False Then
    VAR_ST_1 = DateValue(VAR_ST_4) + TimeValue(Now)
End If
' ---------------------------------------
' ---------------------------------------
If GOT_DATE = "" Then
    If IsNumeric(Mid(VAR_ST_4, 1, 2)) Then
        If DateValue(VAR_ST_4) > 0 And TimeValue(VAR_ST_4) > 0 Then
            VAR_ST_4 = DateValue(VAR_ST_4) + TimeValue(VAR_ST_4)
            NOW_VAR = VAR_ST_4
        End If
        If DateValue(VAR_ST_4) = False And TimeValue(VAR_ST_4) > 0 Then
            VAR_ST_4 = DateValue(Now) + TimeValue(VAR_ST_4)
            NOW_VAR = VAR_ST_4
        End If
        If DateValue(VAR_ST_4) = True And TimeValue(VAR_ST_4) = False Then
            VAR_ST_4 = DateValue(VAR_ST_4) + TimeValue(Now)
            NOW_VAR = VAR_ST_4
        End If
    End If
End If

' ---------------------------------------------------
' 19 IS FULL LENGTH DATE AND TIME -- LEAVE HERE FOR LATER
' ---------------------------------------------------
' If Len(VAR_ST_1) < 19 Then
' NOW_VAR WOULD BE ADJUST THE SEOCND ARRAY DATE VALUE
' BUT DONE ANOTHER WAY AS FIND HARD TO READ
' ---------------------------------------------------
If NOW_USER = "NOW USER" Then
    ' DIFF_GET_TIME_VAR = DateDiff("S", DateValue(VAR_ST_1) + TimeValue(VAR_ST_1), NOW_VAR) ' "n" MINUTE
    DIFF_GET_TIME_VAR = DateDiff("S", VAR_ST_1, NOW_VAR)  ' "n" MINUTE
End If
If NOW_USER = "2 * DATE" Then
    DIFF_GET_TIME_VAR = DateDiff("S", VAR_ST_1, VAR_ST_4) ' "n" MINUTE
End If
DIFF_GET_TIME_VAR_2 = DIFF_GET_TIME_VAR / 60 ' MINUTE

tMm = DIFF_GET_TIME_VAR ' IN SECONDS
F1 = Int((tMm / 60 ^ 2) / 24)  ' DAY
If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)
F2 = Int((tMm / 60 ^ 2))       ' HOUR
F3 = Int(tMm / 60)             ' MINUTE
F3_2 = Int(tMm / 60) Mod 60    ' MINUTE MOD
F4 = tMm Mod 60                ' SECOND MOD

' ----------------------------------------------------------------------
' Sun 26-Jan-2020 14:15:00
' ----------------------------------------------------------------------
'
' ----------------------------------------------------------------------
If DateValue(VAR_ST_1) = DateValue(NOW_VAR) Then
    DISPLAY_STR = Str(DateValue(VAR_ST_1)) + " --" + Str(TimeValue(VAR_ST_1)) + " -- " + Str(TimeValue(NOW_VAR))
Else
    DISPLAY_STR = Str(DateValue(VAR_ST_1)) + "--" + Str(TimeValue(VAR_ST_1)) + " -- " + Str(DateValue(NOW_VAR)) + "--" + Str(TimeValue(NOW_VAR))
    ' LEAD SPACE WITH STR() GIVE NOT REQUIRE TRIM(STR()) WITH DATEVALUE() TIMEVALUE()
End If
DISPLAY_STR = "  " + DISPLAY_STR
MNU_TIME_DIFFERENCE.Caption = "[__ TIME DIFFERENCE__ " + Format(F2, "0") + " SECOND -- " + Format(F3, "0") + " MINUTE -- " + Format(F2, "0") + " HOUR -- " + NOW_USER + " -- " + DISPLAY_STR + " __]"
MMM_SHOW_THE_TIME_ARR(4).BackColor = Label1(4).BackColor
MMControl8.Command = "prev"
MMControl8.Command = "Play"

VARTIME_DIFFERENCE = Format(F1, "0") + " DAY " + Format(F2, "0") + " HOUR " + Format(F3_2, "0") + " MINUTE"
If F1 = 0 Then VARTIME_DIFFERENCE = Replace(VARTIME_DIFFERENCE, Format(F1, "0") + " DAY ", "")
If F2 = 0 Then VARTIME_DIFFERENCE = Replace(VARTIME_DIFFERENCE, Format(F2, "0") + " HOUR ", "")
End Sub




Private Sub Timer_GIVEN_FILENAME_AND_CMD_LINE_STRING_FOR_COMPARE_Timer()

Dim FILENAME_PATH
Dim DIR_PATH
Dim FR1
Dim r
Dim FILE_TO_KILL

FILENAME_PATH = App.Path + "\# DATA\" + GetComputerName + "\Data _ Hammer Button Compare\"

If Dir(FILENAME_PATH, vbDirectory) = "" Then
    Exit Sub
End If

DIR_PATH = FILENAME_PATH
File1.Path = DIR_PATH
File1.Pattern = "*.TXT"
File1.Refresh

FILE_TO_KILL = ""
For r = 0 To File1.ListCount
    If InStr(File1.List(r), "COMPARE BUTTON HAMMER") > 0 Then
        FILE_TO_KILL = File1.List(r)
    End If
Next

On Error GoTo EXIT_SUB

If FILE_TO_KILL <> "" Then
    Kill DIR_PATH + FILE_TO_KILL
    
    ' -------------------
    ' CHECK CRC32 SAME OR
    ' -------------------
    If CRC_1 <> CRC_2 Then
        Call Label1_Click(3)
    End If
    EXIT_TRUE = True
    Unload Me
End If


EXIT_SUB:

End Sub


Private Sub Label_FindWindowPart_VB_CLIPVIEWER_AUTO_BUTTON_COMPARE_Click()
    Dim FILENAME_PATH
    Dim DIR_PATH
    Dim FR1
    Dim r
    
    FILENAME_PATH = App.Path + "\# DATA\" + GetComputerName + "\Data _ Hammer Button Compare\"
    
    If Dir(FILENAME_PATH, vbDirectory) = "" Then
        CreateFolderTree FILENAME_PATH
    End If
    
    ' RESET ALL WORK BEFORE BEGIN
    ' ---------------------------------------------------
    DIR_PATH = FILENAME_PATH
    File1.Path = DIR_PATH
    File1.Pattern = "*.TXT"
    
    For r = 0 To File1.ListCount
        If InStr(File1.List(r), "COMPARE BUTTON HAMMER") > 0 Then
            Kill DIR_PATH + File1.List(r)
        End If
    Next
    
    If RUN_AT_FORM_LOAD = True Then
        RUN_AT_FORM_LOAD = False
        Exit Sub
    End If
    
    For r = 1 To FindWindowPart_VB_CLIPVIEWER_AUTO_BUTTON_COMPARE
        FR1 = FreeFile
        Open FILENAME_PATH + "COMPARE BUTTON HAMMER " + Trim(Str(r)) + ".TXT" For Output As #FR1
        Close #FR1
    Next
    

End Sub
Private Sub Label_FindWindowPart_VB_CLIPVIEWER_AUTO_COMP_2()

If FILENAME_AND_COMMAND_STRING_GIVE <> True Then
    Exit Sub
End If

Dim FILENAME_PATH
Dim DIR_PATH
Dim FR1
Dim r

FILENAME_PATH = App.Path + "\# DATA\" + GetComputerName + "\Data _ Hammer Button Compare\"

If Dir(FILENAME_PATH, vbDirectory) = "" Then
    Exit Sub
End If

If FindWindowPart_VB_CLIPVIEWER_AUTO_BUTTON_COMPARE > 1 Then
    Exit Sub
End If

' RESET ALL WORK AS EXIT IF ONE INSTANCE APP LEFT
' -----------------------------------------------
DIR_PATH = FILENAME_PATH
File1.Path = DIR_PATH
File1.Pattern = "*.TXT"

For r = 0 To File1.ListCount
    If InStr(File1.List(r), "COMPARE BUTTON HAMMER") > 0 Then
        Kill DIR_PATH + File1.List(r)
    End If
Next



End Sub


Private Sub LabelFILE1_Click()

On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText Mid(LabelFILE1.Caption, 6)
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

If Err.Number > 0 Then Exit Sub

Me.WindowState = vbMinimized


End Sub

Private Sub LabelFILE2_Click()
On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText Mid(LabelFILE2.Caption, 6)
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

If Err.Number > 0 Then Exit Sub

Me.WindowState = vbMinimized


End Sub

Private Sub MNU_CAP_TEXT_Click()

Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4

VAR_ST_1 = Clipboard.GetText

VAR_ST_4 = VAR_ST_1

VAR_ST_1 = UCase(VAR_ST_1)

If VAR_ST_1 <> VAR_ST_4 Then
    
    Do
        On Error Resume Next
        Clipboard.Clear
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> "" Then Sleep 500
    Loop Until VAR_ST_2 = ""
    Sleep 100
    Do
        On Error Resume Next
        Clipboard.SetText VAR_ST_1, vbCFText
        Sleep 100
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> VAR_ST_1 Then Sleep 800
    Loop Until VAR_ST_2 = VAR_ST_1

    Me.WindowState = vbMinimized

End If

'Beep

End Sub


Private Sub MNU_CAPS_TXT_LOW_Click()
Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4

VAR_ST_1 = Clipboard.GetText

VAR_ST_4 = VAR_ST_1

VAR_ST_1 = LCase(VAR_ST_1)

If VAR_ST_1 <> VAR_ST_4 Then
    
    Do
        On Error Resume Next
        Clipboard.Clear
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> "" Then Sleep 500
    Loop Until VAR_ST_2 = ""
    Sleep 100
    Do
        On Error Resume Next
        Clipboard.SetText VAR_ST_1, vbCFText
        Sleep 100
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> VAR_ST_1 Then Sleep 800
    Loop Until VAR_ST_2 = VAR_ST_1

    Me.WindowState = vbMinimized

End If

'Beep

End Sub


Private Sub MNU_CLIP_TIME_TESTER_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText Time$

End Sub

Private Sub MNU_KILL_OTHER_CLIPVIEWER_Click()
    Call FindWindowPart_VB_CLIPVIEWER_CLOSE
End Sub



Private Sub LabelCRC1_Click()

TEXT_CLIPPER = "SIZE " + Format(Len(Text2.Text), "#####0000") + " -- " + Replace(LabelCRC1.Caption, "_1 =", "")

Clipboard.Clear
Clipboard.SetText TEXT_CLIPPER

Me.WindowState = vbMinimized

End Sub

Private Sub LabelCRC3_Click()

'TEXT_CLIPPER = "Text Size " + Str(Len(Text2.Text)) + " -- " + Replace(LabelCRC1.Caption, "_1 =", "")

On Error Resume Next

DONT_UPDATE_CLIPBOARD_THIS_ONE = True
Clipboard.Clear
DONT_UPDATE_CLIPBOARD_THIS_ONE = True
Clipboard.SetText LabelCRC3.Caption + "*"

If Err.Number > 0 Then Exit Sub

If DONT_MIN = True Then
    DONT_MIN = False
    Exit Sub
End If

Me.WindowState = vbMinimized


End Sub

Private Sub LabelCRC4_Click()
    DONT_MIN = True
    Call LabelCRC3_Click

End Sub

Private Sub LabelCRC5_Click()
    DONT_MIN = True
    Call LabelCRC3_Click

End Sub

Private Sub MNU_MESSENGER_GIVE_CRC32_Click()
    Call LabelCRC3_Click
End Sub

Private Sub MESSENGER_GIVE_CRC32_Click()

End Sub

Private Sub MNU_API_RESET_Click()
' -------------------------------------------------------
' ALLOW API CLIPBOARDER IF REQUEST MADE BY BUTTON OR MENU
' WHEN NOT ENABLE BEFORE AS TEXT MODE WILL BE ON
' -------------------------------------------------------

FRM_CLIPTEST_02_ENABLE = True

'FRM_ClipTest_02.ctlClipboard1.EndClipboardViewer
'FRM_ClipTest_02.ctlClipboard1.StartClipboardViewer

' EXIT_TRUE = True
' Beep
FRM_ClipTest_02.ctlClipboard1.EndClipboardViewer
'Unload FRM_ClipTest_02
DoEvents
'Load FRM_ClipTest_02
FRM_ClipTest_02.ctlClipboard1.StartClipboardViewer

End Sub

Private Sub Mnu_Exit_Click()
    If IsIDE = True Then End
    EXIT_TRUE = True
    Unload Me
End Sub

Private Sub MNU_GIVE_ME_68_DASH_LINE_2_Click()
' MNU_GIVE_ME_68_DASH_LINE_2

Dim STRING_OUT

STRING_OUT = String(69, "-") + vbCrLf

On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText STRING_OUT
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

If Err.Number > 0 Then Exit Sub

Me.WindowState = vbMinimized

End Sub

Private Sub MNU_GIVE_ME_68_DASH_LINE_REMARK_Click()

Dim STRING_OUT

STRING_OUT = "' " + String(68, "-") + vbCrLf

On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText STRING_OUT
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

If Err.Number > 0 Then Exit Sub

Me.WindowState = vbMinimized

End Sub


Private Sub MNU_TRIM_END_OF_LINE_SPACE_Click()

'----
'How to split a string with linefeed character ( <LF>) as a delimiter in VB6?-VBForums
'http://www.vbforums.com/showthread.php?765013-How-to-split-a-string-with-linefeed-character-(-lt-LF-gt-)-as-a-delimiter-in-VB6
'----

Dim TEXT_BOARD_02, STRING_OUT As String
Dim XXTX As String

TEXT_BOARD_02 = Text2.Text

Dim StrArray
Dim R_COUNT

StrArray = Split(TEXT_BOARD_02, vbCrLf)
TEXT_BOARD_02 = ""
For R_COUNT = 0 To UBound(StrArray)
    TEXT_BOARD_02 = TEXT_BOARD_02 + RTrim(StrArray(R_COUNT)) + vbCrLf
Next
'MsgBox TEXT_BOARD_02
'End

On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText TEXT_BOARD_02
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

Beep

If Err.Number > 0 Then Exit Sub

Me.WindowState = vbMinimized

End Sub

Private Sub MNU_TRIM_SPACE_ON_EACH_LINE_AND_WITHIN_Click()

'----
'How to split a string with linefeed character ( <LF>) as a delimiter in VB6?-VBForums
'http://www.vbforums.com/showthread.php?765013-How-to-split-a-string-with-linefeed-character-(-lt-LF-gt-)-as-a-delimiter-in-VB6
'----

Dim TEXT_BOARD_02, STRING_OUT As String
Dim XXTX As String

TEXT_BOARD_02 = Text2.Text

Dim StrArray
Dim R_COUNT

StrArray = Split(TEXT_BOARD_02, vbCrLf)
TEXT_BOARD_02 = ""
For R_COUNT = 0 To UBound(StrArray)
    TEXT_BOARD_02 = TEXT_BOARD_02 + Replace(RTrim(StrArray(R_COUNT)), " ", "") + vbCrLf
Next
'MsgBox TEXT_BOARD_02
'End

On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText TEXT_BOARD_02
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

Beep

If Err.Number > 0 Then Exit Sub

Me.WindowState = vbMinimized

End Sub


Private Sub MNU_TRIM_SPACE_ON_EACH_LINE_AND_WITHIN_UPTO_FIRST_COLON_Click()

'----
'How to split a string with linefeed character ( <LF>) as a delimiter in VB6?-VBForums
'http://www.vbforums.com/showthread.php?765013-How-to-split-a-string-with-linefeed-character-(-lt-LF-gt-)-as-a-delimiter-in-VB6
'----

Dim TEXT_BOARD_02, STRING_OUT As String
Dim XXTX As String

TEXT_BOARD_02 = Text2.Text

Dim StrArray
Dim R_COUNT
Dim TS_01
Dim FIND_STRING_01
Dim FIND_STRING_02

StrArray = Split(TEXT_BOARD_02, vbCrLf)
TEXT_BOARD_02 = ""
For R_COUNT = 0 To UBound(StrArray)
    TS_01 = InStr(StrArray(R_COUNT), ":")
    If TS_01 > 0 Then
        FIND_STRING_01 = Replace(LTrim(Mid(StrArray(R_COUNT), 1, TS_01)), " ", "")
        FIND_STRING_02 = RTrim(Mid(StrArray(R_COUNT), TS_01 + 1))
    Else
        FIND_STRING_01 = Trim(StrArray(R_COUNT))
        FIND_STRING_02 = ""
    End If
    
    TEXT_BOARD_02 = TEXT_BOARD_02 + FIND_STRING_01 + FIND_STRING_02 + vbCrLf
Next
'MsgBox TEXT_BOARD_02
'End

On Error Resume Next
Do
    Err.Clear
    Clipboard.Clear
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

On Error Resume Next
Do
    Err.Clear
    Clipboard.SetText TEXT_BOARD_02
    If Err.Number > 0 Then Sleep 200
Loop Until Err.Number = 0
On Error GoTo 0

Beep

If Err.Number > 0 Then Exit Sub

Me.WindowState = vbMinimized

End Sub



Private Sub Mnu_UrlLoad_Click()
    If Mid$(LastURL, 2, 2) = ":\" Then
        Shell "Explorer.exe /e," + LastURL, vbNormalFocus
    Else
        vLaunch (LastURL)
    End If
End Sub

Private Sub VB_RUN_NOT_WHEN_IDE_AND_THEN_SHOW()
' -----------------------------------------------------------------
' IS SHORT CUT NOT SET ADMINISTATOR FOR THIS PARTICALUR CODE IN EXE
' MODE IT WILL NOT RUN LOAD ITSELF ON THE VB COMPILER
' BUT AND WILL RUN OTHER PROGRAM
' AS TALK OTHER PROGRAM ARE OKAY BUT THEN MAYBE SET AS ADMIN
' AMSWER GIVEN
' AUG 28 SUN
' -----------------------------------------------------------------

If 1 = 1 And IsIDE = True Then
    MNU_VB.Enabled = False
    Beep
    MsgBox "NOT WHEN ISIDE = TRUE"
    Exit Sub
End If

Dim ReturnHwnd As Long
Dim i
'VB ONLY WANTS THE 1ST OF THE 2 HWND
'ReturnHwnd = FindWindowPartVB("ClipBoard Logger - Microsoft Visual Basic[")

'DONT NEED ABOVE USE THIS
x1 = FindWindow("wndclass_desked_gsk", vbNullString)
'------------------------------------------------
'FIND THE WINDOW PRIZE TREASURE CHEST BUST BLOSSOM
'TRAIN SPOTTER
'------------------------------------------------
'-----------------------------------------------------------
'CHECK CLASS NAME IS VB IDE WINDOW WE WANT OR IF ANOTHER NOT
'-----------------------------------------------------------
X2 = GetWindowTitle(x1)
If InStr(X2, " - Microsoft Visual Basic [") > 0 Then
    'RUN A WINDOW SPY
    WIN_SPY_NAME = "ClipBoard Logger"
    If InStr(X2, WIN_SPY_NAME) > 0 Then
    
        MsgBox "DON'T RUN VB IDE - LOADED"
        i = GetWindowState(x1)
        If i = vbMinimized Then
            SHOWVAR = SW_SHOWDEFAULT
            ShowWindow ReturnHwnd, SHOWVAR
        End If
        
        EXIT_TRUE = True
        Unload Me
        Exit Sub
    End If
End If

'BETTER LINE NEXT DON'T USE VB MENU IF ISIDE
'------------------------------------------------
'TEMP WORK AROUND
'OVER DRIVE
'OVER RIDE
'------------------------------------------------
'FINDWINDOW PART PROBLEM IN EXE AND WHEN IN WIN 7
'------------------------------------------------
'WIN 7 PROBLEM MUST USE EXTRA LAST LEFT SQUARE BRACKET OF SERACH END ABOVE
'------------------------------------------------
If ReturnHwnd > 0 Then
    If MsgBox("ReturnHwnd" + Str(ReturnHwnd) + " -- " + "MeHwnd" + Str(Me.hWnd) + vbCrLf + "VB Code " + vbCrLf + WIN_SPY_NAME + vbCrLf + " already Running - Sure Want to Run VB IDE", vbYesNo) = vbYes Then
        WANT_TO_RUN_ANYWAY = True
    End If
End If

If ReturnHwnd > 0 Then
    i = GetWindowState(ReturnHwnd1)
    If i = vbMinimized Then
'        SHOWVAR = SW_RESTORE
'        SHOWVAR = SW_SHOW
'        SHOWVAR = SW_SHOWNA
'        SHOWVAR = SW_SHOWDEFAULT
'        SHOWVAR = SW_SHOWNORMAL
'        SHOWVAR = SW_SHOWMAXIMIZED
        SHOWVAR = SW_SHOWDEFAULT
        ShowWindow ReturnHwnd, SHOWVAR
        'GUESS MAYBE
        'SetForegroundWindow (ReturnHwnd)
        DoEvents
    End If
   
    'MADE REDUNDANT CODE BY CONDICTION HERE AND ABOVE
    If WANT_TO_RUN_ANYWAY = False Then
        MsgBox "EXIT AS FOUND WINODW OF VB AND PUT TO SHOW FOCUS"
        EXIT_TRUE = True
        Unload Me
        Exit Sub
    End If
End If

Dim CODER_VBP_FILE_NAME_2
Dim VB_1, VB_2, VB_3
VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VB_1) <> "" Then VB_3 = VB_1
If Dir(VB_2) <> "" Then VB_3 = VB_2
'------------------------------------------------------
'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
'------------------------------------------------------
Dim OBJSHELL
Set OBJSHELL = CreateObject("Wscript.Shell")
CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
OBJSHELL.RUN """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
Set OBJSHELL = Nothing
EXIT_TRUE = True
Unload Me
End Sub

Private Sub Mnu_VB_ME_Click()
    Call VB_RUN_NOT_WHEN_IDE_AND_THEN_SHOW
    
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3
    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    Call CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    
    Dim OBJSHELL
    Set OBJSHELL = CreateObject("WSCRIPT.SHELL")
    OBJSHELL.RUN """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set OBJSHELL = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Sub CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    ' --------------------------------------------------------------------
    ' FIND IF THE IS NOT CORRECT VERSION AND SIMPLE PUT CORRECT
    ' ONE MY COMPUTER SEEM PROBLEM WITH BASIC
    ' TALK HAS WRONG VERISON ONCE EVERY FEW DAY
    ' AND EDIT IT TO WHAT NORM VERISON SUPPOSED TO BE AND FINE
    ' DISCOVERY -- IT SEEM TO HAPPEN JUST AFTER A COMPILE SOMETIME_
    ' [ Monday 15:37:30 Pm_17 June 2019 ]
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    Dim r, FR
    Dim VAR_STRING As String

    FR = FreeFile
    Open CODER_VBP_FILE_NAME_2 For Binary As FR
        VAR_STRING = Space(LOF(FR))
        Get #FR, , VAR_STRING
    Close FR
    ' --------------------------------------------------------------
    ' # .VBP # .VBP" -- # MSCOMC -- # OCX -- # MSCOMCTL # CTL # .CTL
    ' --------------------------------------------------------------
    ' FOUND TO BE ERROR
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    ' CORRECT VALUE
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    XR1 = InStr(UCase(VAR_STRING), UCase("A1}#2.1; MSCOMCTL.OCX"))
    XR2 = InStr(UCase(VAR_STRING), UCase("A1}#2.#")) ' ---- "A1}#2.#0; mscomctl.OCX"
    XR3 = InStr(UCase(VAR_STRING), UCase("#0; MSCOMCTL.OCX"))
    If XR1 = 0 And XR2 > 0 And XR3 > 0 Then
        GET_01 = InStr(UCase(VAR_STRING), UCase("MSCOMCTL.OCX")) - XR2
        Mid(VAR_STRING, XR2, GET_01 + Len("MSCOMCTL.OCX")) = "A1}#2.1#0; MSCOMCTL.OCX"
        ' MsgBox "MSCOMCTL.OCX" + vbCrLf + "WRONG VERSION -- CHANGE TO" + vbCrLf + "#2.1# -- MSCOMCTL.OCX"), vbMsgBoxSetForeground
        GO_NEXT_IN = True
    End If
    If GO_NEXT_IN = True Then
        If Dir(CODER_VBP_FILE_NAME_2) <> "" Then
            Kill CODER_VBP_FILE_NAME_2
        End If
        FR = FreeFile
        Open CODER_VBP_FILE_NAME_2 For Binary As FR
            Put #FR, , VAR_STRING
        Close FR
    End If
End Sub
Private Sub MNU_FOLDER_DIGITAL_STILL_CAMERA_Click()
    ' Beep
    Me.WindowState = vbMinimized
    Dim FOLDER_Path
    FOLDER_Path = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG"
    Dir1.Path = FOLDER_Path
    Dim TD
    For r = 0 To Dir1.ListCount - 1
        FIND_PATH_1 = Dir1.List(r)
        FIND_PATH_2 = Mid(FIND_PATH_1, InStrRev(FIND_PATH_1, "\"))
        If Mid(FIND_PATH_2, 1, 3) = "\20" Then
            TD = Dir1.List(r)
        End If
    Next
    FOLDER_Path = TD
    
    Shell "EXPLORER /SELECT, """ + FOLDER_Path + """", vbNormalFocus
    EXIT_TRUE = True
    ' Unload Me
    'End
End Sub
Private Sub MNU_CAMERA_VIDEO_FOLDER_4G_Click()
    Me.WindowState = vbMinimized
    Shell "EXPLORER \\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4", vbMaximizedFocus
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    ' Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub




Private Sub Text2_Change()

'If IsIDE = True Then
'    If KeyCode = 27 Then End
'End If

' Label1(7).Caption = "COMPARE VEIWER"
Call MNU_FIND_DIFFER_Click

If Len(Trim(Text2)) >= 1 Then
    CHAR_STRING = Mid(Trim(Text2.Text), 1, 1)
    If CHAR_STRING <> "" Then
        MNU_CHAR_CODE.Caption = "[__ 1ST CHARACTOR CODE -- """ + CHAR_STRING + """ -- DEC " + Format(Asc(CHAR_STRING), "000") + " -- HEX " + Hex(Asc(CHAR_STRING)) + " __]" ' Ë
    Else
        MNU_CHAR_CODE.Caption = "[__ 1ST CHARACTOR CODE -- EMPTY OF __]"
    End If
Else
    MNU_CHAR_CODE.Caption = "[__ 1ST CHARACTOR CODE -- EMPTY OF __]"
End If

If WAIT_FORM_BOOT_BEFORE_GO = False Then
    Exit Sub
End If

Call MNU_TIME_DIFFERENCE_CODE_Click

Dim SAME_VALUE_TRUE_FALSE
Dim CRC_1
Dim CRC_2
Dim TextCRC3 As String
Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32


Dim HH, HH_2, HH_3, HH1, HH2, HH3, HH4, HH5, HH6, HH7, HH8, HH9

' Tue 10-Sep-2019 20:21:27
On Error Resume Next
HH = "Tue 10-Sep-2019 20:21:27"
If DateValue(HH) > 0 Then
    HH_2 = DateValue(HH)
End If
If TimeValue(HH) > 0 Then
    HH_2 = HH2 + TimeValue(HH)
End If
If HH_2 > 0 Then
    ' MNU_SHOW_THE_TIME_03(0).Caption = ""
End If
On Error GoTo 0


If Text2.Text = "" Then
    TextCRC3 = "00000000"
    TEXT_CLIPPER = "Text Size 0" + " -- CRC32 00000000"
End If

If Text2.Text <> "" Then
    CRC_1 = Hex(m_CRC.CalculateString(Text2.Text))
    CRC_2 = Hex(m_CRC.CalculateString(Text3.Text))
    CRC_1_LC = Hex(m_CRC.CalculateString(LCase(Text2.Text)))
    CRC_2_LC = Hex(m_CRC.CalculateString(LCase(Text3.Text)))
    ' LABEL_COMPARE_CASE_SENSITIVE.FontSize = 11
    LABEL_COMPARE_CASE_SENSITIVE.Alignment = 0
    ' ALREADY ALINE ARANGER
    ' ----------------------------------------------------
    ' LABEL_COMPARE_CASE_SENSITIVE.Left = Label1(3).Left
    ' LABEL_COMPARE_CASE_SENSITIVE.Width = Label1(3).Width

    If CRC_1_LC = CRC_2_LC Then
        LABEL_COMPARE_CASE_SENSITIVE.Caption = "Compare Case Not Sensitive -- SAME"
        LABEL_COMPARE_CASE_SENSITIVE.BackColor = Label1(4).BackColor
    Else
        LABEL_COMPARE_CASE_SENSITIVE.Caption = "Compare Case Not Sensitive -- DIFFERENT"
        LABEL_COMPARE_CASE_SENSITIVE.BackColor = RGB(174, 224, 237) ' Label1(5).BackColor
    End If
    
    If InStrRev(Text2.Text, "~") Then
        ' TEXT SIZE CHARACTOR COUNT ADJUSTER -3 HERE -- InStrRev(Text2.Text, "~") - 3)
        ' ----------------------------------------------------------------------------
        TextCRC3 = ""
        If InStrRev(Text2.Text, "~") > 1 Then
        TextCRC3 = Mid(Text2.Text, 1, InStrRev(Text2.Text, "~") - 3) + vbCrLf
        End If
        ' Debug.Print "--" + TextCRC3 + "--" + vbCrLf + Str(Len(TextCRC3))
        TEXT_CLIPPER = "SIZE " + Format(Len(TextCRC3), "#####0000") + " -- CRC32 " + Hex(m_CRC.CalculateString(TextCRC3))
    Else
        TextCRC3 = Text2.Text
        TEXT_CLIPPER = "SIZE " + Format(Len(TextCRC3), "#####0000") + " -- CRC32 " + Hex(m_CRC.CalculateString(TextCRC3))
    End If
End If

If InStr(Hex(m_CRC.CalculateString(TextCRC3)), "6") > 0 Then
    LabelCRC3.BackColor = Label1(5).BackColor
    LabelCRC4.BackColor = Label1(5).BackColor
    LabelCRC5.BackColor = Label1(5).BackColor
Else
    LabelCRC3.BackColor = Label1(4).BackColor
    LabelCRC4.BackColor = Label1(4).BackColor
    LabelCRC5.BackColor = Label1(4).BackColor
End If

If InStr(Hex(m_CRC.CalculateString(TextCRC3)), "6") = 0 Then
    If InStrRev(Text2.Text, "~") > 0 Then
        MMControl1.Command = "prev"
        MMControl1.Command = "Play"
    End If
End If
If InStr(Hex(m_CRC.CalculateString(TextCRC3)), "6") = 0 Then
    If MMCONTROL4_FIRST_LOAD_SOUND_OFF = False Then
'    If InStr(Text2.Text, "Text Size ") > 0 Then
'    If InStr(Text2.Text, "CRC32") > 0 Then
    If Len(TextCRC3) > 100 Then
    If InStrRev(Text2.Text, "~") = 0 Then
        MMControl4.Command = "prev"
        MMControl4.Command = "Play"
    End If
    End If
    End If
'    End If
'    End If
    MMCONTROL4_FIRST_LOAD_SOUND_OFF = False
End If

LabelCRC1.Caption = "CRC32_1 = " + CRC_1
LabelCRC2.Caption = "CRC32_2 = " + CRC_2
LabelCRC3.Caption = TEXT_CLIPPER
LabelCRC4.Caption = TEXT_CLIPPER

Call SET_TEXT_2_AND_TEXT_COLOUR_AND_TEXT_IF_SAME


'Label_SIZE_1.Caption = "SIZE =" + Str(Len(Text2.Text)) + " -- " + SAME_VALUE_TRUE_FALSE
'Label_SIZE_2.Caption = "SIZE =" + Str(Len(Text3.Text)) + " -- " + SAME_VALUE_TRUE_FALSE


End Sub

Private Sub Text3_Change()

Dim SAME_VALUE_TRUE_FALSE
Dim CRC_2
Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32

If RUN_ONCE_01 = False Then
    Label1(2).BackColor = Label1(5).BackColor
End If

RUN_ONCE_01 = False


If Text3.Text = "" Then
    TEXT_CLIPPER = "Text Size 0" + " -- CRC32 00000000"
End If

If Text3.Text <> "" Then
    CRC_2 = Hex(m_CRC.CalculateString(Text3.Text))
    LabelCRC2.Caption = "CRC32_2 = " + CRC_2
End If

Call SET_TEXT_2_AND_TEXT_COLOUR_AND_TEXT_IF_SAME
'
'If Text2.Text = Text3.Text Then
'    SAME_VALUE_TRUE_FALSE = "SAME IS TRUE"
'    Label1(3).BackColor = Label1(5).BackColor
'Else
'    SAME_VALUE_TRUE_FALSE = "SAME IS FALSE"
'    Label1(3).BackColor = Label1(4).BackColor
'End If

'Label_SIZE_2.Caption = "SIZE =" + Str(Len(Text3.Text)) + " -- " + SAME_VALUE_TRUE_FALSE

End Sub

Sub SET_TEXT_2_AND_TEXT_COLOUR_AND_TEXT_IF_SAME()

Dim SAME_VALUE_TRUE_FALSE

If Text2.Text = Text3.Text Then
    SAME_VALUE_TRUE_FALSE = "COMPARE IS TRUE"
    Label1(3).BackColor = Label1(5).BackColor
    Label_SIZE_1.BackColor = Label1(4).BackColor
    Label_SIZE_2.BackColor = Label1(4).BackColor
    LabelCRC1.BackColor = Label1(4).BackColor
    LabelCRC2.BackColor = Label1(4).BackColor
    Label1(6).BackColor = Label1(5).BackColor
    Label1(6).Caption = "Compare is Equal"
Else
    SAME_VALUE_TRUE_FALSE = "COMPARE IS FALSE"
    Label1(3).BackColor = Label1(4).BackColor
    Label1(3).BackColor = RGB(174, 224, 237)
    Label_SIZE_1.BackColor = Label_COLOR_RED.BackColor
    Label_SIZE_2.BackColor = Label_COLOR_RED.BackColor
    LabelCRC1.BackColor = Label_COLOR_RED.BackColor
    LabelCRC2.BackColor = Label_COLOR_RED.BackColor
    Label1(6).BackColor = Label_COLOR_RED.BackColor
    Label1(6).Caption = "Compare is Not Same"
End If

Label_SIZE_1.Caption = "SIZE =" + Str(Len(Text2.Text)) + " -- " + SAME_VALUE_TRUE_FALSE
Label_SIZE_2.Caption = "SIZE =" + Str(Len(Text3.Text)) + " -- " + SAME_VALUE_TRUE_FALSE

End Sub

Sub SET_TEXT_2_AND_TEXT_COLOUR_AND_TEXT_IF_SAME_WHEN_FILE()

Dim SAME_VALUE_TRUE_FALSE

If CRC_1 = CRC_2 Then
    SAME_VALUE_TRUE_FALSE = "COMPARE IS TRUE"
    Label1(3).BackColor = Label1(5).BackColor
    Label_SIZE_1.BackColor = Label1(4).BackColor
    Label_SIZE_2.BackColor = Label1(4).BackColor
    LabelCRC1.BackColor = Label1(4).BackColor
    LabelCRC2.BackColor = Label1(4).BackColor
    Label1(6).BackColor = Label1(5).BackColor
    Label1(6).Caption = "Compare is Equal"
Else
    SAME_VALUE_TRUE_FALSE = "COMPARE IS FALSE"
    Label1(3).BackColor = Label1(4).BackColor
    Label1(3).BackColor = RGB(174, 224, 237)
    Label_SIZE_1.BackColor = Label_COLOR_RED.BackColor
    Label_SIZE_2.BackColor = Label_COLOR_RED.BackColor
    LabelCRC1.BackColor = Label_COLOR_RED.BackColor
    LabelCRC2.BackColor = Label_COLOR_RED.BackColor
    Label1(6).BackColor = Label_COLOR_RED.BackColor
    Label1(6).Caption = "Compare is Not Same"
End If

If FILE_NAME_SIZE_1 > 0 Or FILE_NAME_SIZE_2 > 0 Then
    Label_SIZE_1.Caption = "SIZE =" + Str(FILE_NAME_SIZE_1) + " -- " + SAME_VALUE_TRUE_FALSE
    Label_SIZE_2.Caption = "SIZE =" + Str(FILE_NAME_SIZE_2) + " -- " + SAME_VALUE_TRUE_FALSE
Else
    Label_SIZE_1.Caption = "SIZE =" + Str(Len(Text2.Text)) + " -- " + SAME_VALUE_TRUE_FALSE
    Label_SIZE_2.Caption = "SIZE =" + Str(Len(Text3.Text)) + " -- " + SAME_VALUE_TRUE_FALSE
End If

End Sub

Private Sub Timer_CLIPBOARD_API_Timer()

    If FRM_CLIPTEST_02_ENABLE = False Then Exit Sub

    ' Call MNU_API_RESET_Click
    
'    FRM_ClipTest_02.ctlClipboard1.EndClipboardViewer
    Unload FRM_ClipTest_02
    DoEvents
    Load FRM_ClipTest_02
'    FRM_ClipTest_02.ctlClipboard1.StartClipboardViewer

End Sub

Private Sub Timer_EXIT_Timer()
    
    If Date = "27/01/2020" Then Exit Sub
    If Date = "28/01/2020" Then Exit Sub
    
    If IsIDE = False Then Timer_EXIT.Enabled = False
    
    If GetForegroundWindow = Me.hWnd Then
        If GetAsyncKeyState(27) Then End 'ESCAPE KEY
    End If

End Sub


Public Sub Timer_RESET_API_CLIPPER_Timer()

    If FRM_CLIPTEST_02_ENABLE = False Then Exit Sub

    ' FRM_ClipTest_02.ctlClipboard1.EndClipboardViewer
    Unload FRM_ClipTest_02
    DoEvents
    Load FRM_ClipTest_02
    'FRM_ClipTest_02.ctlClipboard1.StartClipboardViewer
    Timer_RESET_API_CLIPPER.Enabled = False
'    FRM_ClipTest.Timer_RESET_API_CLIPPER

End Sub


Public Sub Timer_Test_Logic_Timer()

Timer_Test_Logic.Enabled = False
Timer_Test_Logic.Interval = 10000

Exit Sub

'REMOVE THIS ONE AUG 08 2K SIXTEEN


'HOOK_CLIPBOARD_API_lOADED = True

If IsIDE = True Then
    Timer_Test_Logic.Interval = 4000
Else
    Timer_Test_Logic.Interval = 1000
End If

Timer_API_Test_Logick_Var1 = Now

End Sub


Private Sub MNU_COMPARE_Click()
    
Call Label1_Click(0)

End Sub

Private Sub MNU_GIVE_ME_TIME_WITHER_UTC_Click()

Dim i

i = MNU_GIVE_ME_TIME_WITHER_UTC.Caption

' DEFAULT AT STARTOR
'----------------------------------------------------
If TO_SETTER = False Then
    TO_SETTER = True
    i = "GIVE ME TIME AND UNI_ UTC Time Toggle = YES"
    i = "GIVE ME TIME AND UNI_ UTC Time Toggle = NOT"
    MNU_GIVE_ME_TIME_WITHER_UTC.Caption = i
    Call frmListMenu.SET_MENU_PADD_WORK
    Exit Sub
End If

If InStr(i, "NOT") Then
    MNU_GIVE_ME_TIME_WITHER_UTC.Caption = Replace(i, "NOT", "YES")
    Exit Sub
End If
If InStr(i, "YES") Then
    MNU_GIVE_ME_TIME_WITHER_UTC.Caption = Replace(i, "YES", "NOT")
End If


End Sub


Sub SET_MENU_PADD_WORK_NOT_USE_RELOCATE_TO_OWN_FORM()

Dim i_Menu_Count, i_Form_Counter
Dim i_Menu_Not_Visa_Count

Dim Control As Control, Label_44, LABEL_48

Dim R_NEXT

Dim Text_Checker_Form_Menu As String

Dim MENU_ITEM_VAR
Dim i

For i = 0 To Forms.Count - 1
    
    For Each Control In Forms(i).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            If Control.Visible = True Then
                i_Menu_Count = i_Menu_Count + 1
            End If
            i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count + 1

        End If
    Next
Next

i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count - i_Menu_Count

MNU_MENU_ITEM_COUNT.Caption = "Menu Item Count = " + Str(i_Menu_Count) + " &&" + Str(i_Menu_Not_Visa_Count) + " Not Visible"
'Mnu_Form_Count.Caption = "Form Counter = " + Str(Forms.Count - 1) '  + " Really 7"
'Mnu_Form_Count.Visible = False

i_Menu_Count = 0



For i = 0 To Forms.Count - 1
    Text_Checker_Form_Menu = ""
    frmListMenu.GetMenuInfo_Not_Indented GetMenu(Forms(i).hWnd), 0, "", Text_Checker_Form_Menu
    Text_Checker_Form_Menu = UCase(Text_Checker_Form_Menu)
    For Each Control In Forms(i).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            MENU_ITEM_VAR = Replace(Control.Caption, "[__ ", "")
            MENU_ITEM_VAR = Replace(MENU_ITEM_VAR, " __]", "")
            MENU_ITEM_VAR = UCase(Trim(MENU_ITEM_VAR))
            If InStr(Text_Checker_Form_Menu, "SUB MENU ----" + MENU_ITEM_VAR) = 0 Then
                
                'i_Menu_Count = i_Menu_Count + 1
                If InStr(Trim(Control.Caption), "[__ ") = 0 Then
                    Label_44 = Trim(Control.Caption)
                    'LABEL_48 = Replace(LABEL_44, " ", "_")
                    LABEL_48 = Label_44
                    LABEL_48 = Replace(LABEL_48, "___", "__")
                    LABEL_48 = "[__ " + LABEL_48 + " __]"
                    LABEL_48 = Replace(LABEL_48, "[__ [__ ", "[__ ")
                    LABEL_48 = Replace(LABEL_48, " __] __]", " __]")
                    If LABEL_48 <> Label_44 Then
                        Control.Caption = LABEL_48
                    End If
                End If
            End If
        End If
    Next
Next

'Stop

''MNU_BRing_Front
''---------------
'i_Form_Counter = Forms.Count
'i_Form_Counter = 0
''for each f
''For i = 0 To Forms.Count - 1
''    Load Forms(i)
''    Show Forms(i)
''Next
'
'Dim Form As Form
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    Show Form
'    'Set Form = Nothing
'Next Form
'
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    'Set Form = Nothing
'Next Form

i_Form_Counter = Forms.Count - 1
Me.Refresh
End Sub




Private Sub MNU_GIVE_ME_TIME_Click()

'Timer_SHOW_THE_TIME_Timer()
'---------------------------
'---------------------------

Dim XI
Dim A_DATE_TIME_UTC
Dim A_DATE_TIME As String

'--------------------------------
'get 24 hour format and add am pm
'--------------------------------
'or put back as 12 hour format
'WATCH CLOCK GO BACK PROBLEM
'--------------------------------
A_DATE_TIME = "[ " + Format(Now, "DDDD HH:MM:SS") + " " + A_DATE_TIME_PM + "_" + Format(Now, "DD MMMM YYYY") + " ]"
If Hour(Now) = 16 Then
    A_DATE_TIME = "[ " + Format(Now, "DDDD HH:MM:SS AM/PM") + "_" + Format(Now, "DD MMMM YYYY") + " ]"
End If
A_DATE_TIME = Replace(A_DATE_TIME, "PM", "Pm")
A_DATE_TIME = Replace(A_DATE_TIME, "AM", "Am")
'------------------------------------------------------
'REPLACE LAST DIGIT AS 0
'----------------------------
XI = InStr(A_DATE_TIME, "m_")
Mid(A_DATE_TIME, XI - 3, 1) = "0"
'----------------------------
'------------------------------------------------------
'------------------------------------------------------
Dim GMTBST
'Convert_To_UT
'-------------
'-----------------------------------------------------------------------------------------------
'ONE MINUTE ADJUST OFFSET IF DELAY IN TWO TIME WHEN SAME CALCULATOR MESSUREMENT THEN MAYBE ERROR
'OFFSET BY HALF 30 MINUTE
'-----------------------------------------------------------------------------------------------
If ConvertToUT(Now) + TimeSerial(0, 30, 0) < Now Then
    GMTBST = "BST BDT DST DLS 1 Hour Ahead UTC+1 CET"
Else
    GMTBST = "GMT UTC+0 WET"
End If

'--------------------------------
'get 24 hour format and add am pm
'--------------------------------
'or put back as 12 hour format
'--------------------------------
A_DATE_TIME_UTC = "[ " + Format(ConvertToUT(Now), "DDDD HH:MM:SS") + " " + A_DATE_TIME_PM + "_" + Format(ConvertToUT(Now), "DD MMMM YYYY") + " ]"
If Hour(Now) = 16 Then
    A_DATE_TIME_UTC = "[ " + Format(ConvertToUT(Now), "DDDD HH:MM:SS AM/PM") + "_" + Format(ConvertToUT(Now), "DD MMMM YYYY") + " ]"
End If
A_DATE_TIME_UTC = Replace(A_DATE_TIME_UTC, "PM", "Pm")
A_DATE_TIME_UTC = Replace(A_DATE_TIME_UTC, "AM", "Am")
'------------------------------------------------------
'REPLACE LAST DIGIT AS 0
'----------------------------
XI = InStr(A_DATE_TIME_UTC, "m_")
Mid(A_DATE_TIME_UTC, XI - 3, 1) = "0"
'----------------------------
If Day(ConvertToUT(Now)) = 7 - 1 Then A_DATE_TIME_UTC = Replace(A_DATE_TIME_UTC, Format(ConvertToUT(Now), "DD MMMM YYYY"), "Sixer " + Format(ConvertToUT(Now), "MMMM YYYY"))
If Day(Now) = 7 - 1 Then A_DATE_TIME = Replace(A_DATE_TIME, Format(Now, "DD MMMM YYYY"), "Sixer " + Format(Now, "MMMM YYYY"))


If InStr(MNU_GIVE_ME_TIME_WITHER_UTC.Caption, "=_YES") Then
    A_DATE_TIME = A_DATE_TIME + " The UK" + vbCrLf + A_DATE_TIME_UTC + " " + GMTBST
Else
    A_DATE_TIME_UTC = ""
End If
'------------------------------------------------------

Clipboard.Clear
Clipboard.SetText A_DATE_TIME

Beep
Me.WindowState = vbMinimized

End Sub

Public Function ConvertToUT(Time As Date) As Date
' convert system time to universal time and adjust for DST
Dim TZ As TIME_ZONE_INFORMATION

'01 of 02 -- SEARCH SAME REM LINE
'------------------------------------------------------------------------------------
'ERROR OF ORGINAL SOURCE -- SUPPOSED TO BE TEMP AS RETURN STRING IN FORM OF --- (GMT)
'OR OPTION HAS THAT IN TEXT STRING AND EXTRA TIME INFO
'BUT TEMP AS DATE -- IS USED THROUGH OUT THIS CODE
'------------------------------------------------------------------------------------
'ALSO MY COMPUTER IS UNABLE TO RETURN INFO ABOUT TIME ZONE -- WIN 10
'------------------------------------------------------------------------------------
'WORK AROUND ADD A ISNUMERIC DETECT

'TIMEZONE IS RETURN AS STRING AND NUMERIC IS ADJUST IS MADE AT GMT
'BUT NOT CORRECT TOWARD THE ORGINAL SOURCE CODE METHOD

'------------------------------------------------------------------------------------
'METHOD WORKED OUT
'DON'T DO STRING CONVERT METHOD TO NUMERIC
'AND MEAN ADD ANOTHER FUNCTION HERE
'------------------------------------------------------------------------------------

Dim temp As Date
Dim temp_str As String
Dim TimeZone_Numeric

Dim Rtn As Long
Rtn = GetTimeZoneInformation(TZ)
If Rtn > TIME_ZONE_ID_UNKNOWN Then

    If Rtn = TIME_ZONE_ID_STANDARD Then
    temp = DateAdd("n", TZ.Bias, Time)
    Else
    temp = DateAdd("n", (TZ.Bias + TZ.DaylightBias), Time)
    End If

End If
ConvertToUT = temp
'Temp = TimeZone
temp = TimeZone_Numeric

'Debug.Print IsNumeric(Mid("00:00:00", 1, 2))
'----------------------------------------
'Debug.Print IsNumeric(TimeZone_Numeric)
'----------------------------------------
'A DATE VAR IS RETRUNED AS NUMERIC = TRUE
'----------------------------------------
'----------------------------------------
'Debug.Print IsNumeric(TimeZone)

'If IsNumeric(IsNumeric(Mid(TimeZone, 1, 2))) Then
'    temp = TimeZone
'Else
'    temp_str = TimeZone
'End If

End Function


Sub TIME_CLIPBOARD_CHANGE_SUB()

Dim AATDIFF

If TIME_CLIPBOARD_CHANGE = 0 Then
    TIME_CLIPBOARD_CHANGE = Now
End If

AATDIFF = DateDiff("s", TIME_CLIPBOARD_CHANGE, Now)
Lab_TIME_CLIPBOARD_CHANGE.Caption = Trim(Str(AATDIFF)) + " Second API Check Worker"


End Sub



Private Sub Timer_SHOW_THE_TIME_Timer()

Call TIMER_END_BUTTON_MENU_005_TIMER
Call TIME_CLIPBOARD_CHANGE_SUB

Dim A_DATE_TIME As String ', PATH_VOLUME_NAME, R_FOR
'Dim A_DATE_TIME_TEMP As String
Dim XI

Dim RT
Dim A_DATE_TIME_02
Dim A_DATE_TIME_03
Dim A_DATE_TIME_04
Dim GET_TIME_VAR


'LCASE THE SECOND PM AM DIGIT
'BETTER

'A_DATE_TIME = "[__ " + Format(Now, "DDDD YYYY-MMM-DD-HH-MM-SS ") + A_DATE_TIME_PM + " __]"
'If Hour(Now) = 16 Then
'A_DATE_TIME = "[__ " + Format(Now, "DDDD YYYY-MMM-DD-HH-MM-SS AM/PM") + " __]"
'End If
'MsgBox Format(Now, "YYYY-MMM-DD -- HH:MM:SS - HH:MM:SS-AM/PM DDD")


GET_TIME_VAR = Now

A_DATE_TIME_PM = Mid(Format(GET_TIME_VAR, "HH-MM-SS Am/Pm"), 10)

A_DATE_TIME = "[__ " + Format(GET_TIME_VAR, "DDDD HH:MM:SS") + " " + A_DATE_TIME_PM + "__" + Format(GET_TIME_VAR, "DD MMMM YYYY") + " __]"
If Hour(GET_TIME_VAR) = 16 Then
    A_DATE_TIME = "[__ " + Format(GET_TIME_VAR, "DDDD HH:MM:SS AM/PM") + "__" + Format(GET_TIME_VAR, "DD MMMM YYYY") + " __]"
End If
A_DATE_TIME = Replace(A_DATE_TIME, "PM", "Pm")
A_DATE_TIME = Replace(A_DATE_TIME, "AM", "Am")
'------------------------------------------------------
'REPLACE LAST DIGIT AS 0
'----------------------------
XI = InStr(A_DATE_TIME, "m__")
Mid(A_DATE_TIME, XI - 3, 1) = "0"
'----------------------------

A_DATE_TIME = Replace(A_DATE_TIME, "m__", "m_")
A_DATE_TIME = Replace(A_DATE_TIME, "__", "")

For RT = 1 To 7
    A_DATE_TIME_02 = Replace(A_DATE_TIME, Format(GET_TIME_VAR + RT, "DDDD"), Format(GET_TIME_VAR + RT, "DDD"))
Next

If MNU_SHOW_THE_TIME_03(3).Caption <> A_DATE_TIME Then
    MNU_SHOW_THE_TIME_03(3).Caption = A_DATE_TIME
End If
If MNU_SHOW_THE_TIME_03(2).Caption <> A_DATE_TIME_02 Then
    MNU_SHOW_THE_TIME_03(2).Caption = A_DATE_TIME_02
End If

Dim r
Dim HH, HH_1, HH_2, HH_3, HH_4, HH1, HH2, HH3, HH4, HH5, HH6, HH7, HH8, HH9

HH = UCase("Tue 09-Sep-2019 20:30:00")
For r = 1 To 8
    HH = Trim(Replace(HH, UCase(Format(Now + r, "DDD")), ""))
Next
If DateValue(HH) > 0 Then
    HH_1 = DateValue(HH)
End If
If TimeValue(HH) > 0 Then
    HH_1 = HH_1 + TimeValue(HH)
End If
HH = UCase("Tue 11-Sep-2019 19:30:00")
For r = 1 To 8
    HH = Trim(Replace(HH, UCase(Format(Now + r, "DDD")), ""))
Next
If DateValue(HH) > 0 Then
    HH_2 = DateValue(HH)
End If
If TimeValue(HH) > 0 Then
    HH_2 = HH_2 + TimeValue(HH)
End If

HH1 = DateDiff("s", HH_1, HH_2)
If HH1 > 0 Then
    HH1 = Int(HH_1)
    HH2 = Int(HH_2) ' NOW
    HH3 = HH2 - HH1
    HH4 = HH_1 - Int(HH_1)
    HH5 = HH_2 - Int(HH_2)
    HH7 = HH5 - HH4
    HH8 = HH3 - HH7
    HH9 = Trim(Str(HH3)) + " --" + Format(DateSerial(2000 + HH3, 0, 0) + HH7, "HH:MM:SS")
End If

'HH1 = Int(HH_1)
'HH2 = Int(HH_2) ' NOW
'HH3 = HH1 - HH2
'HH9 = HH7 - HH8
'
'' HH10 = HH2 + (HH5 + HH9)
'
'If HH3 > HH4 Then HH5 = HH3 - HH4
'If HH3 < HH4 Then HH5 = HH4 - HH3
'If HH3 = HH4 Then HH5 = 0
'HH7 = GET_TIME_VAR - Int(GET_TIME_VAR)
'HH8 = HH_2 - Int(HH_2)
'If HH7 > HH8 Then HH9 = HH7 - HH8
'If HH7 < HH8 Then HH9 = HH8 - HH7
'If HH7 = HH8 Then HH9 = 0
'
'' HH2=HH2-



A_DATE_TIME_03 = Format(GET_TIME_VAR, "DDD DD-MMM-YYYY HH:MM:SS")
'REPLACE LAST DIGIT AS 0
'----------------------------
Mid(A_DATE_TIME_03, Len(A_DATE_TIME_03), 1) = "0"
'----------------------------

'If MNU_SHOW_THE_TIME_03(0).Caption <> A_DATE_TIME_03 Then
'    MNU_SHOW_THE_TIME_03(0).Caption = A_DATE_TIME_03
'End If

If MNU_SHOW_THE_TIME_03(0).Caption <> MNU_SHOW_THE_TIME_03(2).Caption Then
    MNU_SHOW_THE_TIME_03(0).Caption = MNU_SHOW_THE_TIME_03(2).Caption
End If

' YOU
' ~
' THU 08-OCT-2020 16:29:40
' TEXT SIZE  9 -- CRC32 EBD907AC

A_DATE_TIME_04 = Format(GET_TIME_VAR, "DD-MMM-YYYY HH:MM:SS DDD")
If UCase(MNU_SHOW_THE_TIME_03(1).Caption) <> UCase(A_DATE_TIME_04) Then
    MNU_SHOW_THE_TIME_03(1).Caption = UCase(A_DATE_TIME_04)
End If

' Thu 18-Oct-2018 13:28:00

'If Val(Second(GET_TIME_VAR)) Mod 10 <> 0 Then Exit Sub

'MNU_SHOW_THE_TIME.Caption = A_DATE_TIME

'MNU_SHOW_THE_TIME.Enabled = False
'MNU_SHOW_THE_TIME.Enabled = True

End Sub

Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)

        End With
       
    End If
    
    If Menu_Height <> ARCHIVE_Menu_Height Then
        TIMER_TO_RESIZE.Enabled = True
    End If
    
    ARCHIVE_Menu_Height = Menu_Height
   
End Function

Private Sub TIMER_TO_RESIZE_Timer()
    TIMER_TO_RESIZE.Enabled = False
    'NOT_RESIZE_EVENTER = False
    Call Form_Resize
    'NOT_RESIZE_EVENTER = True

End Sub

Sub TIMER_Form_Resize_TIMER()
    'Call ME_TOP_LEFT_CENTER
'    Call SET_FORM_CENTER
    Call Form_Resize
    TIMER_Form_Resize.Enabled = False
End Sub


Sub SET_FORM_CENTER()
    Dim Control As Control
    Dim XC
    Dim LET_TEXT_BOX_GO
    
    
    For Each Control In Controls
        XX_1 = Mid(UCase(Control.Name), 1, 3)
        XX_2 = UCase(Control.Name)
        If XX_1 <> "MNU" Or XX_2 = "MMM_SHOW_THE_TIME_ARR" Then
            LET_TEXT_BOX_GO = False
            If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
            On Error Resume Next
            XC = Control.Visible
            On Error GoTo 0
            If XC = True Or LET_TEXT_BOX_GO = True Then
                If Control.width + Control.Left > WX Then WX = Control.width + Control.Left
                If Control.height + Control.Top > HY Then HY = Control.height + Control.Top
            End If
        End If
    Next


    Call ME_POSITION
    
End Sub

Private Sub Timer1_Timer()
    Call TIMER_RESIZE_CHECK
End Sub

Sub TIMER_RESIZE_CHECK()

    If Me.Left < 0 Or Me.Top < 0 Then
        DO_FORM_RESIZE_ONCE = False
        TIMER_Form_Resize.Enabled = True
    End If

End Sub

Sub Form_Resize()

    On Error GoTo ENDA
    
    If Me.Left < 0 Then
        DO_FORM_RESIZE_ONCE = False
        'CALL Form_ResizE
    End If
    If Me.Top < 0 Then
        DO_FORM_RESIZE_ONCE = False
    End If
        
    If DO_FORM_RESIZE_ONCE = True Then Exit Sub
    
    DO_FORM_RESIZE_ONCE = True
    
    If SHOW_WINDOW_ALWAYS_ON_TOP_AFTER_FORM_LOAD = True Then
    If Me.Visible = True Then
        SetWindowPos Me.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        SHOW_WINDOW_ALWAYS_ON_TOP_AFTER_FORM_LOAD = False
    Else
        DO_FORM_RESIZE_ONCE = False
    End If
    End If
    If Me.Visible = False Then
        DO_FORM_RESIZE_ONCE = False
    End If
    If Me.Top < 0 Then
        DO_FORM_RESIZE_ONCE = False
        TIMER_Form_Resize.Enabled = True
    End If

Exit Sub
ENDA:

DO_FORM_RESIZE_ONCE = False
TIMER_Form_Resize.Enabled = True

End Sub


Sub ME_TOP_LEFT_CENTER()
    
    ' -------------------------------------------
    
    Dim x, y, X3, X2, X4, X5, X7, X8, X9
    Dim EXTRA_SIZE_ER_COLUMN1_2_AND_3
    Dim LET_TEXT_BOX_GO
    x = 1
    y = 1
    On Error Resume Next
    Dim Control As Control
    

    For Each Control In Controls
        Err.Clear
        X4 = Control.Visible
        If Err.Number = 0 Then
            If Control.Enabled = False Then
                Control.Visible = False
            End If
        End If
    Next
    
    ' RICH TEXT BOX ARE NOT ENABLED UNTIL FORM FULLY LOADER
    For Each Control In Controls
        Err.Clear
        X4 = Control.Name
        If Err.Number = 0 Then
            If UCase(X4) = "TEXT2" Then
                Control.Visible = True
                Control.Enabled = True
            End If
            If UCase(X4) = "TEXT3" Then
                Control.Visible = True
                Control.Enabled = True
            End If
        End If
    Next
    
    ' OFFSET ALL ON
    ' --------------------------
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left
        If Err.Number = 0 Then
            'Control.Left = Control.Left - 850
            Control.Left = Control.Left - 300
        End If
    Next
    
    Text3.width = Text2.width
    
    X3 = 100
    X7 = 2000
    ' LOWER X LEFT SIDE IN
    ' --------------------------
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left
        If Err.Number = 0 Then
        'If Control.Visible = True Then
        If X4 < X7 Then
            Control.Left = X3
        'End If
        End If
        End If
    Next
    
    
    ' SET WIDTH COLUMN 1
    ' -----------------------------------------------------------
    EXTRA_SIZE_ER_COLUMN1_2_AND_3 = 40
    X3 = Label1(3).Left + Label1(3).width + EXTRA_SIZE_ER_COLUMN1_2_AND_3
    X8 = 1000 ' + OR -
    X7 = Label1(3).Left + Label1(3).width
    
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left + Control.width
        If Err.Number = 0 Then
        LET_TEXT_BOX_GO = False
        If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
        LET_HAPPEN = True
        If (Control.Visible = True Or LET_TEXT_BOX_GO = True) And LET_HAPPEN = True Then
        If X4 > X7 - X8 Then
        ' Debug.Print Control.Name
        If X4 < X7 + X8 Then
            ' Debug.Print Control.Name
            Control.width = X3
        End If
        End If
        End If
        End If
    Next
    
    ' SET WIDTH COLUMN 2 SAME AS COLUMN 1
    ' -----------------------------------------------------------
    X3 = Label1(3).Left + Label1(3).width + EXTRA_SIZE_ER_COLUMN1_2_AND_3
    X8 = 1000 ' + OR -
    X7 = Label1(6).Left + Label1(6).width
    X9 = Label1(6).width
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left + Control.width
        X5 = Control.width
        If Err.Number = 0 Then
        LET_TEXT_BOX_GO = False
        If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
        If InStr(Control.Caption, "GIVE TIME 03 BUTTON") > 0 Then LET_HAPPEN = False Else LET_HAPPEN = True
        LET_HAPPEN = True
        If (Control.Visible = True Or LET_TEXT_BOX_GO = True) And LET_HAPPEN = True Then
        If X5 > X9 - X8 Then
        If X5 < X9 + X8 Then
        If X4 > X7 - X8 Then
        ' Debug.Print Control.Name
        If X4 < X7 + X8 Then
            ' Debug.Print Control.Name
            Control.width = X3
        End If
        End If
        End If
        End If
        End If
        End If
    Next
    
    ' SET WIDTH COLUMN 3 SAME AS COLUMN 1 AND COLUMN 2
    ' -----------------------------------------------------------
    X3 = Label1(3).Left + Label1(3).width + EXTRA_SIZE_ER_COLUMN1_2_AND_3
    X8 = 200 ' + OR -
    X7 = LabelCRC3.Left
    X9 = LabelCRC3.width
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left
        X5 = Control.width
        If Err.Number = 0 Then
        LET_TEXT_BOX_GO = False
        If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
        If InStr(Control.Caption, "GIVE TIME 03 BUTTON") > 0 Then LET_HAPPEN = False Else LET_HAPPEN = True
        LET_HAPPEN = True
        If (Control.Visible = True Or LET_TEXT_BOX_GO = True) And LET_HAPPEN = True Then
        If X5 > X9 - X8 Then
        If X5 < X9 + X8 Then
        If X4 > X7 - X8 Then
        ' Debug.Print Control.Name
        If X4 < X7 + X8 Then
            ' Debug.Print Control.Name
            Control.width = X3
            'Control.Width = X3 - 200
            'Control.Left = X4 + 200
        End If
        End If
        End If
        End If
        End If
        End If
    Next
    LabelCRC5.Left = LabelCRC3.Left + LabelCRC3.width + 40
    
    ' INSERT NEW COLUMN 3 IN FRONT
    ' CHANGE CURRENT 3 WIDTH - NEXT ROUTINE
    ' -----------------------------------------------------------
    
    ' SET LEFT COLUMN 2
    ' -----------------------------------------------------------
    X3 = Label1(3).Left + Label1(3).width + 100
    X8 = 500 ' + OR -
    X7 = Label1(3).Left + Label1(3).width
    
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left
        If Err.Number = 0 Then
        LET_TEXT_BOX_GO = False
        If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
        If InStr(Control.Caption, "GIVE TIME 03 BUTTON") > 0 Then LET_HAPPEN = False Else LET_HAPPEN = True
        LET_HAPPEN = True
        If (Control.Visible = True Or LET_TEXT_BOX_GO = True) And LET_HAPPEN = True Then
        If X4 > X7 - X8 Then
        ' Debug.Print Control.Name
        If X4 < X7 + X8 Then
            ' Debug.Print Control.Name
            Control.Left = X3
        End If
        End If
        End If
        End If
    Next
    
    
    ' HERE SET THE GAP BETWEEN COLUMN ONE AND TWO HORIZONTAL
    ' - 30
    X3 = Label1(3).Left + Label1(3).width - 30
    X8 = 1000 ' + OR -
    X7 = Label1(3).Left + Label1(3).width
    
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left + Control.width
        If Err.Number = 0 Then
        LET_TEXT_BOX_GO = False
        If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
        If InStr(Control.Caption, "GIVE TIME 03 BUTTON") > 0 Then LET_HAPPEN = False Else LET_HAPPEN = True
        LET_HAPPEN = True
        If (Control.Visible = True Or LET_TEXT_BOX_GO = True) And LET_HAPPEN = True Then
        If X4 > X7 - X8 Then
        ' Debug.Print Control.Name
        If X4 < X7 + X8 Then
            If X4 < X3 Then
                Do
                    Control.width = Control.width + 1
                    X5 = Control.Left + Control.width
                Loop Until X5 = X3
            End If
            If X4 > X3 Then
                Do
                    Control.width = Control.width - 1
                    X5 = Control.Left + Control.width
                Loop Until X5 = X3
            End If
        End If
        End If
        End If
        End If
    Next
    
    
    ' + GAP_VALUE = 100 THIS SET THE GAP MIDDLE COLUMN BETWEEN COLUMN 2 AND 3
    ' -----------------------------------------------------------
    GAP_VALUE = 100 ' ---------------
    X8 = 1500 ' + OR -
    X7 = Label1(6).Left + Label1(6).width
    LabelFILE2_TOP = LabelFILE2.Top ' DIFF WHEN AFTER HERE LOWER DOWN SCREEN NOT RTB OR TOP TEXT AH
    X3 = Label1(6).Left + Label1(6).width + GAP_VALUE ' + 100
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left
        If Err.Number = 0 Then
        LET_TEXT_BOX_GO = False
        If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
        If InStr(Control.Caption, "GIVE TIME 03 BUTTON") > 0 Then LET_HAPPEN = False Else LET_HAPPEN = True
        LET_HAPPEN = True
        If (Control.Visible = True Or LET_TEXT_BOX_GO = True) And LET_HAPPEN = True Then
        If X4 > X7 - X8 Then
        ' Debug.Print Control.Name
        If X4 < X7 + X8 Then
            ' Debug.Print Control.Name
            Control.Left = X3
        End If
        End If
        End If
        End If
    Next
    
    ' RICH TEXT BOX ARE NOT ENABLED UNTIL FORM FULLY LOADER
    ' RICH TEXT BOX CONTROL SOMETIME LOSER IT VISIBLE TO NOT - OFF
    ' USE ENABLE PROPERTIE INSTEAD
    ' -----------------------------------
    ' END OF COLUMN 2
    ' SET WITDH COLUMN 2
    ' -----------------------------------
    X3 = Label1(6).Left + Label1(6).width
    X8 = 4500 ' + OR -
    X7 = Label1(6).Left + Label1(6).width
    
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left + Control.width
        If Err.Number = 0 Then
        LET_TEXT_BOX_GO = False
        If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
        If InStr(Control.Caption, "GIVE TIME 03 BUTTON") > 0 Then LET_HAPPEN = False Else LET_HAPPEN = True
        LET_HAPPEN = True
        If (Control.Visible = True Or LET_TEXT_BOX_GO = True) And LET_HAPPEN = True Then
        If X4 > X7 - X8 Then
        ' Debug.Print Control.Name
        If X4 < X7 + X8 Then
            If X4 < X3 Then
                Do
                    Control.width = Control.width + 1
                    X5 = Control.Left + Control.width
                Loop Until X5 = X3
            End If
            If X4 > X3 Then
                Do
                    Control.width = Control.width - 1
                    X5 = Control.Left + Control.width
                Loop Until X5 = X3
            End If
        End If
        End If
        End If
        End If
    Next
    
    X3 = MNU_SHOW_THE_TIME_03(0).Left + MNU_SHOW_THE_TIME_03(0).width
    X8 = 100 ' + OR -
    X7 = LabelCRC5.Left + 100
    
    For Each Control In Controls
        Err.Clear
        X4 = Control.Left + Control.width
        If Err.Number = 0 Then
        LET_TEXT_BOX_GO = False
        If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
        If InStr(Control.Caption, "GIVE TIME 03 BUTTON") > 0 Then LET_HAPPEN = False Else LET_HAPPEN = True
        If (Control.Visible = True Or LET_TEXT_BOX_GO = True) And LET_HAPPEN = True Then
        If X4 > X7 Then '- X8 Then
        ' Debug.Print Control.Name
        ' If X4 < X7 + X8 Then
            ' Debug.Print Control.Name
            If Control.Left + Control.width < X3 Then
                Do
                    Control.width = Control.width + 1
                    X5 = Control.Left + Control.width
                Loop Until X5 = X3 ' Or 1 = 1
            End If
            If Control.Left + Control.width > X3 Then
                Do
                    Control.width = Control.width - 1
                    X5 = Control.Left + Control.width
                Loop Until X5 = X3 ' Or 1 = 1
            End If
        ' End If
        End If
        End If
        End If
    Next
    
    ' END COLUMN 4
    ' ---------------------------------------------------------
    X3 = MNU_SHOW_THE_TIME_03(0).Left + MNU_SHOW_THE_TIME_03(0).width
    X4 = MNU_SHOW_THE_TIME_03(0).FontSize
    X5 = MNU_SHOW_THE_TIME_03(0).Top
    
    For R1 = 0 To MMM_SHOW_THE_TIME_ARR.UBound
        LET_TEXT_BOX_GO = False
        If InStr(UCase(MMM_SHOW_THE_TIME_ARR(R1).Name), "MMM_SHOW_THE_TIME_ARR") > 0 Then
            MMM_SHOW_THE_TIME_ARR(R1).Left = X3 + 80
            MMM_SHOW_THE_TIME_ARR(R1).width = 780 + 400 ' WIDTH NOT MUCH OVER OR NEW CODE
            HEIGHT_GAP = 28
            If MNU_SHOW_THE_TIME_03.UBound >= R1 Then
                X7 = MNU_SHOW_THE_TIME_03(R1).height
                MMM_SHOW_THE_TIME_ARR(R1).height = X7
                MMM_SHOW_THE_TIME_ARR(R1).Top = MNU_SHOW_THE_TIME_03(R1).Top
                X5 = MNU_SHOW_THE_TIME_03(R1).Top + X7 + HEIGHT_GAP
            Else
                MMM_SHOW_THE_TIME_ARR(R1).height = MMM_SHOW_THE_TIME_ARR(R1).height
                MMM_SHOW_THE_TIME_ARR(R1).Top = X5
                X5 = X5 + X7 + HEIGHT_GAP
            End If
            MMM_SHOW_THE_TIME_ARR(R1).Caption = "000" + Trim(Str(R1))
            ' MMM_SHOW_THE_TIME_ARR(R1).FontSize = X4
            MMM_SHOW_THE_TIME_ARR(R1).Alignment = 2 ' CENTER
            If R1 = MMM_SHOW_THE_TIME_ARR.UBound Then
                ' MMM_SHOW_THE_TIME_ARR(R1).Visible = False
            End If
        End If
    Next
    
    
    ' ---------------------------------------
    
    Text3.width = Text3.width + 1
    LabelFILE1.width = LabelFILE1.width + 10
    LabelFILE2.width = LabelFILE2.width + 10
    Label3.width = Label3.width
    Label5.width = Label5.width
    Label2.width = Label2.width
    LabelCRC5.width = LabelCRC5.width + 10
    
    Label1(1).width = Text2.width / 2
    Label1(2).width = Text3.width
    Label1(7).width = Text2.width / 2 - 40
    Label1(7).Left = Label1(1).width + Label1(1).Left + 40
    Label1(7).BackColor = RGB(255, 255, 255)
    Label1(7).Caption = "COMPARE VEIWER"
    ' ---------------------------------------
    
    For Each Control In Controls
        XX_1 = Mid(UCase(Control.Name), 1, 3)
        XX_2 = UCase(Control.Name)
        If XX_1 <> "MNU" Or XX_2 = "MMM_SHOW_THE_TIME_ARR" Then
            LET_TEXT_BOX_GO = False
            If InStr(UCase(Control.Name), "TEXT") > 0 Then LET_TEXT_BOX_GO = True
            If Control.Visible = True Then
                If Control.width + Control.Left > WX Then WX = Control.width + Control.Left
                If Control.height + Control.Top > HY Then
                    HY = Control.height + Control.Top
                End If
            End If
        End If
    Next
    
    Call ME_POSITION

    If 11 = 22 Then
    ' -----------------------------------
     For i = 0 To Screen.FontCount - 1
         A = A + Screen.Fonts(i) + vbCrLf
     Next i
    ' -----------------------------------
     Clipboard.Clear
     Clipboard.SetText A
'    Source Code Pro
'    Source Code Pro Medium
'    Source Code Pro Semibold
'    Source Code Pro Black
    ' -----------------------------------
    End If
    ' -----------------------------------
    ' Fri 25-Sep-2020 19:51:41
    ' -----------------------------------
    ' ----
    ' Change Fontstyle Using VB - VB6 | Dream.In.Code
    ' https://www.dreamincode.net/forums/topic/69063-change-fontstyle-using-vb/
    ' ----


'    Dim CONT As Control
'    On Error Resume Next
'    For Each CONT In Me.Controls
'        ' CONT.Font.Name = "ARIAL"
'        ' CONT.Font.Name = "SOURCE CODE PRO BLACK" ' -- "ARIAL"
'        CONT.Font.Name = "ARIAL BLACK"
'        CONT.FontSize = 12
'        CONT.FontBold = True
'    Next
    



End Sub

Sub ME_POSITION()
    Dim OFFSET_WX, OFFSET_HY
    OFFSET_WX = 180
    OFFSET_HY = -280
    HY = HY + ((Menu_Height * Screen.TwipsPerPixelY) - Menu_Height) + OFFSET_HY
    WX = WX + (Screen.TwipsPerPixelX) + OFFSET_WX
    On Error Resume Next
    ' WHEN CALL BY HALIFAX ROUTINE
    ' A FORM NOT POSITION MOVE WHEN MINIMIZE MAXIMIZE
    Me.width = WX ' Text3.Left + Text3.Width + 200
    On Error GoTo 0
    WX = ((Screen.width) / 2 - ((Me.width / 2)))
    Me.Left = WX
    Me.Top = 200
    Me.Visible = True
End Sub


Private Sub MNU_CLIPBOARDER_REPLACE_ER_AND_Click()
' CLIPBOARD REPLACE "AND"


Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4
Dim CHANGE_VAR_ARRAY(1)
Dim CHANGE_VAR_1
Dim CHANGE_VAR_2
Dim CH_1
Dim CH_2
Dim CH_3
Dim CMOS_AND_TTL
Dim R1
Dim R2
Dim x

'Debug.Print "----"
VAR_ST_1 = Clipboard.GetText
VAR_ST_4 = VAR_ST_1
' VAR_ST_1 = "and ff and hh and yy and 123456789 kk"
CHANGE_VAR_ARRAY(1) = LCase(" AND ")
' CHANGE_VAR_ARRAY(2) = LCase(" AN ") ' THIS COULD LOOK SHITTER
Dim IIN
For IIN = 1 To UBound(CHANGE_VAR_ARRAY)
    CHANGE_VAR_1 = CHANGE_VAR_ARRAY(IIN)
    x = Len(CHANGE_VAR_1) + 1
    For R1 = 0 To Len(CHANGE_VAR_1) + 1
        For R2 = 1 To Len(CHANGE_VAR_1)
            
            If R1 > 0 And R1 < x Then
                CHANGE_VAR_2 = CHANGE_VAR_1
                CH_2 = UCase(Mid(CHANGE_VAR_2, R2, 1))
                CH_3 = UCase(Mid(CHANGE_VAR_2, R1, 1))
                Mid(CHANGE_VAR_2, R2, 1) = CH_2
                Mid(CHANGE_VAR_2, R1, 1) = CH_3
            End If
            If R1 = 0 Then CHANGE_VAR_2 = CHANGE_VAR_1
            If R1 = x Then CHANGE_VAR_2 = UCase(CHANGE_VAR_1)
            If InStr(CMOS_AND_TTL, CHANGE_VAR_2) = 0 Then
                ' Debug.Print CHANGE_VAR_2
                VAR_ST_1 = Replace(VAR_ST_1, CHANGE_VAR_2, " & ")
                VAR_ST_1 = Replace(VAR_ST_1, vbCr + Mid(CHANGE_VAR_2, 2), vbCr + "& ")
                VAR_ST_1 = Replace(VAR_ST_1, vbLf + Mid(CHANGE_VAR_2, 2), vbLf + "& ")
                VAR_ST_1 = Replace(VAR_ST_1, Mid(CHANGE_VAR_2, 1, Len(CHANGE_VAR_2) - 1) + vbCr, " &" + vbCr)
                VAR_ST_1 = Replace(VAR_ST_1, Mid(CHANGE_VAR_2, 1, Len(CHANGE_VAR_2) - 1) + vbLf, " &" + vbLf)
            End If
            
            CMOS_AND_TTL = CMOS_AND_TTL + "--" + CHANGE_VAR_2 + "__"
        Next
    Next
Next

If VAR_ST_1 <> VAR_ST_4 Then
    
    Do
        On Error Resume Next
        Clipboard.Clear
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> "" Then Sleep 500
    Loop Until VAR_ST_2 = ""
    Sleep 500
    Do
        On Error Resume Next
        Clipboard.SetText VAR_ST_1, vbCFText
        Sleep 500
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> VAR_ST_1 Then Sleep 800
    Loop Until VAR_ST_2 = VAR_ST_1

    Me.WindowState = vbMinimized

End If

Beep

End Sub

Public Sub Timer_CHECK_PROJECT_DATE_IN_PROCESS_Timer()
    If Me.CHECK_PROJECT_DATE_IN_PROCESS = True Then
        Label_CHECK_PROJECT_DATE_IN_PROCESS.BackColor = &HFF00&
        ' HEX(Label1(4).BackColor) ' &H0000FF00& -- BRIGHT GREEN
        ' HEX(Label1(5).BackColor) ' &HC0FFFF&   -- PALE YELLOW
    Else
        Label_CHECK_PROJECT_DATE_IN_PROCESS.BackColor = &HC0FFFF
    End If
    Label_CHECK_PROJECT_DATE_IN_PROCESS.Refresh
    DoEvents
End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
  
  'TESTING
  'IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************






'NumToString

'Description:
' The function NumToString writes out any number, up to about 920 trillion, the limit of the Currency variable type, in English words. It only writes out the integer portion of the number. The DollarToString function does the same, but places the word dollars after the string and also writes out the fractional part of the value as the cents.

'1/8/2001: Added DateToString and NumToStringTh (ie, "23" -> "twenty-third") helper functions.

'7/20/2002: Corrected error in DollarToString that limited the range of values.
  
'Code:
 Public Function DollarToString(ByVal nAmount As Currency) As String
'Example:
'  DollarToString(5.99) = " five dollars and ninety-nine cents"
    Dim nDollar As Currency
    Dim nCent As Currency
    
    nDollar = Int(nAmount)
    nCent = (Abs(nAmount) - Int(Abs(nAmount))) * 100
    
    DollarToString = NumToString(nDollar) & " dollar"
    
    If Abs(nDollar) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
    DollarToString = DollarToString & " and" & _
    NumToString(nCent) & " cent"
    
    If Abs(nCent) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
End Function

Public Function NumToString(ByVal nNumber As Currency) As String
'Example: NumToString(123) = " one hundred twenty-three"
    Dim bNegative As Boolean
    Dim bHundred As Boolean

    If nNumber < 0 Then
        bNegative = True
    End If

    nNumber = Abs(Int(nNumber))
    If nNumber < 1000 Then
        If nNumber \ 100 > 0 Then
            NumToString = NumToString & _
                NumToString(nNumber \ 100) & " hundred"
                NumToString = NumToString & " and"

            bHundred = True
        End If
        nNumber = nNumber - ((nNumber \ 100) * 100)
        Dim bNoFirstDigit As Boolean
        bNoFirstDigit = False
        Select Case nNumber \ 10
            Case 0
                Select Case nNumber Mod 10
                    Case 0
                        If Not bHundred Then
                            NumToString = NumToString & " zero"
                        End If
                    Case 1: NumToString = NumToString & " one"
                    Case 2: NumToString = NumToString & " two"
                    Case 3: NumToString = NumToString & " three"
                    Case 4: NumToString = NumToString & " four"
                    Case 5: NumToString = NumToString & " five"
                    Case 6: NumToString = NumToString & " six"
                    Case 7: NumToString = NumToString & " seven"
                    Case 8: NumToString = NumToString & " eight"
                    Case 9: NumToString = NumToString & " nine"
                End Select
                bNoFirstDigit = True
            Case 1
                Select Case nNumber Mod 10
                    Case 0: NumToString = NumToString & " ten"
                    Case 1: NumToString = NumToString & " eleven"
                    Case 2: NumToString = NumToString & " twelve"
                    Case 3: NumToString = NumToString & " thirteen"
                    Case 4: NumToString = NumToString & " fourteen"
                    Case 5: NumToString = NumToString & " fifteen"
                    Case 6: NumToString = NumToString & " sixteen"
                    Case 7: NumToString = NumToString & " seventeen"
                    Case 8: NumToString = NumToString & " eighteen"
                    Case 9: NumToString = NumToString & " nineteen"
                End Select
                bNoFirstDigit = True
            Case 2: NumToString = NumToString & " twenty"
            Case 3: NumToString = NumToString & " thirty"
            Case 4: NumToString = NumToString & " forty"
            Case 5: NumToString = NumToString & " fifty"
            Case 6: NumToString = NumToString & " sixty"
            Case 7: NumToString = NumToString & " seventy"
            Case 8: NumToString = NumToString & " eighty"
            Case 9: NumToString = NumToString & " ninety"
        End Select
        If Not bNoFirstDigit Then
            If nNumber Mod 10 <> 0 Then
                NumToString = NumToString & "-" & _
                    Mid(NumToString(nNumber Mod 10), 2)
            End If
        End If
    Else
        Dim nTemp As Currency
        nTemp = 10 ^ 12 'trillion
        Do While nTemp >= 1
            If nNumber >= nTemp Then
                NumToString = NumToString & _
                    NumToString(Int(nNumber / nTemp))
                Select Case Int(Log(nTemp) / Log(10) + 0.5)
                    Case 12: NumToString = NumToString & " trillion"
                    Case 9: NumToString = NumToString & " billion"
                    Case 6: NumToString = NumToString & " million"
                    Case 3: NumToString = NumToString & " thousand"
                End Select
                nNumber = nNumber - (Int(nNumber / nTemp) * nTemp)
            End If
            nTemp = nTemp / 1000
        Loop
    End If
    If bNegative Then
        NumToString = " negative " & NumToString
    End If
End Function

Public Function NumToStringTh(ByVal nNumber As Currency) As String
'Example: NumToStringTh(123) = " one hundred twenty-third"

    Dim sNum As String
    Dim sExtra As String
    Dim nSpace As String

    'Convert the number
    sNum = NumToString(nNumber)
    'Find the location of the last space or dash
    nSpace = Len(sNum)
    
    Do Until Mid(sNum, nSpace, 1) = " " Or _
        Mid(sNum, nSpace, 1) = "-"
        nSpace = nSpace - 1
    Loop

    sExtra = Mid(sNum, nSpace + 1)
    sNum = Mid(sNum, 1, nSpace)
    
    'Conver the last word ("one" -> "first", etc)
    Select Case sExtra
        Case "hundred":   NumToStringTh = sNum & "hundredth"
        Case "zero" 'no such thing as 'zeroth'
            NumToStringTh = sNum & "zero"
        Case "one":       NumToStringTh = sNum & "first"
        Case "two":       NumToStringTh = sNum & "second"
        Case "three":     NumToStringTh = sNum & "third"
        Case "four":      NumToStringTh = sNum & "fourth"
        Case "five":      NumToStringTh = sNum & "fifth"
        Case "six":       NumToStringTh = sNum & "sixth"
        Case "seven":     NumToStringTh = sNum & "seventh"
        Case "eight":     NumToStringTh = sNum & "eighth"
        Case "nine":      NumToStringTh = sNum & "ninth"
        Case "ten":       NumToStringTh = sNum & "tenth"
        Case "eleven":    NumToStringTh = sNum & "eleventh"
        Case "twelve":    NumToStringTh = sNum & "twelfth"
        Case "thirteen":  NumToStringTh = sNum & "thirteenth"
        Case "fourteen":  NumToStringTh = sNum & "fourteenth"
        Case "fifteen":   NumToStringTh = sNum & "fifteenth"
        Case "sixteen":   NumToStringTh = sNum & "sixteenth"
        Case "seventeen": NumToStringTh = sNum & "seventeenth"
        Case "eighteen":  NumToStringTh = sNum & "eighteenth"
        Case "nineteen":  NumToStringTh = sNum & "nineteenth"
        Case "twenty":    NumToStringTh = sNum & "twentieth"
        Case "thirty":    NumToStringTh = sNum & "thirtieth"
        Case "forty":     NumToStringTh = sNum & "fortieth"
        Case "fifty":     NumToStringTh = sNum & "fiftieth"
        Case "sixty":     NumToStringTh = sNum & "sixtieth"
        Case "seventy":   NumToStringTh = sNum & "seventieth"
        Case "eighty":    NumToStringTh = sNum & "eightieth"
        Case "ninety":    NumToStringTh = sNum & "ninetieth"
        Case "trillion":  NumToStringTh = sNum & "trillionth"
        Case "billion":   NumToStringTh = sNum & "billionth"
        Case "million":   NumToStringTh = sNum & "millionth"
        Case "thousand":  NumToStringTh = sNum & "thousandth"
        Case Else 'This shouldn't happen, but just in case
            NumToStringTh = NumToString(nNumber)
    End Select
    
End Function

Public Function DateToString(FromDate As String) As String
'Example: ' DateToString(#1/1/2001#) = "January the first, two thousand one"

    DateToString = Format(FromDate, "mmmm") & " the" & _
        NumToStringTh(DatePart("d", FromDate)) & "," & _
        NumToString(DatePart("yyyy", FromDate))
        
End Function

 
  
'Sample Usage:
  
     'Debug.Print DollarToString(1234.11)
    'Debug.Print NumToString(-54321)

 
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------
