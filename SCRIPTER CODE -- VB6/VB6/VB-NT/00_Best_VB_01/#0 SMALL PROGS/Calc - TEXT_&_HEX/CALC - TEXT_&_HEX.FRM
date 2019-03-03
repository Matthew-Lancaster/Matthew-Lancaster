VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   9015
   ClientLeft      =   8550
   ClientTop       =   1845
   ClientWidth     =   17565
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CALC - TEXT_&_HEX.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9015
   ScaleWidth      =   17565
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer_ClipBoard_Change_Flash_Colour 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6675
      Top             =   2865
   End
   Begin VB.Timer Timer_ClipBoard_Status 
      Interval        =   1000
      Left            =   6255
      Top             =   2865
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8940
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   53
      Text            =   "CALC - TEXT_&_HEX.frx":030A
      ToolTipText     =   "HEX TO TEXT -- CLICK HITT HERE TO READ THE VALUE GET CLIPBOARD IN"
      Top             =   3375
      Width           =   8505
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SET Clipboard To HEX With UniCode &&H00 Inserted Between Each HEX For Symbolic *.LNK Linker -- Method Work Around Save Time"
      Height          =   2595
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5925
      Width           =   1845
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFDFDF&
      Caption         =   "Info Tip for Multi Finder.exe"
      Height          =   825
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5745
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Launch Multi Finder.exe"
      Height          =   540
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5205
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009FFFFF&
      Caption         =   "GET CLIPBOARD"
      Height          =   525
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5385
      Width           =   1845
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Text            =   "CALC - TEXT_&_HEX.frx":0315
      ToolTipText     =   "HEX TO TEXT -- CLICK HITT HERE TO READ THE VALUE GET CLIPBOARD IN"
      Top             =   900
      Width           =   9795
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   7665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Text            =   "CALC - TEXT_&_HEX.frx":0320
      ToolTipText     =   "TEXT TO HEX DISPLAY AS SEPERATE SPACE TO ENABLE WRAP AROUND DISPLAY __ CLIPBOARD THE VALUE IN LABEL"
      Top             =   2025
      Width           =   9795
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "REPLACE SEPERATOR HEX WITH ---- ;"
      Height          =   480
      Left            =   8505
      TabIndex        =   27
      Top             =   1050
      Width           =   4080
   End
   Begin VB.CommandButton Number_HEX_TO_TEXT 
      BackColor       =   &H0080FF80&
      Caption         =   "HEX TO TEXT"
      Height          =   570
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4815
      Width           =   1845
   End
   Begin VB.CommandButton Number_TEXT_TO_HEX 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "TEXT TO HEX"
      Height          =   540
      Left            =   7665
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4245
      Width           =   1845
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      Height          =   480
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   480
      Index           =   8
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   480
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "C"
      Height          =   480
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton CancelEntry 
      Caption         =   "CE"
      Height          =   480
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   480
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   480
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   480
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      Height          =   480
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      Height          =   480
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   480
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   480
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   480
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "X"
      Height          =   480
      Index           =   2
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      Height          =   480
      Index           =   0
      Left            =   2640
      TabIndex        =   15
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1080
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      Height          =   480
      Left            =   1320
      TabIndex        =   18
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      Height          =   480
      Index           =   4
      Left            =   2040
      TabIndex        =   16
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Percent 
      Caption         =   "%"
      Height          =   480
      Left            =   2640
      TabIndex        =   17
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current ClipBoard Changed"
      Height          =   780
      Left            =   7665
      TabIndex        =   55
      Top             =   3375
      Width           =   1230
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clip Board View Interval Update -- Active Idle Coluring"
      Height          =   1815
      Left            =   5940
      TabIndex        =   54
      Top             =   3360
      Width           =   1650
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label_BackColor2 
      BackColor       =   &H00D7FFFF&
      Caption         =   "Back Color2 Option"
      Height          =   285
      Left            =   13515
      TabIndex        =   51
      Top             =   8295
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Color Box Indicate Control Code"
      Height          =   285
      Left            =   10170
      TabIndex        =   50
      Top             =   7980
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Some Unicode Like VBCRLF  Have Bee Removed Frm Top Line For Display Look at Nex Line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2865
      Left            =   3225
      TabIndex        =   49
      ToolTipText     =   $"CALC - TEXT_&_HEX.frx":032B
      Top             =   105
      Width           =   2940
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VBCRLF LABEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10170
      TabIndex        =   48
      ToolTipText     =   $"CALC - TEXT_&_HEX.frx":03C3
      Top             =   8280
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "REPLACE SEPERATOR HEX WITH ---- &&h00 Unicode Worker"
      Height          =   285
      Left            =   9540
      TabIndex        =   47
      Top             =   7620
      Width           =   5970
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label_BackColor1 
      BackColor       =   &H00C6FFDD&
      Caption         =   "Back Color1 Option"
      Height          =   285
      Left            =   13515
      TabIndex        =   44
      Top             =   7980
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CLIPBOARD WILL GET AT START -- PROCESS WILL DECIDE WHAT IS A HEX OR A TEXT"
      Height          =   1800
      Left            =   9540
      TabIndex        =   43
      Top             =   4950
      Width           =   7920
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Info -- Here Not Working Completely Yet -- Extra Idea Use Quicker Work Aournd At Moment Not Swappable That Easy"
      Height          =   2010
      Left            =   15585
      TabIndex        =   41
      Top             =   6765
      Width           =   1875
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USE TOOLTIP OVER THE BUTTON LABEL SET or GTE CLIPBOARD"
      Height          =   360
      Left            =   9540
      TabIndex        =   40
      Top             =   4560
      Width           =   7920
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "\/\/\/\/ **** ---- INFO ---- **** \/\/\/\/"
      Height          =   300
      Left            =   9540
      TabIndex        =   39
      Top             =   4230
      Width           =   7920
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "REPLACE SEPERATOR HEX WITH ---- ,"
      Height          =   285
      Left            =   9540
      TabIndex        =   38
      Top             =   7065
      Width           =   5970
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "REPLACE SEPERATOR HEX WITH ---- - DASH"
      Height          =   285
      Left            =   9540
      TabIndex        =   37
      Top             =   7350
      Width           =   5970
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "REPLACE SEPERATOR HEX WITH ---- ;"
      Height          =   285
      Left            =   9540
      TabIndex        =   36
      Top             =   6780
      Width           =   5970
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"CALC - TEXT_&_HEX.frx":045B
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5865
      Left            =   120
      TabIndex        =   35
      Top             =   3045
      Width           =   5790
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "HEX TO TEXT"
      Height          =   285
      Left            =   6255
      TabIndex        =   34
      Top             =   525
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "HEX WITH SEPERATOR"
      Height          =   750
      Left            =   6255
      TabIndex        =   33
      Top             =   2070
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "HEX && TEXT LONG"
      Height          =   645
      Left            =   6255
      TabIndex        =   32
      Top             =   930
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "HEX"
      Height          =   285
      Left            =   6255
      TabIndex        =   31
      Top             =   1680
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "TEXT && HEX"
      Height          =   285
      Left            =   6255
      TabIndex        =   30
      Top             =   120
      Width           =   1380
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9585
      TabIndex        =   26
      ToolTipText     =   "HEX WITH SPERATER SPACE TRUNCICATED IN DISPLAY IF LONG -- CLICK HITT HERE TO READ THE CLIPBOARD IN"
      Top             =   8295
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7665
      TabIndex        =   25
      ToolTipText     =   "HEX -- TRUCICATED IN DISPLAY IF LONG __ CLICK HITT HERE TO READ THE WHOLE VALUE WITHOUT SPACE SEPERATOR TO GET CLIPBOARD"
      Top             =   1650
      Width           =   9795
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7665
      TabIndex        =   24
      ToolTipText     =   "HEX TO TEXT CONVERSION -- CLIPBOARD THE VALUE IN LABEL"
      Top             =   495
      Width           =   9795
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9585
      TabIndex        =   23
      ToolTipText     =   "HEX DISPLAY AS SEPERATE SPACE TO ENABLE WRAP AROUND DISPLAY __ CLIPBOARD THE VALUE IN LABEL"
      Top             =   7950
      Visible         =   0   'False
      Width           =   540
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7680
      TabIndex        =   20
      ToolTipText     =   $"CALC - TEXT_&_HEX.frx":0505
      Top             =   105
      Width           =   9795
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFCCCC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   105
      Width           =   3000
   End
   Begin VB.Menu Mnu_Exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB APP Folder"
   End
   Begin VB.Menu MNU_ABOUT 
      Caption         =   " By Matt.Lan@BTinternet.Com 2k16 Hex Conversion Addtion on Top of Microsoft  Sample Source Code"
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'               Copyright (C) 1994 Microsoft Corporation
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------
Option Explicit

Dim CLIPBOARD_FLASH_COUNT
Dim Set_To_Beep
Dim Clipboard_Contain_Text
Dim ClipBoard_Status_Force_Through
Dim OX, OY

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
             
' ---------------

Dim R, i, I_2, XX

' ---------------

Dim Op1, Op2                ' Previously input operand.
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Number of operands.
Dim LastInput               ' Indicate type of last keypress event.
Dim OpFlag                  ' Indicate pending operation.
Dim TempReadout

' Click event procedure for C (cancel) key.
' Reset the display and initializes variables.
Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

' Click event procedure for CE (cancel entry) key.
Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

Private Sub Command1_Click()

'Same as Call Label1_Click

Call READ_CLIPBOARD_HEX_CONVERSION

End Sub

Private Sub Command2_Click()
'Command2.Caption = "Launch " + vbCrLf + "Multi Finder.exe"
Dim i, i1, i2, PID
i1 = "C:\Program Files (x86)\MultiFind\MultiFind.exe"
i = i1
i2 = "C:\Program Files\MultiFind\MultiFind.exe"
If Dir(i1) = "" Then
    i = i2
End If

If Dir(i) = "" Then
    MsgBox "PROGRAM NOT EXIST AT" + vbCrLf + vbCrLf + i1 + vbCrLf + vbCrLf + i2
End If

If Dir(i) <> "" Then
    PID = Shell(i, vbNormalFocus)
    Exit Sub
End If

MsgBox "PROGRAM NOT RETURN VALUE FOR LAUNCH " + vbCrLf + vbCrLf + i + vbCrLf + vbCrLf + "DOES EXIST BUT POSIABLE ADMIN RIGHT_ PROBLEM"

End Sub

Private Sub Command3_Click()
MsgBox "NOT HERE YET"
End Sub

Private Sub Command4_Click()

i = ""
i = i + "1.. Load It Up"
i = i + vbCrLf + "2.. Note About the Read Only and Not Read Only Option Splitter"
i = i + vbCrLf + "3.. You Able to Set Extension as Mask Away"
i = i + vbCrLf + "5.. Sett Multi Folder Batch is Little Differcult"
i = i + vbCrLf + "6.. Idea is Make a New Page and then Search It Proper"
i = i + vbCrLf + "7.. Clear Result Remember When Pull Up Again"
i = i + vbCrLf + "8.. Clear Critia Not Same as Make a New Page Search"
i = i + vbCrLf + "9.. Do the Search With the File and Hex String -- See The Result Action Panic Over"

MsgBox i, vbMsgBoxSetForeground

End Sub

Private Sub Command5_Click()

Dim String_Var
String_Var = Replace(Text1, " ", "00")

Clipboard.Clear
Clipboard.SetText String_Var
Beep
    
Text3.Text = String_Var
    

End Sub

' Click event procedure for decimal point (.) key.
' If last keypress was an operator, initialize
' readout to "0." Otherwise, append a decimal
' point to the display.
Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

' Initialization routine for the form.
' Set all variables to initial values.
Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
    'Decimal.Caption = Format(0, ".")
    
    Call READ_CLIPBOARD_HEX_CONVERSION
        
    Me.Caption = "Hexadecimal_2_Text_And_Back_MULTI_Convertor"
    
    Me.BackColor = Label_BackColor1.BackColor
    'Me.BackColor = &H8000000F ' ORGINAL
    'Me.BackColor = Label_BackColor2.BackColor
    
    
    MNU_ABOUT.Caption = "Project Idea By Matt.Lan@BTinternet.Com 2k16 07919180203-- Hexadecimal && Text Conversion"

    Label11.Caption = "Project Design for Use on Sybolic Link ShortCut When Use a Binary Search and Replace APP HERE That Require Proper Input in HEX" + vbCrLf + vbCrLf + "C:\Program Files (x86)\MultiFind\MultiFind.exe"
    Label11.Caption = Label11.Caption + vbCrLf + vbCrLf + "Better Still write My Own Search and Replace"
    Label11.Caption = Label11.Caption + vbCrLf + "BUT And Administrtor is Being a Problem of Scan All Folder WIN 7 Up"
    Label11.Caption = Label11.Caption + vbCrLf + "Write Build Your Own Tool Kitt"
    Label11.Caption = Label11.Caption + vbCrLf + vbCrLf + "Saturday 5TH November 2k16"
    Label11.Caption = Label11.Caption + vbCrLf + "From 08:29am to 12:20am -- Here Writing 3h 40m"
    'Label11.Caption = Label11.Caption + vbCrLf + "From 08:29am to 03:00pm -- Maximum  Six 6h 10m Total"
    Label11.Caption = Label11.Caption + vbCrLf + "From 08:29am to 04:39pm -- Maximum  8h 0m 8Sec Total"
    Label11.Caption = Label11.Caption + vbCrLf + "Working on this One Long Time"
    
    Label11.Caption = Label11.Caption + vbCrLf + "Good Morning Sun Noon 11:42am & Good Afternoon This Morning"
    
    Label11.Caption = Label11.Caption + vbCrLf + "Coder Addition on Top of a Microsoft Sample Source Code Called Calc.vbp"
    Label11.Caption = Label11.Caption + vbCrLf + "Calc Code Free Royalty But And Not any Reward If Maybe Damage Your Computer"
    Label11.Caption = Label11.Caption + vbCrLf + "Rename a Path Branch Tree of Symbolic Link NightMare Zoombie -- Good Morning"



    Label18.Caption = "1.. CLIPBOARD WILL BE GET AT START AND PROCESED "
    Label18.Caption = Label18.Caption + vbCrLf + "2.. PROCESS WILL DECIDE WHAT IS A HEX OR WHAT NOT IS A HEX"
    Label18.Caption = Label18.Caption + vbCrLf + "By Count Each Byte Pair For Hex Value" + vbCrLf + "The IsNumeric Function Won't Work on HEX And Process Naught Value Are Discounted"
    Label18.Caption = Label18.Caption + vbCrLf + "3.. The Content of Text Box Are Left Modifible -- Take Care"
    Label18.Caption = Label18.Caption + vbCrLf + "Finally Content about the Setting_s_plural -- Set and Setting Tripper"
    '----
    'Calculate Duration Between Two Dates – Results
    'https://www.timeanddate.com/date/durationresult.html?&d1=5&m1=11&y1=2016&d2=5&m2=11&y2=2016&h1=08&i1=40&s1=00&h2=11&i2=46&s2=25
    '----
    
    Command2.Caption = "Launch" + vbCrLf + "Multi Finder.exe"
    Command4.Caption = "Info Tip On Using" + vbCrLf + "Multi Finder.exe"
    
    Label1.ToolTipText = "Click Hitt Here To Read The Clipboard In -- If Hex On Clipboard Condition For It Will Be Processed -- Hex Spaces Will Be Stripped On The Input For Process"
    Command1.ToolTipText = Label1.ToolTipText
    Number_TEXT_TO_HEX.ToolTipText = "Don't Really Matter Unless Had Change and Do Again"
    Number_HEX_TO_TEXT.ToolTipText = Number_TEXT_TO_HEX.ToolTipText
    
    Label21.Caption = "Some Control Code" + vbCrLf + " Like Like Like vbCrLf Have Been Removed From The Input For Don't Display Well"
    Label21.Caption = Label21.Caption + vbCrLf + vbCrLf + "Colour Green in This Box Indicate Yes or NO"
    Label21.Caption = Label21.Caption + vbCrLf + vbCrLf + "Look at Next Text-Box Line Down But 2" + vbCrLf + "For The Full Result"
    Label21.Caption = Label21.Caption + vbCrLf + "-HEX && TEXT LONG-"



    Command5.Caption = "SET Clipboard To HEX With UniCode Null &&H00 Void Inserted Between Each HEX For Symbolic *.LNK Linker -- Method Work Around Save Time"
    
    ClipBoard_Status_Force_Through = True
    Call Timer_ClipBoard_Status_Timer

End Sub


Sub READ_CLIPBOARD_HEX_CONVERSION()

Label1.Caption = ""
Label4.Caption = ""

Beep

If Clipboard.GetFormat(vbCFText) = True Then
    Label1.Caption = Clipboard.GetText
    Label20.Caption = Label1.Caption
    'LABEL20 WILL BE USED HERE AFTER
    'FOR DISPLAY REMOVE CONTROL CODE
    '-------------------------------
    
    Text3.Text = Label1.Caption

    
    Label1.Caption = Replace(Label1.Caption, vbLf, "")
    Label1.Caption = Replace(Label1.Caption, vbCr, "")
    If Len(Label20.Caption) <> Len(Label1.Caption) Then
        Label1.BackColor = Label22.BackColor
        Label21.BackColor = Label22.BackColor
        'info box
        Label6.BackColor = Label22.BackColor
        Label8.BackColor = Label22.BackColor
    End If
    
    'Label1.Caption = Replace(Label1.Caption, vbCr, "")
End If
Text2 = Label20.Caption

'16 THE MAX A VALUE ABLE READ OF A HEX
'TO DETECT IT IS A HEX
'SO READ THE FIRST 16 BYTE DOUBLE
'I = "&h" + Mid(Label20.Caption, 1, 16)
'Debug.Print Val(I)
'NOT A TRUE IS NUMERIC ONLY A VAL ABOVE OR BELOW 0


'ANOTHER METHOD DO IT BIT BY BIT AND IGNOR A NAUGHT 00 VALUE
'-----------------------------------------------------------
Label2 = "": Label3 = ""
i = Label20.Caption
i = Replace(i, " ", "")
XX = 0
For R = 1 To Len(i) Step 2
    I_2 = "&H" + Mid(i, R, 2)
    If Mid(i, R, 2) <> "00" Then
        If Val(I_2) = 0 Then
        XX = 1
    End If
    End If
Next
    
If XX = 0 Then
    Call Number_HEX_TO_TEXT_Click
    'Call Number_TEXT_TO_HEX_Click
Else
    Call Number_TEXT_TO_HEX_Click
End If

End Sub

Private Sub Label1_Click()

Call READ_CLIPBOARD_HEX_CONVERSION

End Sub


Private Sub Label2_Click()
Clipboard.Clear
Clipboard.SetText Label2
End Sub

Private Sub Label3_Click()
Clipboard.Clear
Clipboard.SetText Label3
Beep

Text3.Text = Label3

'ClipBoard_Status_Force_Through = True
'Call Timer_ClipBoard_Status_Timer

End Sub



Private Sub Label4_Click()
Clipboard.Clear
Clipboard.SetText Label4.Caption
Beep
Text3.Text = Label4.Caption

'ClipBoard_Status_Force_Through = True
'Call Timer_ClipBoard_Status_Timer
End Sub

Private Sub Mnu_Exit_Click()
Unload Me
End Sub

Private Sub MNU_VB_FOLDER_Click()
MsgBox "NOT WORKER YET"
End Sub

Private Sub MNU_VB_ME_Click()
MsgBox "NOT WORKER YET"
End Sub

' Click event procedure for number keys (0-9).
' Append new number to the number in the display.
Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub


Private Sub Number_TEXT_TO_HEX_Click()
Dim HH
i = Label20.Caption
Label4.Caption = ""

'WITH SPACE
Label2.Caption = "": Label5.Caption = ""

For R = 1 To Len(Label20.Caption)
    HH = Hex$(Asc(Mid(i, R, 1)))
    If Len(HH) = 1 Then
        HH = "0" + HH
    End If
    Label4.Caption = Label4.Caption + HH
    Label2.Caption = Label2.Caption + HH + " "
    Label5.Caption = Label5.Caption + HH + " "
Next

Text1 = Label2

End Sub

Private Sub Number_HEX_TO_TEXT_Click()
Dim R, i, I_2
i = Label20.Caption
i = Replace(i, " ", "")
Label3.Caption = ""
For R = 1 To Len(i) Step 2
    I_2 = "&H" + Mid(i, R, 2)
    Label3.Caption = Label3.Caption + Chr(Val(I_2))
Next
End Sub

' Click event procedure for operator keys (+, -, x, /, =).
' If the immediately preceeding keypress was part of a
' number, increments NumOps. If one operand is present,
' set Op1. If two are present, set Op1 equal to the
' result of the operation on Op1 and the current
' input string, and display the result.
Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = CDbl(Op1) + CDbl(Op2)
            Case "-"
                Op1 = CDbl(Op1) - CDbl(Op2)
            Case "X"
                Op1 = CDbl(Op1) * CDbl(Op2)
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Can't divide by zero", 48, "Calculator"
                Else
                   Op1 = CDbl(Op1) / CDbl(Op2)
                End If
            Case "="
                Op1 = CDbl(Op2)
            Case "%"
                Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

' Click event procedure for percent key (%).
' Compute and display a percentage of the first operand.
Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub

Private Sub Text1_Click()

Clipboard.Clear
Clipboard.SetText Label4.Caption

Beep

Text3 = Label4.Caption

Exit Sub

'On Error Resume Next
'XX_CNT = 0
'Do
'    XX_CNT = XX_CNT + 1
'    SLEEP 20
'
'    Text3 = Clipboard.GetText
'Loop Until Err.Number = 0 Or XX_CNT > 500
'
'If Err.Number > 0 Then
'    Text3 = Clipboard.GetText
'
'    If Err.Number > 0 Then MsgBox "RETRY 500 -- ERROR READ CLIPBOARD.GETTEXT -- USELESS INFO", vbMsgBoxSetForeground
'
'End If


'ClipBoard_Status_Force_Through = True
'Call Timer_ClipBoard_Status_Timer
End Sub


Private Sub Text2_Click()
Clipboard.Clear
Clipboard.SetText Label4.Caption
Beep

Text3 = Label4.Caption

'ClipBoard_Status_Force_Through = True
'Call Timer_ClipBoard_Status_Timer


End Sub

Public Sub FindCursor(X, Y)
    Dim P As POINTAPI
    GetCursorPos P
    X = P.X
    Y = P.Y
End Sub


Private Sub Text3_Change()

CLIPBOARD_FLASH_COUNT = 10
Timer_ClipBoard_Change_Flash_Colour.Enabled = True

End Sub

Private Sub Text3_Click()

Beep

End Sub

Private Sub Timer_ClipBoard_Change_Flash_Colour_Timer()

Timer_ClipBoard_Change_Flash_Colour.Interval = 400

CLIPBOARD_FLASH_COUNT = CLIPBOARD_FLASH_COUNT - 1

If CLIPBOARD_FLASH_COUNT Mod 2 = 0 Then
    Text3.BackColor = Label22.BackColor
    Label24.BackColor = RGB(255, 255, 255)
Else
    Text3.BackColor = RGB(255, 255, 255)
    Label24.BackColor = Label22.BackColor
End If

If CLIPBOARD_FLASH_COUNT < 0 Then
    Timer_ClipBoard_Change_Flash_Colour.Enabled = False
End If

End Sub

Private Sub Timer_ClipBoard_Status_Timer()

Dim FLAG_VAR_EXIT_SUB

Dim Stat_Active_Idle, Contain_Text

If ClipBoard_Status_Force_Through = True Then Set_To_Beep = True

If ClipBoard_Status_Force_Through = False Then
    
    Dim X, Y
    
    Call FindCursor(X, Y)
    
    
    If OX = X And OY = Y Then
        OX = X: OY = Y
        Label23.BackColor = RGB(255, 255, 255)
        FLAG_VAR_EXIT_SUB = True 'Exit Sub
    Else
        Label23.BackColor = Label22.BackColor
        OX = X: OY = Y
    End If
    
End If

If GetForegroundWindow <> Me.hWnd Then
    FLAG_VAR_EXIT_SUB = True 'Exit Sub
End If

'----------------------------------------------------------
'DONT POLL INTERVAL THE CLIPBOARD IF MACHINE IDLE
'LEAVE THE STATE STATUS FLAG VARIABLE AS IT IS
'----------------------------------------------------------
If FLAG_VAR_EXIT_SUB = False Or ClipBoard_Status_Force_Through = True Then
    If Clipboard.GetFormat(vbCFText) = False Then
        Clipboard_Contain_Text = "False"
    Else
        Clipboard_Contain_Text = "True"
    End If
End If

If Label23.BackColor = Label22.BackColor Then
    Stat_Active_Idle = "Active = True"
Else
    Stat_Active_Idle = "Idle = True"
End If
    
i = "Clip Board View Interval Update" + vbCrLf + "Active <> Idle " + vbCrLf + Stat_Active_Idle + vbCrLf + "Colouring"
i = i + vbCrLf + "Contain Text " + vbCrLf + Clipboard_Contain_Text
If i <> Label23.Caption Then Label23.Caption = i
'----------------------------------------------------------
If FLAG_VAR_EXIT_SUB = True Then Exit Sub
ClipBoard_Status_Force_Through = False
If Clipboard_Contain_Text = "False" Then Exit Sub
'----------------------------------------------------------

If Text3.Text <> Clipboard.GetText Then
    Text3.Text = Clipboard.GetText
        
        
    
    
    '-------------------------------------------------
    'Beep happen only when change not Force to Checker
    'Beep happen on First Change
    '-------------------------------------------------
    If Set_To_Beep = True Then
        Beep
    End If
End If

End Sub


Private Sub Timer1_Timer()



End Sub
