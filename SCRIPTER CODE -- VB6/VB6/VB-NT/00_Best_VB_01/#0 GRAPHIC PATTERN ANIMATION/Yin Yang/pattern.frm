VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   Caption         =   "Pattern"
   ClientHeight    =   6312
   ClientLeft      =   60
   ClientTop       =   1032
   ClientWidth     =   9156
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   526
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   763
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4170
      Left            =   0
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   4176
      ScaleWidth      =   5112
      TabIndex        =   0
      Top             =   0
      Width           =   5115
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   7680
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_CROP 
      Caption         =   "Crop"
   End
   Begin VB.Menu Mnu_Save 
      Caption         =   "Save"
   End
   Begin VB.Menu MNU_PAUSE 
      Caption         =   "PAUSE"
   End
   Begin VB.Menu MNU_STEP 
      Caption         =   "STEP"
   End
   Begin VB.Menu MNU_STEP_NUMBERING 
      Caption         =   "STEP #"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_DRAW_TIME 
      Caption         =   "DRAW TIME"
   End
   Begin VB.Menu MNU_CROP_HOZI__UP 
      Caption         =   "CROP_X_HOZI__UP"
   End
   Begin VB.Menu MNU_CROP_HOZI__DOWN 
      Caption         =   "CROP_X_HOZI__DOWN"
   End
   Begin VB.Menu MNU_CROP_VERTI_LEFT 
      Caption         =   "CROP_Y_VERTI_LEFT"
   End
   Begin VB.Menu MNU_CROP_VERTI_RIGHT 
      Caption         =   "CROP_Y_VERTI_RIGHT"
   End
   Begin VB.Menu MNU_X_MAX 
      Caption         =   "X MAX"
   End
   Begin VB.Menu MNU_Y_MAX 
      Caption         =   "Y MAX"
   End
   Begin VB.Menu MNU_X_MIN 
      Caption         =   "X MIN"
   End
   Begin VB.Menu MNU_Y_MIN 
      Caption         =   "Y MIN"
   End
   Begin VB.Menu MNU_BOX_W 
      Caption         =   "BOX_W"
   End
   Begin VB.Menu MNU_BOX_H 
      Caption         =   "BOX_H"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public TIME_BEGIN As Date

Public STEP_NUMBER
Public X_MAX_IT
Public Y_MAX_IT
Public X_MIN_IT
Public Y_MIN_IT

Public FILE_NAME

Public I, CROPPER_HOZ_, CROPPER_VERT

Public F1_H, F1_W

Public FF1
Public FF2
Public FT1
Public TT1, TT2, TT3

Public XXL1, XXH1_, YYL1, YYH1_

Public RA1, RA2, RA3

Public NOIN As Integer
Public ATIM As Double
Public HANDLE As Integer
Public RGB_W_ As Integer
Public RGB_HW As Double
Public RGB_KW As Double
Public TW1 As Double
Public TW2 As Double
Public TW3 As Double
Public RR4 As Double
Public X3, Y3, X8
Public X4 As Single
Public Y4 As Single
Public X5 As Single
Public Y5 As Single
Public PI As Double
Public TR

Dim H1_, HE1
Dim H2_, HE2_
Dim H3_, HE3_

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long



Private Sub Form_Activate()

DoEvents
On Error Resume Next
Me.Top = 10

End Sub

Private Sub Form_Load()

TIME_BEGIN = Now


'HANDLE = Shell("C:\VBDOS\VBDOS /ah /RUN C:\COOLCID\CIBAS /Es /S:0 ", vbNormalFocus)

FILE_NAME = App.Path + "\X__Y_SETTING.TXT"
If Dir(FILE_NAME) <> "" Then
    FR1 = FreeFile
    Open FILE_NAME For Input As #FR1
        Line Input #FR1, CROPPER_VERT_4
        Line Input #FR1, CROPPER_HOZ__4
    Close #FR1
    CROPPER_VERT = Val(CROPPER_VERT_4)
    CROPPER_HOZ_ = Val(CROPPER_HOZ__4)
End If



PI = 4 * Atn(1)
TW1 = 50
'TW2 = 2.4
'TW3 = 2.8
'TW1 = 0.1
'TW2 = 0.14
'TW3 = 0.18

'------------------------
F1_H = 11700
F1_W = 11700
'------------------------

Form1.Height = F1_H
Form1.Width = F1_W

'FF1 = 10000 / 315
'FF2 = 9650 / 315

'Picture1.Align = 1
'Picture1.Align = 2
'Picture1.Align = 3
'Picture1.Align = 4
'Picture1.Align = 0


'Picture1.DrawWidth = 100
'Picture1.DrawWidth = 50

Picture1.BackColor = RGB(255, 255, 255) '__WHITE
'Picture1.BackColor = RGB(0, 0, 0)      '__BLACK
'Picture1.BackColor = RGB(0, 0, 100)    '__BLUE DARKER SHADE

'Picture1.ScaleHeight = Form1.Height
'Picture1.ScaleLeft = 0
'Picture1.ScaleTop = 0
'Picture1.ScaleWidth = Form1.Width



Form1.Height = F1_H
Form1.Width = F1_W

'Picture1.Width = (Form1.Width / Screen.TwipsPerPixelX) - 15
'Picture1.ScaleWidth = 100
'Picture1.ScaleHeight = 100

Picture1.Width = (Form1.Width / Screen.TwipsPerPixelX) - 40
Picture1.Height = (Form1.Height / Screen.TwipsPerPixelY) - 90
Picture1.Top = 0
Picture1.Left = 0

Picture1.DrawWidth = 120 '70

X3 = (Picture1.Width * Screen.TwipsPerPixelX) / 2 '* Screen.TwipsPerPixelX '/ FF1 ' 315 'LEFT 31.7
Y3 = (Picture1.Height * Screen.TwipsPerPixelX) / 2 ' * Screen.TwipsPerPixelY '/ FF2 ' 315 'UP

X3 = X3 - 150
Y3 = Y3 - 150
X8 = X3 * 1.016 'RADIUS
X8 = X3 * 1.016 'RADIUS

XXL1 = Picture1.Width * Screen.TwipsPerPixelX
YYL1 = Picture1.Height * Screen.TwipsPerPixelY


'Me.WindowState = vbNormal
Me.Top = 10
Me.Left = Me.Width / 2


I = 0
Call CROPPER

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'PSet (X, Y)                 ' DRAW A POINT.
'X3 = X: Y3 = Y
'DrawWidth = 90              ' USE WIDER BRUSH.  ' USE WIDER BRUSH.
'ForeColor = RGB(KW, HW, W)  ' Set drawing color.

End Sub

Private Sub Form_Resize()

'DoEvents
'Me.Top = 10

End Sub

Private Sub Form_Unload(Cancel As Integer)

Clipboard.Clear
Clipboard.SetText MNU_STEP.Caption + vbCrLf + MNU_DRAW_TIME.Caption

End Sub

Private Sub Mnu_Crop_Click()

'Picture1.Top = XXL1
'Picture1.Left = YYL1

Picture1.Top = 0
Picture1.Left = 0

DW = Picture1.DrawWidth + 9
RW1 = (YYH1_ - YYL1) / Screen.TwipsPerPixelY
RW2 = (XXH1_ - XXL1) / Screen.TwipsPerPixelX
RW1 = RW1 + DW
RW2 = RW2 + DW

'Picture1.Height = RW1 + 120 '(((YYH1_ - YYL1) + Picture1.DrawWidth + 9)) - 150
'Picture1.Width = RW2 + 120  '(((XXH1_ - XXL1) + Picture1.DrawWidth + 9)) - 150
'Picture1.Align = 0

If CROPPER_HOZ_ = 0 Then CROPPER_HOZ_ = RW1
If CROPPER_VERT = 0 Then CROPPER_VERT = RW1

Picture1.Height = CROPPER_HOZ_
Picture1.Width = CROPPER_VERT

Form1.Height = F1_H
Form1.Width = F1_W

End Sub



Sub CROPPER()

If CROPPER_HOZ_ = 0 Then
    CROPPER_HOZ_ = Picture1.Height
End If
If CROPPER_VERT = 0 Then
    CROPPER_VERT = Picture1.Width
End If
If I = 1 Then CROPPER_HOZ_ = CROPPER_HOZ_ - 1
If I = 2 Then CROPPER_HOZ_ = CROPPER_HOZ_ + 1

If I = 3 Then CROPPER_VERT = CROPPER_VERT - 1
If I = 4 Then CROPPER_VERT = CROPPER_VERT + 1

Me.MNU_CROP_VERTI_LEFT.Caption = "CROP_Y_VERTI_LEFT__" + Str(CROPPER_VERT)
Me.MNU_CROP_VERTI_RIGHT.Caption = "CROP_Y_VERTI_RIGHT__" + Str(CROPPER_VERT)
Me.MNU_CROP_HOZI__UP.Caption = "CROP_X_HOZI_UP__" + Str(CROPPER_HOZ_)
Me.MNU_CROP_HOZI__DOWN.Caption = "CROP_X_HOZI__DOWN__" + Str(CROPPER_HOZ_)


FR1 = FreeFile
Open FILE_NAME For Output As #FR1
    Print #FR1, CROPPER_VERT
    Print #FR1, CROPPER_HOZ_
Close #FR1


F1_H = F1_H - 1
F1_W = F1_W - 1

F1_H = F1_H + 1
F1_W = F1_W + 1

Picture1.Height = CROPPER_HOZ_
Picture1.Width = CROPPER_VERT
If I = 0 Then Picture1.Refresh

On Error Resume Next
Form1.Height = F1_H
Form1.Width = F1_W

Call SET_MENU_PBOX_DIMENSION_LABEL

End Sub


Private Sub MNU_CROP_VERTI_LEFT_Click()

I = 3
Call CROPPER

End Sub

Private Sub MNU_CROP_VERTI_RIGHT_Click()

I = 4
Call CROPPER

End Sub

Private Sub MNU_CROP_HOZI__UP_Click()

I = 1
Call CROPPER

End Sub

Private Sub MNU_CROP_HOZI__DOWN_Click()

I = 2
Call CROPPER

End Sub

Private Sub MNU_PAUSE_Click()

If Timer1.Enabled = False Then Timer1.Enabled = True: Exit Sub
If Timer1.Enabled = True Then Timer1.Enabled = False: Exit Sub

End Sub

Private Sub Mnu_Save_Click()
Timer1.Enabled = False

'Clipboard.Clear
'Clipboard.SetData (Picture1.Image  PictureBox)
'Picture1.Picture SavePicture(App.Path + "KK.BMP")

RR = Format(Now, "yyyy mm dd - hh mm ss")
SavePicture Picture1.Image, App.Path + "\YIN YANG " + RR + ".BMP"

Shell "EXPLORER /SELECT, """ + App.Path + "\YIN YANG " + RR + ".BMP""", vbNormalFocus

End
End Sub

Private Sub MNU_STEP_Click()
Call Timer1_Timer
Timer1.Enabled = False
End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    I = CODER_VBP_FILE_NAME_2
    If InStr(I, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(I, " - 64 BIT.vbp", ".vbp")
    If InStr(I, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(I, " - 32 BIT.vbp", ".vbp")
    
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
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub
Private Sub Mnu_VB_ME_Click()
    Beep
    If GetSpecialfolder(38) = "" Then
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    If IsIDE = True Then
        If IsIDE = True Then
            MSGQ = vbYesNoCancel
        Else
            MSGQ = vbOKOnly
        End If
        IR = MsgBox("NOT IN IDE ARE YOU SURE", MSGQ, vbMsgBoxSetForeground)
        If IR <> vbYes Then Exit Sub
    End If
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Beep
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub
'-------------------------------------------
'-------------------------------------------
'---------------------------------------------------
'Private Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Function GetSpecialfolder(ByVal CSIDL As Long) As String
    Dim R As Long
    Dim IDL As ITEMIDLIST
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If R = NOERROR Then
        Path$ = Space$(512)
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function
'---------------------------------------------------



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Timer1.Enabled = False

'Clipboard.Clear
'Clipboard.SetData (Picture1.Image  PictureBox)

'Picture1.Picture SavePicture(App.Path + "KK.BMP")
'SavePicture Picture1.Image, App.Path + "\YINYANG1.BMP"

'Shell "EXPLORER " + App.Path
End Sub

Private Sub Timer1_Timer()

'----------------------
H1_ = -1: HE1_ = Abs(H1_)
H2_ = -3: HE2_ = Abs(H2_)
H3_ = -5: HE3_ = Abs(H3_)
'----------------------
'NOT USED
'H1_ = -5: HE1_ = Abs(H1_)
'H2_ = -8: HE2_ = Abs(H2_)
'H3_ = -10: HE3_ = Abs(H3_)
'----------------------

TG_ = 255
HG12 = 10 ' LOW AND DARK
RGB_W_ = RGB_W_ + TW1

If RGB_W_ >= TG_ - HE1_ Then TW1 = H1_
If RGB_W_ < HE1_ + HG12 Then TW1 = HE1_
RGB_HW = RGB_HW + TW2
If RGB_HW >= TG_ - HE2_ Then TW2 = H2_
If RGB_HW < HE2_ + HG12 Then TW2 = HE2_
RGB_KW = RGB_KW + TW3
If RGB_KW >= TG_ - HE3_ Then TW3 = H2_
If RGB_KW < HE3_ + HG12 Then TW3 = HE3_

'------------------------------------

'TT1 = 0.09
'TT2 = 0.04
'TT3 = 0.02

TT1 = 0.01
TT2 = 0.012
TT3 = 0.014
RA1 = RA1 + TT1
RA2 = RA2 + TT2
RA3 = RA3 + TT3

'COLOUR CYCLE ABOVE
'------------------
'RR1 = 127 * Sin(RA1) + 127
'RR2 = 127 * Sin(RA2) + 127
'RR3 = 127 * Cos(RA3) + 127

RR1 = 100 * Sin(RA1) + 150
RR2 = 100 * Sin(RA2) + 150
RR3 = 100 * Cos(RA3) + 150

'Picture1.ForeColor = RGB(Int(RR1), Int(RR2), Int(RR3)) 'W   ' SET DRAWING COLOR.

'PICTURE1.FORECOLOR = RGB(255, 0, 0) 'RGB(RGB_KW, RGB_HW, RGB_W_)   ' SET DRAWING COLOR.

Picture1.ForeColor = RGB(RGB_KW, RGB_HW, RGB_W_)   ' SET DRAWING COLOR.

'QUAD ORGINAL
'RR4 = RR4 + 0.005 '* 3

'QUAD ORGINAL MOIDIFIED
'RR4 = RR4 + 0.005 * 5
RR4 = RR4 + 0.003



'2015-06-022
'RR4 = RR4 + 0.0005

'TRIPLE
'RR4 = RR4 + 0.005 * 2

'TEST
'RR4 = RR4 + 0.000001

'COLOUR
'ORGINAL QUAD + SPEED
'TR = TR + 0.01 * 2
'TR = TR + 0.01 * 1
TR = TR + 0.001

'If TR > 62.9 Then Stop: Timer1.Enabled = False

LW = (X8 / 1.2) * Sin(TR)
Y4 = LW * Sin(RR4) + X3
X4 = LW * Cos(RR4) + Y3

'X5 = 200 * Sin(RR3 - 3) + X3
'Y5 = 200 * Cos(RR3 - 3) + Y3

'X3 = Form1.Height / 2 '/ FF1 ' 315 'LEFT 31.7
'Y3 = Form1.Width / 2 '/ FF2 ' 315 'UP

OFFSET2 = 0
OFFSET_X_HOZ_2 = 120 + OFFSET2
OFFSET_Y_VERT2 = 355 + OFFSET2

Picture1.DrawWidth = 140 '70
Picture1.PSet (X4 + OFFSET_X_HOZ_2, Y4 + OFFSET_Y_VERT2) ' DRAW A POINT.


If X_MIN_IT = 0 Then X_MIN_IT = X4 + OFFSET_X_HOZ_2
If Y_MIN_IT = 0 Then Y_MIN_IT = Y4 + OFFSET_Y_VERT2

If X4 + OFFSET_X_HOZ_2 > X_MAX_IT Then X_MAX_IT = X4 + OFFSET_X_HOZ_2
If Y4 + OFFSET_Y_VERT2 > Y_MAX_IT Then Y_MAX_IT = Y4 + OFFSET_Y_VERT2
If X4 + OFFSET_X_HOZ_2 < X_MIN_IT Then X_MIN_IT = X4 + OFFSET_X_HOZ_2
If Y4 + OFFSET_Y_VERT2 < Y_MIN_IT Then Y_MIN_IT = Y4 + OFFSET_Y_VERT2

VAR1 = "X MAX: " + Str(Int(X_MAX_IT / Screen.TwipsPerPixelX)) '+ Picture1.DrawWidth / 2)
VAR2 = "Y MAX: " + Str(Int(Y_MAX_IT / Screen.TwipsPerPixelY)) ' + Picture1.DrawWidth / 2)
VAR3 = "X MIN: " + Str(Int(X_MIN_IT / Screen.TwipsPerPixelX)) ' - Picture1.DrawWidth / 2)
VAR4 = "Y MIN: " + Str(Int(Y_MIN_IT / Screen.TwipsPerPixelY)) ' - Picture1.DrawWidth / 2)
'------------------------------------------
'REDUCE FLICKER DONT CHANGE IF NOT A CHANGE
'------------------------------------------
If MNU_X_MAX.Caption <> VAR1 Then MNU_X_MAX.Caption = VAR1
If MNU_Y_MAX.Caption <> VAR1 Then MNU_Y_MAX.Caption = VAR2
If MNU_X_MIN.Caption <> VAR1 Then MNU_X_MIN.Caption = VAR3
If MNU_Y_MIN.Caption <> VAR1 Then MNU_Y_MIN.Caption = VAR4


If X4 < XXL1 Then XXL1 = X4
If X4 > XXH1_ Then XXH1_ = X4
If Y4 < YYL1 Then YYL1 = Y4
If Y4 > YYH1_ Then YYH1_ = Y4

'Line (X4, Y4)-(X5, Y5)
'Circle (X4, Y4), 10
'ForeColor = RGB(0, 255, 0) 'RGB(KW, HW, W)   ' SET DRAWING COLOR.

STEP_NUMBER = STEP_NUMBER + 1
If STEP_NUMBER >= 3148 Then
    Timer1.Enabled = False
    'Exit Sub
End If
'MNU_STEP_NUMBERING.Caption = "STEP #+" + Str(STEP_NUMBER)
MNU_STEP.Caption = "STEP #+" + Str(STEP_NUMBER)
MNU_DRAW_TIME.Caption = "Draw Time =" + Str(DateDiff("s", TIME_BEGIN, Now))

Call SET_MENU_PBOX_DIMENSION_LABEL

End Sub

Sub SET_MENU_PBOX_DIMENSION_LABEL()

    VAR12 = Int(X_MAX_IT / Screen.TwipsPerPixelX)
    VAR22 = Int(Y_MAX_IT / Screen.TwipsPerPixelY)
    VAR32 = Int(X_MIN_IT / Screen.TwipsPerPixelX)
    VAR42 = Int(Y_MIN_IT / Screen.TwipsPerPixelY)
    
    Me.MNU_BOX_H.Caption = "BOX_H_" + Str(Picture1.Height) + "-" + Trim(Str(VAR12)) + "=" + Trim(Str(Picture1.Height - VAR12))
    Me.MNU_BOX_W.Caption = "BOX_W_" + Str(Picture1.Width) + "-" + Trim(Str(VAR22)) + "=" + Trim(Str(Picture1.Width - VAR22))

End Sub

