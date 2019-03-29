VERSION 5.00
Begin VB.Form ShowPicFrm2 
   BorderStyle     =   0  'None
   Caption         =   "  "
   ClientHeight    =   3015
   ClientLeft      =   2250
   ClientTop       =   4020
   ClientWidth     =   5175
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   2985
      ScaleWidth      =   3660
      TabIndex        =   0
      Top             =   45
      Width           =   3660
   End
   Begin VB.Menu mnpopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu umn 
         Caption         =   "On Top"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu umn 
         Caption         =   "Locked"
         Index           =   1
      End
      Begin VB.Menu umn 
         Caption         =   "Copy"
         Index           =   2
      End
      Begin VB.Menu umn 
         Caption         =   "Print"
         Index           =   3
      End
      Begin VB.Menu umn 
         Caption         =   "Close"
         Index           =   4
      End
   End
End
Attribute VB_Name = "ShowPicFrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'showpicfrm2
'*********************************
Option Explicit
'*********************************

Private Const HWND_NOTOPMOST& = -2
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVESIZE& = SWP_NOMOVE Or SWP_NOSIZE
Private Const SWP_FRAMECHANGED& = &H20
Private Const WS_BORDER& = &H800000
Private Const WS_CAPTION& = &HC00000
Private Const GWL_STYLE& = (-16)
'---------------------------------
Private Type POINTAPI
  x As Long
  y As Long
End Type
'---------------------------------
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd&, ByVal nIndex&)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&)
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hIA&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
'---------------------------------
Dim xDiff!, yDiff!, CmdPic$, Ctrl&, z&, zz&
Dim StartUp$, tpx&, tpy&, SW&, SH&
Dim AG%, Style&, pt As POINTAPI
Dim Locked As Boolean, TopMost As Boolean
'*********************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub

If KeyCode = 17 Then Ctrl = 1
If Ctrl = 0 Then Exit Sub

'# Set the form TopMost/NoTopmost; lock/unlock the form
Select Case KeyCode
Case 76: If Locked = False Then Locked = True: Exit Sub '# "L"
Case 85: If Locked = True Then Locked = False: Exit Sub '# "U"
Case 84: Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVESIZE):   TopMost = True:  Exit Sub '"T"
Case 78: Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVESIZE): TopMost = False: Exit Sub '"N"
Case Else: If Locked Then Exit Sub
End Select

z = z + 1
If z Mod 4 = 0 Then zz = zz + 1 '# use acceleration

'# Move the form using arrow keys
Select Case KeyCode
Case 37:  Left = Left - tpx * zz
Case 38:  Top = Top - tpy * zz
Case 39:  Left = Left + tpx * zz
Case 40:  Top = Top + tpy * zz
End Select
End Sub

'*********************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
z = 0: zz = 1
If KeyCode = 17 Then Ctrl = 0
End Sub


'*********************************
Private Sub Form_Load()

'THIS CONTROLS ---- FMAIN ----- SHUT IT DOWN WHILE HERE TAKES PROCESS TIME

ShowPicFrm2Loaded = True
On Error GoTo Handler

'# Hide the titlebar, but let the popup menu appear
Style = GetWindowLong(Me.hwnd, GWL_STYLE)
Style = Style And Not WS_CAPTION

Call SetWindowLong(Me.hwnd, GWL_STYLE, Style)
Call SetWindowPos(Me.hwnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED)

With Screen
  tpx = .TwipsPerPixelX
  tpy = .TwipsPerPixelY
  SW = .Width
  SH = .Height
End With

zz = 1
'CmdPic = Trim$(Command)
'CmdPicpic = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\0-Plus 800K\Img_1479.jpg"

CmdPic = A1 + G1
If LenB(CmdPic) = 0 Then
  Beep
  Unload Me
  Exit Sub
End If
'MsgBox CmdPic

If Left$(CmdPic, 1) = Chr$(34) Then
  CmdPic = Mid$(CmdPic, 2, Len(CmdPic) - 2)
End If

'# Get the StartUp position
'Call ReadRegistry
'Left = Val(Left$(StartUp, InStr(StartUp, ",") - 1))
'Top = Val(Mid$(StartUp, InStr(StartUp, ",") + 1))

'# Load the picture and position the picturebox
'With Picture1
'  .Picture = LoadPicture(CmdPic)
'  If .Width > SW Then               '# Breite größer als Bildschirm
'    Left = 0: Width = SW            '# Width larger than screen
'    .Left = (Width - .Width) \ 2
'    AG = AG + 1
'  Else
'    .Left = 0: Width = .Width
'  End If
'  If .Height > SH Then              '# Höhe größer als Bildschirm
'    Top = 0: Height = SH            '# Height larger than screen
'    .Top = (Height - .Height) \ 2
'    AG = AG + 2
'  Else
'    .Top = 0: Height = .Height
'  End If
'End With

Dim ADJUSTX, ADJUSTY, AX, AY, RATIO
Dim COMPSIZE, TAG1, TAG2, TAG3, COUNTER
Dim PASSES, Comps$

With Picture1
  .Picture = LoadPicture(CmdPic)
    
    .Left = 0
    .Top = 0
    
    ADJUSTX = (Screen.Width) '/ Screen.TwipsPerPixelX)
    ADJUSTY = (Screen.Height) ' / Screen.TwipsPerPixelY)
    
    AX = .Width
    AY = .Height
    
      
    'ratio = 1 + (1 / 2000) 'Slower = More Passes
    RATIO = 1 + (1 / 200) 'Faster = Less Passes
      
    COMPSIZE = 1
      
    If AX = 0 Or AY = 0 Then Exit Sub
      
      TAG1 = 0
      TAG2 = 0
      Do
'      DoEvents
      TAG3 = 0
      If AX > ADJUSTX Or AY > ADJUSTY Then
      AX = AX / RATIO: AY = AY / RATIO: TAG1 = 1
      COMPSIZE = COMPSIZE / RATIO
      TAG3 = 1
      End If
      
      If AX < ADJUSTX - 0.1 And AY < ADJUSTY - 0.1 Then
      AX = AX * RATIO: AY = AY * RATIO: TAG2 = 1
      COMPSIZE = COMPSIZE * RATIO
      TAG3 = 1
      End If
      
      'DoEvents
      COUNTER = COUNTER + 1
      If COUNTER > 90000 Then Stop
      Loop Until (TAG1 = 1 And TAG2 = 1) Or (TAG3 = 0)
      
      ADJUSTX = AX
      ADJUSTY = AY
          
       .Width = ADJUSTX
       .Height = ADJUSTY
       
    Me.Height = .Height
    Me.Width = .Width
    
       
            
      
      PASSES = COUNTER
      Comps$ = Format$(COMPSIZE, "##0.0000")


End With



'MsgBox AG
Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVESIZE)
ShowPicFrm2.Show

TopMost = True


Exit Sub

Handler:
msgbox Err.Description, 64
Unload Me

End Sub

'*********************************
Private Sub Form_Unload(Cancel%)
'If AG = 0 Then Call WriteRegistry

ShowPicFrm2Loaded = False


'FileInfoTimer.Enabled = T1


End Sub

'*********************************
Private Sub Picture1_DblClick()
If Locked = False Then

Unload Me
End If
End Sub

'*********************************
Private Sub Picture1_MouseDown(Button%, Shift%, x!, y!)



fMain.MNU_FULL_SCREEN_YES.Checked = False
Unload Me

Exit Sub
If Button = 2 Then
  umn(0).Checked = TopMost
  umn(1).Checked = Locked
  Call GetCursorPos(pt)
  PopupMenu mnpopup, 4, , (pt.y - 7) * tpy - Top
  Exit Sub
End If
If Locked Then Exit Sub
xDiff = x: yDiff = y
End Sub

'*********************************
'# Because a form cannot be larger than the screen, but a picturebox can,
'# we have to move the picturebox rsp. the form depending on the size of
'# the picture and the current position of the picture.

Private Sub Picture1_MouseMove(Button%, Shift%, x!, y!)
Exit Sub

Static AX&, AY&, LF&, lp&, TF&, TP&

If Locked Then Exit Sub
If Button <> 1 Then Exit Sub

AX = x - xDiff: AY = y - yDiff
With Picture1
  Select Case AG                        '# AG tells about the relation between picture size and screen size
  Case 0: Move Left + AX, Top + AY      '# Nur Form verschieben/move only form
  
  Case 1                                '# Breite größer als Bildschirm/width larger than screen
    If AX > 0 Then                      '# Maus nach rechts/mouse to the right
      If Left < 0 Then                  '# Form ist links außerhalb/form is left outside
        Move Left + AX, Top + AY        '# Form bewegen/move form
      Else
        If .Left + AX > 0 Then          '# Bild wäre nach rechts verschoben/picture would be moved to the right
          .Move 0, .Top                 '# Auf Null halten und/keep picture on zero and
          Move Left + AX, Top + AY      '# Form verschieben/move form
        Else
          Move 0, Top + AY              '# Form auf Null halten und/keep form on zero and
          .Move .Left + AX, .Top        '# Bild verschieben/move picture
        End If
      End If
    Else                                '# Maus nach links/mouse to the left
      If Left + AX > 0 Then
        Move Left + AX, Top + AY
      Else
        If .Left + .Width > Width Then
          .Move .Left + AX, .Top
          Move 0, Top + AY
        Else
          .Move Width - .Width, .Top
          Move Left + AX, Top + AY
        End If
      End If
    End If
  
  Case 2                                '# Höhe größer als Bildschirm/height larger than screen
    If AY > 0 Then                      '# Maus nach unten/mouse down
      If Top < 0 Then                   '# Form ist oben außerhalb/form is outside above
        Move Left + AX, Top + AY        '# Form bewegen/move form
      Else
        If .Top + AY > 0 Then           '# Bild wäre nach unten verschoben/picture would be moved down
          .Move .Left, 0                '# Auf Null halten und/keep on zero and
          Move Left + AX, Top + AY      '# Form verschieben/move form
        Else
          Move Left + AX, 0             '# Form auf Null halten und/keep form on zero and
          .Move .Left, .Top + AY        '# Bild verschieben/move picture
        End If
      End If
    Else                                '# Maus nach oben/mouse up
      If Top + AY > 0 Then
        Move Left + AX, Top + AY
      Else
        If .Top + .Height > Height Then
          .Move .Left, .Top + AY
          Move Left + AX, 0
        Else
          .Move .Left, Height - .Height
          Move Left + AX, Top + AY
        End If
      End If
    End If
  
  Case 3                                '# beides größer als Bildschirm/both larger than screen
    If AX > 0 Then                      '# Maus nach rechts/mouse to the right
      If Left < 0 Then                  '# Form ist links außerhalb/form left outside
        LF = Left + AX
      Else
        If .Left + AX > 0 Then          '# Bild wäre nach rechts verschoben
          lp = 0: LF = Left + AX
        Else
          LF = 0: lp = .Left + AX
        End If
      End If
    Else                                '# Maus nach links/mouse to the left
      If Left + AX > 0 Then
        LF = Left + AX
      Else
        If .Left + .Width > Width Then
          lp = .Left + AX: LF = 0
        Else
          lp = Width - .Width: LF = Left + AX
        End If
      End If
    End If
    If AY > 0 Then                      '# Maus nach unten/mouse down
      If Top < 0 Then                   '# Form ist oben außerhalb/form is outside above
        TF = Top + AY
      Else
        If .Top + AY > 0 Then           '# Bild wäre nach unten verschoben
          TP = 0: TF = Top + AY
        Else
          TF = 0: TP = .Top + AY
        End If
      End If
    Else                                '# Maus nach oben/mouse up
      If Top + AY > 0 Then
        TF = Top + AY
      Else
        If .Top + .Height > Height Then
          TP = .Top + AY: TF = 0
        Else
          TP = Height - .Height: TF = Top + AY
        End If
      End If
    End If
    Move LF, TF
    .Move lp, TP
  End Select
End With
End Sub

'*********************************
Private Sub ReadRegistry()
'StartUp = GetSetting(App.Title, "StartUp", "Position", "600,600")
End Sub

'*********************************
Private Sub WriteRegistry()
Exit Sub
Dim ddl$, ddt$
If Left < 0 Then ddl = "0" Else ddl = Trim$(Left)
If Top < 0 Then ddt = "0" Else ddt = Trim$(Top)
Call SaveSetting(App.title, "StartUp", "Position", ddl & "," & ddt)
End Sub

'*********************************
Private Sub umn_Click(index As Integer)
'Exit Sub

Dim chk&
Select Case index
Case 0: umn(0).Checked = Not umn(0).Checked
        chk = IIf(umn(0).Checked, HWND_TOPMOST, HWND_NOTOPMOST)
        SetWindowPos hwnd, chk, 0, 0, 0, 0, SWP_NOMOVESIZE
        TopMost = umn(0).Checked
Case 1: Locked = Not Locked: umn(1).Checked = Locked
Case 2: Clipboard.SetData Picture1.Picture
Case 3: Printer.PaintPicture Picture1.Picture, 0, 0
        Printer.EndDoc
Case 4: Unload Me
End Select
End Sub

'*********************************

