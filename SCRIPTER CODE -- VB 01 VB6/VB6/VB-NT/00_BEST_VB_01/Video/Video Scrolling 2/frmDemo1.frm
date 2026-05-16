VERSION 5.00
Begin VB.Form frmDemo1 
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmDemo1.frx":0000
   ScaleHeight     =   4005
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   2400
   End
   Begin VB.PictureBox picCredits 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3825
      Left            =   2900
      Picture         =   "frmDemo1.frx":57F42
      ScaleHeight     =   3765
      ScaleWidth      =   3660
      TabIndex        =   0
      Top             =   90
      Width           =   3725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here To Exit"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3645
      Width           =   2655
   End
End
Attribute VB_Name = "frmDemo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
'
'   SCROLLING AND FADING CREDITS
'
'   Written by Dirk Rijmenants (c) 2005
'------------------------------------------------------------

'Option Explicit

Dim CreditLine()    As String
Dim CreditLeft()    As Long
Dim ColorFades(100) As Long
Dim ScrollSpeed     As Integer
Dim ColText         As Long
Dim FadeIn          As Long
Dim FadeOut         As Long

Dim cDiff1          As Long
Dim cDiff2          As Double
Dim cDiff3          As Double

Dim TotalLines      As Integer
Dim LinesOffset     As Integer
Dim Yscroll         As Long
Dim CharHeight      As Integer
Dim LinesVisible    As Integer

Private Sub Form_Click()
End
End Sub

Private Sub Form_Load()
Dim FileO       As Integer
Dim fileName    As String
Dim tmp         As String
Dim i           As Integer

Dim Rcol1       As Long
Dim Gcol1       As Long
Dim Bcol1       As Long

Dim Rcol2       As Long
Dim Gcol2       As Long
Dim Bcol2       As Long

Dim Rfade       As Long
Dim Gfade       As Long
Dim Bfade       As Long

Dim PercentFade As Integer
Dim TimeInterval As Integer
Dim AlignText  As Integer

'################################################################
'
'   Preset the text and fade properties
'
'   To change text and background color or font you have to
'   change these properties on the picCredits picturebox
'   You can also use any image as background
'
    PercentFade = 30
    'The PercentFade sets the percentage of the box is used
    ' to fade in and out (set to zero when image used as background)
    TimeInterval = 10
    ScrollSpeed = 15
'   you might need to experiment with the ScrollSpeed and TimeInterval
    AlignText = 2 '( 1=left 2=center 3=right )
    
'################################################################

'set the number of line to be printed in the box
LinesVisible = (picCredits.Height / picCredits.TextHeight("A")) + 1

'add empty lines at beginning to start off
For i = 1 To LinesVisible
    ReDim Preserve CreditLine(TotalLines) As String
    CreditLine(TotalLines) = tmp
    TotalLines = TotalLines + 1
Next

'Load the credits file
'  You could write the text also in the CreditLine(n) array,
'  but you than have to add the value TotalLines yourself!
FileO = FreeFile
fileName = App.Path & "\Credits.txt"
If Dir(fileName) = "" Then
        GoTo errHandler
        End If
On Error GoTo errHandler
Open fileName For Input As FileO
While Not EOF(FileO)
    Line Input #FileO, tmp
    ReDim Preserve CreditLine(TotalLines) As String
    CreditLine(TotalLines) = tmp
    TotalLines = TotalLines + 1
    Wend
Close #FileO

'set timer interval
Me.Timer1.Interval = TimeInterval

'set the number of line to be printed in the box
LinesVisible = (picCredits.Height / picCredits.TextHeight("A")) + 1

'Next, we calculate a lot of time-eating stuff in advance.
'This is done before, to speedup timer sub ;-)

'set the fade-in and fade-out regions
CharHeight = picCredits.TextHeight("A")
If PercentFade <> 0 Then
    FadeOut = ((picCredits.Height / 100) * PercentFade) - CharHeight
    FadeIn = (picCredits.Height - FadeOut) - CharHeight - CharHeight
    Else
    FadeIn = picCredits.Height
    FadeOut = 0 - CharHeight
    End If
    
'set the percent values, ready for instant use later
ColText = picCredits.ForeColor
cDiff1 = (picCredits.Height - (CharHeight - 10)) - FadeIn
cDiff2 = 100 / cDiff1
cDiff3 = 100 / FadeOut

'calculate the left-position of each line, to center it
ReDim CreditLeft(TotalLines - 1)
For i = 0 To TotalLines - 1
    Select Case AlignText
    Case 1
        CreditLeft(i) = 100
    Case 2
        CreditLeft(i) = (picCredits.Width - picCredits.TextWidth(CreditLine(i))) / 2
    Case 3
        CreditLeft(i) = picCredits.Width - picCredits.TextWidth(CreditLine(i)) - 100
    End Select
Next i

'calculate 100 fade values from backcolor to forecolor
'(another time-eating thing done in advance)
Rcol1 = picCredits.ForeColor Mod 256
Gcol1 = (picCredits.ForeColor And vbGreen) / 256
Bcol1 = (picCredits.ForeColor And vbBlue) / 65536
Rcol2 = picCredits.BackColor Mod 256
Gcol2 = (picCredits.BackColor And vbGreen) / 256
Bcol2 = (picCredits.BackColor And vbBlue) / 65536
For i = 0 To 100
    Rfade = Rcol2 + ((Rcol1 - Rcol2) / 100) * i: If Rfade < 0 Then Rfade = 0
    Gfade = Gcol2 + ((Gcol1 - Gcol2) / 100) * i: If Gfade < 0 Then Gfade = 0
    Bfade = Bcol2 + ((Bcol1 - Bcol2) / 100) * i: If Bfade < 0 Then Bfade = 0
    ColorFades(i) = RGB(Rfade, Gfade, Bfade)
Next

'hit the throttle
Me.Timer1.Enabled = True
Exit Sub

errHandler:
Close FileO
MsgBox "Could not load Credits", vbCritical, " Credits Demo"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'End
End Sub

Private Sub Form_Unload(Cancel As Integer)

End Sub

Private Sub picCredits_Click()
'End
End Sub

Private Sub Timer1_Timer()
Dim Ycurr       As Long
Dim TextLine    As Integer
Dim ColPrct     As Long
Dim i           As Integer
'clear pic for next draw
picCredits.Cls
Yscroll = Yscroll - ScrollSpeed
'calculate beginscroll
If Yscroll < (0 - CharHeight) Then
    Yscroll = 0
    LinesOffset = LinesOffset + 1
    If LinesOffset > TotalLines - 1 Then LinesOffset = 0
    'the offset sets the first line of the serie to be printed
    'this offset goes to the next line after each completely
    'scrolled line
    End If
'set Y for first  line
picCredits.CurrentY = Yscroll
Ycurr = Yscroll
'print only the visible lines
For i = 1 To LinesVisible
    If Ycurr > FadeIn And Ycurr < picCredits.Height Then
        'calculate fade-in forecolor
        ColPrct = cDiff2 * (cDiff1 - (Ycurr - FadeIn))
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        picCredits.ForeColor = ColorFades(ColPrct)
    ElseIf Ycurr < FadeOut Then
        'calculate fade-out forecolor
        ColPrct = cDiff3 * Ycurr
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        picCredits.ForeColor = ColorFades(ColPrct)
    Else
        'normal forecolor
        picCredits.ForeColor = ColText
    End If
    'get next line with offset
    TextLine = (i + LinesOffset) Mod TotalLines
    'set the X aligne value
    picCredits.CurrentX = CreditLeft(TextLine)
    'print that line
    picCredits.Print CreditLine(TextLine)
    'set Y to print next line
    Ycurr = Ycurr + CharHeight
Next i
End Sub

Private Sub Label1_Click()
'just for the demo
Unload Me
End
End Sub

