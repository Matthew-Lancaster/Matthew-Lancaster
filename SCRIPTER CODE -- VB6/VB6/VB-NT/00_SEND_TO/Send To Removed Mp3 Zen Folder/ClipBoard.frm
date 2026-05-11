VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Com1"
      Height          =   900
      Left            =   4020
      TabIndex        =   3
      Top             =   45
      Width           =   630
   End
   Begin MSComctlLib.ProgressBar ProBar1 
      Height          =   525
      Left            =   150
      TabIndex        =   2
      Top             =   2625
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   926
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   2160
      Left            =   645
      TabIndex        =   0
      Top             =   450
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   3810
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"ClipBoard.frx":0000
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      Height          =   450
      Left            =   3780
      TabIndex        =   5
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   450
      Left            =   3780
      TabIndex        =   4
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   360
      Left            =   555
      TabIndex        =   1
      Top             =   45
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public at2$

Private Sub Command1_Click()
'Clipboard.Clear
'Clipboard.SetText at2$, vbCFText
'End
End Sub




Private Sub Form_Load()
'rtfRTF 0 (Default) RTF. The file loaded must be a valid .rtf file.
'rtfText 1 Text. The RichTextBox control loads any text file.


'tt$ = "C:\"
'tt$ = GetFolder(Me.hWnd, "Scan Path:", tt$)

'tt$ = "C:\HTTrack\01 Matts Web Httrack.txt"

'RTB1.LoadFile tt$, rtfText
'Open tt$ For Binary As #1
'Input(number, [#]filenumber)

'yy$ = Input(LOF(1), 1)
'Close #1

'Stop
'Clipboard.Clear
'Clipboard.SetText au$, vbCFText
'Clipboard.SetText RTB1.Text

at$ = Command$
'at$ = """E:\04 Music ---\03 My Music Zen\01 Banging Tunes\2005\0001 - Dj Snort\Dj Snort-Tube Ride.mp3"""

'at$ = "E:\04 Music ---\03 My Music Zen\01 Banging Tunes\2004\0004 - Morgan\Morgan-Thump.mp3"

'MsgBox at$

If Mid$(at$, Len(at$) - 4, 1) <> "." Then
MsgBox "Got to Be a File With Extension" + vbCrLf + at$
End
End If

'at$ = ""

If at$ <> "" And Mid$(at$, 1, 1) = """" Then
    at$ = Mid$(at$, 2)
    at$ = Mid$(at$, 1, Len(at$) - 1)
End If
'Label1.Caption = at$

clip1$ = "E:\04 Music ---\03 My Music Zen\"
clip2$ = "E:\04 Music ---\02 My Music Zen Removed\"

r4$ = Mid$(at$, Len(clip1$) + 1)
r5$ = clip2$ + r4$

oo = Len(clip2$)
For r = 1 To 10
oo = InStr(oo + 1, r5$, "\")
If oo = 0 Then Exit For
tt$ = Mid$(r5$, 1, oo)
On Error Resume Next
MkDir tt$
On Error GoTo 0
Next


If at$ = "" Then MsgBox "No File ": End

    Dim Fs22
    Set Fs22 = CreateObject("Scripting.FileSystemObject")
    
    errs2 = 0
    On Local Error Resume Next
    Fs22.moveFile at$, r5$
    On Local Error GoTo 0

    If errs2 <> 0 Then
        MsgBox ("Error Moving File to Temp")
        End
    End If



'MsgBox R$

'If Dir$(R$) = "" And R$ <> "" Then
'Open R$ For Output As #1
'Print #1, "------------------------------------------------"
'Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS a/p")
'Print #1, at$
'Print #1, R$
'Print #1, "------------------------------------------------"
'Close #1
'End If

End
Exit Sub


'RTB1.LoadFile tt$, rtfText



End Sub
