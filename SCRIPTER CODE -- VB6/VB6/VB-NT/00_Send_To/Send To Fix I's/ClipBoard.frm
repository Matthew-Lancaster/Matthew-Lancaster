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
Clipboard.Clear
Clipboard.SetText at2$, vbCFText
End
End Sub

Private Sub Form_Activate()
'rtfRTF 0 (Default) RTF. The file loaded must be a valid .rtf file.
'rtfText 1 Text. The RichTextBox control loads any text file.


'tt$ = "C:\"
'tt$ = GetFolder(Me.hWnd, "Scan Path:", tt$)

'tt$ = "C:\HTTrack\01 Matts Web Httrack.txt"

'RTB1.LoadFile tt$, rtfText


'Exit Sub

at$ = Command$
'at$ = """e:\email2\email2\Wrap Alerts.dbx"""
'at$ = "e:\email2\email2\HostNed Log's & Reports.dbx"

Label1.Caption = at$






If at$ <> "" And Mid$(at$, 1, 1) = """" Then
at$ = Mid$(at$, 2)
at$ = Mid$(at$, 1, Len(at$) - 1)
End If
Label1.Caption = at$

'Exit Sub


'If at$ = "" Then at$ = "e:\email2\email2\Email2 04.rar"

'RTB1.LoadFile at$, rtfText

Open at$ For Binary As #1
at2$ = Input$(LOF(1), 1)
Close #1

Clipboard.Clear
'Clipboard.SetText au$, vbCFText
Clipboard.SetText at2$

'ar$ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'ar2$ = ar$ + LCase$(ar$)
'ar2$ = ar2$ + " "



Exit Sub

Form1.Refresh

rr = Len(at2$)
ProBar1.Max = rr

Label2.Caption = Trim(Str$(rr))
Do
et = InStr(et + 1, at2$, Chr$(26))
If et = 0 Then Exit Do
Mid$(at2$, et, 1) = " "
'asc(Mid$(at2$, et, 1))
DoEvents
ProBar1.Value = et
Label3.Caption = Trim(Str$(et))

Loop Until 1 = 2

Do
et = InStr(et + 1, at2$, Chr$(0))
If et = 0 Then Exit Do
Mid$(at2$, et, 1) = " "
'asc(Mid$(at2$, et, 1))
DoEvents
ProBar1.Value = et
Label3.Caption = Trim(Str$(et))

Loop Until 1 = 2

rr = Len(at2$)
au$ = at2$

GoTo endo

rr = Len(at2$)
Label2.Caption = Trim(Str$(rr))
Do
et = InStr(et + 1, at2$, " ")
If et = 0 Then Exit Do
et2 = InStr(et + 1, at2$, " ")
au$ = au$ + Mid$(at2$, et, et2 - et)
DoEvents
ProBar1.Value = 1000 / ((rr - et) + 1)
Label3.Caption = Trim(Str$(et))

Loop Until 1 = 2

at2$ = au$

rr = Len(at2$)
For r = 1 To rr
DoEvents
If InStr(ar2$, Mid$((at2$), r, 1)) > 0 Then
    au$ = au$ + Mid$((at2$), r, 1)
    x = 1
Else
If x = 1 Then
au$ = au$ + " "
x = 0
End If
End If


DoEvents
ProBar1.Value = 1000 / ((rr - r) + 1)
Label3.Caption = Trim(Str$(r))
'ProBar1..Value = 1000 / ((rr - r) + 1)

Next

e = Len(at2$)
e = Len(au$)

'at2$ = RTB1.Text


endo:



Clipboard.Clear
'Clipboard.SetText au$, vbCFText
Clipboard.SetText au$


End

End Sub

Private Sub Form_Load()
'rtfRTF 0 (Default) RTF. The file loaded must be a valid .rtf file.
'rtfText 1 Text. The RichTextBox control loads any text file.

Exit Sub

at$ = Command$
'at$ = """e:\email2\email2\Wrap Alerts.dbx"""
'at$ = "e:\email2\email2\HostNed Log's & Reports.dbx"

Label1.Caption = at$






If at$ <> "" And Mid$(at$, 1, 1) = """" Then
at$ = Mid$(at$, 2)
at$ = Mid$(at$, 1, Len(at$) - 1)
End If
Label1.Caption = at$

'Exit Sub


'If at$ = "" Then at$ = "e:\email2\email2\Email2 04.rar"

'RTB1.LoadFile at$, rtfText

Open at$ For Binary As #1
at2$ = Input$(LOF(1), 1)
Close #1

Clipboard.Clear
'Clipboard.SetText au$, vbCFText
Clipboard.SetText at2$







Exit Sub

tt$ = "C:\"
'tt$ = GetFolder(Me.hWnd, "Scan Path:", tt$)

tt$ = "C:\HTTrack\01 Matts Web Httrack.txt"

RTB1.LoadFile tt$, rtfText
'Open tt$ For Binary As #1
'Input(number, [#]filenumber)

'yy$ = Input(LOF(1), 1)
'Close #1

Stop
Clipboard.Clear
'Clipboard.SetText au$, vbCFText
Clipboard.SetText RTB1.Text

Exit Sub


RTB1.LoadFile tt$, rtfText



End Sub
