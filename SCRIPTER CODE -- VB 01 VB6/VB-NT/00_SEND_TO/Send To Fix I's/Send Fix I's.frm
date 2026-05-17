VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
      TextRTF         =   $"Send Fix I's.frx":0000
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
Public AT2$, QQ$

Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText AT2$, vbCFText
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





End

End Sub

Private Sub Form_Load()
'rtfRTF 0 (Default) RTF. The file loaded must be a valid .rtf file.
'rtfText 1 Text. The RichTextBox control loads any text file.

at$ = Command$
'at$ = """e:\email2\email2\Wrap Alerts.dbx"""
'at$ = "e:\email2\email2\HostNed Log's & Reports.dbx"

Label1.Caption = at$

If at$ <> "" And Mid$(at$, 1, 1) = """" Then
at$ = Mid$(at$, 2)
at$ = Mid$(at$, 1, Len(at$) - 1)
End If
Label1.Caption = at$


QQ$ = at$

If QQ$ = "" Then
QQ$ = App.Path + "\#1Darling.txt"
ee = 1
End If

Open QQ$ For Binary As #1
    tt$ = Input$(LOF(1), #1)
Close #1


tt2$ = LCase(tt$)

Do
    DoEvents
    vf = InStr(tt$, " i ")
    If vf = 0 Then vf = InStr(tt$, vbLf + "i ")
    If vf = 0 Then vf = InStr(tt$, vbCr + "i ")
    If vf > 0 Then
        Mid$(tt$, vf + 1, 1) = "I"
        tt1 = tt1 + 1
    End If
    
Loop Until vf = 0

vf = 0
Do
    DoEvents
    vf = InStr(vf + 1, tt2$, " wat ")
    If vf = 0 Then vf = InStr(tt2$, vbLf + "wat ")
    If vf = 0 Then vf = InStr(tt2$, vbCr + "wat ")
    If vf > 0 Then
'        rr$ = Mid$(tt$, 1, vf + 8)
'        rr$ = Mid$(rr$, vf)
        tt$ = Mid$(tt$, 1, vf + 1) + "hat " + Mid$(tt$, vf + 5)
        tx2 = tx2 + 1
    End If
    
ty = ty + 1
    
Loop Until vf = 0 Or ty > 100

rf = "Done" + vbCrLf + "Corrected " + Str(tt1) + " I's" + vbCrLf
rf = rf + "Corrected " + Str(tx2) + " Wat's" + vbCrLf

MsgBox rf

If ee = 0 Then
Open QQ$ For Output As #1
    Print #1, tt$
Close #1
End If
















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
AT2$ = Input$(LOF(1), 1)
Close #1

Clipboard.Clear
'Clipboard.SetText au$, vbCFText
Clipboard.SetText AT2$







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
