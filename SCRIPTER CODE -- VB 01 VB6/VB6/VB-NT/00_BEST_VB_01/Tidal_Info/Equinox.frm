VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form Equinox 
   BackColor       =   &H0000FFFF&
   Caption         =   "Equinox - Tidal Info"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   312
   ClientWidth     =   6672
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   6672
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   630
      Left            =   1050
      TabIndex        =   2
      Top             =   600
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1101
      _Version        =   393217
      BackColor       =   65535
      ReadOnly        =   -1  'True
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"Equinox.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Cmdgoto 
      BackColor       =   &H0000FFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   330
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3150
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Equinox"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   0
      Width           =   825
   End
End
Attribute VB_Name = "Equinox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set WSHShell = CreateObject("WScript.Shell")

'TXT EXE
'WSHShell.Run """" + "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\DXGameControllers\dxwebsetup -- 2012.txt" + """"
'Set WSHShell = Nothing

SUN_PER_T = -1

If Month(Now) >= 6 Then
seas$ = "Autumnal "
Else
seas$ = "Vernal "
End If

peas$ = seas$ + "Equinox " + Format$(EquinoxT, "DDDD DD/MMM/YYYY")
If EquinoxT = Int(Now + speedtime) Then
peas$ = seas$ + "Equinox Today..........."
End If
'Label1.Caption = peas$

If Equinox5 < Int(Now + speedtime) Then
peas$ = seas$ + "Equinox Was " + Format$(Equinox5, "DDDD DD/MMM/YYYY")
End If
Label1.Caption = peas$

Label1.Caption = Label1.Caption + vbCr + seas$ + "Equinox SunRise =" + Format$(tooleq1, "HH:MM:SS ampm")
Label1.Caption = Label1.Caption + vbCr + seas$ + "Equinox SunSet =" + Format$(tooleq2, "HH:MM:SS ampm")



Label1.Caption = Label1.Caption + vbCr + "----------------------------"


If Month(Now) >= 5 Then seas$ = "Autumnal " Else seas$ = "Vernal "


peas$ = seas$ + "Equinox " + Format$(Equinox6, "DDDD DD/MMM/YYYY")
If Equinox6 = Int(Now + speedtime) Then
peas$ = seas$ + "Equinox Today..........."
End If
'Label1.Caption = peas$

If Equinox6 < Int(Now + speedtime) Then
peas$ = seas$ + "Equinox Was " + Format$(Equinox6, "DDDD DD/MMM/YYYY")
End If
Label1.Caption = Label1.Caption + vbCr + peas$







Label1.Caption = Label1.Caption + vbCr + seas$ + "Equinox SunRise =" + Format$(tooleq3, "HH:MM:SS ampm")
Label1.Caption = Label1.Caption + vbCr + seas$ + "Equinox SunSet =" + Format$(tooleq4, "HH:MM:SS ampm")


Label1.Caption = Label1.Caption + vbCr + "----------------------------"

Label1.Caption = Label1.Caption + vbCr + "The Autumnal Equinox Is Three Days Earlier than the Date Specified Here"
'Else
Label1.Caption = Label1.Caption + vbCr + "The Vernal Equinox Is Three Days Later than the Date Specified Here"

Label1.Caption = Label1.Caption + vbCr + "The Exact Time Of Which I Cannot Workout..."

Label1.Caption = Label1.Caption + vbCr + "These Dates and Times are when the AM and PM times are almost the Same"

Label1.Caption = Label1.Caption + vbCr + "----------------------------"


Label1.Caption = Label1.Caption + vbCr + "Longitude" + str$(welong) + " Latitude " + str$(welat)

'555ROIDSGOLDRIM

If CompName$ = "555ROIDSGOLDRIM" Then Postcode$ = "Matt's Street BN41 1DQ"
If CompName$ = "MAGICRAM2EODUR" Then Postcode$ = "Matt's Street BN41 1DQ"
If CompName$ = "EDDIESCIENTIST" Then Postcode$ = "Eddie's Street BN41 2RY"
If CompName$ = "NASA1234567890" Then Postcode$ = "Marianne's Street BN2 6QE"
If CompName$ = "MEACHELLE12345" Then Postcode$ = "Meachelle's Street BN2 6QD"
Label1.Caption = Label1.Caption + vbCr + Postcode$
Label1.Caption = Label1.Caption + vbCr + "----------------------------"
Label1.Caption = Label1.Caption + vbCr + "To Find the Exact Time Try This Link"

Label1.Visible = False
'RTB1.Visible = False
RTB1.Height = Label1.Height + 450
RTB1.Width = Label1.Width
RTB1.Left = Label1.Left
RTB1.Top = Label1.Top
RTB1.Text = Label1.Caption

Cmdgoto.Top = Label1.Height + 450
Cmdgoto.Left = 1
Cmdgoto.Width = Label1.Width
Cmdgoto.Caption = "More From This Achived Web http://www.astrologycom.com/solstinox.html"
'Cmdgoto.Caption = "More
Equinox.Height = Label1.Height + 400 + Cmdgoto.Height + 450
Equinox.Width = Label1.Width + 150




End Sub

Private Sub cmdGoTo_Click()
   'UserDocument.Hyperlink.NavigateTo _
   '   "http://www.astrologycom.com/solstinox.html"
'Shell drive2$ + ":\Program Files\Internet Explorer\iexplore.exe http://www.astrologycom.com/solstinox.html", vbNormalFocus
'Shell drive2$ + ":\Program Files\Internet Explorer\iexplore.exe file:///" + App.Path + "/Equinox Web Page/Astrology_on_the_Web_Times_of_Solstices_and_Equinoxes.html", vbNormalFocus
Shell drive2$ + ":\Program Files\Internet Explorer\iexplore.exe " + App.Path + "\Equinox Web Page\Astrology_on_the_Web_Times_of_Solstices_and_Equinoxes.html", vbNormalFocus

End Sub



