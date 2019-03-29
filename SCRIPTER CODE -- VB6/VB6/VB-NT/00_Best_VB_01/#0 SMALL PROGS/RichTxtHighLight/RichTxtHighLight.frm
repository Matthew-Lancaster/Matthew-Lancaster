VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form RTForm 
   Caption         =   "RichTxt High Light Names"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   11010
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RB1 
      Height          =   4410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   7779
      _Version        =   393217
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"RichTxtHighLight.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "RTForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Hh(), UU, Fo1, Qt2, TT, QQ, Test, GRR


Private Sub Form_Load()

RTForm.Show
DoEvents
RB1.FileName = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Bigg Blogs\Staff Today.txt"
RB1.Width = RTForm.Width - 120
RB1.Height = RTForm.Height - 400
ReDim Hh(400)

a = a + 1: Hh(a) = "john"
a = a + 1: Hh(a) = "peter"
a = a + 1: Hh(a) = "steve"
a = a + 1: Hh(a) = "cook"
a = a + 1: Hh(a) = "guy"
a = a + 1: Hh(a) = "frank"
a = a + 1: Hh(a) = "nasia"
a = a + 1: Hh(a) = "carol"
a = a + 1: Hh(a) = "shaz"
a = a + 1: Hh(a) = "sharon"
a = a + 1: Hh(a) = "linda"
a = a + 1: Hh(a) = "bruce"
a = a + 1: Hh(a) = "eman"
a = a + 1: Hh(a) = "liz"
a = a + 1: Hh(a) = "paul pig"
a = a + 1: Hh(a) = "ibrag"
a = a + 1: Hh(a) = "assin"
a = a + 1: Hh(a) = "seve"
a = a + 1: Hh(a) = "baker"
a = a + 1: Hh(a) = "lukat"
a = a + 1: Hh(a) = "jenny"
a = a + 1: Hh(a) = "emma"
a = a + 1: Hh(a) = "jo"
a = a + 1: Hh(a) = "andy"
a = a + 1: Hh(a) = "mum"
a = a + 1: Hh(a) = "dad"

a = a + 1: Hh(a) = "paul"
a = a + 1: Hh(a) = "paul short"
a = a + 1: Hh(a) = "paul tall"
a = a + 1: Hh(a) = "luv"
a = a + 1: Hh(a) = "cath"
a = a + 1: Hh(a) = "martha"
a = a + 1: Hh(a) = "justine"
a = a + 1: Hh(a) = "janice"
a = a + 1: Hh(a) = "abby"
a = a + 1: Hh(a) = "anita"
a = a + 1: Hh(a) = "alice"
a = a + 1: Hh(a) = "alison"
a = a + 1: Hh(a) = "mark"
a = a + 1: Hh(a) = "sarah"
a = a + 1: Hh(a) = "frank"
a = a + 1: Hh(a) = "amilea"
a = a + 1: Hh(a) = "akan"
a = a + 1: Hh(a) = "tanya"
a = a + 1: Hh(a) = "ricci"
a = a + 1: Hh(a) = "rick"
a = a + 1: Hh(a) = "black"
a = a + 1: Hh(a) = "man"
a = a + 1: Hh(a) = "wendy"
a = a + 1: Hh(a) = "asen"
a = a + 1: Hh(a) = "shar"
a = a + 1: Hh(a) = "fara"
a = a + 1: Hh(a) = "patrick"
a = a + 1: Hh(a) = "carter"
a = a + 1: Hh(a) = "fowler"
a = a + 1: Hh(a) = "funnel"
a = a + 1: Hh(a) = "reige"
a = a + 1: Hh(a) = "rage"
a = a + 1: Hh(a) = "tesco"
a = a + 1: Hh(a) = "naterlie"
a = a + 1: Hh(a) = "katie"
a = a + 1: Hh(a) = "#6"
a = a + 1: Hh(a) = "dark"
a = a + 1: Hh(a) = "ring"
a = a + 1: Hh(a) = "louise"
a = a + 1: Hh(a) = "kiss"
a = a + 1: Hh(a) = "max"
a = a + 1: Hh(a) = "dr ojo"
a = a + 1: Hh(a) = "ojo"
a = a + 1: Hh(a) = "ray"
a = a + 1: Hh(a) = "pape"
a = a + 1: Hh(a) = "sophia"
a = a + 1: Hh(a) = "sophie"
a = a + 1: Hh(a) = "mary"
a = a + 1: Hh(a) = "colan"
a = a + 1: Hh(a) = "coughlan"
a = a + 1: Hh(a) = "body guard"
a = a + 1: Hh(a) = "nick"
a = a + 1: Hh(a) = "edd"
a = a + 1: Hh(a) = "sonny"
a = a + 1: Hh(a) = "coretta"
a = a + 1: Hh(a) = "tree" 'treena
a = a + 1: Hh(a) = "rache"
a = a + 1: Hh(a) = "rikk"
a = a + 1: Hh(a) = "william"
a = a + 1: Hh(a) = "charlo"
a = a + 1: Hh(a) = "amanda"
a = a + 1: Hh(a) = "mel"
a = a + 1: Hh(a) = "track" 'tracksuit kid
a = a + 1: Hh(a) = "horn" ' hornsby kid
a = a + 1: Hh(a) = "blond"
a = a + 1: Hh(a) = ""
a = a + 1: Hh(a) = ""
a = a + 1: Hh(a) = ""
a = a + 1: Hh(a) = ""
a = a + 1: Hh(a) = ""
a = a + 1: Hh(a) = ""
a = a + 1: Hh(a) = ""
a = a + 1: Hh(a) = ""
a = a + 1: Hh(a) = ""

For rr = UBound(Hh) To 1 Step -1
    If Hh(rr) <> "" Then Exit For
Next

ReDim Preserve Hh(rr)

'add trim end LF's

RB1.Text = Trim(RB1.Text)
GRR = LCase(RB1.Text)
TT = 1
xx = 0
Do
    vbc = False
    UU = Len(GRR)
    oo = 0
    DoEvents
    For i = 1 To UBound(Hh)
        QQ = InStr(TT, GRR, Hh(i))
        If QQ < UU And QQ > 0 Then
            UU = QQ: oo = 1: vbc = False
        End If
        TestForDate
        If QQ < UU And QQ > 0 Then
            UU = QQ: oo = 1: vbc = True
        End If
    Next
    If oo = 0 Then Exit Do
    
    Qt2 = 0
    Test = "-- Wed 19-Nov-08 18:41:26 --"
    
    
    TT = UU + 1
    Qt = InStr(TT, GRR, " ")
    If vbc = True Then Qt = TT + Len(Test)
    If InStr(TT, GRR, vbCrLf) < Qt Then Qt = InStr(TT, GRR, vbCrLf)
    If InStr(TT, GRR, vbLf) < Qt Then Qt = InStr(TT, GRR, vbLf)
    
    If TT > 1 Then
        jj = LCase(Mid$(RB1.Text, TT - 1, Qt - TT + 1))
    Else
        jj = LCase(Mid$(RB1.Text, TT, Qt - TT + 1))
    End If
    'Mid$(RB1.Text, TT, 1) = UCase(Mid$(RB1.Text, TT, 1))
    
    Mid$(jj, 1, 1) = UCase(Mid$(jj, 1, 1))
    
    RB1.SelStart = TT - 2
    RB1.SelLength = Qt - TT + 1
    RB1.SelText = jj
    
    RB1.SelStart = TT - 2
    RB1.SelLength = Qt - TT + 1
    RB1.SelBold = True
    RB1.SelFontSize = 20
    
    
    If oldjj = UCase(jj) Then
        RB1.SelColor = vbRed
    Else
        RB1.SelColor = vbBlue
    End If
    
    If vbc = True Then RB1.SelColor = QBColor(5)

    oldjj = UCase(jj)

Loop Until QQ = 0 ' Or TT > 50000



End Sub

Sub TestForDate()

    Test = "-- Wed 19-Nov-08 18:41:26 --"
    Fo1 = InStr(TT, GRR, "-- ")
    fo = InStr(Fo1 + 2, GRR, " ")
    fo = InStr(fo + 2, GRR, "-")
    fo = InStr(fo + 1, GRR, "-")
    fo = InStr(fo + 1, GRR, " ")
    If Fo1 + 16 = fo Then QQ = Fo1
    'jj = Mid$(RB1.Text, fo, 20)

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
