VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Double Words"
   ClientHeight    =   5205
   ClientLeft      =   945
   ClientTop       =   2400
   ClientWidth     =   9675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   9675
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4305
      Top             =   1575
   End
   Begin MSComctlLib.ProgressBar ProBar1 
      Height          =   300
      Left            =   6630
      TabIndex        =   11
      Top             =   3795
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "Form1.frx":08CA
      Left            =   4815
      List            =   "Form1.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   4755
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear   List   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3870
      Picture         =   "Form1.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3780
      Width           =   675
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2310
      Picture         =   "Form1.frx":1710
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3780
      Width           =   645
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Philo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1650
      Picture         =   "Form1.frx":1B52
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3780
      Width           =   645
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3825
      Top             =   1575
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   3375
      Top             =   1575
   End
   Begin VB.ListBox Sort_List1 
      Height          =   255
      Left            =   2565
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   4935
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clip Selected"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2955
      Picture         =   "Form1.frx":241C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3780
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit Workin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   765
      Picture         =   "Form1.frx":325E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3780
      Width           =   870
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2925
      Top             =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "And Again"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   -15
      Picture         =   "Form1.frx":40A0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3765
      Width           =   765
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "Form1.frx":4EE2
      Left            =   -15
      List            =   "Form1.frx":4EE4
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4755
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   "Pass 1 of 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4605
      TabIndex        =   16
      Top             =   4380
      Width           =   1710
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Size Char"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7785
      TabIndex        =   15
      Top             =   4695
      Width           =   1770
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Char"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      TabIndex        =   14
      Top             =   4395
      Width           =   1755
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   "Doubles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4605
      TabIndex        =   13
      Top             =   4095
      Width           =   3120
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Pass 1 of 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   12
      Top             =   4080
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4605
      TabIndex        =   5
      Top             =   3780
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Working"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5370
      TabIndex        =   3
      Top             =   3780
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu mnu_Double 
         Caption         =   "Open Double.txt"
      End
      Begin VB.Menu mnu_TestDups 
         Caption         =   "Test Dupes"
      End
      Begin VB.Menu Mnu_Reset 
         Caption         =   "Reset For Next Click"
      End
      Begin VB.Menu Mnu_Fake 
         Caption         =   "Fake Run Update"
      End
      Begin VB.Menu Mnu_CopyAll 
         Caption         =   "Copy All"
      End
      Begin VB.Menu Mnu_Black 
         Caption         =   "Notepad Black Listed"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TF$, Timer4Flag, B$, TT, XC, ODD1$, TipCount, GG$, CH$

Dim dfwd$(1800)
Dim Start As Integer
Dim JKd As Integer
Dim KG As Long
Public TotalList$, CountDouble

Dim kjh As Integer
Dim qw As Integer
Dim lm As Integer
Dim xx As Integer
Dim plonk As Integer
Dim indy As Integer
Dim sort As Integer
Dim ads As Integer
Dim tr1$
Dim tr2$
Dim TR3$
Dim TR4$
Dim t1$
Dim t2$
Dim t3$
Dim t4$

Public HOllow1$
Public Hollow2$
Public Clip1$
Public ExitNow

Public WxHex5$

Public BackOff
Public BackEcho

Public Hack

Public Black$


Private Sub Form_Load()

On Local Error Resume Next
re = FreeFile
Open App.Path + "\BlackList.txt" For Input As #re
Black$ = Input(LOF(re), re)
Close #re
On Local Error GoTo 0


Hack = True

ReDim SelectedEighty(10000)

'Call Tube

Exit Sub


End Sub

Private Sub Command2_Click()

ExitNow = True

End Sub

Private Sub Command3_Click()

'SortFrm.Command1.Visible = True

pn$ = ""

For r = List2.ListCount - 1 To 0 Step -1
    pn$ = List2.List(r) + vbCrLf + pn$
    List2.RemoveItem (r)
    CountDouble = CountDouble + 1
Next

Label4 = "Double Words =" + Str(CountDouble)

'Clipboard.Clear
'Clipboard.SetText pn$


f3 = FreeFile
Open "c:\callerid\2double.txt" For Append As #f3
Print #f3, pn$;
Close #f3

pn$ = ""

'SortFrm.Show

Call Main_Sort_1st

'If SortFrmPop = 0 Then
'SortFrm.Hide
'End If

'Form1.WindowState = 1

End Sub
Sub Tube()

Dim word$()

ReDim word$(90000)

Dim PutNut1$
Dim PutNut2$

f3 = FreeFile
Open "c:\callerid\2double.txt" For Input As #f3
PutNut1$ = Input$(LOF(f3), f3)
Close #f3
PutNut2 = ""

tg = InStr(PutNut1$, "DOUBLE WORD SAME FIRST LETTER")
X1 = tg
CountDouble = 0
Do
tg = InStr(tg + 1, PutNut1$, vbCrLf)
tg2 = InStr(tg + 1, PutNut1$, vbCrLf)

'get rid if in black list
'If tg2 > 0 And tg > 0 Then zz$ = Mid$(PutNut1$, tg + 2, tg2 - tg - 2)
'th = InStr(Black$, vbCrLf + zz$ + vbCrLf)
'If th > 0 Then Stop

'Run this to sort any dupes
GoTo jump4
If tg2 > 0 And tg > 0 Then zz$ = Mid$(PutNut1$, tg + 2, tg2 - tg - 2)
vv = 0
Do
DoEvents
If tg = 0 Or tg2 = 0 Then Exit Do
th = InStr(tg + 1, PutNut1$, vbCrLf + zz$ + vbCrLf)
If th > 0 Then
'Stop 'right$(Mid$(PutNut1$, 1, th -1),7)
vv = vv + 1
If vv > 1 Then ddflagdupe = ddflagdupe + 1
tg3 = InStr(th + 1, PutNut1$, vbCrLf)
PutNut1$ = Mid$(PutNut1$, 1, th - 1) + vbCrLf + Mid$(PutNut1$, tg3 + 2)
flagdupe = flagdupe + 1
End If
Loop Until th = 0

jump4:

CountDouble = CountDouble + 1
Loop Until tg = 0
CountDouble = CountDouble - 2
Label4 = "Double Words =" + Str(CountDouble)

'Do the black list redo as well

If flagdupe > 0 Then
    MsgBox Str(flagdupe) + " Dupes Found" + vbCrLf + Str(ddflagdupe) + " Double Dupes Found"
    f3 = FreeFile
    Open "c:\callerid\2double.txt" For Output As #f3
    Print #f3, PutNut1$;
    Close #f3
End If




alp1$ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
alp1$ = alp1$ + LCase$(alp1$)


a$ = " " + a$ + " "
On Local Error Resume Next
Do
Err.Clear
a$ = " " + Clipboard.GetText() + " "
DoEvents
Loop Until Err.Number = 0
On Local Error GoTo 0



TF$ = "Dummy Dont Run"


If TF$ <> "" Then Exit Sub

If TF$ = "" And 1 = 1 Then
    TF$ = "D:\# MY DOCS\# 01 My Documents\00 Email Texts\All in One -- Ioannis.txt" ' go back to that one
    TF$ = "E:\01 Matts Web Httrack\BBC Double Words\BBC Double Words.rar"
    TF$ = "c:\callerid\2ABBREV.TXT"
    TF$ = "D:\# MY DOCS\# 01 My Documents\00 Email Texts\00_Scan_Merge_MS Dos 6 TXT DIZ NFO LOG.txt"
    TF$ = "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Pirate\00_Scan_Merge_MS Dos 6 TXT DIZ NFO LOG.txt"
    
    Open TF$ For Binary As #1
    a$ = " " + Input$(LOF(1), 1) + " "
    Close #1
End If

If TF$ = "" And 1 = 2 Then
    a$ = ""
    TF$ = "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Terrorist Files\"
    TF$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\"
    TF$ = "D:\VB6\VB-NT\00_Best_VB_01\Date1991\00 Data\"
    TF$ = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\# DATA\"
    TF$ = "D:\# MY DOCS\# 01 My Documents\00 Email Texts\My Web Sites"
    TF$ = "D:\MY WEBS\00 Documents 3"
    TF$ = "C:\CallerID" 'later
    TF$ = "D:\# MY DOCS\# 01 My Documents\00 Email Texts" 'later
    TF$ = "D:\MY WEBS\01 Swaps" 'later
    
    ScanPath.chkSubFolders.Value = vbChecked
    ScanPath.txtPath.Text = TF$
    ScanPath.cboMask.Text = "*.txt"
    Call ScanPath.cmdScan_Click
    
    For r = 1 To ScanPath.ListView1.ListItems.Count
        DoEvents
        a1$ = ScanPath.ListView1.ListItems.Item(r).SubItems(1)
        b1$ = ScanPath.ListView1.ListItems.Item(r)
        Label7 = Str(r) + " of " + Str(ScanPath.ListView1.ListItems.Count)
    
        Open a1$ + b1$ For Binary As #1
        
        GoTo jump
        '<----------- Email Texts
        If LOF(1) >= 2000000 And LOF(1) < 15000000 Then tip = 1 Else tip = 0
        If InStr(b1$, "WN-") Or InStr(UCase$(b1$), "2DIARY") = 1 Then
            tip = 0
        End If
        'If r = 1 Then tip = 1 Else tip = 0
        If tip = 1 Then a$ = a$ + " " + Input$(LOF(1), 1) + " "
        '----->
jump:
        
        a$ = a$ + " " + Input$(LOF(1), 1) + " "
        
        Close #1
    Next

'a$ = Mid$(a$, 1, Len(a$) / 10)
End If

y = 0
F = 0
E = 1
pit = 0
tg = 1
B$ = " " + Trim(UCase$(a$)) + " "

ProBar1.Min = 0
TT = Len(B$)
ProBar1.Max = TT
Label6 = Str(TT)

pass$ = " of 7"
Label3.Caption = "Pass 1" + pass$

Do
DoEvents
XC = InStr(XC + 1, B$, vbCrLf)
If XC > 0 Then Mid$(B$, XC, 2) = "  ": ProBar1.Value = XC
Loop Until XC = 0


Label3.Caption = "Pass 2" + pass$

Do
DoEvents
XC = InStr(XC + 1, B$, Chr(0))
If XC > 0 Then
Mid$(B$, XC, 1) = " "
ProBar1.Value = XC
End If
Loop Until XC = 0

Label3.Caption = "Pass 3" + pass$

Do
DoEvents
XC = InStr(XC + 1, B$, Chr(9))
If XC > 0 Then
Mid$(B$, XC, 1) = " "
ProBar1.Value = XC
End If
Loop Until XC = 0

Label3.Caption = "Pass 4" + pass$

Do
DoEvents
XC = InStr(XC + 1, B$, vbCr)
If XC > 0 Then
Mid$(B$, XC, 1) = " "
ProBar1.Value = XC
End If
Loop Until XC = 0

Label3.Caption = "Pass 5" + pass$

Do
DoEvents
XC = InStr(XC + 1, B$, vbLf)
If XC > 0 Then Mid$(B$, XC, 1) = " ": ProBar1.Value = XC
Loop Until XC = 0

B$ = " " + Trim((B$)) + " "

Label3.Caption = "Pass 6" + pass$

ODD1$ = ""
TT = Len(B$)
ProBar1.Max = TT
XC = 1

slow = False

If slow = True Then
    Timer4.Enabled = True
End If

If slow = False Then
    Do
        DoEvents
        Call Timer4_Timer
    Loop Until Timer4Flag = True
End If
'--------------------------------------------


End Sub


Sub Tube2()

'Exit Sub

TT = Len(B$)

Label3.Caption = "Pass 7" + pass$

XC = 0
ProBar1.Max = Len(B$)
Label6 = Str(TT)
Do
DoEvents
XC = InStr(XC + 1, B$, " ")
xc2 = InStr(XC + 1, B$, " ")
xc3 = InStr(xc2 + 1, B$, " ")
xc4 = InStr(xc3 + 1, B$, " ")
xc5 = InStr(xc4 + 1, B$, " ")
If XC = 0 Or xc2 = 0 Then Exit Do

ProBar1.Value = XC

a1$ = Mid$(B$, XC + 1, 1)
b1$ = Mid$(B$, xc2 + 1, 1)
If xc3 > 0 Then c1$ = Mid$(B$, xc3 + 1, 1)
If xc4 > 0 Then d1$ = Mid$(B$, xc4 + 1, 1)
tip = 0
dd1$ = ""
dd2$ = ""
dd3$ = ""
dd4$ = ""
If a1$ = b1$ Then
    dd1$ = Mid$(B$, XC, xc2 - XC)
    dd2$ = Mid$(B$, xc2, xc3 - xc2)
    tip = 1
End If
If b1$ = c1$ And tip = 1 Then
    dd3$ = Mid$(B$, xc3, xc4 - xc3)
    tip = 2
End If
If c1$ = d1$ And tip = 2 Then
    dd4$ = Mid$(B$, xc4, xc5 - xc4)
    tip = 3
End If
    
    
If tip > 0 Then
dd1$ = Trim(dd1$)
dd2$ = Trim(dd2$)
dd3$ = Trim(dd3$)
dd4$ = Trim(dd4$)

If Len(dd1$) > 1 Then Mid$(dd1$, 2) = LCase(Mid$(dd1$, 2))
If Len(dd2$) > 1 Then Mid$(dd2$, 2) = LCase(Mid$(dd2$, 2))
If Len(dd3$) > 1 Then Mid$(dd3$, 2) = LCase(Mid$(dd3$, 2))
If Len(dd4$) > 1 Then Mid$(dd4$, 2) = LCase(Mid$(dd4$, 2))

'ff = Asc(Mid$(dd1$, 1, 1))
'If ff <= 64 Or ff >= 123 Then tip = 0
'If Len(dd1$) = 1 And Len(dd2$) = 1 Then tip = 0

End If
    
    
    
    
If tip > 0 Then

For t = 1 To 4
    DoEvents
    If t = 1 Then thgx$ = dd1$
    If t = 2 Then thgx$ = dd2$
    If t = 3 Then thgx$ = dd3$
    If t = 4 Then thgx$ = dd4$
    
    If thgx$ <> "" Then
        
        If t > 1 Then
            rh = InStr(thgx$, ".")
            If rh > 0 Then
                If t = 2 Then dd3$ = "": dd4$ = ""
                If t = 3 Then dd4$ = ""
            End If
        End If
        
        
        For r = 1 To Len(thgx$)
            rh = InStr("!!""#$%&'(),-./*+:;<=>?[\]^_`{|}~,—”’³¿å¹ž€êœæö‰", Mid$(thgx$, r, 1))
            If rh > 0 Then
                thgx$ = Mid$(thgx$, 1, r - 1) ' + Mid$(thgx$, r)
                Exit For
            End If
        Next
        If t = 1 Then dd1$ = thgx$
        If t = 2 Then dd2$ = thgx$
        If t = 3 Then dd3$ = thgx$
        If t = 4 Then dd4$ = thgx$
    End If
Next

thg$ = dd1$ + " " + dd2$
If dd3$ <> "" Then thg$ = thg$ + " " + dd3$
If dd4$ <> "" Then thg$ = thg$ + " " + dd4$

If dd1$ = dd3$ Then tip = 0
If dd1$ = dd2$ Then tip = 0
If dd2$ = dd3$ Then tip = 0
If dd2$ = dd4$ Then tip = 0
If dd1$ = dd4$ Then tip = 0
If dd3$ <> "" And dd3$ = dd4$ Then tip = 0
If Right$(dd1$, 1) = "s" And InStr(dd1$, dd2$) Then tip = 0
If Right$(dd2$, 1) = "s" And InStr(dd2$, dd1$) Then tip = 0
If Len(dd1$) = 1 And Len(dd2$) = 1 Then tip = 0
If tip > 0 Then
    If InStr(PutNut1$, thg$) > 0 Then tip = 0
    If InStr(Black$, thg$) > 0 Then tip = 0
        DoEvents
        If tip > 0 Then
            Form1.List1.AddItem thg$
            TotaList$ = TotalList$ + thg$
            Label2.Caption = Str(List1.ListCount)
        End If
    End If
End If

Loop Until XC = 0


ExitNow = False

PutNut1$ = ""
Black$ = ""
B$ = ""

Label1.Visible = False
    
For r = List1.ListCount - 2 To 0 Step -1
    If List1.List(r) = List1.List(r + 1) Then List1.RemoveItem (r)
Next
    
Call Mnu_CopyAll_Click
    
Label2.Caption = Str(List1.ListCount)
    
List2.SetFocus


End Sub


Private Sub Command4_Click()

Key = 0
fg5 = FreeFile
Open "D:\TEMP\Notepa~1.txt" For Input As #fg5
Do
    Line Input #fg5, ew$
    Key = Key + 1
Loop Until InStr(UCase$(ew$), UCase$("From: ""BORG")) = 1 Or EOF(fg5)
If EOF(fg5) Then Key = 0
Close #fg5

If Key = 0 Then
    fg5 = FreeFile
    Open "D:\TEMP\Notepa~1.txt" For Input As #fg5
    Do
        Line Input #fg5, ew$
        Key = Key + 1
    Loop Until InStr(UCase$(ew$), UCase$("BORG")) > 0 Or EOF(fg5)
    If EOF(fg5) Then Key = 0
    Close #fg5
End If

If Key = 0 Then
    fg5 = FreeFile
    Open "D:\TEMP\Notepa~1.txt" For Input As #fg5
    Do
        Line Input #fg5, ew$
        Key = Key + 1
    Loop Until InStr(UCase$(ew$), UCase$("FROM: ""KAN")) > 0 Or EOF(fg5)
    If EOF(fg5) Then Key = 0
    Close #fg5
End If

If Key = 0 Then
    fg5 = FreeFile
    Open "D:\TEMP\Notepa~1.txt" For Input As #fg5
    Do
        Line Input #fg5, ew$
        Key = Key + 1
    Loop Until InStr(UCase$(ew$), UCase$("KAN")) > 0 Or EOF(fg5)
    If EOF(fg5) Then Key = 0
    Close #fg5
End If



Key = Key - 8
    'Shell "c:\bat\t.bat /#:" + Trim(Str(Key)) + " D:\TEMP\Notepa~1.txt", vbNormalFocus
    Shell "c:\windows\t.bat D:\TEMP\Notepa~1.txt /#:" + Trim(Str(Key)), vbNormalFocus

End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
List1.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Form1.WindowState = 1 Then
End
Else
Form1.WindowState = 1
Cancel = True
End If
End Sub

Private Sub Label2_Click()
'Label2
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If BackOff = True Or Hack = False Then Exit Sub

Hack = False

If Button = 1 Then
    t = List1.ListIndex
    SelectedEighty(t) = True

If InStr(TotalList$, List1.List(t)) = 0 Then
    If InStr(Black$, List1.List(t)) = 0 Then
        List2.AddItem List1.List(t)
    End If
End If
'    BackOff = True
'    For r = 0 To List1.ListCount - 1
'        If SelectedEighty(r) = True Then
'            List1.Selected(r) = True
'        Else
'            List1.Selected(r) = False
'        End If
'    Next
    
    'List1.ListIndex = List1.ListCount - 1
        
    'If List1.ListCount > t + 13 Then
    '    List1.ListIndex = t - 1
    '    List1.ListIndex = 1
        'List1.ListIndex = t
    'End If
    BackOff = False
    
End If

If Button = 2 Then
    t = List1.ListIndex
    SelectedEighty(t) = False

    BackOff = True
    For r = 0 To List1.ListCount - 1
        If SelectedEighty(r) = True Then
            List1.Selected(r) = True
        Else
            List1.Selected(r) = False
        End If
    Next
    'List1.ListIndex = List1.ListCount - 1
    'If List1.ListCount > t + 13 Then
    '    List1.ListIndex = t - 1
    '    List1.ListIndex = t
    'End If
    BackOff = False
End If

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Hack = True
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    t = List2.ListIndex
    'SelectedEighty(t) = True

Black$ = Black$ + List2.List(t) + vbCrLf

re = FreeFile
Open App.Path + "\BlackList.txt" For Append As #re
Print #re, List2.List(t)
Close #re

'List1.AddItem List2.List(t)
List2.RemoveItem t
Label2.Caption = Str(List2.ListCount)

End If
End Sub

Private Sub Mnu_BringBack_Click()
    
For r = 0 To List1.ListCount - 1
    If SelectedEighty(r) = True Then
        List1.Selected(r) = True
    End If
Next

End Sub

Private Sub Mnu_Black_Click()
Shell "notepad " + App.Path + "\BlackList.txt", vbNormalFocus
End Sub

Private Sub Mnu_CopyAll_Click()
List2.Clear

For r = 0 To List1.ListCount - 1

List2.AddItem List1.List(r)

Next

'If List2.ListCount > 0 Then Form1.WindowState = 1

End Sub

Private Sub mnu_Double_Click()
Shell "notepad c:\callerid\2double.txt", vbNormalFocus
End Sub

Private Sub Mnu_Fake_Click()
Fake = True
Call Main_Sort_1st
Fake = False
End Sub

Private Sub Mnu_Reset_Click()
ReDim SelectedEighty(10000)
End Sub

Private Sub Timer1_Timer()
Call Tube
Timer1.Enabled = False
Exit Sub

GoTo skippy
Set m_CRC = New clsCRC

m_CRC.Algorithm = CRC32
    
wxhex4$ = Hex(m_CRC.CalculateFile("C:\callerid\2DoubleTrip.txt"))

If wxhex4$ <> WxHex5$ And WxHex5$ <> "" Then
Fake = True
Call Main_Sort_1st
Fake = False
End If
WxHex5$ = wxhex4$

skippy:

'hash1 = FreeFile
'Open "E:\01 Matts Web Httrack\BBC Double Words\BBC Double Words.rar" For Binary As #hash1
'ttg$ = Input$(10000, hash1)
'Close hash1

On Local Error Resume Next
Do
Err.Clear
ttg$ = Clipboard.GetText()
DoEvents
'err.description
Loop Until Err.Number = 0

If ttg$ <> Clip1$ And List1.ListCount = 0 Then
    Label1.Visible = True
    Clip1$ = ttg$
    
    
    If InStr(Clip1$, "alt.philosophy") > 0 And _
    InStr(Clip1$, "http://groups.google.com/group/alt.philosophy") > 0 And _
    InStr(Clip1$, "alt.philosophy@googlegroups.com") > 0 And _
    InStr(Clip1$, "Today's topics:") > 0 And _
    Len(Clip1$) < 1500000 Then
        freet8 = FreeFile
        Open "D:\TEMP\NotepadPhilo.txt" For Output As #freet8
        Print #freet8, Clip1$
        Close #freet8
        'Shell "notepad D:\TEMP\NotepadPhilo.txt", vbNormalFocus
        'Shell "c:\windows\t.com D:\TEMP\Notepa~1.txt", vbNormalFocus
        Call Command4_Click
    End If


    If InStr(Clip1$, "alt.support.schizophrenia") > 0 And _
    InStr(Clip1$, "http://groups.google.com/group/alt.support.schizophrenia") > 0 And _
    InStr(Clip1$, "alt.support.schizophrenia@googlegroups.com") > 0 And _
    InStr(Clip1$, "Today's topics:") > 0 And _
    Len(Clip1$) < 1500000 Then
        freet8 = FreeFile
        Open "D:\TEMP\NotepadPhilo.txt" For Output As #freet8
        Print #freet8, Clip1$
        Close #freet8
        'Shell "notepad D:\TEMP\NotepadPhilo.txt", vbNormalFocus
        'Shell "c:\windows\t.com D:\TEMP\Notepa~1.txt", vbNormalFocus
        Call Command4_Click
    End If
    
    
    
    JKd = 0
    Start = 1
    'List1.Clear
    ReDim SelectedEighty(10000)
    'For r = 1 To 26
    '    Hollow2$ = Chr$(r + 64) + Chr$(r + 64) + "--"
    '    Call Tube
    '    Hollow2$ = Chr$(r + 64) + Chr$(r + 64) + Chr$(r + 64) + Chr$(r + 64)
    '    Call Tube
    '    DoEvents
    
    'If ExitNow = True Then Exit For
    'If List1.ListCount > 3 Then Form1.WindowState = 0
    'Next
    List1.Clear
    List2.Clear
    Call Tube
    Label1.Visible = False
    
    Label1.Visible = False
    For r = List1.ListCount - 2 To 0 Step -1
        If List1.List(r) = List1.List(r + 1) Then List1.RemoveItem (r)
    Next
    
    Label2.Caption = Str(List1.ListCount)
    
    Call Mnu_CopyAll_Click
    
    Label2.Caption = Str(List2.ListCount)
    
    If List2.ListCount = 0 Then Timer3.Enabled = True

End If

'Timer1.Enabled = False

Ende:
Label1.Visible = False

End Sub

Private Sub Timer2_Timer()
Form1.WindowState = 1
Timer2.Enabled = False
End Sub
Private Sub Timer3_Timer()
Form1.WindowState = 1
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()

    XC = InStr(XC + 1, B$, " ")
    If XC = 0 Or XC = TT Then XC = 0: GoTo Ende


    oxc5 = XC
    xc5 = XC
    Do
        xc5 = InStrRev(B$, " ", xc5 - 1)
        If xc5 + 1 < oxc5 Then Exit Do
        If xc5 = 0 Then Exit Do
        oxc5 = xc5
    Loop Until 1 = 1

    If xc5 > 0 Then
        dd1$ = Mid$(B$, xc5 + 1, 1)
        ff = Asc(dd1$)
        'all caps
        If ff <= 64 Or ff >= 91 Then tip = 0 Else tip = 1
        If tip = 1 Then
            
            If ODD1$ <> "" And ODD1$ = dd1$ Then tip = 0 Else tip = 1
            If tip = 0 Then TipCount = TipCount + 1 Else TipCount = 0
            
            If tip = 0 Then
                GG$ = " :"
                If vv1$ <> dd1$ Then
                   GG$ = ""
                End If
                
                vv1$ = dd1$
                If TipCount = 1 Then
                    CH$ = CH$ + GG$ + ohh$ + Mid$(B$, xc5, XC - xc5)
                Else
                    CH$ = CH$ + Mid$(B$, xc5, XC - xc5)
                End If
            End If
            Label6 = Str(Len(CH$))
        End If
        
        ODD1$ = dd1$: ohh$ = Mid$(B$, xc5, XC - xc5)
    End If

ProBar1.Max = TT
ProBar1.Value = XC
Label5 = Str(TT - XC)

'Loop Until xc = 0
Ende:

If XC = 0 Then
    Timer4.Enabled = False
    B$ = " " + Trim(CH$) + " "
    CH$ = ""
    Timer4Flag = True
    Call Tube2
End If

End Sub
