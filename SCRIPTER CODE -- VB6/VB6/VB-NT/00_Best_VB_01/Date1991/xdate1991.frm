VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Make A Date"
   ClientHeight    =   8340
   ClientLeft      =   612
   ClientTop       =   720
   ClientWidth     =   13764
   Icon            =   "xdate1991.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   13764
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer_CLIPBOARD 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1356
      Top             =   648
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   1065
      Left            =   450
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   12945
      _ExtentX        =   22839
      _ExtentY        =   1884
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1725
      Left            =   2790
      TabIndex        =   8
      Top             =   3435
      Visible         =   0   'False
      Width           =   9210
      _ExtentX        =   16235
      _ExtentY        =   3048
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ListBox List6 
      BackColor       =   &H00400000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1224
      ItemData        =   "xdate1991.frx":0E42
      Left            =   5745
      List            =   "xdate1991.frx":0E44
      TabIndex        =   10
      Top             =   3225
      Visible         =   0   'False
      Width           =   2955
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1275
      Left            =   2505
      TabIndex        =   12
      Top             =   990
      Visible         =   0   'False
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   2244
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   16711680
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ListBox List7 
      Enabled         =   0   'False
      Height          =   1584
      Left            =   8880
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   2010
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ListBox List5 
      Enabled         =   0   'False
      Height          =   1584
      Left            =   4605
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1230
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ListBox List4 
      Enabled         =   0   'False
      Height          =   816
      Left            =   375
      TabIndex        =   7
      Top             =   3150
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      Height          =   1584
      Left            =   375
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1215
      Visible         =   0   'False
      Width           =   4095
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   945
      Left            =   3255
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1291
      _ExtentY        =   1672
      _Version        =   393217
      TextRTF         =   $"xdate1991.frx":0E46
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1935
      Top             =   735
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   -15
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   13740
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00400000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      ItemData        =   "xdate1991.frx":0ECB
      Left            =   975
      List            =   "xdate1991.frx":0ECD
      TabIndex        =   0
      Top             =   4260
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   15
      TabIndex        =   13
      Top             =   4845
      Width           =   13560
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Season"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   7620
      Visible         =   0   'False
      Width           =   13755
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Star Sign"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   13755
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Click An Item in List To view Description Here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15
      TabIndex        =   2
      Top             =   8025
      Visible         =   0   'False
      Width           =   13740
   End
   Begin VB.Menu MenuExit 
      Caption         =   "EXIT"
   End
   Begin VB.Menu Mnu_Redo 
      Caption         =   "RE-DO"
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "EDITOR"
      Begin VB.Menu Menu_EditListBOTH 
         Caption         =   "Edit Both Lists"
      End
      Begin VB.Menu Menu_EditList 
         Caption         =   "Edit List"
      End
      Begin VB.Menu Menu_EditList_S 
         Caption         =   "Edit List Secure"
      End
      Begin VB.Menu Mnu_Logg_Folder 
         Caption         =   "Logg Folder"
      End
      Begin VB.Menu Menu_ResetList 
         Caption         =   "Reset List"
      End
      Begin VB.Menu Mnu_ClearClip 
         Caption         =   "Clear ClipBoard"
      End
   End
   Begin VB.Menu Menu_View 
      Caption         =   "VIEW HTML"
      Begin VB.Menu Menu_HtmlLarge 
         Caption         =   "Date1991 Large Html"
      End
      Begin VB.Menu Menu_HtmlWeb 
         Caption         =   "Date1991 Web Html"
      End
      Begin VB.Menu Menu_HtmlSmall 
         Caption         =   "Date1991 Small Html"
      End
      Begin VB.Menu Menu_htmlSmallEmail 
         Caption         =   "Date1991 Small Email"
      End
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "HELP"
      Begin VB.Menu Menu_About 
         Caption         =   "About"
      End
   End
   Begin VB.Menu MNU_CLEAR_CLIP_SET 
      Caption         =   "CLEAR CLIP SET"
   End
   Begin VB.Menu Mnu_Clip_Them 
      Caption         =   "CLIP SELECT SET"
   End
   Begin VB.Menu Mnu_DelPrvi 
      Caption         =   "DEL PREVIOUS YEAR SET"
   End
   Begin VB.Menu Mnu_Select 
      Caption         =   "SELECT ALL FAV SET"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JOBS
'Job clocks go back sunday before end oct

'This Needs flicker free list box planet source code

Dim MInfo2$(10)
Dim OL As Date
'SEPT. EQUINOX:  Today, Sept. 22nd at 2118 UT (5:18 pm EDT), the sun crosses the celestial equator.
'This event marks the beginning of autumn in the northern hemisphere and spring in the southern hemisphere.
'It's also the beginning of aurora season around the poles. Happy equinox!


Public DateClip$, KK$, ListCollect, DoAgain, Finished, RunOnce, Redo
'Limit Lenght On Trunccate Sig Of Date Year Dates
'Search This
Public NowNextData
Dim DateChk
Dim FullMoonBeforeEaster
Dim Rvm1()
Dim Rvm2()

Public R

Dim Mls$()
Dim Tagg5()
Dim Lastest()
Dim LastestInx()
Dim FormatDate$
Dim TaggStr
Dim TaggCnt
Dim String1
Dim String2

Dim SeasonsDate(): Dim Hatt
Dim Tagger()
Dim Tagge2()
Dim XvRs1()
Dim XvRs2()

Dim j As Double
Dim enow As Date
Dim F As Double
Dim pj As Double
Dim ml() As Integer
Public rs5$
Public rs6 As Integer
Public rs8$
Public rs9 As Integer
Public qz As Integer
Public qj As Double
Public xcd As Date
Public xcd2 As Double
Public xcd3 As Date

Public tool1 As Date
Public tool2 As Date
Public tool3 As Date
Public tool5 As Double
Public tool6 As Double
Public toolnow As Date

Dim yr As String * 6
Dim wek As String * 6
Dim wek2 As String * 8

Dim lpg As Double

Const Synod = 29.53058867

Const BaseNewMoonDateString As String = "18/11/1998 9:36:00 pm"

Public xcddate As Date

Public TheDate As Date

Public thedate2 As Date
Public wclops As Date

Public XFlip As Date
Public ftag As Integer
Public nextlastmoon$
Public Diar As Date

Public we$

Public p As Long
Dim wclops2 As Date

Public Web

Public ll1$
Public ll2$

Public star2$


Public WebStop



Private Sub Form_GotFocus()
Call Mnu_Select_Click
End Sub

Private Sub Form_Load()


Dim FileSpec, TT As Long
FileSpec = App.Path + "\" + App.EXEName + ".vbp"

If IsIDE = False And Dir$(FileSpec) <> "" Then
    'TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec + """", vbMinimizedNoFocus)
    'End
End If

WebStop = True

GoTo SkipHere
' SOME CONVERT WORK FOR TEACHER DATE
' ----------------------------------
Open "D:\VB6\VB-NT\00_Best_VB_01\Date1991\00 Data\Teachers Dates.txt" For Input As #1
Do
    Line Input #1, lx
    milk = 0
    If Mid$(lx, 7, 4) = "0000" Then
    day2 = Mid$(lx, 1, 2)
    mon2 = Mid$(lx, 4, 2)
    milk = 1
    End If
    datev = lx
    If lx <> "" And milk = 0 Then
        If Mid$(lx, 1, 1) <> "#" Then
            ty = InStr(lx, " ") - 3
            day2 = Mid(lx, 1, ty)
            datev = Format(Val(day2), "00") + "-" + mon2 + "-0000 " + Trim(Mid$(lx, ty + 3))
        End If
    End If
    If milk = 0 Then
        datefx = datefx + datev + vbCrLf
    End If
Loop Until EOF(1)
Close #1

Open "D:\VB6\VB-NT\00_Best_VB_01\Date1991\00 Data\Teachers Dates2.txt" For Output As #1
    Print #1, datefx
Close #1

End
SkipHere:
' ----------------------------------------------------------------------

Form1.Left = 0
Form1.Top = 420
Form1.Width = Screen.Width
Form1.Height = Screen.Height - 1000
List1.Left = 0
List2.Left = 0
List1.Width = Form1.Width - 90
List2.Width = Form1.Width - 90
List1.Height = Form1.Height - 2400
ListView1.Left = 0
ListView2.Left = 0
ListView2.Top = 0
ListView1.Top = ListView2.Height + ListView2.Top
ListView2.Width = Form1.Width - 90
Label1.Left = 0
Label2.Left = 0
Label3.Left = 0
Label1.Top = Form1.Height - Label1.Height - 700
Label2.Top = Label1.Top - Label2.Height
Label3.Top = Label2.Top - Label3.Height
Label1.Width = Form1.Width
Label2.Width = Form1.Width
Label3.Width = Form1.Width
Label5.Top = Form1.Height / 2 - Label5.Height / 2
Label5.Left = 0
Label5.Width = Form1.Width
PBar1.Top = Form1.Height / 2 - Label5.Height / 2 + 1900
PBar1.Left = 0
PBar1.Width = Form1.Width

Me.Show

Me.WindowState = vbMaximized

ListView2.Visible = True
    
With ListView2
    .ColumnHeaders.Add , "#", "#", 550, lvwColumnLeft
    .ColumnHeaders.Add , "Day1", "Day", 600, lvwColumnLeft
    .ColumnHeaders.Add , "Day2", "Day", 600, lvwColumnLeft
    .ColumnHeaders.Add , "Days To Go", "Days", 950, lvwColumnLeft
    .ColumnHeaders.Add , "Date", "Start", 1200, lvwColumnLeft
    .ColumnHeaders.Add , "Years", "Years", 750, lvwColumnLeft
    .ColumnHeaders.Add , "Same Day Count", "Same", 750, lvwColumnLeft
    .ColumnHeaders.Add , "Age For Nx Same", "Age For", 850, lvwColumnLeft
    .ColumnHeaders.Add , "On a Leap", "Born", 650, lvwColumnLeft
    .ColumnHeaders.Add , "Born New Moon", "Born", 800, lvwColumnLeft
    .ColumnHeaders.Add , "Nx Last Bday on Full Moon", "Next Last", 1100, lvwColumnLeft
    .ColumnHeaders.Add , "Description", "Description", 10000, lvwColumnLeft
    .View = lvwReport
End With
With ListView1
    .ColumnHeaders.Add , "#", "#", 550, lvwColumnLeft
    .ColumnHeaders.Add , "Day1", "Day", 600, lvwColumnLeft
    .ColumnHeaders.Add , "Day2", "Day", 600, lvwColumnCenter
    .ColumnHeaders.Add , "Days To Go", "Days", 950, lvwColumnCenter
    .ColumnHeaders.Add , "Date", "Start", 1200, lvwColumnCenter
    .ColumnHeaders.Add , "Years", "Years", 750, lvwColumnLeft
    .ColumnHeaders.Add , "Same Day Count", "Same", 750, lvwColumnLeft
    .ColumnHeaders.Add , "Age For Nx Same", "Age For", 850, lvwColumnLeft
    .ColumnHeaders.Add , "On a Leap", "Born", 650, lvwColumnCenter
    .ColumnHeaders.Add , "Born New Moon", "Born", 800, lvwColumnCenter
    .ColumnHeaders.Add , "Nx Last Bday on Full Moon", "Next Last", 1100, lvwColumnCenter
    .ColumnHeaders.Add , "Description", "Description", 30000, lvwColumnLeft
    .View = lvwReport
End With
With ListView2
        Set LV = .ListItems.Add(, , "#")
        LV.SubItems(1) = "Day"
        LV.SubItems(2) = "Day"
        LV.SubItems(3) = "Days"
        LV.SubItems(4) = "Start"
        LV.SubItems(5) = "Years"
        LV.SubItems(6) = "Same"
        LV.SubItems(7) = "Age For"
        LV.SubItems(8) = "Born"
        LV.SubItems(9) = "Born"
        LV.SubItems(10) = "Next Last"
        LV.SubItems(11) = "Description"
End With
With ListView2
        Set LV = .ListItems.Add(, , "")
        LV.SubItems(2) = "Now"
        LV.SubItems(3) = "To"
        LV.SubItems(4) = "Date"
        LV.SubItems(5) = "From"
        LV.SubItems(6) = "Dayer"
        LV.SubItems(7) = "Next"
        LV.SubItems(8) = "Leap"
        LV.SubItems(9) = "Days"
        LV.SubItems(10) = "B-Day"
End With
With ListView2
        Set LV = .ListItems.Add(, , "")
        LV.SubItems(3) = "Go"
        LV.SubItems(4) = "Start"
        LV.SubItems(5) = "Count"
        LV.SubItems(7) = "Same"
        LV.SubItems(8) = "Year"
        LV.SubItems(9) = "After A"
        LV.SubItems(10) = "On A Full"
End With
With ListView2
        Set LV = .ListItems.Add(, , "")
        LV.SubItems(5) = "Date"
        LV.SubItems(7) = "Dayer"
        LV.SubItems(8) = "New"
        LV.SubItems(10) = "Moon"
End With
With ListView2
        Set LV = .ListItems.Add(, , "")
        LV.SubItems(9) = "Moon"
End With
DoEvents

DoAgain = False
Open App.Path + "\DayNow.txt" For Input As #1
    Line Input #1, TTDATE
Close #1

If Date$ <> TTDATE Then DoAgain = True

Open App.Path + "\DayNow.txt" For Output As #1
    Print #1, Date$
Close #1

Dim F
Set FS2 = CreateObject("Scripting.FileSystemObject")
Set F = FS2.getfolder(App.Path + "\00 Data\")
g = F.Size

Open App.Path + "\Datasize.txt" For Input As #1
Line Input #1, ll$
Close #1
If g <> Val(ll$) Then DoAgain = True

Open App.Path + "\Datasize.txt" For Output As #1
    Print #1, F.Size
Close #1

List2.AddItem "Orig|Day |Days   |Start     |Years|Same |Age For|Born-Days  |Next/Last|Description"
List2.AddItem "Day |This|To     |Date      |From |Dayer|Next   |on a|After |B-Day On |"
List2.AddItem "    |Time|Go     |          |Start|Count|Same   |Leap|A New |A Full   |"
List2.AddItem "    |    |       |          |Date |     |Dayer  |Year|Moon  |Moon     |"
List4.AddItem "Orig|Day |Days   |Start     |Years|Same |Age For|Born-Days  |Next/Last|Description"
List4.AddItem "Day |This|To     |Date      |From |Dayer|Next   |on a|After |B-Day On |"
List4.AddItem "    |Time|Go     |          |Start|Count|Same   |Leap|A New |A Full   |"
List4.AddItem "    |    |       |          |Date |     |Dayer  |Year|Moon  |Moon     |"

Form1.Show
Form1.WindowState = 1

If DoAgain = False Then
    Open App.Path + "\List1.txt" For Input As #1
        Do
            If Not (EOF(1)) Then
                Line Input #1, ll$
                List1.AddItem ll$
                With ListView1
                Set LV = .ListItems.Add(, , .ListItems.Count + 1)
                LV.SubItems(1) = Trim(Mid$(ll$, 1, 3))
                LV.SubItems(2) = Trim(Mid$(ll$, 6, 3))
                LV.SubItems(3) = Trim(Mid$(ll$, 11, 7))
                LV.SubItems(4) = Trim(Mid$(ll$, 19, 10))
                LV.SubItems(5) = Trim(Mid$(ll$, 30, 5))
                LV.SubItems(6) = Trim(Mid$(ll$, 36, 5))
                LV.SubItems(7) = Trim(Mid$(ll$, 42, 7))
                LV.SubItems(8) = Trim(Mid$(ll$, 50, 4))
                LV.SubItems(9) = Trim(Mid$(ll$, 55, 6))
                LV.SubItems(10) = Trim(Mid$(ll$, 62, 9))
                LV.SubItems(11) = Trim(Mid$(ll$, 72))
                End With
            End If
        Loop Until EOF(1)
    Close #1
End If
    

If DoAgain = True Then
    PBar1.Visible = True
    Call Explish
    ' END OF CODE SUB ROUTINE
    ' Open App.Path + "\List1.txt" For Output As #FR1
End If

PBar1.Visible = False
Label5.Visible = False
ListView1.Width = Screen.Width - 90
ListView1.Height = (Screen.Height - 1000) - ListView2.Height - 700
ListView1.Visible = True
ListView2.Visible = True
Form1.ListView1.SetFocus
Call Mnu_Select_Click
Finished = True
RunOnce = True

'OutPut - OutPour - Inpour DownPour - Poor
If DoAgain = True Then
    FR1 = FreeFile
    Open App.Path + "\List1.txt" For Output As #FR1
        For R = 1 To List1.ListCount - 1
            Print #FR1, List1.List(R)
        Next
    Close FR1
End If

Me.WindowState = vbMaximized

If Command$ <> "" Then Unload Me

End Sub

Private Sub Form_Resize()
    Call Mnu_Select_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Mnu_Clip_Them_Click
'End


If Redo = False Then End

End Sub


Sub OutPutSig()

On Local Error GoTo ErrSub

FR1 = FreeFile
Open "D:\# MY DOCS\# 01 My Documents\CALLerID\2DOUBLE.TXT" For Input As #FR1
tr$ = Input$(LOF(FR1), FR1)
Close FR1

On Local Error GoTo 0


tg = InStr(tr$, "DOUBLE WORD SAME FIRST LETTER")
tg = InStr(tg, tr$, vbCrLf)
tr$ = Mid$(tr$, tg + 2)

trsx = 0
tg = 0
Do
tg = InStr(tg + 1, tr$, vbCrLf)
If tg > 0 Then trsx = trsx + 1
Loop Until tg = 0

Randomize Int(Now) + 1
Counter = 5
Dim DWdup$
Dim DW(100)
For R = 1 To Counter
Do
rt = Int(Rnd(1) * Len(tr$)) + 1
tg = InStr(rt, tr$, vbCrLf)
Loop Until InStr(DWdup$, Str(tg)) = 0
DWdup$ = DWdup$ + Str(tg)
tg2 = InStrRev(tr$, vbCrLf, rt)
DW(R) = Mid$(tr$, tg2 + 2, (tg - tg2) - 2)

Next



If Dir("D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures\Signature-DW of the Day.txt") <> "" Then
    FR1 = FreeFile
    Open "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures\Signature-DW of the Day.txt" For Output As #FR1
    
    Print #FR1, "Double Words of The Day - Dbase "; Str(trsx)
    For R = 1 To Counter
    Print #FR1, DW(R) 'dw(counter)
    Next
    Close #FR1
End If

Exit Sub

ErrSub:
    DoEvents
    Resume

End Sub


Sub SeasonsEqiSol()

ReDim SeasonsDate(20)
Dim Cuoo

Dim srs As New clsSunRiseSet
srs.City = "London, England"
Hatt = 1

For rty2 = -1 To 20 '-this set how many year covers
    For rty4 = 6 To 12 Step 6 '-Short Longest Day
        For rty3 = -2 To 5 '-this norm cover year
            iy = 0
            If rty4 = 6 Then toolnow = DateSerial(Year(Now) + rty2, 6, rty3 + 21)
            If rty4 = 12 Then toolnow = DateSerial(Year(Now) + rty2, 12, rty3 + 21)
            tool3 = toolnow
            srs.DateDay = toolnow
            srs.CalculateSun
            tool1 = srs.Sunrise
            tool2 = srs.Sunset
            tool5 = DateDiff("s", tool1, tool2)
            'toolnow = DateSerial(Year(Now) + rty2, 6, rty3 + 22)
            toolnow = toolnow + 1
            srs.DateDay = toolnow
            srs.CalculateSun
            tool1 = srs.Sunrise
            tool2 = srs.Sunset
            
            toolx = toolnow + 1
            
            tool6 = DateDiff("s", tool1, tool2)
            If tool5 < tool6 Then toolnow = toolx: Exit For
        Next
        If rty4 = 6 Then Longest = toolnow
        If rty4 = 12 Then Shortest = toolnow - 2
    Next


    'LONGEST DAY - Summer Solstice Jun
    'SHORTEST DAY - Winter Solstice Dec
    
    
    '---- Tue 10-Jun-2008 04:06:33p ----  ----
    'Phil Plait 's Bad Astronomy: Misconceptions
    'Http://www.badastronomy.com/bad/misc/badseasons.html
    
    '2008
    '  Vernal Equinox Mar 20 2008 05:48 UT
    '  Summer Solstice Jun 20 2008 23:59 UT
    '  Autumnal Equinox Sep 22 2008 15:44 UT
    '  Winter Solstice Dec 21 2008 12:04 UT
    
    'Spring
    'Vernal Equinox Mar 20 2008 05:48 UT
    Cuoo1 = Format$((Longest - 93), "DD-MM-yyyy")
    Debug.Print Cuoo1 + " -- Vernal Equinox Spring 93 Days Before Longest Day Summer Solstice"
    
    'Summer Solstice
    '  Summer Solstice Jun 20 2008 23:59 UT
    Cuoo2 = Format$((Longest), "DD-MM-yyyy")
    Debug.Print Cuoo2 + " -- Summer Solstice Longest Day"
    
    'autumn
    '  Autumnal Equinox Sep 22 2008 15:44 UT
    
    Cuoo3 = Format$(Longest + 93, "DD-MM-yyyy")
    Debug.Print Cuoo3 + " -- Autumnal Equinox 93 Days After Longest Day Summer Solstice"
    
    'Winter
    '  Winter Solstice Dec 21 2008 12:04 UT
    Cuoo4 = Format$((Shortest), "DD-MM-yyyy")
    Debug.Print Cuoo4 + " -- Winter Solstice Shortest Day"
    
    
    If Hatt + 5 > UBound(SeasonsDate) Then
        ReDim Preserve SeasonsDate(Hatt + 10)
    End If
    SeasonsDate(Hatt) = Cuoo1: Hatt = Hatt + 1
    SeasonsDate(Hatt) = Cuoo2: Hatt = Hatt + 1
    SeasonsDate(Hatt) = Cuoo3: Hatt = Hatt + 1
    SeasonsDate(Hatt) = Cuoo4: Hatt = Hatt + 1

Next

Debug.Print "Time Zone London"
Debug.Print "That's the Mathmatics Way to Do It"
Debug.Print "One Thing to Look out for Is the 21st and 22nd of December Winter Solstice Shortest Day"
Debug.Print "Solstice and Equinox Calculating 4 Cross Quarter Seasons"
Debug.Print "My Code First Found the Solution When Calculating Sunrise-Set Difference Of Long Short Days When Switch to Look at the Shortest Day Because was Easier to Calc as there was a Sharper Arc to Pin a Change in Differencing Time Arc"

Debug.Print "---- Double Checker Verify Source - Dead Link not Checked"
Debug.Print "Http://www.badastronomy.com/bad/misc/badseasons.html"
Debug.Print "Calculates the Actual Time in the Day From the Equinox or Solstice"
Debug.Print "----"

Hatt = Hatt - 1
ReDim Preserve SeasonsDate(Hatt)

Dim SeasonsDate2() As Integer

ReDim SeasonsDate2(UBound(SeasonsDate))


For RI = 1 To UBound(SeasonsDate)

    String2 = ""
    FormatDate$ = "DDD-DD-MMM-YYYY"
    If Month(SeasonsDate(RI)) = 3 Then
        String1 = " Vernal Equinox"
        Call NextLastAnMiddCode(DateValue(SeasonsDate(RI)))
    End If
    If Month(SeasonsDate(RI)) = 6 Then
        String1 = " Summer Solstice"
        Call NextLastAnMiddCode(DateValue(SeasonsDate(RI)))
    End If
    
    If Month(SeasonsDate(RI)) = 6 Then
        String1 = "Mid Summers Day (made famous by William Shakespeare)"
        Call NextLastAnMiddCode(DateValue(SeasonsDate(RI)))
    End If
    If Month(SeasonsDate(RI)) = 6 Then
        String1 = "Longest Day"
        Call NextLastAnMiddCode(DateValue(SeasonsDate(RI)))
    End If
    
    If Month(SeasonsDate(RI)) = 9 Then
        String1 = " Autumnal Equinox"
        Call NextLastAnMiddCode(DateValue(SeasonsDate(RI)))
    
    
        xcd = PreviousFullMoon(DateValue(SeasonsDate(RI)))
        String1 = " Harvest Moon - Full Moon Closest to the Autumnal Equinox"
        FormatDate$ = "DD-MM-YYYY - HH:MM:SS"
        Call NextLastAnMiddCode(xcd)
    
    
    
    End If
    If Month(SeasonsDate(RI)) = 12 Then
        String1 = " Winter Solstice"
        Call NextLastAnMiddCode(DateValue(SeasonsDate(RI)))
    End If
    

Next


FF = FreeFile
Open App.Path + "\SEASONS.txt" For Output As #FF
For RI = 1 To UBound(SeasonsDate2) - 1
    ii1 = DateValue(SeasonsDate(RI))
    ii2 = DateValue(SeasonsDate(RI + 1))
    Diff = DateDiff("d", ii1, ii2)
    ii3 = ii2 - (Diff / 2) ' WHY /2 -- CENTER OF SEASON /?
    SeasonsDate2(RI) = Diff 'Int(ii3) ' DIFFERANCE OF DAYS
    
    
    
    String2 = ""
    FormatDate$ = "DDD-DD-MMM-YYYY"
    If Month(SeasonsDate(RI)) = 3 Then
        String1 = " Spring -- Cross-Quarter Season Begin"
        Debug.Print "1" + Diff
    End If
    
    If Month(SeasonsDate(RI)) = 6 Then
        String1 = " Summer -- Cross-Quarter Season Begin"
        Debug.Print "2" + Diff
    End If
    
    If Month(SeasonsDate(RI)) = 9 Then
        String1 = " Autumn -- Cross-Quarter Season Begin"
        Debug.Print "3" + Diff
    End If
    
    If Month(SeasonsDate(RI)) = 12 Then
        String1 = " Winter -- Cross-Quarter Season Begin"
        Debug.Print "4" + Diff
    End If
    
    
    Print #FF, SeasonsDate(RI);
    Print #FF, " -- ";
    If RI + 1 < UBound(SeasonsDate2) Then
        Print #FF, SeasonsDate(RI + 1)
        Else
        Print #FF,
    End If
    Print #FF, String1
    Print #FF, Trim(Str(Diff)) + " -- DIFFERANCE OF DAYS"
    Print #FF, " -- "
    Print #FF, " -- "
    
    'TEMP DISABLE 2016
'    Call NextLastAnMiddCode(SeasonsDate2(ri))

Next
Close #FF

'TEMP INTRODUCTION
' End


ReDim Preserve SeasonsDate2(UBound(SeasonsDate2) - 1)
'-----------------------------------
'--------Start farmers moon

Dim SeasCount
MoonCount = 0
'OldYear = 0
dsf = 1900
'xcd = DateSerial(dsf, 1, 1)
xcd = SeasonsDate2(1)
dsf = xcd
Tag1 = 0
SeasCount = 1
Do
    xcd = NextFullMoon(xcd)
    If SeasCount + 1 > UBound(SeasonsDate2) Then
        Exit Do
    End If
    If xcd > SeasonsDate2(SeasCount + 1) Then
        MoonCount = 0: SeasCount = SeasCount + 1
    End If
    MoonCount = MoonCount + 1
    
    If MoonCount >= 4 Then
        If R + 5 > UBound(Mls$) Then
            ReDim Preserve Mls$(R + 1000)
        End If
        
        String1 = " The Maine Farmer's Almanac Blue Moon"
        String2 = " - 4 Full Moons in Season"
        FormatDate$ = "DD-MM-YYYY - HH:MM:SS"
        Call NextLastAnMiddCode(xcd)
    
    End If
    'OldYear = Year(xcd)
Loop Until 1 = 2







    







End Sub


Sub NextLastAnMiddCode(xcd5)
        R = R + 3
        'If R >= 6508 Then Stop
        'If R < 4 Then MsgBox "R @ 1": Stop
        abc = R Mod 3: If abc <> 1 Then MsgBox "R out of sequence ": Stop
        Mls$(R) = Format$(xcd5, "DD-MM-yyyy")
        'Mls$(R-3)
        
        String1 = Trim(String1)
        String2 = Trim(String2)
        
        If String2 <> "" Then String2 = " --- " + String2
        If InStr(TaggStr, String1) = 0 Then
            TaggCnt = TaggCnt + 1
            TaggStr = TaggStr + Format$(TaggCnt, " 0000 ") + String1
        End If
        
        mv2 = InStr(TaggStr, String1)
        mv3 = Val(Mid$(TaggStr, mv2 - 6, 5))
        Test = Mid$(TaggStr, mv2 - 6, 50)
        
        atvar = String1 + " ---- @ " + Format$(xcd5, FormatDate$) + String2
        
        'If Year(xcd5) > 2009 Then Stop
        If Tagg5(mv3) = 0 And xcd5 > Now Then
            Mls$(R + 1) = "--- Next -- " + atvar
            If LastestInx(mv3) = 0 Then
                ListCollect = ListCollect + "Your not Giving a Last For" + vbCrLf + Test + vbCrLf + vbCrLf
                MsgBox ListCollect
                Stop
            End If
            If LastestInx(mv3) > 0 Then
                Mls$(LastestInx(mv3) + 1) = "--- Last -- " + Lastest(mv3)
                'Mls$ (LastestInx(mv3) )
                'Mls$ (R )
                Tagg5(mv3) = 1
            End If
        Else
            If xcd5 > Now Then Drop$ = "Next" Else Drop$ = "Last"
            Lastest(mv3) = atvar
            LastestInx(mv3) = R
            'Mls$(R + 1) = "--- " + Drop$ + " -- " + Lastest(mv3)
            Mls$(R + 1) = "--- " + Lastest(mv3)
        End If
End Sub


Private Sub Explish()

ListView1.Enabled = False

ReDim Mls$(10000)
ReDim Tagg5(1000)
ReDim Lastest(1000)
ReDim LastestInx(1000)

ftag = 1

Call OutPutSig

ReDim Preserve Rvm1(1000)
ReDim Preserve Rvm2(1000)

Diar = 0

Form1.Show

If RunOnce = False And HOUR(Now) = 0 And Minute(Now) < 5 Then Form1.WindowState = 1
DoEvents

enow = Now

w1 = enow
d$ = Format$(enow, "yyyy")
F = enow
jk = 0

R = -2

For i2 = 1 To 3
If i2 = 1 Then Open App.Path + "\00 Data\1991Dat2.txt" For Input As #1
If i2 = 2 Then Open App.Path + "\00 Data\Teachers Dates.txt" For Input As #1
If i2 = 3 Then Open App.Path + "\00 Data\1991Dat3 Secure.txt" For Input As #1
Do
    Line Input #1, A$
    A$ = Trim(A$)
    
    gg$ = Trim(Mid$(A$, 12))
    If Mid$(A$, 1, 1) <> "'" And Mid$(A$, 1, 1) <> "#" And A$ <> "" And gg$ <> "" Then
        milk = 0
        If Mid$(A$, 7, 4) = "0000" Then
            For Rw = -1 To 1
                DateChk = DateSerial(Year(Now) + Rw, Val(Mid$(A$, 4, 2)), Val(Mid$(A$, 1, 2)))
                String1 = "Event:- " + Mid$(A$, 12)
                String2 = ""
                FormatDate$ = "DDD-DD-MMM-YYYY"
                Call NextLastAnMiddCode(DateChk)
            Next
            milk = 1
        End If
        If milk = 0 Then
            R = R + 3
            Mls$(R) = Mid$(A$, 1, 10)
            Mid$(Mls$(R), 3, 1) = "-"
            Mid$(Mls$(R), 6, 1) = "-"
        
            Mls$(R + 1) = Mid$(A$, 12)
            Mls$(R + 2) = ""
        End If
    End If
Loop Until EOF(1)
Close #1
Next i2

Open App.Path + "\00 Data\1991Dat3 Secure.txt" For Input As #1
KK$ = Input$(LOF(1), 1)
Close #1



'Year zero Moon Phases
'Year      New Moon       First Quarter       Full Moon       Last Quarter          ?T

'    1                                                         Jan  5  08:09      02h56m
'        Jan 13  10:57     Jan 21  09:04     Jan 28  01:02     Feb  3  23:54
'        Feb 12  05:17     Feb 19  19:21     Feb 26  10:53     Mar  5  17:40
'        Mar 13  20:48     Mar 21  02:41     Mar 27  20:41     Apr  4  12:12
'        Apr 12  09:18     Apr 19  08:04     Apr 26  07:15     May  4  06:09
'        May 11  19:19     May 18  12:48     May 25  19:23     Jun  2  22:30

'First full moon year Dot -- Jan  5  08:09
'First full moon before easter since year dot -- Mar 27  20:41

'http://eclipse.gsfc.nasa.gov/phase/phases0001.html

'NASA - Moon Phases: 2001 to 2100
'http://web.archive.org/web/20140727124725/http://eclipse.gsfc.nasa.gov/phase/phases2001.html

'Easter Sunday can fall on any date from March 22 to April 25
'The cycle of Easter dates repeats after exactly 5700000 years, ...
'http://www.assa.org.au/edm.html#Method

'Use a calculator to divide the year by 19
'Obtain the 3 digits after the decimal point
'Find the Paschal Full Moon (PFM) date in Table A below
'Then continue below with Step 2
'Year 2285 example: 2285 ÷ 19 = 120.263157, so use the .263 part to find the PFM date in Table A.   The result is March 21, 2285 A.D.
'Year 2005 example: 2005 ÷ 19 = 105.52632, so use the .526 part to find the PFM date in Table A.   The result is March 21, 2285 A.D.

'Table A: PFM Date for Years 326 to 2599 (M=March, A=April)
'Fraction
'after dividing
'year by 19

Dim Year2(9, 2) As Integer
Year2(1, 1) = 0: Year2(1, 2) = 325
Year2(2, 1) = 326: Year2(2, 2) = 1582
Year2(3, 1) = 1583: Year2(3, 2) = 1699
Year2(4, 1) = 1700: Year2(4, 2) = 1899
Year2(5, 1) = 1900: Year2(5, 2) = 2199
Year2(6, 1) = 2200: Year2(6, 2) = 2299
Year2(7, 1) = 2300: Year2(7, 2) = 2399
Year2(8, 1) = 2400: Year2(8, 2) = 2499
Year2(9, 1) = 2500: Year2(9, 2) = 2599

Dim east1(19) As Double
east1(1) = 0
east1(2) = 0.052
east1(3) = 0.105
east1(4) = 0.157
east1(5) = 0.21
east1(6) = 0.263
east1(7) = 0.315
east1(8) = 0.368
east1(9) = 0.421
east1(10) = 0.473
east1(11) = 0.526
east1(12) = 0.578
east1(13) = 0.631
east1(14) = 0.684
east1(15) = 0.736
east1(16) = 0.789
east1(17) = 0.842
east1(18) = 0.894
east1(19) = 0.947

Dim eastm$(19, 6)
eastm$(1, 1) = "A5": eastm$(1, 2) = "A12": eastm$(1, 3) = "A13": eastm$(1, 4) = "A14": eastm$(1, 5) = "A15": eastm$(1, 6) = "A16"
eastm$(2, 1) = "M25": eastm$(2, 2) = "A1": eastm$(2, 3) = "A2": eastm$(2, 4) = "A3": eastm$(2, 5) = "A4": eastm$(2, 6) = "A5"
eastm$(3, 1) = "A13": eastm$(3, 2) = "M21": eastm$(3, 3) = "M22": eastm$(3, 4) = "M23": eastm$(3, 5) = "M24": eastm$(3, 6) = "M25"
eastm$(4, 1) = "A2": eastm$(4, 2) = "A9": eastm$(4, 3) = "A10": eastm$(4, 4) = "A11": eastm$(4, 5) = "A12": eastm$(4, 6) = "A13"
eastm$(5, 1) = "M22": eastm$(5, 2) = "M29": eastm$(5, 3) = "M30": eastm$(5, 4) = "M31": eastm$(5, 5) = "A1": eastm$(5, 6) = "A2"
eastm$(6, 1) = "A10": eastm$(6, 2) = "A17": eastm$(6, 3) = "A18": eastm$(6, 4) = "A18": eastm$(6, 5) = "M21": eastm$(6, 6) = "M22"
eastm$(7, 1) = "M30": eastm$(7, 2) = "A6": eastm$(7, 3) = "A7": eastm$(7, 4) = "A8": eastm$(7, 5) = "A9": eastm$(7, 6) = "A10"
eastm$(8, 1) = "A18": eastm$(8, 2) = "M26": eastm$(8, 3) = "M27": eastm$(8, 4) = "M28": eastm$(8, 5) = "M29": eastm$(8, 6) = "M30"
eastm$(9, 1) = "A7": eastm$(9, 2) = "A14": eastm$(9, 3) = "A15": eastm$(9, 4) = "A16": eastm$(9, 5) = "A17": eastm$(9, 6) = "A18"
eastm$(10, 1) = "M27": eastm$(10, 2) = "A3": eastm$(10, 3) = "A4": eastm$(10, 4) = "A5": eastm$(10, 5) = "A6": eastm$(10, 6) = "A7"
eastm$(11, 1) = "A15": eastm$(11, 2) = "M23": eastm$(11, 3) = "M24": eastm$(11, 4) = "M25": eastm$(11, 5) = "M26": eastm$(11, 6) = "M27"
eastm$(12, 1) = "A4": eastm$(12, 2) = "A11": eastm$(12, 3) = "A12": eastm$(12, 4) = "A13": eastm$(12, 5) = "A14": eastm$(12, 6) = "A15"
eastm$(13, 1) = "M24": eastm$(13, 2) = "M31": eastm$(13, 3) = "A1": eastm$(13, 4) = "A2": eastm$(13, 5) = "A3": eastm$(13, 6) = "A4"
eastm$(14, 1) = "A12": eastm$(14, 2) = "A18": eastm$(14, 3) = "M21": eastm$(14, 4) = "M22": eastm$(14, 5) = "M23": eastm$(14, 6) = "M24"
eastm$(15, 1) = "A1": eastm$(15, 2) = "A8": eastm$(15, 3) = "A9": eastm$(15, 4) = "A10": eastm$(15, 5) = "A11": eastm$(15, 6) = "A12"
eastm$(16, 1) = "M21": eastm$(16, 2) = "M28": eastm$(16, 3) = "M29": eastm$(16, 4) = "M30": eastm$(16, 5) = "M31": eastm$(16, 6) = "A1"
eastm$(17, 1) = "A9": eastm$(17, 2) = "A16": eastm$(17, 3) = "A17": eastm$(17, 4) = "A17": eastm$(17, 5) = "A18": eastm$(17, 6) = "M21"
eastm$(18, 1) = "M29": eastm$(18, 2) = "A5": eastm$(18, 3) = "A6": eastm$(18, 4) = "A7": eastm$(18, 5) = "A8": eastm$(18, 6) = "A9"
eastm$(19, 1) = "A17": eastm$(19, 2) = "M25": eastm$(19, 3) = "M26": eastm$(19, 4) = "M27": eastm$(19, 5) = "M28": eastm$(19, 6) = "M29"

'd$ = "2285"

'GoTo JUMP4


xhag = R

qd = Val(d$) - 1

ReDim Tagger(50)
ReDim Tagge2(50)
For i = 1980 To 3000
    TheDate = DateSerial(i, 3, 22)
    OL = Int(NextFullMoon(TheDate))
    W = Int(PreviousFullMoon(OL))
    '6
    DateChk = OL - 2
    
    If Day(DateChk) = 13 Then
        String1 = "Good Friday, On Friday the 13th"
        String2 = ""
        FormatDate$ = "DDD-DD-MMM-YYYY"
        Call NextLastAnMiddCode(DateChk)
    End If
Next
here:
GoTo JUMP4


iy = 0

For rs = 1 To 9
    If qd > Year2(rs, 1) And qd < Year2(rs, 2) Then qx = rs
Next

qx = qx - 1
If qx = 8 Then qx = 6
If qx = 7 Then qx = 5

qj = qd / 19
wsd$ = Str$(qj)
qz = InStr(wsd$, ".")
azx$ = Mid$(wsd$, qz, 4)
afg = Val(azx$)

For rs = 1 To 19
    If afg = east1(rs) Then ygh2 = rs
Next

ftad$ = eastm$(ygh2, qx)

If Mid$(ftad$, 1, 1) = "M" Then loj = 3 Else loj = 4
loj2 = Val(Mid$(ftad$, 2))

'God, the center and focus of religious faith, a holy being or ultimate reality to whom worship and prayer are
'addressed. Especially in monotheistic religions (see Monotheism), God is considered the creator or source of
'everything that exists and is spoken of in terms of perfect attributes-for instance, infinitude, immutability, eternity,
'goodness, knowledge (omniscience), and power (omnipotence). Most religions traditionally ascribe to God
'certain human characteristics that can be understood either literally or metaphorically, such as will, love, anger,
'and forgiveness.

'This is Full Moon before Easter
        W = DateSerial(Val(qd), loj, loj2)

'ol = PreviousFullMoon(W)
'Easter Full Moon is next full moon after Mar 22
'TheDate = DateSerial(2007, 3, 22)
'ol = NextFullMoon(TheDate)

'First full moon year Dot -- Jan  5  08:09
'First full moon before easter since year dot -- Mar 27  20:41

'Easter Sunday is Next Sunday After Easter Full Moon

'Debug.Print Format$(DateSerial(2000, 3, 27), "DDD")
' Monday

'Debug.Print Format$(DateSerial(2000, 3, 27 + 6), "DDD")
' Sunday Easter year Dot - 0000-04-02 -YYYY-MM-DD

'famous words
'As Jesus apparently died on year 0029 with a question mark
'You might be interested to know, if it was, the first Easter then was
'April 15th which was the same day as the full moon...

'Lets see

'Apr 17  02:45 ' first moon after mar 22nd 0029 YYYY

'Debug.Print Format$(DateSerial(2000, 4, 17), "DDD")
'Monday - Monday again same as year dot
'Is Later Thou

'Debug.Print Format$(DateSerial(2000, 4, 17 + 6), "DDD")
'Sunday Easter year 0029 - 0000-04-23 -YYYY-MM-DD

'15 dates for easter Includes Mothers Day
'Easter is the Zero Date to Find
'+60 for Corpus Christi
'-47 for Shrove Tuesday
'Total 107 Days

'This is Full Moon before Easter
W = DateSerial(Val(qd), loj, loj2)

FullMoonBeforeEaster = W



'--  WORK TO HAVE MATH OF TH\N TABLE
'    GoTo JUMP_OUTMATH
JUMP4:
'
'JUMP_OUTMATH:

'Dim OL As Date



'Easter Full Moon is next full moon after Mar 22

'-----------------------------------------------
'TRY HERE SORT THE NOW NEXT BUG WORKAROUND TEMP
'-----------------------------------------------

GoTo JUMP_RELIGOUS_DATES

''This is Full Moon before Easter
'W = DateSerial(Val(qd), loj, loj2)
'
'FullMoonBeforeEaster = W
'
'For i = W To W + 10
'    gv$ = Format$(i, "DDD")
'    If gv$ = "Sun" Then Exit For
'Next
'
'If poodog = 0 Then xcd = W
'
'W = i
'
'If poodog2 = 0 Then xcd2 = W: poodog2 = 1
'
'If poodog = 0 Then chill = i

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'Easter Full Moon is next full moon after Mar 22
For i3 = Year(Now) To 2020
TheDate = DateSerial(i3, 3, 22)

i = Int(NextFullMoon(TheDate))
FullMoonBeforeEaster = i
'Same
DateChk = i - 47
String1 = "Shrove Tue - 1st Biblical Day Year -47 Days From Easter Sunday +60 for Corpus Christi -- Tott 107 Days - Tott 17+ Biblical Days"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'1
DateChk = i - 47
String1 = "*Shrove Tuesday* is the day before Lent begins, when confession leads one to be shriven of sins for which the penitent is given absolution."
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'2
DateChk = i - 46
String1 = "Ash Wednesday - First day of the penitential season of Lent ""Remember that you are dust, and unto dust you shall return."""
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'3
DateChk = i - 42
String1 = "First sunday in lent"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'Same
'------------------------------------------------------------
'Fourth sunday in lent
'1st sunday an three weeks
Lentmothers = Format$(i - 42, "DD-MM-yyyy")
DateChk = DateValue(Lentmothers) + 21
String1 = "Mothering Sunday 4th Sunday in Lent"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'4
'Holy Week is sometimes called the "Great Week" by Roman Catholic and Orthodox Christians
DateChk = i - 7
String1 = "Holy Week - The week immediately preceding Easter, beginning with Palm Sunday."
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'Same
DateChk = i - 7
String1 = "Palm Sunday -- Custom of blessing palms -- In commemoration of the triumphal entry of Jesus into Jerusalem"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'5
DateChk = i - 3
String1 = "Maundy Thursday or Holy Thursday, the Thursday before Easter Sunday, observed by Christians in commemoration of Christ's Last Supper"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'6
DateChk = i - 2
String1 = "Good Friday, Friday immediately preceding Easter, celebrated by Christians as the anniversary of Christ's crucifixion."
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)


'7
DateChk = i - 1
String1 = "Holy Saturday commemorates the burial of Christ"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'8
DateChk = FullMoonBeforeEaster
String1 = "Full Moon Before Easter -- Easter Sunday is Next Sunday After Easter Full Moon"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'9
'Easter Sunday is Next Sunday After Easter Full Moon
DateChk = i
String1 = "Easter Sunday -- Easter Full Moon is the next Full Moon after Mar 22"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'10
DateChk = i + 1
String1 = "Easter Monday Holiday"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'11
DateChk = i + 39
String1 = "Ascension Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'12
DateChk = i + 42
String1 = "Sunday after Ascension Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'13
DateChk = i + 49
String1 = "Pentecost"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'Same
DateChk = i + 49
String1 = "WhitSuntide, Week beginning with WhitSunday or Pentecost -- Whitsuntide (Anglo-Saxon hweta sunnandaeg, white Sunday)"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'14
DateChk = i + 50
String1 = "WhitMonday -- White Sunday, derives from the white garments worn by those baptized on Whitsunday."
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'15
DateChk = i + 51
String1 = "WhitTuesday"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'16
'Whitsuntide, week beginning with Whitsunday or Pentecost (the seventh Sunday and 50th day after Easter). The name Whitsuntide (Anglo-Saxon hweta sunnandaeg, "white Sunday")
'"white Sunday") derives from the white garments worn by those baptized on Whitsunday.
DateChk = i + 56
String1 = "Trinity Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'17
DateChk = i + 60
String1 = "Corpus Christi -- The Last Biblical Day of Year Besides Christmas - Makes 17 Christian Days In Year as My Count"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next
'r
'--------------------------------------------------------------------------------------
JUMP_RELIGOUS_DATES:
'TheDate = i
'TheDate = ConvertToUT(TheDate)

Dim TT As Long

TT = -657434
'smallest date 01-01-100
TheDate = TT



'xcd = NextFullMoon(TheDate)
'eggchoc = eggchoc + 1
'qd = qd + 1
'If eggchoc < 5 Then GoTo here

'Label(8) = "Age Of Lunation = " & Age(now) & " Days"
'Label(8) = "Age Of Lunation = " & Age(thedate) & " Days"
'Label(9) = "Angle Of Illumination = " & Angle(now) & "°"
'Label(9) = "Angle Of Illumination = " & Angle(thedate) & "°"
'Label(10) = "Percent Of Illumination = " & Illum(thedate) & "%"
'Label(11) = "Lunation Number = " & Lunation(now)
'Label(12) = "Moon Phase = " & MoonDescription(now)
'Label(13) = "Previous New Moon = " & Format(PreviousNewMoon(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(14) = "Previous First Quarter = " & Format(PreviousFirstQuarter(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(15) = "Previous Full Moon = " & Format(PreviousFullMoon(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(16) = "Previous Last Quarter = " & Format(PreviousLastQuarter(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(17) = "Next New Moon = " & Format(NextNewMoon(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(18) = "Next First Quarter = " & Format(NextFirstQuarter(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(19) = "Next Full Moon = " & Format(NextFullMoon(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(20) = "Next Last Quarter = " & Format(NextLastQuarter(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"


xcd = xcd2
xcd = DateSerial(2007, 1, 1)

Do
xcd = NextNewMoon(xcd)

DateChk = xcd
String1 = "New Moon *"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY HH:MM:SSa/p"
Call NextLastAnMiddCode(DateChk)

Loop Until xcd > DateSerial(2015, 1, 1)

xcd = DateSerial(2007, 1, 1)

Do
xcd = NextFullMoon(xcd)

DateChk = xcd
String1 = "Full Moon *"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY HH:MM:SSa/p"
Call NextLastAnMiddCode(DateChk)

Loop Until xcd > DateSerial(2015, 1, 1)


'-------------------------------------------------------------------------------
'The Harvest Moon is the first full moon after the first frost.
'-------------------------------------------------------------------------------

Tag1 = 0: tag2 = 0

dsf = 1900
Tag1 = 0
Do
    xcd = DateSerial(dsf, 2, 1)
    xcd = NextFullMoon(xcd)

    If Month(xcd) <> 2 Then
        xcd = DateSerial(dsf, 2, 1)
        DateChk = xcd
        String1 = "Black Moon"
        String2 = "No Full Moon in Feb"
        FormatDate$ = "DDD-DD-MMM-YYYY HH:MM:SSa/p"
        Call NextLastAnMiddCode(DateChk)
    
    End If
    dsf = dsf + 1
Loop Until dsf > 2200

Dim OldMonth
dsf = 1900
xcd = DateSerial(dsf, 1, 1)
Tag1 = 0
Do
    xcd = NextFullMoon(xcd)
    If OldMonth = Month(xcd) Then
        DateChk = xcd
        String1 = "Calendar Blue Moon"
        String2 = "2nd Full Moon in Month"
        FormatDate$ = "DDD-DD-MMM-YYYY HH:MM:SSa/p"
        Call NextLastAnMiddCode(DateChk)
    End If
    OldMonth = Month(xcd)
Loop Until Year(xcd) > 2200

Dim MoonCount
Dim OldYear
dsf = 1900
xcd = DateSerial(dsf, 1, 1)
Tag1 = 0
Do
    xcd = NextFullMoon(xcd)
    If Year(xcd) <> OldYear Then MoonCount = 0
    MoonCount = MoonCount + 1
    
    If MoonCount > 12 Then
        DateChk = xcd
        String1 = "Real Blue Moon"
        String2 = "13 Full Moons in Year"
        FormatDate$ = "DDD-DD-MMM-YYYY HH:MM:SSa/p"
        Call NextLastAnMiddCode(DateChk)
    End If
    OldYear = Year(xcd)
Loop Until Year(xcd) > 2200



'If the "Calendar Blue Moon" (1946-1999)

'xcd2 = DateSerial(2003, 7, 13)

xcd = xcd2

w1 = enow

'------------------------------------------
'------------------------------------------
'------------------------------------------
'------------------------------------------
'------------------------------------------

ReDim Tagger(50)
ReDim XvRs1(50)
ReDim XvRs2(50)
     
Dim srs As New clsSunRiseSet
srs.City = "London, England"

For rty2 = -5 To 20
For rty3 = -2 To 5
iy = 0
toolnow = DateSerial(Year(Now) + rty2, 12, rty3 + 21)
tool3 = toolnow
srs.DateDay = toolnow
srs.CalculateSun
tool1 = srs.Sunrise
tool2 = srs.Sunset
tool5 = DateDiff("s", tool1, tool2)
toolnow = DateSerial(Year(Now) + rty2, 12, rty3 + 22)
srs.DateDay = toolnow
srs.CalculateSun
tool1 = srs.Sunrise
tool2 = srs.Sunset
tool6 = DateDiff("s", tool1, tool2)
If tool5 < tool6 Then toolnow = tool3: Exit For
Next
'1st

    DateChk = toolnow
    String1 = "Shortest Day"
    String2 = ""
    FormatDate$ = "DDD-DD-MMM-YYYY"
    Call NextLastAnMiddCode(DateChk)

Next

ReDim Tagger(50)
ReDim XvRs1(50)
ReDim XvRs2(50)

'------------------------------
Call SeasonsEqiSol
' TEMPHERE
' End
'------------------------------

ReDim Tagger(50)
ReDim Tagge2(50)
ReDim XvRs1(50)
ReDim XvRs2(50)
For rty2 = -5 To 20
'srs.CalculateSun
'    Debug.Print srs.Sunrise, srs.Sunset
bert1 = 0
xtool1 = 0
bert2 = 1
xtool2 = 0
For rty3 = 1 To 40
iy = 0
toolnow = DateValue("01-12-" + LTrim$(Str$(Val(Format$(Now, "YYYY") + rty2)))) + rty3
tool3 = toolnow
srs.DateDay = toolnow
srs.CalculateSun
tool1 = srs.Sunrise '08.06.44
If tool1 - Int(tool1) > bert1 Then bert1 = tool1 - Int(tool1): xtool1 = tool1

tool2 = srs.Sunset '16.00.09
If tool2 - Int(tool2) < bert2 Then bert2 = tool2 - Int(tool2): xtool2 = tool2

'tool5 = DateDiff("s", tool1, tool2)
Next

R = R + 3
Mls$(R) = Format$(xtool1, "DD-MM-yyyy")
'mls$(R + 1) = "--- Latest Sunrise Winter" + Format$(xtool1, " hh:mm:ssa/p") + "--"
dizzy$ = "Latest Sunrise Winter"
iy = iy + 1
If xtool1 > Now And Tagger(iy) = 0 Then
    Mls$(R + 1) = "--- Next " + dizzy$ + " " + Format$(xtool1, " hh:mm:ssa/p")
    Tagger(iy) = 1
    Mls$(XvRs1(iy)) = "--- Last " + dizzy$ + " " + Format$(XvRs2(iy), " hh:mm:ssa/p")
Else
    XvRs1(iy) = R + 1: XvRs2(iy) = xtool1
    Mls$(R + 1) = "--- " + dizzy$ + " " + Format$(xtool1, " hh:mm:ssa/p")
End If


R = R + 3
Mls$(R) = Format$(xtool2, "DD-MM-yyyy")
'mls$(R + 1) = "--- Earliest Sunset Winter" + Format$(xtool2, " hh:mm:ssa/p") + "--"
dizzy$ = "Earliest Sunset Winter"
iy = iy + 1
If xtool2 > Now And Tagger(iy) = 0 Then
    Mls$(R + 1) = "--- Next " + dizzy$ + " " + Format$(xtool2, " hh:mm:ssa/p")
    Tagger(iy) = 1
    Mls$(XvRs1(iy)) = "--- Last " + dizzy$ + " " + Format$(XvRs2(iy), " hh:mm:ssa/p")
Else
    XvRs1(iy) = R + 1: XvRs2(iy) = xtool2
    Mls$(R + 1) = "--- " + dizzy$ + " " + Format$(xtool2, " hh:mm:ssa/p")
End If

Next

'------------------------
ReDim Tagger(50)
ReDim XvRs1(50)
ReDim XvRs2(50)

For rty2 = -5 To 5
'srs.CalculateSun
'    Debug.Print srs.Sunrise, srs.Sunset
bert1 = 1
xtool1 = 0
bert2 = 0
xtool2 = 0
For rty3 = 1 To 50
    iy = 0
    toolnow = DateValue("01-06-" + LTrim$(Str$(Val(Format$(Now, "YYYY") + rty2)))) + rty3
    tool3 = toolnow
    srs.DateDay = toolnow
    srs.CalculateSun
    tool1 = srs.Sunrise '08.06.44
    If tool1 - Int(tool1) < bert1 Then bert1 = tool1 - Int(tool1): xtool1 = tool1

    tool2 = srs.Sunset '16.00.09
    If tool2 - Int(tool2) > bert2 Then bert2 = tool2 - Int(tool2): xtool2 = tool2

Next

R = R + 3
Mls$(R) = Format$(xtool1, "DD-MM-yyyy")
dizzy$ = "Earliest Sunrise Summer"
iy = iy + 1
If xtool1 > Now And Tagger(iy) = 0 Then
    Mls$(R + 1) = "--- Next " + dizzy$ + " " + Format$(xtool1, " hh:mm:ssa/p")
    Tagger(iy) = 1
    Mls$(XvRs1(iy)) = "--- Last " + dizzy$ + " " + Format$(XvRs2(iy), " hh:mm:ssa/p")
Else
    XvRs1(iy) = R + 1: XvRs2(iy) = xtool1
    Mls$(R + 1) = "--- " + dizzy$ + " " + Format$(xtool1, " hh:mm:ssa/p")
End If

R = R + 3
Mls$(R) = Format$(xtool2, "DD-MM-yyyy")
dizzy$ = "Latest Sunset Summer"
iy = iy + 1
If xtool2 > Now And Tagger(iy) = 0 Then
    Mls$(R + 1) = "--- Next " + dizzy$ + " " + Format$(xtool2, " hh:mm:ssa/p")
    Tagger(iy) = 1
    Mls$(XvRs1(iy)) = "--- Last " + dizzy$ + " " + Format$(XvRs2(iy), " hh:mm:ssa/p")
Else
    XvRs1(iy) = R + 1: XvRs2(iy) = xtool2
    Mls$(R + 1) = "--- " + dizzy$ + " " + Format$(xtool2, " hh:mm:ssa/p")
End If

Next

'------------------------------------------

For lp = -2 To 2

lx = 0

W = DateSerial(Val(Format$(enow, "yyyy") + lx + lp), 11, 11)
For i = W To W + 10
gv$ = Format$(i, "DDD")
If gv$ = "Sun" Then Exit For
Next
If W + 4 < i Then i = i - 7
W = i

Datechkstr = Format$(i, "DD-MM-") + LTrim$(Str$(Val(Format$(enow, "yyyy")) + lx + lp))
DateChk = DateValue(Datechkstr)
String1 = "Rememberance Sunday Veterans Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)


'R = R + 3

DateChk = DateSerial(Year(Now) + lp, 11, 11)
String1 = "Rememberance Veterans Day 11th Day 11th Hour"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)



'------------------------------------------------------------

'R = R + 3

lx = 0

'Do

W = DateSerial(Val(Format$(enow, "yyyy") + lx + lp), 10, 1)
mond = 0
For i = W To W + 31
gv$ = Format$(i, "DDD")
If gv$ = "Mon" Then mond = mond + 1
If mond = 4 Then Exit For
Next
W = i

'If Int(w1) > W Then lx = 1

'Loop Until Int(w1) <= W

Datechkstr = Format$(i, "DD-MM-") + LTrim$(Str$(Val(Format$(enow, "yyyy")) + lx + lp))
DateChk = DateValue(Datechkstr)
String1 = "Armistice Day Forerunner for Veterans Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

Next


'------------------------------------------------------------


For lp = -2 To 2
'w1 = DateSerial(Val(Format$(enow, "yyyy") + lx + lp), Month(Now), Day(Now))
lx = 0
'Do
W = DateSerial(Val(Format$(enow, "yyyy") + lx + lp), 3, 31)
For i = W To W - 10 Step -1
gv$ = Format$(i, "DDD")
If gv$ = "Sun" Then Exit For
Next
W = i

'If Int(w1) > W Then lx = 1

'Loop Until Int(w1) <= W

Datechkstr = Format$(i, "DD-MM-") + LTrim$(Str$(Val(Format$(enow, "yyyy")) + lx + lp))
DateChk = DateValue(Datechkstr)
String1 = "British Summer Time Begins - BST"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next

'------------------------------------------------------------

For lp = -2 To 2
For lt = 1 To 11 Step 3
W = DateSerial(Val(Format$(enow, "yyyy") + lx + lp), lt, 1)

DateChk = W
If lt = 1 Then String1 = "-- 1st Quarter Of Year -- "
If lt = 4 Then String1 = "-- 2nd Quarter Of Year --"
If lt = 7 Then String1 = "-- 3rd Quarter Of Year --"
If lt = 10 Then String1 = "-- 4th Quarter Of Year --"

String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next
Next
'r
'------------------------------------------------------------
'w1 = enow

'------------------------------------------------------------
'R = R + 3

W = DateSerial(Year(Now), Month(Now), 1)
W = W - 1

xd = W
For i = W To W - 10 Step -1
gv$ = Format$(i, "DDD")
If gv$ <> "Sun" And gv$ <> "Sat" Then Exit For
Next
W = i

Mls$(R) = Format$(W, "DD-MM-YYYY")
If xd = W Then
    String2 = ""
Else
    String2 = "*Before End Of Month*"
End If

Datechkstr = Format$(W, "DD-MM-YYYY")
DateChk = DateValue(Datechkstr)
String1 = "Pay Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

'------------------------------------------------------------

For lp = -2 To 2
lx = 0

'Do
W = DateSerial(Val(Format$(enow, "yyyy") + lx + lp), 9, 31)


'If w1 > W Then
'W = DateSerial(Val(Format$(enow, "yyyy")) + 1, 9, 31): lx = 1
'End If
For i = W To W - 10 Step -1
gv$ = Format$(i, "DDD")
If gv$ = "Sat" Then Exit For
Next
W = i

'If Int(w1) > W Then lx = 1

'Loop Until Int(w1) <= W

Datechkstr = Format$(i, "DD-MM-") + LTrim$(Str$(Val(Format$(enow, "yyyy")) + lx + lp))
DateChk = DateValue(Datechkstr)
String1 = "Event:- Grand Parents Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)


W = DateSerial(Val(Format$(enow, "yyyy") + lp), 11, 5)
Datechkstr = Format$(W, "DD-MM-YYYY")
DateChk = DateValue(Datechkstr)
String1 = "Event:- Remember Remember 5th November Guy Fawkes FireWorks Night"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)


Next

'------------------------------------------------------------

R = R + 3
iiiy = DateSerial(Year(Now) - 1, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact One Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 2, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact Two Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 3, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact Three Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 4, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact Five Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 5, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact Five Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 10, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact Ten Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 20, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact Twenty Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 50, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact Fifty Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 100, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact 100 Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 200, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact 200 Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 300, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact 300 Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 500, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact 500 Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 1000, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact 1000 Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now) - 1500, Month(Now), Day(Now))
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exact 1500 Years Ago -- "

R = R + 3
iiiy = DateSerial(Year(Now), Month(Now), Day(Now) + 1)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Next Day -- "

R = R + 3
iiiy = DateSerial(100, Month(Now), Day(Now) + 1)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Next Day Year 100 -- -- -- -- -- -- ·-·´¯`·.¸.¸.·´¯`·..¸><(((º> --- (¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)"

R = R + 3
iiiy = DateSerial(Year(Now), Month(Now), Day(Now) + 2)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Second Day -- "

R = R + 3
iiiy = DateSerial(Year(Now), Month(Now), Day(Now) + 3)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Third Day -- "

R = R + 3
iiiy = DateSerial(Year(Now), Month(Now), Day(Now) + 7)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Next Week -- "

R = R + 3
iiiy = DateSerial(Year(Now), Month(Now), Day(Now) + 14)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Next Two Weeks -- "

R = R + 3
iiiy = DateSerial(Year(Now), Month(Now) + 1, 1)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Next Start Month -- "

R = R + 3
iiiy = DateSerial(Year(Now), Month(Now) + 2, 1)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Next Start Two Months -- "

R = R + 3
iiiy = DateSerial(Year(Now), Month(Now) + 3, 1)
Mls$(R) = Format$(iiiy, "DD-MM-yyyy")
Mls$(R + 1) = "-- Exactly The Next Start Three Months -- "


'------------------------------------------------------------


For lp = -2 To 2
lx = 0

W = DateSerial(Val(Format$(enow, "yyyy") + lx + lp), 8, 31)

For i = W To W - 10 Step -1
gv$ = Format$(i, "DDD")
If gv$ = "Mon" Then Exit For
Next
W = i

Datechkstr = Format$(i, "DD-MM-") + LTrim$(Str$(Val(Format$(enow, "yyyy")) + lx + lp))
DateChk = DateValue(Datechkstr)
String1 = "Event:- Late Summer Holiday"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next

'------------------------------------------------------------
For lp = -2 To 2
xp = 0
W = DateSerial(Val(Format$(enow, "yyyy") + lp), 6, 30)
Do
    gv$ = Format$(W, "DDD")
    If gv$ = "Sun" Then xp = xp + 1
    If xp = 2 Then Exit Do
    W = W - 1
Loop Until 1 = 2

DateChk = W
String1 = "Event:- Father's Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next

'------------------------------------------------------------

'------------------------------------------------------------
For lp = -2 To 2
W = DateSerial(Val(Format$(enow, "yyyy") + lp), 4, 30)
Do
    gv$ = Format$(W, "DDD")
    If gv$ = "Sun" Then Exit Do
    W = W - 1
Loop Until 1 = 2

DateChk = W
String1 = "Event:- London Marathon Last Sunday In April"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next

'------------------------------------------------------------

For lp = -2 To 2
W = DateSerial(Val(Format$(enow, "yyyy") + lp), 11, 1)
xp = 0
Do
    gv$ = Format$(W, "DDD")
    If gv$ = "Thu" Then xp = xp + 1
    If xp = 4 Then Exit Do
    W = W + 1
Loop Until 1 = 2

DateChk = W
String1 = "US Thanksgiving Day - Fourth Thurs in Month Nov"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
DateChk = W + 1
String1 = "Buy Nothing Day in US - Day after Thanksgiving."
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next



'#US Thanksgiving Day - 4th Thurs. in month. Buy Nothing Day in US - day after Thanksgiving.

'------------------------------------------------------------
For lx = -2 To 2
W = DateSerial(Val(Format$(enow, "yyyy") + lx), 2, 14)

DateChk = W
String1 = "Event:- St Valentine's Day"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next

'------------------------------------------------------------

R = R + 3
lx = 0

Do
W = DateSerial(Val(Format$(enow, "yyyy") + lx), 5, 1)

For i = W To W + 10
gv$ = Format$(i, "DDD")
If gv$ = "Mon" Then Exit For
Next
W = i

If Int(w1) > W Then lx = 1

Loop Until Int(w1) <= W

Mls$(R) = Format$(i, "DD-MM-") + LTrim$(Str$(Val(Format$(enow, "yyyy")) + lx))
Mls$(R + 1) = "Event:- May Bank Holiday"
'------------------------------------------------------------
'Mls$(R + 1) = "Start of Summer Bank Holiday"
'only on Ireland first monday June
'------------------------------------------------------------



'------------------------------------------------------------
For lp = -2 To 2
lx = 0
W = DateSerial(Val(Format$(enow, "yyyy") + lx + lp), 12, 1)
Do
    gv$ = Format$(W, "DDD")
    If gv$ = "Sun" Then Exit Do
    W = W - 1
Loop Until 1 = 2

DateChk = W
String1 = "Event:- First Sunday In Advent"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)
Next




'-----------------------------------------------------------------------------------------------------------

'Dim X1, X2, x3 As Date, x4, x5 As Date
ReDim Tagger(50)
ReDim XvRs1(50)
ReDim XvRs2(50)

kid = 0
'W = enow
'W9 = enow
W = DateSerial(1990, 1, 1)
W9 = DateSerial(1990, 1, 1)
w7 = Val(Format$(W, "YYYY"))
w5 = Val(Format$(W, "MM"))
w6 = Val(Format$(W, "DD"))
yuk = w5
huk = w7
'For i = 1 To 8000
Do
iy = 0
w7 = Val(Format$(W, "YYYY"))
w5 = Val(Format$(W, "MM"))
w6 = 13
gf = DateSerial(w7, w5, w6)
If gf >= W9 Then
gv$ = Format$(W, "DDD")
gv2$ = Format$(W, "dd")

'fri 13

If gv$ = "Fri" And gv2$ = "13" Then

wclops = W

'TheDate = ConvertToUT(wclops)

'wclops = TheDate


'X1 = NextNewMoon(wclops)
'X2 = NextFirstQuarter(wclops)
x3 = NextFullMoon(wclops)

'If wclops = Int(x3) Then fullmoon2 = 1
If Format$(x3, "DD/MM/YYYY") = Format$(wclops, "DD/MM/YYYY") Then fullmoon2 = 1

If gv$ = "Fri" And gv2$ = "13" And fullmoon2 = 1 Then
fullmoon2 = 0

DateChk = wclops
String1 = "Friday The 13th On a Full Moon"
String2 = ""
FormatDate$ = "DDD-DD-MMM-YYYY"
Call NextLastAnMiddCode(DateChk)

kid = 1
End If
End If
End If
yuk = yuk + 1
If yuk = 13 Then huk = huk + 1: yuk = 1

W = DateSerial(huk, yuk, 13)

Loop Until W > DateSerial(2020, 1, 1)









'------------------------------------------------------------
ReDim Tagger(50)
ReDim XvRs1(50)
ReDim XvRs2(50)

kid = 0
W = DateSerial(Year(Now) - 1, 1, 1)
W9 = DateSerial(Year(Now) - 1, 1, 1)
w7 = Val(Format$(W, "YYYY"))
w5 = Val(Format$(W, "MM"))
w6 = Val(Format$(W, "DD"))
yuk = w5
huk = w7
Do
    iy = 0
    w7 = Val(Format$(W, "YYYY"))
    w5 = Val(Format$(W, "MM"))
    w6 = 13
    gf = DateSerial(w7, w5, w6)
    If gf >= W9 Then
        gv$ = Format$(W, "DDD")
        gv2$ = Format$(W, "dd")

        If gv$ = "Fri" And gv2$ = "13" Then

            DateChk = W
            String1 = "Friday The 13th"
            String2 = ""
            FormatDate$ = "DDD-DD-MMM-YYYY"
            Call NextLastAnMiddCode(DateChk)

            kid = 1
        End If
    End If
    yuk = yuk + 1
    If yuk = 13 Then huk = huk + 1: yuk = 1
    W = DateSerial(huk, yuk, 13)
Loop Until W > DateSerial(Year(Now) + 2, 1, 1)


'------------------------------------------------------------
'------------------------------------------------------------


kid = 0
W = enow
W9 = enow
w7 = Val(Format$(W, "YYYY"))
w5 = Val(Format$(W, "MM"))
w6 = Val(Format$(W, "DD"))
huk = w7
For i = 1 To 8000

    w7 = Val(Format$(W, "YYYY"))
    w8 = Val(Format$(W, "YY"))
    w5 = 6
    w6 = 6
    W = DateSerial(w7, w5, w6)
    If W >= W9 And (w8 <> 6 And w8 <> 66) Then

        wclops = W

        x3 = NextFullMoon(wclops)

        If Format$(x3, "DD/MM/YYYY") = Format$(wclops, "DD/MM/YYYY") Then fullmoon2 = 1

        If fullmoon2 = 1 Then
            fullmoon2 = 0

            R = R + 3
            abc = R Mod 3: If abc <> 1 Then MsgBox "R out of sequence ": Stop
            Mls$(R) = Format$(wclops, "DD-MM-yyyy")

            If kid = 0 Then Mls$(R + 1) = ">>-Next 666 On a Full Moon at: " + Format$(x3, "HH:MM:SS")
            If kid = 1 Then Mls$(R + 1) = ">>-666 On a Full Moon at: " + Format$(x3, "HH:MM:SS")

            kid = 1
        End If
    End If
    huk = huk + 1

    If huk >= 10000 Then Exit For
    W = DateSerial(huk, 6, 6)
Next


wclops = DateSerial(6666, 6, 6)
x3 = NextFullMoon(wclops)
R = R + 3
Mls$(R) = Format$(wclops, "DD-MM-yyyy")
Mls$(R + 1) = ">>--Next 6666 On a Full Moon at: " + Format$(x3, "HH:MM:SS")

'------------------------------------------------------------
'------------------------------------------------------------

W = enow
W9 = enow
w7 = Val(Format$(W, "YYYY"))
w5 = Val(Format$(W, "MM"))
w6 = Val(Format$(W, "DD"))
huk = w7
For i = 1 To 8000

    w7 = Val(Format$(W, "YYYY"))
    w8 = Val(Format$(W, "YY"))
    w5 = 6
    w6 = 6
    W = DateSerial(w7, w5, w6)
    If W >= W9 And (w8 = 6 Or w8 = 66) Then

    wclops = W
    x3 = NextFullMoon(wclops)
    
    If Format$(x3, "DD/MM/YYYY") = Format$(wclops, "DD/MM/YYYY") Then fullmoon2 = 1
    If Format$(x3, "DD/MM/YYYY") = Format$(DateSerial(6666, 6, 6), "DD/MM/YYYY") Then fullmoon2 = 0

    If fullmoon2 = 1 Then
        fullmoon2 = 0

        R = R + 3
        Mls$(R) = Format$(wclops, "DD-MM-yyyy")
        If xzag = 0 Then
                Mls$(R + 1) = ">>-Next '6 or 66' 666 On a Full Moon at: " + Format$(x3, "HH:MM:SS")
                xzag = 1
            Else
                Mls$(R + 1) = ">>-'6 or 66' 666 On a Full Moon at: " + Format$(x3, "HH:MM:SS")
            End If
        End If
    End If
    huk = huk + 1

    If huk >= 10000 Then Exit For
    W = DateSerial(huk, 6, 6)
Next

R = R + 3
Mls$(R) = Format$(DateSerial(Year(Now), 6, 6), "DD-MM-yyyy")
Mls$(R + 1) = ">>-This 666 ----"

'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------

bij = 0
W = enow
W9 = enow
w7 = Val(Format$(W, "YYYY"))
w5 = Val(Format$(W, "MM"))
w6 = Val(Format$(W, "DD"))
huk = w7
For i = 1 To 2600

    w7 = Val(Format$(W, "YYYY"))
    w5 = 6
    w6 = 6
    W = DateSerial(w7, w5, w6)
    If W <= W9 Then

        wclops = W

        x3 = NextFullMoon(wclops)
        If Format$(x3, "DD/MM/YYYY") = Format$(wclops, "DD/MM/YYYY") Then fullmoon2 = 1

        If fullmoon2 = 1 Then
            fullmoon2 = 0
            R = R + 3
            Mls$(R) = Format$(wclops, "DD-MM-yyyy")
            Mls$(R + 1) = ">>-Last 666 On a Full Moon at: " + Format$(x3, "HH:MM:SS")
            bij = 1
        End If
    End If

    huk = huk - 1

    W = DateSerial(huk, 6, 6)
    If bij = 1 Then Exit For
Next


'------------------------------------------------------------
'------------------------------------------------------------
bij = 0
W = enow
W9 = enow
w7 = Val(Format$(W, "YYYY"))
w5 = Val(Format$(W, "MM"))
w6 = Val(Format$(W, "DD"))
huk = w7
For i = 1 To 2600

    w7 = Val(Format$(W, "YYYY"))
    w8 = Val(Format$(W, "YY"))
    w5 = 6
    w6 = 6
    W = DateSerial(w7, w5, w6)
    If W <= W9 And (w8 = 66 Or w8 = 6) Then

        wclops = W

        x3 = NextFullMoon(wclops)
        If Format$(x3, "DD/MM/YYYY") = Format$(wclops, "DD/MM/YYYY") Then fullmoon2 = 1

        If fullmoon2 = 1 Then
            fullmoon2 = 0

            R = R + 3
            Mls$(R) = Format$(wclops, "DD-MM-yyyy")
            Mls$(R + 1) = ">>-Last '6 or 66' 666 On a Full Moon at: " + Format$(x3, "HH:MM:SS")
        End If
    End If

    huk = huk - 1

    W = DateSerial(huk, 6, 6)
    If huk < 30 Then Exit For
Next




'------------------------------------------------------------
ReDim Tagger(50)
ReDim Tagge2(50)
ReDim XvRs1(50)
ReDim XvRs2(50)
ReDim XvRs4(50)
ReDim XvRs5(50)


'MInfo2$(6) = "Lunar Eclipse"
MInfo2$(1) = "Solar Eclipse"
MInfo2$(2) = "T -Total"
MInfo2$(3) = "A -Annular"
MInfo2$(4) = "H -Hybrid(Annular / Total)"
MInfo2$(5) = "P -Partial"
MInfo2$(6) = "Lunar Eclipse"
MInfo2$(7) = "t -Total(Umbral)"
MInfo2$(8) = "p -Partial(Umbral)"
MInfo2$(9) = "n -Penumbral"


Dim wfh As Date
questart = R

bfree = FreeFile
Open App.Path + "\00 Data\Eclipse.txt" For Input As #bfree
Do
Line Input #bfree, as2$
If Trim$(as2$) = "" Then valyear = 0: valyear2 = 0
If valyear = 0 Then valyear = Val(Mid$(as2$, 1, 8))
If valyear >= 2000 Then valyear2 = valyear
If valyear2 >= Year(Now) - 3 Then
For ti = 0 To 3
If ti = 0 Then cg = 0
If ti = 1 Then cg = 18
If ti = 2 Then cg = 36
If ti = 3 Then cg = 53
as3$ = Mid$(as2$, 9 + cg, 15)
as4$ = Mid$(as2$, 23 + cg, 1)
If Trim$(as4$) <> "" Then
wfh = DateValue(Mid$(as3$, 1, 6) + "," + Str$(valyear2))
wfh = wfh + TimeValue(Mid$(as3$, 9, 5))
wfh2 = DateValue(Mid$(as3$, 1, 6) + "," + Str$(valyear2))
wfh2 = wfh2 + TimeValue(Mid$(as3$, 9, 5))
For ry7 = 2 To 9
If ry7 = 6 Then ry7 = 7
If Mid$(MInfo2$(ry7), 1, 1) = as4$ Then
egg = 0
If valyear2 < Year(Now) + 8 Then egg = 1
If as4$ = "T" Or as4 = "t" Then egg = 1
If egg = 1 Then
master$ = ""
If ry7 > 6 Then master$ = MInfo2$(6) Else master$ = MInfo2$(1)
master$ = master$ + Mid$(MInfo2$(ry7), 2)
master$ = master$ + Format$(wfh, " DD-MM-YYYY HH:MM:SSa/p") + " GMT"
iy = 0

R = R + 3
abc = R Mod 3: If abc <> 1 Then MsgBox "R out of sequence ": Stop
Mls$(R) = Format$(wfh, "DD-MM-yyyy")

'MInfo2$(6) = "Lunar Eclipse"
'MInfo2$(1) = "Solar Eclipse"

dizzy$ = "# " + master$
iy = iy + 1
If InStr(dizzy$, "Lunar Eclipse") > 0 Then
    If wfh > Now And Tagger(iy) = 0 Then
        Mls$(R + 1) = "--- Next " + dizzy$
        Tagger(iy) = 1
        Mls$(XvRs1(iy)) = "--- Last " + dizzy2$
    Else
        XvRs1(iy) = R + 1: XvRs2(iy) = wfh: dizzy2$ = dizzy$
        Mls$(R + 1) = "--- " + dizzy$
    End If
End If
    

If InStr(dizzy$, "Solar Eclipse") > 0 Then
    If wfh > Now And Tagge2(iy) = 0 Then
        Mls$(R + 1) = "--- Next " + dizzy$
        Tagge2(iy) = 1
        Mls$(XvRs4(iy)) = "--- Last " + dizzy3$
    Else
        XvRs4(iy) = R + 1: XvRs5(iy) = wfh: dizzy3$ = dizzy$
        Mls$(R + 1) = "--- " + dizzy$
    End If
End If


End If
End If
Next
End If
Next
End If
Loop Until EOF(bfree)
Close #bfree










'------------------------------------------------------------

Dim ftg2 As Date

R = R + 3
abc = R Mod 3: If abc <> 1 Then MsgBox "R out of sequence ": Stop

peg1 = R
Open App.Path + "\00 Data\2000Dat2 Celeb Bday Events.txt" For Input As #1
Do
Line Input #1, A$
ftg2 = DateSerial(Year(Now), Val(Mid$(A$, 4, 2)), Val(Mid$(A$, 1, 2)))
If ftg2 >= Int(Now) And ftg2 < Int(Now + 2) Then
Mls$(R) = Mid$(A$, 1, 10)
Mls$(R + 1) = Mid$(A$, 12)
Mls$(R + 2) = c$
R = R + 3
End If
Loop Until EOF(1)
Close #1

Open App.Path + "\00 Data\1991Dat5 BBC.txt" For Input As #1
Do
Line Input #1, A$
ftg2 = DateSerial(Year(Now), Val(Mid$(A$, 4, 2)), Val(Mid$(A$, 1, 2)))
If ftg2 >= Int(Now) And ftg2 < Int(Now + 2) Then
Mls$(R) = Mid$(A$, 1, 10)
Mls$(R + 1) = Mid$(A$, 12)
Mls$(R + 2) = c$
R = R + 3
abc = R Mod 3: If abc <> 1 Then MsgBox "R out of sequence ": Stop

End If
If EOF(1) Then Exit Do
Loop Until EOF(1)
Close #1

Open App.Path + "\00 Data\All in One Dates 13271.txt" For Input As #1
Do
Line Input #1, A$
ftg2 = DateSerial(Year(Now), Val(Mid$(A$, 4, 2)), Val(Mid$(A$, 1, 2)))
If ftg2 >= Int(Now) And ftg2 < Int(Now + 1) Then
Mls$(R) = Mid$(A$, 1, 10)
Mls$(R + 1) = Mid$(A$, 12)
Mls$(R + 2) = c$
R = R + 3
End If
If EOF(1) Then Exit Do
Loop Until EOF(1)
Close #1

ReDim Preserve Mls$(R - 3)


'------------------------------------------------------------
u = 1
Dim Geek As Long
Dim MattGeek$
Do
A$ = Mls$(u)
B$ = Mls$(u + 1)
If Trim(B$) = "" Then
    Stop
    MsgBox "You Got an Empty Field"
End If
If Trim(A$) = "" Then
    Stop
    MsgBox "You Got an Empty Field"
End If
MattGeek$ = Format$(Val(Mid$(A$, 7, 4)), "0000") + "-" + Mid$(Mls$(u), 4, 2) + "-" + Mid$(Mls$(u), 1, 2)
'If Val(Mid$(A$, 7, 4)) < 1800 Then Stop
Geek = DateSerial(2000, Val(Mid$(Mls$(u), 4, 2)), Val(Mid$(Mls$(u), 1, 2)))
If Year(DateValue(A$)) > Year(Now) Then
Geek = DateSerial(Year(DateValue(A$)), Val(Mid$(Mls$(u), 4, 2)), Val(Mid$(Mls$(u), 1, 2)))
'If Year(DateValue(A$)) = 2032 Then Stop
End If
Geek2$ = Format(Geek, "0000000")
List3.AddItem (Geek2$ + "-" + A$ + "-" + B$)

List5.AddItem (MattGeek$ + "-" + A$ + "-" + B$)

u = u + 3
Loop Until u >= R - 3

List3.Refresh
List5.Refresh
DoEvents
On Error GoTo ErrSub2
xt = 0

fr3 = FreeFile
Open "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Date1991 List Future.txt" For Output As #fr3

For rt = 0 To List5.ListCount - 1
    bv$ = Mid$(List5.List(rt), 12)
    xtz = 0
    gdate = 0
    If Val(Mid$(bv$, 7, 4)) = Year(Now) Then
        'gdate = DateValue(Mid$(bv$, 1, 6) + Trim(Str(Year(Now))))
        gdate = DateValue(Mid$(bv$, 1, 10))
    End If
    gdate2 = DateValue(Mid$(bv$, 1, 10))
    If gdate + 30 > Now And gdate2 < DateSerial(Year(Now) + 3, 1, 1) Then
        List7.AddItem Format$(DateValue(Mid$(bv$, 1, 6) + Trim(Str(Year(Now)))), "YYYY-MM-DD") + "-" + Mid$(bv$, 1, 6) + Trim(Str(Year(Now))) + "--" + Mid$(bv$, 1, 10) + "--" + Mid$(bv$, 11)
    End If
Next

'2nd---------
For rt = 0 To List5.ListCount - 1
    bv$ = Mid$(List5.List(rt), 12)
    xtz = 0
    If Val(Mid$(bv$, 7, 4)) = Year(Now) Then
    If Mid$(bv$, 1, 6) + Trim(Str(Year(Now) + 1)) <> "29-02-2009" Then
        gdate = DateValue(Mid$(bv$, 1, 6) + Trim(Str(Year(Now) + 1)))
    Else
        gdate = DateValue("28-02-2009")
    End If
    
    gdate2 = DateValue(Mid$(bv$, 1, 10))
    If gdate < Now + 150 And gdate2 < DateSerial(Year(Now) + 3, 1, 1) Then
        If gdate = DateValue("28-02-2009") Then
        List7.AddItem Format$(gdate, "YYYY-MM-DD") + "-" + Mid$(bv$, 1, 6) + Trim(Str(Year(Now))) + "--" + Mid$(bv$, 1, 10) + "----" + "Xx Leap Year Problem xX -" + Mid$(bv$, 11)
        Else
        
        List7.AddItem Format$(DateValue(Mid$(bv$, 1, 6) + Trim(Str(Year(Now) + 1))), "YYYY-MM-DD") + "-" + Mid$(bv$, 1, 6) + Trim(Str(Year(Now))) + "--" + Mid$(bv$, 1, 10) + "--" + Mid$(bv$, 11)
        End If
    End If
    End If
Next


dash$ = String$(40, "-")
dash$ = ""
break11$ = vbCrLf + "#---- Last to Next ----" + Format$(Now, "DDD DD-MMM-YY") + vbCrLf

dash$ = String$(32, "-")
break5$ = vbCrLf + dash$ + vbCrLf + "---------- Today ----------" + Format$(Now, "DDD DD-MMM-YY") + vbCrLf + dash$ + vbCrLf

dash$ = String$(36, "-")
break7$ = vbCrLf + dash$ + vbCrLf + "---------- Next Day ----------" + vbCrLf + dash$ + vbCrLf

dash$ = String$(39, "-")
break8$ = vbCrLf + dash$ + vbCrLf + "---------- End Month ----------" + vbCrLf + dash$ + vbCrLf

dash$ = String$(42, "-")
break9$ = vbCrLf + dash$ + vbCrLf + "---------- Begin Month ----------" + vbCrLf + dash$ + vbCrLf

dash$ = String$(33, "-")
break10$ = vbCrLf + dash$ + vbCrLf + "---------- Month ----------" + vbCrLf + dash$ + vbCrLf


DoEvents
For rt = 0 To List7.ListCount - 1
    If lastrdate = 0 Then lastrdate = rdate
    bv$ = List7.List(rt)
    rdate = DateValue(Mid$(bv$, 1, 10))
    If Month(rdate) > Month(lastrdate) And xt = 0 Then xt = 1: Print #fr3, break9$
    If rdate > Now - 1 And xt = 1 Then xt = 2: Print #fr3, break5$
    If rdate > Now And xt = 2 Then xt = 3: Print #fr3, break7$
    If Month(rdate) <> Month(lastrdate) And xt = 4 Then Print #fr3, break10$
    If Month(rdate) > Month(Now - 1) And xt = 3 Then xt = 4: Print #fr3, break8$
    lastrdate = rdate
    bv$ = Format$(rdate, "DDD ") + Mid$(List7.List(rt), 12)
    Print #fr3, bv$
Next

Close #fr3


'Shell "C:\Program Files (x86)\Notepad++\notepad++.exe D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Date1991 List Future.txt", vbNormalFocus
'End

Kill "D:\TEMP\BlueMoon.txt"


FR1 = FreeFile
Open "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Date1991 List.txt" For Output As #FR1
fr2 = FreeFile

fr8 = FreeFile
Open "D:\TEMP\BlueMoon.txt" For Output As #fr8
Close fr8

Open "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Date1991 Next & Last.txt" For Output As #fr2
On Error GoTo 0


'find last most recent blue moon
For rt = 0 To List5.ListCount - 1
    bv$ = Mid$(List5.List(rt), 12)
    
    If InStr(bv$, "Blue Moon") > 0 And InStr(List5.List(rt), "Last") > 0 Then
            easydoes1 = DateValue(Mid$(bv$, 1, 10))
            'If DateDiff("d", DateValue(Mid$(bv$, 1, 10)), Now) > 40 Then easydoes = 0
    End If
    If InStr(bv$, "Harvest") > 0 And InStr(List5.List(rt), "Last") > 0 Then
            easydoes2 = DateValue(Mid$(bv$, 1, 10))
            'If DateDiff("d", DateValue(Mid$(bv$, 1, 10)), Now) > 40 Then easydoes = 0
    End If

Next


xt = 0: xt2 = 0

For rt = 0 To List5.ListCount - 1
    bv$ = Mid$(List5.List(rt), 12)
    If InStr(LCase(bv$), "acid") > 0 Then
    A = A
    End If
    
    cue = DateValue(Mid$(bv$, 1, 10))
    If Val(Mid$(bv$, 7, 4)) >= Year(Now) Then cue = DateValue(Mid$(bv$, 1, 6) + Trim(Str(Year(Now))))
    If DateValue(Mid$(bv$, 1, 10)) >= Int(Now) + 1 And xt = 0 Then xt = 1
    'If DateValue(Mid$(bv$, 1, 10)) And xt2 = 0 Then xt2 = 1
    'day22 = Day(Mid$(List5.List(rt), 1, 7))
    'month22 = Month(Mid$(List5.List(rt), 1, 7))
    If xt = 1 Then
        Print #FR1, break11$
        xt = 2
    End If
    
    If InStr(KK$, Mid$(bv$, 12)) = 0 Then
        'Stop
        Print #FR1, Format$(cue, "DDD ") + bv$
'        Print #fr3, DateValue(Mid$(bv$, 1, 6) + Str(Year(Now))) + "--" + Mid$(bv$, 1, 10) + "--" + Mid$(bv$, 11)
    End If
    

'--------------------

If Mid$(bv$, 11, 4) = "->>-" Or Mid$(bv$, 11, 4) = "----" Then nvq = 1 Else nvq = 0
tth1 = InStr(bv$, "Last ")
tth2 = InStr(bv$, "Next ")

If otth1 > 0 And tth1 = 0 And tth2 > 0 And tth1 < 34 Then
    Print #fr2, break11$
End If

If tth1 < 34 Then otth1 = tth1

'Last's
If tth1 > 0 And tth1 < 18 And nvq = 1 Then
    cc1 = cc1 + 1
    If cc1 > UBound(Rvm1) Then
        ReDim Preserve Rvm1(UBound(rmv1) + 1000)
        ReDim Preserve Rvm2(UBound(rmv1) + 1000)
    End If
    day1$ = Format$(DateValue(Mid$(bv$, 1, 10)), "DDD") + "-"
    Print #fr2, Format$(cc1, "00 --- ") + day1$ + bv$
    Rvm1(cc1) = day1$ + bv$

    
    If InStr(bv$, "Eclipse") > 0 Then
        fr8 = FreeFile
        Open "D:\TEMP\Eclipse.txt" For Output As #fr8
            
            For R = 2 To 9
                If R = 6 Then R = 7
                ii = InStr(bv$, Mid$(MInfo2$(R) + " ", 3))
                If ii > 0 Then ii = ii + Len(Mid$(MInfo2$(R), 3)) - 1: Exit For
            Next
            
            
            tth = Mid$(bv$, 1, ii)
            tth = Mid$(tth, InStr(bv$, "Last #"))
            tthx = DateValue(Mid$(bv$, ii + 2, 20)) + TimeValue(Mid$(bv$, ii + 2, 20))
            Print #fr8, Format(tthx, "DD-MMM-YY HH:MM:SSa/p") + " GMT - " + tth
        Close #fr8
    End If
    
    
    If InStr(bv$, "Cross-Q") > 0 Then
        fr8 = FreeFile
        Open "D:\TEMP\Cross Quarter.txt" For Output As #fr8
            tth = Mid$(bv$, 1, InStr(bv$, "Begin") + 4)
            tth = Mid$(tth, InStr(tth, "t --") + 5)
            tth = Mid$(tth, 1, InStr(tth, " ") - 1) + " - " + Mid$(tth, InStr(tth, " ") + 4) + "s"
            Print #fr8, Format(DateValue(Mid$(bv$, 1, 10)), "DD-MMM-YY") + " - " + tth
        Close #fr8
    End If
    
    If (InStr(bv$, "Equinox") > 0 Or InStr(bv$, "Solstice ----") > 0) And bv$ <> "Moon" Then
        fr8 = FreeFile
        Open "D:\TEMP\Equinox.txt" For Output As #fr8
            tth = Mid$(bv$, 1, InStr(bv$, " ---- @") - 1)
            tth = Mid$(tth, InStr(tth, "t --") + 5)
            Print #fr8, Format(DateValue(Mid$(bv$, 1, 10)), "DD-MMM-YY") + " - " + tth
        Close #fr8
    End If

    

    'THIS GETS THE HARVEST MOON & Blue MOONS
    'MOON
    ACE = DateValue(Mid$(bv$, 1, 10))
    xo = 0
    If InStr(bv$, "Blue Moon") > 0 Then xo = 1
    If InStr(bv$, "New Moon") > 0 Then xo = 1
    If InStr(bv$, "Full Moon") > 0 Then xo = 1
    If ACE = easydoes1 Then xo = 1
    If ACE = easydoes2 Then
    xo = 1
    End If
    
    If InStr(bv$, "Black") > 0 Then xo = 0
    If InStr(bv$, "13th") > 0 Then xo = 0
    If InStr(bv$, "666") > 0 Then xo = 0
    If InStr(bv$, "Easter") > 0 Then xo = 0
    If InStr(bv$, "Calendar") > 0 Then xo = 0
    If DateDiff("d", DateValue(Mid$(bv$, 1, 10)), Now) > 70 Then xo = 0
    
    If xo = 1 Then
        'If DateDiff("d", DateValue(Mid$(bv$, 1, 10)), Now) < 32 Then
        'If InStr(bv$, "Moon") > 0 And DateDiff("d", DateValue(Mid$(bv$, 1, 10)), Now) > easydoes Then
        
        fr8 = FreeFile
        Open "D:\TEMP\BlueMoon.txt" For Append As #fr8
            If InStr(bv$, " ---- @") > 0 Then
                tth = Mid$(bv$, 1, InStr(bv$, " ---- @") - 1)
                tth = Mid$(tth, InStr(tth, "t --") + 5)
            
                If InStr(tth, " *") > 0 Then tth = Mid$(tth, 1, InStr(tth, " *") - 1)
            
                'tth2 = Mid$(bv$, InStr(bv$, " --- - ") + 6)
                'tthx = DateValue(Mid$(bv$, ii + 2, 20)) + TimeValue(Mid$(bv$, ii + 2, 20))
            
                Print #fr8, Format(DateValue(Mid$(bv$, 1, 10)), "DD-MMM-YY") + " - " + tth
            End If
        Close #fr8
        End If
    'End If
    End If
    




'Next's
If tth2 > 0 And tth2 < 18 And nvq = 1 Then
    cc2 = cc2 + 1
    If cc2 > UBound(Rvm2) Then
        ReDim Preserve Rvm1(UBound(rmv1) + 1000)
        ReDim Preserve Rvm2(UBound(rmv1) + 1000)
    End If
    day1$ = Format$(DateValue(Mid$(bv$, 1, 10)), "DDD") + "-"
    Print #fr2, Format$(cc2, "00 --- ") + day1$ + bv$
    
    If InStr(bv$, "Cross-Q") > 0 And halo = "" Then
        halo = bv$
        fr8 = FreeFile
        Open "D:\TEMP\Cross Quarter.txt" For Append As #fr8
            tth = Mid$(bv$, 1, InStr(bv$, "Begin") + 4)
            tth = Mid$(tth, InStr(tth, "t --") + 5)
            tth = Mid$(tth, 1, InStr(tth, " ") - 1) + " - " + Mid$(tth, InStr(tth, " ") + 4) + "s"
            
            Print #fr8, Format(DateValue(Mid$(bv$, 1, 10)), "DD-MMM-YY") + " - " + tth
            
        Close #fr8
    End If
    
    If (InStr(bv$, "Equinox") > 0 Or InStr(bv$, "Solstice ----") > 0) And halo2 = "" Then
        halo2 = bv$
        fr8 = FreeFile
        Open "D:\TEMP\Equinox.txt" For Append As #fr8
            tth = Mid$(bv$, 1, InStr(bv$, " ---- @") - 1)
            tth = Mid$(tth, InStr(tth, "t --") + 5)
            'tth = Mid$(tth, 1, InStr(tth, " ") - 1) + " - " + Mid$(tth, InStr(tth, " ") + 4) + "s"
            
            Print #fr8, Format(DateValue(Mid$(bv$, 1, 10)), "DD-MMM-YY") + " - " + tth
            
        Close #fr8
    End If
    
    If InStr(bv$, "Eclipse") > 0 And halo3 = "" Then
        halo3 = bv$
        fr8 = FreeFile
        Open "D:\TEMP\Eclipse.txt" For Append As #fr8
            For R = 2 To 9
                If R = 6 Then R = 7
                ii = InStr(bv$, Mid$(MInfo2$(R) + " ", 3))
                If ii > 0 Then ii = ii + Len(Mid$(MInfo2$(R), 3)) - 1: Exit For
            Next
            
            
            
            tth = Mid$(bv$, 1, ii)
            tth = Mid$(tth, InStr(bv$, "Next #"))
            tthx = DateValue(Mid$(bv$, ii + 2, 20)) + TimeValue(Mid$(bv$, ii + 2, 20))
            Print #fr8, Format(tthx, "DD-MMM-YY HH:MM:SSa/p") + " GMT - " + tth
        Close #fr8
    End If
    
    
    If InStr(bv$, "Blue Moon") > 0 Then
        fr8 = FreeFile
        Open "D:\TEMP\BlueMoon.txt" For Append As #fr8
            tth = Mid$(bv$, 1, InStr(bv$, " ---- @") - 1)
            tth = Mid$(tth, InStr(tth, "t --") + 5)
            If InStr(bv$, " --- ") > 0 Then
            tth2 = Mid$(bv$, InStr(bv$, " --- ") + 4)
            End If
            If InStr(bv$, " --- - ") > 0 Then tth2 = Mid$(bv$, InStr(bv$, " --- - ") + 6)
            
            Print #fr8, Format(DateValue(Mid$(bv$, 1, 10)), "DD-MMM-YY") + " - " + tth + " -" + tth2
        Close #fr8
    End If
    
    'MOON
    If InStr(bv$, "Moon") > 0 And InStr(bv$, "Blue Moon") = 0 Then
        If DateDiff("d", DateValue(Mid$(bv$, 1, 10)), Now) > 40 Then
        
        fr8 = FreeFile
        Open "D:\TEMP\BlueMoon.txt" For Append As #fr8
            tth = Mid$(bv$, 1, InStr(bv$, " ---- @") - 1)
            tth = Mid$(tth, InStr(tth, "t --") + 5)
            
            If InStr(tth, " *") > 0 Then tth = Mid$(tth, 1, InStr(tth, " *") - 1)
            
            'tth2 = Mid$(bv$, InStr(bv$, " --- - ") + 6)
            'tthx = DateValue(Mid$(bv$, ii + 2, 20)) + TimeValue(Mid$(bv$, ii + 2, 20))
            
            Print #fr8, Format(DateValue(Mid$(bv$, 1, 10)), "DD-MMM-YY") + " - " + tth
        Close #fr8
        End If
        xxv = 1
    End If
    
    Rvm2(cc2) = day1$ + bv$
End If




Next
Close #FR1, #fr2

'End '-endhere

ReDim Preserve Rvm1(cc1)
ReDim Preserve Rvm2(cc2)


NowNextData = NowNextData + "-- Last & Next" + vbCrLf
For R = UBound(Rvm1) - 2 To UBound(Rvm1)
    NowNextData = NowNextData + Rvm1(R) + vbCrLf
Next

NowNextData = NowNextData + "---" + vbCrLf

For R = 1 To 4
    NowNextData = NowNextData + Rvm2(R) + vbCrLf
Next

NowNextData = NowNextData + "--"

'MsgBox NowNextData

freef3 = FreeFile
Open App.Path + "\Date1991 Small Day.txt" For Input As #freef3
freef4 = FreeFile
If Dir("D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures\Signature-On This Day.txt") <> "" Then
    Open "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures\Signature-On This Day.txt" For Output As #freef4
    Print #freef4, NowNextData
    Do
    On Error Resume Next
    Line Input #freef3, lr$
    If Err.Number = 0 Then Print #freef4, lr$
    On Error GoTo 0
    Loop Until EOF(freef3)
    Close #freef3, #freef4
End If

a1$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Date1991 Next & Last.txt"
b1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Date1991 Next & Last.txt"
Set Fs22 = CreateObject("Scripting.FileSystemObject")
Fs22.copyFile a1$, b1$

List5.Clear

For rt = 0 To List3.ListCount - 1
    bv$ = Mid$(List3.List(rt), 1, 7)
    day22 = Day(Mid$(List3.List(rt), 1, 7))
    month22 = Month(Mid$(List3.List(rt), 1, 7))
    geek22 = DateSerial(Year(Now), month22, day22)
    If geek22 >= Int(Now) Then Exit For
Next

For rt2 = rt To List3.ListCount - 1
    totdo3 = totdo3 + 1
    bv$ = Mid$(List3.List(rt2), 1, 7)
    If Year(Val(bv$)) > Year(Now) Then yearstoptemp = rt2: Exit For
Next

totdo = (totdo3) + (rt) + (List3.ListCount - yearstoptemp)
PBar1.Max = totdo
totdo2 = 0

For rt2 = rt To List3.ListCount - 1
    totdo2 = totdo2 + 1
    PBar1.Value = totdo2
    Label5 = Str(totdo2) + "  of " + Str(totdo)
    bv$ = Mid$(List3.List(rt2), 1, 7)
    If Year(Val(bv$)) > Year(Now) Then yearstop = rt2: Exit For
    
    ll1$ = Mid$(List3.List(rt2), 9, 10)
    ll2$ = Mid$(List3.List(rt2), 20)
    day22 = Day(Mid$(List3.List(rt2), 1, 7))
    month22 = Month(Mid$(List3.List(rt2), 1, 7))
    geek22 = DateSerial(Year(Now), month22, day22)
    
    'If Diar > 0 And geek22 > Diar Then Exit For
    Call Pop1
    DoEvents

Next

For rt2 = 0 To rt - 1
    totdo2 = totdo2 + 1
    PBar1.Value = totdo2
    Label5 = Str(totdo2) + "  of " + Str(totdo)
    bv$ = Mid$(List3.List(rt2), 1, 7)
    'If Year(Val(bv$)) > Year(Now) Then Exit For
    
    ll1$ = Mid$(List3.List(rt2), 9, 10)
    ll2$ = Mid$(List3.List(rt2), 20)
    
    Call Pop1
    DoEvents

Next

A = A

For rt2 = yearstop To List3.ListCount - 1
    totdo2 = totdo2 + 1
    If totdo2 > PBar1.Max Then
        PBar1.Max = totdo2
    End If
    PBar1.Value = totdo2
    Label5 = Str(totdo2) + "  of " + Str(totdo)
    bv$ = Mid$(List3.List(rt2), 1, 7)
    'If Year(Val(bv$)) > Year(Now) Then Exit For
    
    ll1$ = Mid$(List3.List(rt2), 9, 10)
    ll2$ = Mid$(List3.List(rt2), 20)
        
    Call Pop1
    DoEvents

Next


If star2$ <> "" Then lkj2$ = star2$

Label2.Caption = "Current sign no." + lkj2$
rs8$ = Label2.Caption
rs9 = 1

If 1 = 2 Then
Open "D:\# MY DOCS\# 01 My Documents\CALLerid\2diary.txt" For Append As #1
Print #1, String$(80, "-")
For R = 0 To List4.ListCount - 1
Print #1, List4.List(R)
Next
Print #1, String$(80, "-")
Close #1
'Shell "c:\bat\t.bat C:\callerid\2diary.txt", vbNormalFocus
End If

rs5$ = Label1.Caption
rs6 = 1

qw5 = Val(Format$(Now, "MM"))
If qw5 = 12 Or qw5 = 1 Or qw5 = 2 Then Label3.Caption = "SEASON OF YEAR = WINTER : DEC JAN FEB"
If qw5 = 3 Or qw5 = 4 Or qw5 = 5 Then Label3.Caption = "SEASON OF YEAR = SPRING : MAR APR MAY"
If qw5 = 6 Or qw5 = 7 Or qw5 = 8 Then Label3.Caption = "SEASON OF YEAR = SUMMER : JUN JUL AUG"
If qw5 = 9 Or qw5 = 10 Or qw5 = 11 Then Label3.Caption = "SEASON OF YEAR = AUTUMN : SEP OCT NOV"

Timer1.Enabled = True

'Label2.Caption = Mid$(rs8$, rs9) + " -- " + rs8$ + " -- " + rs8$ + " --|||"
Label2.Caption = rs8$
'rs9 = rs9 + 1
'If rs9 - 1 > Len(rs8$) Then rs9 = 1
loadprogram = 1

nextlast1 = 1


ListView1.ListItems(1).EnsureVisible
ListView1.Enabled = True


Call Webster

Exit Sub

ErrSub2:
    DoEvents
    Resume



End Sub

Sub Pop1()
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
If ll1$ = "" Then Exit Sub
g1$ = Mid$(ll1$, 1, 2)
g2$ = Mid$(ll1$, 4, 2)
j = DateSerial(Val(Mid$(ll1$, 7, 4)), Val(g2$), Val(g1$))
j2 = DateSerial(Year(Now), Val(g2$), Val(g1$))
If Val(Mid$(ll1$, 7, 4)) > Year(Now) Then
j2 = DateSerial(Val(Mid$(ll1$, 7, 4)), Val(g2$), Val(g1$))
End If

j3 = DateDiff("d", Now, j2)

If j3 < 0 Then
If DateSerial(Year(Now), 2, 29) > Now Then
poi = DateSerial(Year(Now) + 1, 2, 29)
poi2 = DateSerial(Year(Now) + 1, 3, 1)
Else
poi = DateSerial(Year(Now), 2, 29)
poi2 = DateSerial(Year(Now), 3, 1)
End If
If poi <> poi2 Then leapyes1 = 1 Else leapyes1 = 0

j3 = j3 + 366 + leapyes1
End If

yg$ = "        "
pq$ = Trim(Str$(j3)) + "|"
RSet yg$ = pq$

poi = DateSerial(Val(Mid$(ll1$, 7, 4)), 2, 29)
poi2 = DateSerial(Val(Mid$(ll1$, 7, 4)), 3, 1)
If poi <> poi2 Then leapyes$ = "Yes |" Else leapyes$ = "No  |"
gv$ = Format$(j, "DDD")
Week = 0
days_into_future = 365
days_into_future = days_into_future * 600
For v = 0 To days_into_future
W = DateSerial(Val(Mid$(ll1$, 7, 4) + v), Val(g2$), Val(g1$))
gs$ = Format$(W, "DDD")
If Val(Mid$(ll1$, 7, 4)) >= Val(Mid$(Date$, 7, 4)) Then Exit For
If (Val(Mid$(ll1$, 7, 4) + v) >= Val(Mid$(Date$, 7, 4)) + locket) Then Exit For
If gv$ = gs$ Then Week = Week + 1
lpg = W
Next


fd = v + 1
For v = fd To days_into_future
W = DateSerial(Val(Mid$(ll1$, 7, 4) + v), Val(g2$), Val(g1$))
gs$ = Format$(W, "DDD")
If gv$ = gs$ Then week1 = ((v - fd) + 2): Exit For
If (Val(Mid$(ll1$, 7, 4) + v) = Val(Mid$(Date$, 7, 4)) + locket) Then Exit For
Next


'If p < 0 Then p = (1172 + p) + 289


x = Val(Mid$(ll1$, 7, 4))
ahha = 0
gd$ = Format$(j2, "DDD")


'T = DateDiff("yyyy", DateSerial(Val(Mid$(ll1$, 7, 4)), Val(g2$), Val(g1$)), Now)
T = DateDiff("d", DateSerial(Val(Mid$(ll1$, 7, 4)), Val(g2$), Val(g1$)), Now)
T = Int(T / 365)
H = T
lkm$ = ll2$
RSet yr = Trim(Str$(T)) + "|"
If gv$ = gd$ Then ewsa$ = "*" Else ewsa$ = " "
RSet wek = Str$(Week) + ewsa$ + "|"
RSet wek2 = Str$(Val(yr) + week1) + "|"

wclops = DateSerial(Val(Mid$(ll1$, 7, 4)), Val(g2$), Val(g1$))

X1 = NextNewMoon(wclops)
X2 = NextFirstQuarter(wclops)
x3 = NextFullMoon(wclops)
x4 = NextLastQuarter(wclops)
x5 = wclops + 200

If X1 < X2 And X1 < x5 Then x5 = X1 ': newm = 0
If X2 < x3 And X2 < x5 Then x5 = X2 ': newm = 1
If x3 < x4 And x3 < x5 Then x5 = x3 ': newm = 2
If x4 < X1 And x4 < x5 Then x5 = x4 ': newm = 3

'wclops2 = wclops
wclops2 = wclops

If Format$(x5, "DDMMYYYY") = Format$(wclops, "DDMMYYYY") Then
wclops = x5
wclops2 = wclops + TimeSerial(9, 0, 0)
End If

'Label(17) = "Next New Moon = " & Format(NextNewMoon(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(18) = "Next First Quarter = " & Format(NextFirstQuarter(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(19) = "Next Full Moon = " & Format(NextFullMoon(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(20) = "Next Last Quarter = " & Format(NextLastQuarter(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"


wsf = Age(wclops)

'mdion$ = MoonDescription(wclops)
'If mdion$ = "Waning Crescent" Then newm = 3
'If mdion$ = "Waning Gibbous" Then newm = 2
'If mdion$ = "Waxing Gibbous" Then newm = 1
'If mdion$ = "Waxing Crescent" Then newm = 0


newm = MoonAge(wclops2)
newm8 = Synod / 4
For rei = 1 To 4
newm7 = newm8 * rei
If newm7 > newm Then Exit For
Next
newm = rei - 1

'newm9 = MoonPhase(wclops2)
'If newm9 > 0 Then newm = newm9



'mdion$ = MoonDescription(wclops - 1)
'If mdion$ = "Waning Crescent" Then newm5 = 3
'If mdion$ = "Waning Gibbous" Then newm5 = 2
'If mdion$ = "Waxing Gibbous" Then newm5 = 1
'If mdion$ = "Waxing Crescent" Then newm5 = 0

newm2 = 0
If Format$(x5, "DD/MM/YYYY") = Format$(wclops, "DD/MM/YYYY") Then
newm2 = 2
End If
'If newm <> newm5 Then newm2 = 2 Else newm2 = 0






If newm2 = 2 Then nmoon2$ = "*" Else nmoon2$ = " "
If newm = 3 Then nmoon$ = "LQ"
If newm = 2 Then nmoon$ = "O "
If newm = 1 Then nmoon$ = "FQ"
If newm = 0 Then nmoon$ = ". "

'If InStr(UCase$(ll2$), "MUM") Then
'w1 = w1
'End If
'wsf = Age(wclops)

If Int(wsf) >= 10 Then fullmoon$ = nmoon$ + nmoon2$ + Format$(Int(wsf), " 00") + "|"
If Int(wsf) < 10 Then fullmoon$ = nmoon$ + nmoon2$ + Format$(Int(wsf), "  0") + "|"

If star < 12 And InStr(ll2$, "Sign Ruled by") Then star = star + 1: lkj2$ = ll2$
If InStr(ll2$, "Sign Ruled by") And p = 0 Then
star2$ = ll2$
End If
XFlip = DateSerial(Year(Now), Month(wclops), Day(wclops))
'XFlip = DateSerial(Year(Now), 7, 2)

Call nextlastfullmoon


'-------------------------------------------------


List1.AddItem gv$ + " |" + gd$ + " |" + yg$ + ll1$ + "|" + yr + wek + wek2 + leapyes$ + fullmoon$ + nextlastmoon$ + ll2$
ll$ = List1.List(List1.ListCount - 1)

With ListView1
Set LV = .ListItems.Add(, , .ListItems.Count + 1)
LV.SubItems(1) = Trim(Mid$(ll$, 1, 3))
LV.SubItems(2) = Trim(Mid$(ll$, 6, 3))
LV.SubItems(3) = Trim(Mid$(ll$, 11, 7))
LV.SubItems(4) = Trim(Mid$(ll$, 19, 10))
LV.SubItems(5) = Trim(Mid$(ll$, 30, 5))
LV.SubItems(6) = Trim(Mid$(ll$, 36, 5))
LV.SubItems(7) = Trim(Mid$(ll$, 42, 7))
LV.SubItems(8) = Trim(Mid$(ll$, 50, 4))
LV.SubItems(9) = Trim(Mid$(ll$, 55, 6))
LV.SubItems(10) = Trim(Mid$(ll$, 62, 9))
LV.SubItems(11) = Trim(Mid$(ll$, 72))
End With

'ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible

'll2$
'If InStr(KK$, ll2$) > 0 Then Stop
If InStr(KK$, ll2$) = 0 Then
List6.AddItem gv$ + " |" + gd$ + " |" + yg$ + ll1$ + "|" + yr + wek + wek2 + leapyes$ + fullmoon$ + nextlastmoon$ + ll2$
End If

'hh2 = DateDiff("d", efd, diar)

'List1.ListIndex = List1.ListCount - 1

qilim = 0
If pdot2 = 0 Then
'qss = qss + 1
List4.AddItem List1.List(List1.ListCount - 1)
qilim = 1
End If

End Sub


Private Sub nextlastfullmoon()

'---------------------------------------
'---------------------------------------
'---------------------------------------
'---------------------------------------
Dim X1, X2, x3, x4, x5 As Date

'If nextlast1 = 0 Then

'kattack = kattack + 1
'kts(kattack) = XFlip

'If kattack = 390 Then
'wes = wes
'End If

'End If

If Year(wclops) > Year(Now) Then
nextlastmoon$ = "   N/A   |"
Exit Sub
End If

If Year(wclops) < 1900 Then
nextlastmoon$ = "   N/A   |"
Exit Sub
End If


nextlastmoon$ = "         |"

If Diar > 0 Or ftag = 1 Then

xjagsq1 = Year(XFlip)
xjagsq2 = Month(XFlip)
xjagsq3 = Day(XFlip)

'xjagsq1 = 2006
'xjagsq2 = 10
'xjagsq3 = 10

'nextlastmoon$ = "1968 2098|"
xxx$ = we$

For R = 1 To 3

quity = 0

Do
DoEvents
wclops = DateSerial(xjagsq1, xjagsq2, xjagsq3)

If R = 1 Then
x5 = NextFullMoon(wclops)
End If

If R = 3 Then
x5 = NextFullMoon(wclops)
End If

If R = 2 Then
x5 = PreviousFullMoon(wclops + 2)
End If



'If Format$(x5, "DDMMYYYY") = Format$(wclops, "DDMMYYYY") Then wclops = x5

'Label(17) = "Next New Moon = " & Format(NextNewMoon(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(18) = "Next First Quarter = " & Format(NextFirstQuarter(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(19) = "Next Full Moon = " & Format(NextFullMoon(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"
'Label(20) = "Next Last Quarter = " & Format(NextLastQuarter(thedate), "mm/dd/yyyy hh:mm:ss") & " UTC"

'If Format$(x5, "DDMMYYYY") = Format$(wclops, "DDMMYYYY") Then wclops = x5


wsf = Age(wclops)

'If MoonDescription(wclops) = "Waning Crescent" Then newm = 3
'If MoonDescription(wclops) = "Waning Gibbous" Then newm = 2
'If MoonDescription(wclops) = "Waxing Gibbous" Then newm = 1
'If MoonDescription(wclops) = "Waxing Crescent" Then newm = 0

'If MoonDescription(wclops - 1) = "Waning Crescent" Then newm5 = 3
'If MoonDescription(wclops - 1) = "Waning Gibbous" Then newm5 = 2
'If MoonDescription(wclops - 1) = "Waxing Gibbous" Then newm5 = 1
'If MoonDescription(wclops - 1) = "Waxing Crescent" Then newm5 = 0



'If newm <> newm5 Then newm2 = 2 Else newm2 = 0

'If newm2 = 2 And newm = 2 And Int(wsf) <> 1 Then
'If Int(x5) = wclops Then
If Format$(x5, "DD/MM/YYYY") = Format$(wclops, "DD/MM/YYYY") Then
    nmoon2$ = "*"
Else
    nmoon2$ = " "
End If

If R = 1 Then quity = 1

If R = 2 And nmoon2$ = "*" Then Exit Do
If R = 3 And nmoon2$ = "*" Then Exit Do
If R = 2 Then xjagsq1 = xjagsq1 - 1
If R = 3 Then xjagsq1 = xjagsq1 + 1

Loop Until quity = 1 Or xjagsq1 < 100 Or xjagsq1 > 3000
quity = 0

If R = 1 And nmoon2$ = "*" Then
Mid$(nextlastmoon$, 5, 1) = "*"
xjagsq1 = Year(XFlip) - 1
End If

If R = 2 And nmoon2$ = "*" Then
Mid$(nextlastmoon$, 1, 4) = Format$(xjagsq1, "0000")
xjagsq1 = Year(XFlip) + 1
End If

If R = 2 And nmoon2$ = " " Then
Mid$(nextlastmoon$, 1, 4) = "????"
xjagsq1 = Year(XFlip) + 1
End If

If R = 3 And nmoon2$ = "*" Then
Mid$(nextlastmoon$, 6, 4) = Format$(xjagsq1, "0000")

If R = 3 And nmoon2$ = " " Then
Mid$(nextlastmoon$, 6, 4) = "????"
End If

End If



Next

End If 'if diar>0 then




End Sub


Private Sub Label4_Click()

WebStop = False

tr$ = "Program Written By Matt.Lan@btinternet.com 2005-9 Version "
tr$ = tr$ + Trim$(Str$(App.Major)) + "." + Trim$(Str$(App.Minor)) + "." + Trim$(Str$(App.Revision))

MsgBox (tr$)


End Sub

Private Sub List1_Click()


Label1.Caption = Mid$(List1.List(List1.ListIndex), 72)
rs5$ = Label1.Caption
rs6 = 1

'Label1.Caption = Mid$(rs5$, rs6) + " -- " + rs5$ + " -- " + rs5$ + " --|||"
Label1.Caption = rs5$

'rs6 = rs6 + 1
'If rs6 - 1 > Len(rs5$) Then rs6 = 1


RTB1.Text = ""
For R = 0 To List2.ListCount - 1
RTB1.Text = RTB1.Text + List2.List(R) + vbCrLf
Next

ad$ = List1.List(List1.ListIndex)

'RTB1.Text = RTB1.Text + List1.List(List1.ListIndex)
Ttg$ = "---------------------------------------------------------------------------------------" + vbCrLf

RTB1.Text = ""
RTB1.Text = RTB1.Text + Ttg$
RTB1.Text = RTB1.Text + "Description -------=" + Trim(Mid$(ad$, 72)) + vbCrLf
RTB1.Text = RTB1.Text + "Start Day ----------=" + Trim(Mid$(ad$, 1, 4)) + vbCrLf
RTB1.Text = RTB1.Text + "Day This Time ---=" + Trim(Mid$(ad$, 6, 4)) + vbCrLf
RTB1.Text = RTB1.Text + "Days to Go --------=" + Trim(Mid$(ad$, 11, 7)) + vbCrLf
RTB1.Text = RTB1.Text + "Start Date ----------=" + Trim(Mid$(ad$, 19, 10)) + vbCrLf
RTB1.Text = RTB1.Text + "Years From Start Date ---------------------=" + Trim(Mid$(ad$, 30, 5)) + vbCrLf
RTB1.Text = RTB1.Text + "Same Dayer Count --------------------------=" + Trim(Mid$(ad$, 36, 5)) + vbCrLf
RTB1.Text = RTB1.Text + "Age For Next Same Dayer ----------------=" + Trim(Mid$(ad$, 42, 7)) + vbCrLf
RTB1.Text = RTB1.Text + "Leap Year ---------------------------------------=" + Trim(Mid$(ad$, 50, 4)) + vbCrLf
RTB1.Text = RTB1.Text + "Days After a New Moon ---------------------=" + Trim(Mid$(ad$, 55, 6)) + vbCrLf
RTB1.Text = RTB1.Text + "Last Next Before/After On a Full Moon -=" + Trim(Mid$(ad$, 62, 9)) + vbCrLf
RTB1.Text = RTB1.Text + "Description --------------------------------------=" + Trim(Mid$(ad$, 72)) + vbCrLf
RTB1.Text = RTB1.Text + Ttg$
RTB1.Text = RTB1.Text + vbCrLf


On Local Error Resume Next

Ttg$ = "--------------------------------------------------" + vbCrLf
If DateClip$ = "" Then
    DateClip$ = Ttg$
    DateClip$ = DateClip$ + "Today " + Format$(Now, "MMM dd-mmm-yyyy HH:MM:SSp") + vbCrLf
    DateClip$ = DateClip$ + Ttg$
End If

DateClip$ = DateClip$ + RTB1.Text

' NOT LIST 1 -- LISTVEIW INSTEAD
' DO AT MENU ANYWAY
' ------------------------------
' Timer_CLIPBOARD.Enabled = True


End Sub



Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'End
End Sub


'Test a date for moon phase:
'Returns:
' 0: None
' 1: new moon
' 2: Quarter moon
' 3: Full moon
' 4: Three-quarter moon




Public Function MoonPhase(dDate As Date) As Integer
        
    Select Case MoonAge(dDate)
    
        'Day of a new moon
        Case Is > Synod - 1:
            MoonPhase = 1
            
        'Day of a 1/4 moon
        Case Synod / 4 - 1 To Synod / 4:
            MoonPhase = 2
            
        'Day of a full moon
        Case Synod / 2 - 1 To Synod / 2:
            MoonPhase = 3
            
        'Day of a 3/4 moon
        Case 3 * Synod / 4 - 1 To 3 * Synod / 4:
            MoonPhase = 4
            
        'No special day
        Case Else:
            MoonPhase = 0
            
    End Select
End Function

Public Function MoonAge(dDate As Date) As Single
    Dim BaseDate As Date
    BaseDate = CDate(BaseNewMoonDateString)
    
    MoonAge = Remainder((dDate - BaseDate), Synod)

End Function
Public Function Remainder(Number As Variant, DivideBy As _
 Variant) As Variant
    If Number = 0 Then
        Remainder = 0
    Else
        Remainder = Number - DivideBy * Int(Number / DivideBy)
    End If
    
End Function

Private Sub List5_Click()
'list5
End Sub

Private Sub ListView1_Click()

For R = 1 To ListView1.ListItems.Count
If ListView1.ListItems(R).Selected = True Then Exit For
Next
'ListView1.ListItems(R).Selected = False



Label1.Caption = Mid$(List1.List(R - 1), 72)
rs5$ = Label1.Caption
rs6 = 1

'Label1.Caption = Mid$(rs5$, rs6) + " -- " + rs5$ + " -- " + rs5$ + " --|||"
Label1.Caption = rs5$

'rs6 = rs6 + 1
'If rs6 - 1 > Len(rs5$) Then rs6 = 1


RTB1.Text = ""
For r1 = 0 To List2.ListCount - 1
RTB1.Text = RTB1.Text + List2.List(r1) + vbCrLf
Next

ad$ = List1.List(R - 1)

'RTB1.Text = RTB1.Text + List1.List(List1.ListIndex)
Ttg$ = "---------------------------------------------------------------------------------------" + vbCrLf

RTB1.Text = ""
RTB1.Text = RTB1.Text + Ttg$
RTB1.Text = RTB1.Text + "Description -------=" + Trim(Mid$(ad$, 72)) + vbCrLf
RTB1.Text = RTB1.Text + "Start Day ----------=" + Trim(Mid$(ad$, 1, 4)) + vbCrLf
RTB1.Text = RTB1.Text + "Day This Time ---=" + Trim(Mid$(ad$, 6, 4)) + vbCrLf
RTB1.Text = RTB1.Text + "Days to Go --------=" + Trim(Mid$(ad$, 11, 7)) + vbCrLf
RTB1.Text = RTB1.Text + "Start Date ----------=" + Trim(Mid$(ad$, 19, 10)) + vbCrLf
RTB1.Text = RTB1.Text + "Years From Start Date ---------------------=" + Trim(Mid$(ad$, 30, 5)) + vbCrLf
RTB1.Text = RTB1.Text + "Same Dayer Count --------------------------=" + Trim(Mid$(ad$, 36, 5)) + vbCrLf
RTB1.Text = RTB1.Text + "Age For Next Same Dayer ----------------=" + Trim(Mid$(ad$, 42, 7)) + vbCrLf
RTB1.Text = RTB1.Text + "Leap Year ---------------------------------------=" + Trim(Mid$(ad$, 50, 4)) + vbCrLf
RTB1.Text = RTB1.Text + "Days After a New Moon ---------------------=" + Trim(Mid$(ad$, 55, 6)) + vbCrLf
RTB1.Text = RTB1.Text + "Last Next Before/After On a Full Moon -=" + Trim(Mid$(ad$, 62, 9)) + vbCrLf
RTB1.Text = RTB1.Text + "Description --------------------------------------=" + Trim(Mid$(ad$, 72)) + vbCrLf
RTB1.Text = RTB1.Text + Ttg$
RTB1.Text = RTB1.Text + vbCrLf


On Local Error Resume Next

Ttg$ = "--------------------------------------------------" + vbCrLf
If DateClip$ = "" Then
    DateClip$ = Ttg$
    DateClip$ = DateClip$ + "Today " + Format$(Now, "MMM dd-mmm-yyyy HH:MM:SSp") + vbCrLf
    DateClip$ = DateClip$ + Ttg$
End If

DateClip$ = DateClip$ + RTB1.Text

End Sub

Private Sub MNU_CLEAR_CLIP_SET_Click()

DateClip$ = ""

End Sub

Private Sub Timer_CLIPBOARD_Timer()
'Timer_CLIPBOARD.Enabled = True

On Error Resume Next
Err.Clear
Clipboard.Clear
Clipboard.SetText DateClip$

If Err.Number = 0 Then Timer_CLIPBOARD.Enabled = False

End Sub


Private Sub Menu_About_Click()
Call Label4_Click
End Sub

Private Sub Menu_EditList_Click()
'Shell "C:\Program Files (x86)\Notepad++\notepad++.exe " + App.Path + "\00 Data\1991date.txt", vbNormalFocus
Reset
Close
Shell "C:\Program Files (x86)\Notepad++\notepad++.exe """ + App.Path + "\00 Data\1991DAT2.TXT""", vbNormalFocus

End Sub

Private Sub Menu_EditList_S_Click()
Reset
Close
Shell "C:\Program Files (x86)\Notepad++\notepad++.exe """ + App.Path + "\00 Data\1991DAT3 Secure.TXT""", vbNormalFocus

End Sub

Private Sub Menu_EditListBOTH_Click()
Reset
Close
' "C:\Program Files (x86)\Notepad++\notepad++.exe"
Shell "C:\Program Files (x86)\Notepad++\notepad++.exe """ + App.Path + "\00 Data\1991DAT2.TXT""", vbNormalFocus
Shell "C:\Program Files (x86)\Notepad++\notepad++.exe """ + App.Path + "\00 Data\1991DAT3 Secure.TXT""", vbNormalFocus

End Sub

Private Sub Menu_HtmlLarge_Click()
Shell "Explorer " + App.Path + "\Date1991_Secure.html"
End Sub

Private Sub Menu_HtmlSmall_Click()
Shell "Explorer " + App.Path + "\Date1991_Small.html"

End Sub
Private Sub Menu_HtmlSmallEmail_Click()
Shell "C:\Program Files\WinRAR\winrar.exe A -cfg- -m5 -av -cfg- -ibck -ep " + App.Path + "\Date1991_Small.zip " + App.Path + "\Date1991_Small.html -iemlMatt.Lan@btinternet.com;Marianne.Farley@btinternet.com;Lancaster.E@sky.com ", vbNormalFocus
End Sub

Private Sub Menu_HtmlWeb_Click()
Shell "Explorer " + App.Path + "\Date1991.html"

End Sub

Private Sub Menu_ResetList_Click()

Unload Form1
Form2.Timer1.Enabled = True

End Sub

Private Sub MenuExit_Click()
End
End Sub

Private Sub Mnu_ClearClip_Click()
DateClip$ = ""
End Sub

Private Sub Mnu_Clip_Them_Click()
If DateClip$ <> "" Then
    Timer_CLIPBOARD.Enabled = True
    'Clipboard.Clear
    'Clipboard.SetText (DateClip$)
End If
End Sub

Sub Webster()
On Error GoTo ErrSub3

freef1 = FreeFile
Open App.Path + "\" + "Date1991_Secure.html" For Output As #freef1
freef5 = FreeFile
Open App.Path + "\" + "Date1991.html" For Output As #freef5


freef2 = FreeFile
Open App.Path + "\Date1991_Dud.html" For Output As #freef2
freef8 = FreeFile
Open App.Path + "\Date1991_Small.html" For Output As #freef8

Do
Err.Clear
freef3 = FreeFile
Open App.Path + "\Date1991 Small Day.txt" For Output As #freef3
Loop Until Err.Number = 0
xxn$ = ""
On Error GoTo 0



Print #freef1, "<html><head>"

Print #freef1, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
Print #freef1, "<title>Matt Lan's Date1991</title>"
Print #freef1, "<META name=""description"" content=""Matt Lan's Date1991, Last Updated :-" + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + """ >"
Print #freef1, "<META NAME=""ROBOTS"" CONTENT=""INDEX,FOLLOW"">"

Print #freef5, "<html><head>"

Print #freef5, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
Print #freef5, "<title>Matt Lan's Date1991Secure</title>"
Print #freef5, "<META name=""description"" content=""Matt Lan's Date1991, Last Updated :-" + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + """ >"
Print #freef5, "<META NAME=""ROBOTS"" CONTENT=""INDEX,FOLLOW"">"

Print #freef2, "<html><head>"
Print #freef2, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
Print #freef2, "<title>Matt Lan's Date1991 Small</title>"
Print #freef2, "<META name=""description"" content=""Matt Lan's Date1991 Small, Last Updated :-" + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + """ >"
Print #freef2, "<META NAME=""ROBOTS"" CONTENT=""INDEX,FOLLOW"">"

Print #freef8, "<html><head>"
Print #freef8, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
Print #freef8, "<title>Matt Lan's Date1991 Dud</title>"
Print #freef8, "<META name=""description"" content=""Matt Lan's Date1991 Small, Last Updated :-" + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + """ >"
Print #freef8, "<META NAME=""ROBOTS"" CONTENT=""INDEX,FOLLOW"">"

Print #freef1, "</head><body>"
Print #freef2, "</head><body>"
Print #freef8, "</head><body>"
Print #freef5, "</head><body>"

fontnut$ = "<font size=""3"" face=""Arial"" color=""#000000"""

Print #freef1, "<div align=""center"">"
Print #freef1, "  <center>"
Print #freef1, "  <table border=""5"">"
Print #freef1, "    <tr>"
Print #freef1, "      <td colspan=""3"" bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Star Symbols are for On the Cuss of things </b></font></td>"
Print #freef1, "      <td>"
Print #freef1, "      <td>"
Print #freef1, "      <td>"
Print #freef1, "      <td>"
Print #freef1, "      <td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>'NM'=New Moon 'O'=Full Moon 'LQ'=Last Quarter 'FQ'=First Quarter</b></font></td>"
Print #freef1, "      <td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>A Few of Todays and Tommorows Famous Birthdays and Then The Rest</b></font></td>"
Print #freef1, "    </tr>"

Print #freef1, "    <tr>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Start Day</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Day This Time</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Days To Go</b></font></td>"
Print #freef1, "      <td nowrap bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Start Date</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Years From Start Date</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Same Dayer Count</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Age For Next Same Dayer</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Born On A Leap Year</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Days After A New Moon</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Next/Last B-Day On A Full Moon</b></font></td>"
Print #freef1, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Description</b></font></td>"

Print #freef5, "<div align=""center"">"
Print #freef5, "  <center>"
Print #freef5, "  <table border=""5"">"
Print #freef5, "    <tr>"
Print #freef5, "      <td colspan=""3"" bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Star Symbols are for On the Cuss of things </b></font></td>"
Print #freef5, "      <td>"
Print #freef5, "      <td>"
Print #freef5, "      <td>"
Print #freef5, "      <td>"
Print #freef5, "      <td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>'NM'=New Moon 'O'=Full Moon 'LQ'=Last Quarter 'FQ'=First Quarter</b></font></td>"
Print #freef5, "      <td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>A Few of Todays and Tommorows Famous Birthdays and Then The Rest</b></font></td>"
Print #freef5, "    </tr>"

Print #freef5, "    <tr>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Start Day</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Day This Time</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Days To Go</b></font></td>"
Print #freef5, "      <td nowrap bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Start Date</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Years From Start Date</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Same Dayer Count</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Age For Next Same Dayer</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Born On A Leap Year</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Days After A New Moon</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Next/Last B-Day On A Full Moon</b></font></td>"
Print #freef5, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Description</b></font></td>"

For rsd = 1 To 2
If rsd = 1 Then xxh = freef2
If rsd = 2 Then xxh = freef8

Print #xxh, "<div align=""center"">"
Print #xxh, "  <center>"
Print #xxh, "  <table border=""5"">"
Print #xxh, "    <tr>"
Print #xxh, "      <td colspan=""3"" bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Star Symbols are for On the Cuss of things</b></font></td>"
Print #xxh, "      <td>"
Print #xxh, "      <td>"
Print #xxh, "      <td>"
Print #xxh, "      <td>"
Print #xxh, "      <td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>'NM'=New Moon 'O'=Full Moon 'LQ'=Last Quarter 'FQ'=First Quarter</b></font></td>"
Print #xxh, "      <td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>A Few of Todays and Tommorows Famous Birthdays and Then The Rest</b></font></td>"
Print #xxh, "    </tr>"
Print #xxh, "    <tr>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Start Day</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Day This Time</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Days To Go</b></font></td>"
Print #xxh, "      <td nowrap bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Start Date</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Years From Start Date</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Same Dayer Count</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Age For Next Same Dayer</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Born On A Leap Year</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Days After A New Moon</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Next/Last B-Day On A Full Moon</b></font></td>"
Print #xxh, "      <td bgcolor=""#006699"" align=""center""><font size=""2"" face=""Arial"" color=""#FFFFFF""><b>Description</b></font></td>"

For et = 0 To 300
    
    'leap year error avoid data
    On Error Resume Next
back:
    Err.Clear
    exv = DateValue(Mid$(List6.List(et), 19, 6) + Trim(Str$(Year(Now))))
    If et > 300 Then Exit For
    If Err.Number > 0 Then et = et + 1: GoTo back
    On Error GoTo 0
    
    If exv < Int(Now) Then
        exv = DateValue(Mid$(List6.List(et), 19, 6) + Trim(Str$(Year(Now) + 1)))
    End If
    
    If exv > Now + 60 Then Exit For
    
    Print #xxh, "    </tr>"
    Print #xxh, "    <tr>"
    eko = Not eko
    If eko = 0 Then wash$ = "<td bgcolor=""#89C6DC"" align=""center"">"
    If eko = -1 Then wash$ = "<td bgcolor=""#88A6DD"" align=""center"">"
    If eko = 0 Then wash2$ = "<td nowrap bgcolor=""#89C6DC"" align=""center"">"
    If eko = -1 Then wash2$ = "<td nowrap bgcolor=""#88A6DD"" align=""center"">"
    
    a12$ = Mid$(List6.List(et), 1, 4)
    a12$ = Trim$(a12$)
    b2$ = a12$
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 6, 4)
    a12$ = Trim$(a12$)
    b3$ = a12$
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 12, 6)
    a12$ = Trim(Str(Val(a12$)))
    'If Mid$(a12$, 1, 1) = "|" Then a12$ = Mid$(a12$, 2)
 '   a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 19, 10)
    a12$ = Trim$(a12$)
    b4$ = Mid$(a12$, 7)
    Print #xxh, wash2$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 30, 5)
    a12$ = Trim$(a12$)
    bx5 = Val(a12$)
    b5$ = a12$ + "-Yr"
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 36, 5)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 42, 7)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 50, 4)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 55, 6)
    If Mid$(a12$, 1, 1) = "." Then a12$ = "NM" + Mid$(a12$, 2)
    a12$ = Trim$(a12$)
    b6$ = a12$ '+ " Phaze & Days After New Moon "
    If Mid$(b6$, 1, 1) = "." Then b62$ = "New Moon"
    If Mid$(b6$, 1, 2) = "FQ" Then b62$ = "First Quarter"
    If Mid$(b6$, 1, 1) = "O" Then b62$ = "Full Moon"
    If Mid$(b6$, 1, 2) = "LQ" Then b62$ = "Last Quarter"
    xv = InStr(b6$, " ")
    xv2 = Val(Mid$(b6$, xv))
    If xv2 = 1 Then
        b63$ = Str(xv2) + " Day after NM"
    Else
        b63$ = Str(xv2) + " Days after NM"
    End If
    
    b6$ = b62$ + " And" + b63$
    
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 62, 9)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(List6.List(et), 72)
    a12$ = Trim$(a12$)
    b7$ = a12$
    
    
    If InStr(a12$, "Friday The 13th On a ") Then
        wash$ = "<td bgcolor=""#7B75AA"" align=""center"">"
    End If
    
    If InStr(a12$, "6 or 66") Then
        wash$ = "<td bgcolor=""#67D66C"" align=""center"">"
    End If
    If InStr(a12$, "Next 666") Then
        wash$ = "<td bgcolor=""#67D66C"" align=""center"">"
    End If
    If InStr(a12$, "Last 666") Then
        wash$ = "<td bgcolor=""#67D66C"" align=""center"">"
    End If
    If a12$ = "666" Then
        wash$ = "<td bgcolor=""#67D66C"" align=""center"">"
    End If
    
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"

'    If commandstring$ <> "" Then
    If exv < Now And bx5 <= 20 And InStr(b7$, "-- Exact") = 0 Then
    xxn$ = xxn$ + b2$ + "-" + b4$ + " -" + b5$ + vbCrLf ' ------------- " - " + b6$ + vbCrLf
    xxn$ = xxn$ + "-- " + b7$ + vbCrLf
    xxn$ = xxn$ + vbCrLf
    
    'Print #freef3, b2$ + " - " + b4$ + " - " + b5$ + " - " + b6$
    'Print #freef3, "-- " + b7$
    'Print #freef3,
    End If
    
'    End If



Next

Next


'Limit Lenght On Trunccate Sig Of Date Year Dates
xxt$ = xxn$
'xxt5$ = xxn$
'xxn$ = xxt5$: xxt$ = xxn$
Lenght5 = 200
If Len(xxn$) > Lenght5 Then
    xxn$ = Mid$(xxn$, Len(xxn$) - Lenght5)
    xxm = InStr(xxn$, vbCrLf + vbCrLf)
    xxt$ = Mid$(xxn$, xxm + 4)
End If

'MsgBox xxt$

xxn$ = "On This Day - "
xxn$ = xxn$ + Format$(Now, "DDD DD-MM-YY") + vbCrLf
xxn$ = xxn$ + "--" + vbCrLf
xxn$ = xxn$ + xxt$

Print #freef3, xxn$


Print #freef2, "    </tr>"
Print #freef2, "  </table>"
Print #freef2, "  </center>"
Print #freef2, "</div>"

Print #freef8, "    </tr>"
Print #freef8, "  </table>"
Print #freef8, "  </center>"
Print #freef8, "</div>"

Statc$ = "<!-- Start of StatCounter Code -->" + vbCrLf
Statc$ = Statc$ + "<script type=""text/javascript"" language=""javascript"">" + vbCrLf
Statc$ = Statc$ + "var sc_project=585393;" + vbCrLf
Statc$ = Statc$ + "var sc_partition=4;" + vbCrLf
Statc$ = Statc$ + "var sc_security=""8832a018"";" + vbCrLf
Statc$ = Statc$ + "var sc_invisible=1;" + vbCrLf
Statc$ = Statc$ + "</script>" + vbCrLf
Statc$ = Statc$ + "<script type=""text/javascript"" language=""javascript"" src=""http://www.statcounter.com/counter/counter.js""></script><noscript><a href=""http://www.statcounter.com/"" target=""_blank""><img  src=""http://c5.statcounter.com/counter.php?sc_project=585393&amp;java=0&amp;security=8832a018&amp;invisible=1"" alt=""php hit counter"" border=""0""></a> </noscript>" + vbCrLf
Statc$ = Statc$ + "<!-- End of StatCounter Code -->" + vbCrLf

'Grab$ Statc$

grab$ = "<body style=""font-family: Arial"" link=""#00FFFF"" vlink=""#9066FF"" alink=""#FFFFFF"">" + vbCrLf
grab$ = grab$ + "<br><br>" + vbCrLf
grab$ = grab$ + "<div align=""left"">" + vbCrLf
grab$ = grab$ + "  <table border=""5"" style=""font-family: Rockwell; font-size: 14pt"">" + vbCrLf
grab$ = grab$ + "      <tr>" + vbCrLf
grab$ = grab$ + "        <td width=""100%"">" + vbCrLf
grab$ = grab$ + "         <p align=""center""><font size=""4""><a href=""http://matthewlan.com"">Back to Home Page</a></font></td>" + vbCrLf
grab$ = grab$ + "      </tr>" + vbCrLf
grab$ = grab$ + "  </table>" + vbCrLf
grab$ = grab$ + "</div>" + vbCrLf
grab$ = grab$ + Statc$
grab$ = grab$ + "</BODY></HTML>" + vbCrLf


Print #freef2, grab$
Print #freef8, grab$

Close #freef2, #freef8


For rsd = 1 To 2
    If rsd = 1 Then xxh = freef1: xxk = List1.ListCount - 1
    If rsd = 2 Then xxh = freef5: xxk = List6.ListCount - 1
    eko = 0

For et = 0 To xxk
    If rsd = 1 Then xxh = freef1: xxj$ = List1.List(et)
    If rsd = 2 Then xxh = freef5: xxj$ = List6.List(et)
    Print #xxh, "    </tr>"
    Print #xxh, "    <tr>"
    eko = Not eko
    If eko = 0 Then wash$ = "<td bgcolor=""#89C6DC"" align=""center"">"
    If eko = -1 Then wash$ = "<td bgcolor=""#88A6DD"" align=""center"">"
    If eko = 0 Then wash2$ = "<td nowrap bgcolor=""#89C6DC"" align=""center"">"
    If eko = -1 Then wash2$ = "<td nowrap bgcolor=""#88A6DD"" align=""center"">"
    a12$ = Mid$(xxj$, 1, 4)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 6, 4)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 12, 6)
'    If Mid$(a12$, 1, 1) = "|" Then a12$ = Mid$(a12$, 2)
    a12$ = Trim(Str(Val(a12$)))
'    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 19, 10)
    a12$ = Trim$(a12$)
    Print #xxh, wash2$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 30, 5)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 36, 5)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 42, 7)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 50, 4)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 55, 6)
    If Mid$(a12$, 1, 1) = "." Then a12$ = "NM" + Mid$(a12$, 2)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 62, 9)
    a12$ = Trim$(a12$)
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    a12$ = Mid$(xxj$, 72)
    a12$ = Trim$(a12$)
    If InStr(a12$, "Friday The 13th On a ") Then
        wash$ = "<td bgcolor=""#7B75AA"" align=""center"">"
    End If
    If InStr(a12$, "6 or 66") Then
        wash$ = "<td bgcolor=""#67D66C"" align=""center"">"
    End If
    If InStr(a12$, "Next 666") Then
        wash$ = "<td bgcolor=""#67D66C"" align=""center"">"
    End If
    If InStr(a12$, "Last 666") Then
        wash$ = "<td bgcolor=""#67D66C"" align=""center"">"
    End If
    If a12$ = "666" Then
        wash$ = "<td bgcolor=""#67D66C"" align=""center"">"
    End If
    
    Print #xxh, wash$ + "" + fontnut$ + ">" + a12$ + "</font></td>"
    
Next
Next






Print #freef1, "    </tr>"
Print #freef1, "  </table>"
Print #freef1, "  </center>"
Print #freef1, "</div>"
Print #freef1, grab$

Print #freef5, "    </tr>"
Print #freef5, "  </table>"
Print #freef5, "  </center>"
Print #freef5, "</div>"
Print #freef5, grab$

Close #freef1, #freef3, #freef5


'Set Fs22 = CreateObject("Scripting.FileSystemObject")
'Fs22.copyFile a1$, b1$
''D:\MY WEBS\MatthewLan.Com Web\Secure_Folder\Cool

    On Local Error Resume Next
    Kill App.Path + "\Date1991_Small.zip"
    Kill App.Path + "\Date1991_x.zip"
    On Local Error GoTo 0
    Shell "C:\Program Files\WinRAR\winrar.exe A -CFG- -av -ibck -m5 -ep " + App.Path + "\Date1991_Small.zip " + App.Path + "\Date1991_Small.html", vbNormalFocus
    Shell "C:\Program Files\WinRAR\winrar.exe A -CFG- -av -ibck -m5 -ep " + App.Path + "\Date1991_x.zip " + App.Path + "\Date1991.html", vbNormalFocus

    On Local Error Resume Next
    Kill App.Path + "\Date1991_Secure.zip"
    On Local Error GoTo 0

    Shell "C:\Program Files\WinRAR\winrar.exe A -CFG- -av -ibck -m5 -ep " + App.Path + "\Date1991_Secure.zip " + App.Path + "\Date1991_Secure.html", vbNormalFocus

Exit Sub

ErrSub3:
    DoEvents
    Resume

End Sub

Private Sub Timer2_Timer()

End Sub
'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

Private Sub Mnu_Events_Click()
End Sub

Private Sub Mnu_DelPrvi_Click()
'Dim itmFound As ListItem   ' FoundItem variable.
   
Call Mnu_Select_Click


Dim itmFound As ListItem   ' FoundItem variable.
   
For we2 = ListView1.ListItems.Count To 1 Step -1
    DoEvents
    a1$ = ListView1.ListItems.Item(we2).SubItems(11)
    b1$ = ListView1.ListItems.Item(we2).SubItems(4)
    
    If DateValue(b1$) < DateSerial(Year(Now) - 1, 1, 1) Then milk = 1
    If ListView1.ListItems(we2).Selected = True Then milk = 0
    If milk = 1 Then
        ListView1.ListItems.Remove (we2)
        'ListView1.ListItems(we2).Selected = True
        'If ag = 0 Then ListView1.ListItems(we2).EnsureVisible: ag = 1
        'Exit For
    End If

Next
On Error Resume Next
ListView1.SetFocus


On Error Resume Next
ListView1.SetFocus

End Sub

'***********************************************

Private Sub Mnu_Logg_Folder_Click()
Shell "Explorer  /e, " + App.Path + "\00 Data\", vbNormalFocus
End Sub

Private Sub Mnu_Redo_Click()
Open App.Path + "\DayNow.txt" For Output As #1
    Print #1,
Close #1
Redo = True
Unload Form1
Redo = False

Form1.Show
End Sub

Private Sub Mnu_Select_Click()
Dim itmFound As ListItem   ' FoundItem variable.
   
For we2 = 1 To ListView1.ListItems.Count
    DoEvents
    a1$ = ListView1.ListItems.Item(we2).SubItems(11)
    b1$ = ListView1.ListItems.Item(we2).SubItems(4)
    milk = 0
    If a1$ <> "" Then
    If InStr(a1$, "Exact") > 0 Then milk = 1
    If InStr(a1$, "B DAY") > 0 Then milk = 1
    If InStr(a1$, "B Day") > 0 Then milk = 1
    If InStr(a1$, "Birthday") > 0 Then milk = 1
    If InStr(a1$, "School Hols") > 0 Then milk = 1
    If Len(a1$) > 20 Then
        If InStr(a1$, "Next") > 0 And InStr(Mid$(a1$, Len(a1$) - 15), Trim(Str(Year(Now)))) > 0 Then milk = 1
    End If
    If InStr(a1$, "Next -- Event:-") > 0 Then milk = 1
    If InStr(a1$, "Event:-") > 0 And Year(DateValue(b1$)) >= Year(Now) Then milk = 1
    If InStr(a1$, "Next ") > 0 And Year(DateValue(b1$)) >= Year(Now) Then milk = 1
    
    If milk = 1 Then
        ListView1.ListItems(we2).Selected = True
        'If ag = 0 Then ListView1.ListItems(we2).EnsureVisible: ag = 1
        'Exit For
    End If
    End If
Next
On Error Resume Next
ListView1.SetFocus

End Sub

Private Sub Mnu_VB_Click()

If Dir("C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
    Shell """C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    End
End If

If Dir("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    End
End If
End Sub

Private Sub MNU_VB_FOLDER_Click()
    Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus
    End
    Beep
    
    Me.WindowState = vbMinimized
    
End Sub

