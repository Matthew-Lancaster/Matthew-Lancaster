VERSION 5.00
Begin VB.Form z_MENU_Form1 
   BackColor       =   &H00808080&
   Caption         =   "MENU Form"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "7. VCF CARDS REFORMAT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4680
      Width           =   11115
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6. VCF CARDS SPLIT MERGE HARDCODED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   4080
      Width           =   11115
   End
   Begin VB.Label Label23 
      BackColor       =   &H0000FF00&
      Caption         =   "Label7"
      Height          =   495
      Left            =   13200
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5. DEL EMPTYS HARD CODED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   11085
   End
   Begin VB.Label Label22 
      BackColor       =   &H008080FF&
      Caption         =   "Label7"
      Height          =   495
      Left            =   13200
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4. DEL EMPTYS - WITH - CLIPBOARD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   11115
   End
   Begin VB.Label Label3 
      Caption         =   "3. CRC DUPE SEMI CODED WITH TEXT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   11115
   End
   Begin VB.Label Label2 
      Caption         =   "2. CRC DUPE HARD CODED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   11100
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "CLIPBOARD GET STATUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   15315
   End
   Begin VB.Label Label20 
      Caption         =   "CLIPBOARD GET"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   15315
   End
   Begin VB.Label Label1 
      Caption         =   "1. CRC DUPE - WITH - CLIPBOARD or SELECTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11100
   End
   Begin VB.Menu MNU_VBME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu Mnu_Logg_File 
      Caption         =   "Load Logg File"
   End
End
Attribute VB_Name = "z_MENU_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path1
Dim Path2
Dim Path3
Dim RESIZE_AT_LOAD

Private Sub Form_Load()

Me.Caption = "Sort_AnyThing - CRC DUPE - GOOGLE IMAGE"


If mCancelScan = True Then
    Exit Sub: Unload Me: Exit Sub
End If


'---------------------------------------------
Label20.Caption = Clipboard.GetText


'If Fs.FolderExists(Label20.Caption) = False Then
    
Dim DArray(100), XACount
XACount = 1
FR1 = FreeFile
Open App.Path + "\DATA Path Store" For Input As #FR1
    Do
    Line Input #FR1, DArray(XACount)
    XACount = XACount + 1
    If XACount = 101 Then XACount = 1
    Loop Until EOF(FR1)
Close #FR1

XACount = XACount - 1
vx = 1
Do
    If DArray(XACount) <> "" Then
        If Fs.FolderExists(DArray(XACount)) Then Label20.Caption = DArray(XACount): Exit Do
    End If
    XACount = XACount - 1
    If XACount = 0 Then XACount = 100
    vx = vx + 1
Loop Until vx = 100


IDE_PATH_VAR = "D:\# MY DOCS\# 01 My Documents\Z"
IDE_PATH_VAR = "D:\M"

    
Path1 = DArray(XACount)
Path2 = Clipboard.GetText
Path3 = IDE_PATH_VAR

        
If Path1 <> Path2 Then
    If Fs.FolderExists(Clipboard.GetText) = True Then
        FR1 = FreeFile
        Open App.Path + "\DATA Path Store" For Append As #FR1
            Print #FR1, Clipboard.GetText
        Close #FR1
    End If
End If


'-----------------------------------------
pathc = 0

pathc = pathc + 1
If Fs.FolderExists(Path1) = True Then
    textmsgbox1 = "Option " + Trim(Str(pathc)) + ". " + Path1 + " -- From Remembered"
    ax = "1"
    pathx1 = Path1
    pathx2 = "From Remembered For You"
Else
    Z_Choice_Frm.Label1.Enabled = False: Z_Choice_Frm.Label1.Caption = "..."
    'Z_Choice_Frm.Label1.Enabled = False
End If

pathc = pathc + 1
If Fs.FolderExists(Path2) = True Then
    textmsgbox2 = "Option " + Trim(Str(pathc)) + ". " + Path2 + " -- From ClipBoard"
    ax = ax + "2"
    pathx1 = Path2
    pathx2 = "From ClipBoard"
Else
    Z_Choice_Frm.Label2.Enabled = False: Z_Choice_Frm.Label2.Caption = "..."
End If

pathc = pathc + 1
If Fs.FolderExists(Path3) = True Then
    textmsgbox3 = "Option " + Trim(Str(pathc)) + ". " + Path3 + " -- From HardCoded"
    ax = ax + "3"
    pathx1 = Path3
    pathx2 = "From HardCoded"
Else
    Z_Choice_Frm.Label3.Enabled = False: Z_Choice_Frm.Label3.Caption = "..."
End If
'-----------------------------------------


If pathc = 0 Then
    MsgBox "You Got to Copy a Path to Clipboard to Get Started - Run This Again", , ScanPath.Caption: End
End If
'-----------------------------------------

        
        
If pathc > 1 Then
    If textmsgbox1 <> "" Then Z_Choice_Frm.Label1 = textmsgbox1 '+ vbCrLf
    If textmsgbox2 <> "" Then Z_Choice_Frm.Label2 = textmsgbox2
    If textmsgbox3 <> "" Then Z_Choice_Frm.Label3 = textmsgbox3
    Me.WindowState = vbMinimized
    Z_Choice_Frm.Show
Else
    Label20.Caption = pathx1
    If pathx2 = "From HardCoded" Then Path_from_Ide = True
    Call Form_Part2
End If

'If Z_Choice_Frm.Visible = True Then Me.Visible = False: DoEvents



End Sub

Public Sub Clip_Choice(var3)

If var3 = 1 Then Label20.Caption = Path1
If var3 = 2 Then Label20.Caption = Path2
If var3 = 3 Then Label20.Caption = Path3

Call Form_Part2
Me.Visible = True

End Sub



Sub Form_Part2()

        
RESIZE_AT_LOAD = True
Me.WindowState = vbNormal
        
        
        
        
'            Label20.Caption = IDE_PATH_VAR
'            Path_from_Ide = True
'        End If
'    'End If
'End If

'If Fs.FolderExists(Label20.Caption) = False And IsIDE = True Then
'    Label20.Caption = IDE_PATH_VAR
'    Path_from_Ide = True
'End If


'If Fs.FolderExists(Clipboard.GetText) = True Then
'    fr1 = FreeFile
'    Open App.Path + "\DATA Path Store" For Append As #fr1
'        Print #fr1, Clipboard.GetText
'    Close #fr1
'End If


'---------------------------------------------

If Path_from_Ide = True Then
    Source_path = "Path From IDE"
        For Each Control In Controls
            If InStr(Control.Caption, "- WITH - CLIPBOARD") > 0 Then Control.Caption = Replace(Control.Caption, "- WITH - CLIPBOARD", "- WITH - IDE HARD-CODED PATH")
        Next

Else
    Source_path = "Path From ClipBoard"
End If

'LABEL1
For Each Control In Controls
    'Len - Single Digit Labels
    If InStr(Control.Name, "Label") > 0 And Len(Control.Name) = 6 Then
        Control.Width = 12000: Control.Left = 0
    End If
Next


If Fs.FolderExists(Label20.Caption) = False Then
    
    Label21.Caption = Source_path + " - FOLDER DON'T EXIST"
    
    Label21.BackColor = Label22.BackColor
    Label1.BackColor = Label22.BackColor
    Label4.BackColor = Label22.BackColor

Else
    
    
    If Len(Label20.Caption) < 4 And Fs.FolderExists(Label20.Caption) = True Then
        Label21.Caption = Source_path + " - FOLDER WARNING ON THE ROOT"
        Label21.BackColor = vbYellow
    
    Else
        Label21.Caption = Source_path + " - FOLDER STATUS IS OKAY"
        Label21.BackColor = vbGreen
    End If
    If Mid(Label20.Caption, Len(Label20.Caption)) <> "\" Then Label20.Caption = Label20.Caption + "\"
End If

'For Each Control In Controls
'
'    If Mid(Control.Caption, 2, 2) = ". " Then
'        Control.FontSize = 15
'        'Label1.FontSize = 15
'    End If
'
'Next




End Sub


Private Sub Form_Resize()

'If Me.Visible = False Then Exit Sub

If RESIZE_AT_LOAD = True And Me.WindowState <> vbMinimized Then
    Me.WindowState = vbNormal
    Call MNU_Norm_Click
    RESIZE_AT_LOAD = False
End If

'xbigg = 0
'For Each Control In Controls
'    'control.name
'    If LCase(Mid(Control.Name, 1, 3)) <> "mnu" Then
'        If Control.Width > xbigg Then xbigg = Control.Width
'    End If
'Next
'
'Me.Width = xbigg + 140


End Sub

Private Sub Form_Unload(Cancel As Integer)

'mCancelScan = True

Me.Hide
Me.WindowState = vbMinimized

'Unload ScanPath





End Sub

Private Sub Label1_Click()

'1. CRC DUPE WITH CLIPBOARD

If Label1.BackColor = Label22.BackColor Then Exit Sub


MENU_OPTION = 1
'Call CRCDUPE

PathToLoad = z_MENU_Form1.Label20.Caption

ScanPath.txtPath = PathToLoad

ScanPath.Show
DoEvents

'REM 1ST FOLDER
ScanPath.Dir1.Path = ScanPath.txtPath
'REM 1ST FOLDER
D1$ = ScanPath.Dir1.List(0)

ScanPath.First_Folder_Path = D1$


If Dir(App.Path + "\Repeat_Questions_Toggle_Flag.txt") <> "" Then
    FR1 = FreeFile
    Open App.Path + "\Repeat_Questions_Toggle_Flag.txt" For Input As #FR1
        Line Input #FR1, var1
    Close #FR1
End If


ScanPath.Mnu_Repeat_Questions_Toggle.Checked = var1

If ScanPath.Mnu_Repeat_Questions_Toggle.Checked Then
    ScanPath.Mnu_Repeat_Questions_Toggle.Caption = "Answer Question Message Boxes Toggle Checked = True"
Else
    ScanPath.Mnu_Repeat_Questions_Toggle.Caption = "Answer Question Message Boxes Toggle Checked = False"
End If




Me.Hide
Me.WindowState = vbMinimized
Unload Me

End Sub

Private Sub Label2_Click()
'2. CRC DUPE HARD CODED

Me.Hide
MENU_OPTION = 2
'Call CRCDUPE

ScanPath.txtPath = z_MENU_Form1.Label20.Caption

ScanPath.Show
'Unload Me

End Sub

Private Sub Label3_Click()

MENU_OPTION = 3

Me.Hide

ScanPath.txtPath = z_MENU_Form1.Label20.Caption

ScanPath.Show
'Unload Me

End Sub

Private Sub Label4_Click()

If Label4.BackColor = Label22.BackColor Then Exit Sub

MENU_OPTION = 4

Me.Hide

ScanPath.txtPath = z_MENU_Form1.Label20.Caption

ScanPath.Show
'Unload Me

End Sub

Private Sub Label5_Click()

MENU_OPTION = 5

Me.Hide

ScanPath.txtPath = z_MENU_Form1.Label20.Caption

ScanPath.Show

MsgBox "Finish the Code to Verify Not Del Folders With Contents", , ScanPath.Caption



'Call ScanPath.Del_Emptys

'Unload Me


End Sub

Private Sub Label6_Click()

MENU_OPTION = 6
Me.Hide

ScanPath.txtPath = z_MENU_Form1.Label20.Caption

ScanPath.Show

Call VCF_CARDS_SPLIT_MERGE

'Unload Me
'D:\# MY DOCS\# 01 My Documents\00 Blue Tooth Exchange Folder\00 CONTACTS\CONTACT_BOOK_DADS\00 MERGE

End Sub




Sub VCF_CARDS_SPLIT_MERGE()


WDIR = "D:\# MY DOCS\# 01 My Documents\00 Blue Tooth Exchange Folder\00 CONTACTS\00 2015-09-05 -- CONTACTS GMAIL\"
WFILE = "D:\# MY DOCS\# 01 My Documents\00 Blue Tooth Exchange Folder\00 CONTACTS\GMAIL contacts 2015-09.vcf"

FR1 = FreeFile
Open WFILE For Input As #FR1

Do
    DoEvents
    Line Input #FR1, Data1
    If Mid(Data1, 1, 3) = "FN:" Then
    we = we + 1
    ScanPath.lblCount1 = Trim(Str(we))

        FNAME = Trim(Mid(Data1, 4)) + ".vcf"
        FNAME = Replace(FNAME, "/", "-")
        FNAME = Replace(FNAME, ":", "-")
        FNAME = Replace(FNAME, "?", "-")
        'FNAME = Replace(FNAME, "=", "-")
        FR2 = FreeFile
        Open WDIR + FNAME For Output As #FR2
            Print #FR2, "BEGIN:VCARD"
            Print #FR2, "VERSION:3.0"
            Print #FR2, Data1
            Do
                Line Input #FR1, Data2
                Print #FR2, Data2
            Loop Until Data2 = "END:VCARD"
        Close #FR2
    
    End If


Loop Until EOF(FR1)
Close #FR1


Exit Sub

'ScanPath.ListView1.ListItems.Clear
'ScanPath.txtPath = "D:\My Webs\MatthewLan.Com Web\"
'ScanPath.cboMask.Text = "*.*"
'ScanPath.chkSubFolders = vbChecked
'Call ScanPath.cmdScanDir_Click
'ScanPath.Show
'Reset
'For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
'    we = ScanPath.ListView1.ListItems.Count
'    ScanPath.lblCount1 = Trim(Str(we))
'    DoEvents
'
'    Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
'    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
'    A1$ = A1$ + ScanPath.ListView1.ListItems.Item(we)
'    outff = 0
'    If InStr(A1$, "_private") > 0 Then outff = 1
'    If InStr(A1$, "_vti") > 0 Then
'        outff = 1
'    End If
'
'    If outff = 1 Then
'        ScanPath.ListView1.ListItems.Remove (we)
'        On Error Resume Next
'        Fs.Deletefolder A1$, True
'        If Err.Number > 0 Then Stop
'        Err.Description
'        On Error GoTo 0
'
'    End If
'Next



End Sub

Private Sub Label7_Click()

MENU_OPTION = 7
Me.Hide
ScanPath.Show

ScanPath.txtPath = z_MENU_Form1.Label20.Caption

Call VCF_CARDS_REFORMAT

'Unload Me
End Sub



Sub VCF_CARDS_REFORMAT()


WDIR = "D:\# MY DOCS\# 01 My Documents\00 Blue Tooth Exchange Folder\00 CONTACTS\CONTACT_BOOK_DADS\00 MERGE\"


ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath = WDIR
ScanPath.cboMask.Text = "*.vcf"
ScanPath.chkSubFolders = vbUnchecked
Call ScanPath.cmdScan_Click
ScanPath.Show
Reset
For we = 1 To ScanPath.ListView1.ListItems.Count
'    we = ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount1 = Trim(Str(we))
    DoEvents


    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)





    WFILE = B1$
    
    'WFILW = Trim(Mid(Data1, 4)) + ".vcf"
    WFILE = Replace(WFILE, "_", " ")
    If WFILE <> B1$ Then
        If Dir(A1$ + WFILE) <> "" Then
            Kill A1$ + B1$
        Else
            Name A1$ + B1$ As A1$ + WFILE
        End If
    End If
    
    FR1 = FreeFile
    Open A1$ + WFILE For Input As #FR1
    Data3 = ""
    Do
        DoEvents
        Line Input #FR1, Data1
        If Mid(Data1, 1, 3) = "FN:" Or Mid(Data1, 1, 2) = "N:" Then
            'we = we + 1
            'ScanPath.lblCount1 = Trim(Str(we))
        
            'FNAME = Trim(Mid(Data1, 4)) + ".vcf"
            Data1 = Replace(Data1, "/", "-")
'            Data1 = Replace(Data1, ":", "-")
            Data1 = Replace(Data1, "?", "-")
            Data1 = Replace(Data1, "_", " ")
        End If
        Data3 = Data3 + Data1 + vbCrLf
    
    
    Loop Until EOF(FR1)
    Close #FR1

    FR2 = FreeFile
    Open WDIR + WFILE For Output As #FR2
        Print #FR2, Data3
    Close #FR2



Next

Exit Sub

'ScanPath.ListView1.ListItems.Clear
'ScanPath.txtPath = "D:\My Webs\MatthewLan.Com Web\"
'ScanPath.cboMask.Text = "*.*"
'ScanPath.chkSubFolders = vbChecked
'Call ScanPath.cmdScanDir_Click
'ScanPath.Show
'Reset
'For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
'    we = ScanPath.ListView1.ListItems.Count
'    ScanPath.lblCount1 = Trim(Str(we))
'    DoEvents
'
'    Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
'    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
'    A1$ = A1$ + ScanPath.ListView1.ListItems.Item(we)
'    outff = 0
'    If InStr(A1$, "_private") > 0 Then outff = 1
'    If InStr(A1$, "_vti") > 0 Then
'        outff = 1
'    End If
'
'    If outff = 1 Then
'        ScanPath.ListView1.ListItems.Remove (we)
'        On Error Resume Next
'        Fs.Deletefolder A1$, True
'        If Err.Number > 0 Then Stop
'        Err.Description
'        On Error GoTo 0
'
'    End If
'Next



End Sub

Private Sub Mnu_Logg_File_Click()
    Call ScanPath.Mnu_NotePad_Logg_Click
End Sub


Private Sub MNU_Norm_Click()
    
    On Error Resume Next
    
    'If FORM_LOAD_FLAG = False Then Me.WindowState = vbNormal
    
    If Me.WindowState <> vbMinimized Then
        Form1.Width = Screen.Width - 3000
        Form1.Height = Screen.Height - 2000
        
        Form1.Left = (Screen.Width - Me.Width) / 2
        Form1.Top = (Screen.Height - Me.Height) / 2
    End If
    
    'Call Form_Resize

Exit Sub

'
'If Me.WindowState = RESIZE_AT_LOAD2 Then Exit Sub
'If Me.WindowState <> RESIZE_AT_LOAD2 Then RESIZE_AT_LOAD = True
'RESIZE_AT_LOAD2 = Me.WindowState
'
'If Me.WindowState = vbMinimized Then Exit Sub
'
''Me.WindowState = 2
''Form1.Left = 0
''Form1.Top = 440
'
''If GetComputerName = "MATT-555ROIDS" Then
'
'If Me.WindowState = vbNormal Then
''If RESIZE_AT_LOAD = True And Me.WindowState = vbNormal Then
'    'RESIZE_AT_LOAD = False
'
'    Form1.Width = Screen.Width - 3000
'    Form1.Height = Screen.Height - 2000
'    'End If
'
'    Form1.Left = (Screen.Width - Me.Width) / 2
'    Form1.Top = (Screen.Height - Me.Height) / 2
'End If
'
''If RESIZE_AT_LOAD = True And Me.WindowState = vbMaximized Then
'If Me.WindowState = vbMaximized Then
'    'RESIZE_AT_LOAD = False
'
'    OFFSETZ = 420
'    Form1.Width = Screen.Width
'    Form1.Height = Screen.Height - OFFSETZ
'    Form1.Left = 0
'    Form1.Top = OFFSETZ
'
'End If
'
''DoEvents
'

End Sub


Private Sub MNU_VBME_Click()
Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End
End Sub
