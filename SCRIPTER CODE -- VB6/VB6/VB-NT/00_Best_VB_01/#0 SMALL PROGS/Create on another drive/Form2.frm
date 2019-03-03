VERSION 5.00
Begin VB.Form INFAR 
   Caption         =   "INFAR"
   ClientHeight    =   10470
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12555
   LinkTopic       =   "Form2"
   ScaleHeight     =   10470
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   7665
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   7665
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9540
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "REMOVE FRAME STREAM  NONSE"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   8160
      TabIndex        =   10
      Top             =   5520
      Width           =   4365
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8160
      TabIndex        =   9
      Top             =   4800
      Width           =   1755
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "SELECT && ADD RECENT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   8160
      TabIndex        =   8
      Top             =   3600
      Width           =   3585
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9960
      TabIndex        =   7
      Top             =   3120
      Width           =   315
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "DATE RECENT ONE MONTH"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   8160
      TabIndex        =   6
      Top             =   1920
      Width           =   5325
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "CUT QUARTERED"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8160
      TabIndex        =   5
      Top             =   1320
      Width           =   4425
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "COMBINE ALL"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8160
      TabIndex        =   4
      Top             =   720
      Width           =   3525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "FORM UNLOAD"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8160
      TabIndex        =   1
      Top             =   120
      Width           =   3795
   End
   Begin VB.Menu MNU_CREATE_SCRIPT_FILM_AVI_MP4 
      Caption         =   "CREATE SCRIPT FILM AVI MP4"
   End
End
Attribute VB_Name = "INFAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LAB2, FLAGQUARTM, NOTEXECUTEINFRAR


Sub RUN_INFRA()

If NOTEXECUTEINFRAR = True Then Exit Sub

Me.WindowState = vbMinimized


Shell "C:\Program Files\IrfanView\i_view32.exe /killmesoftly", vbHide

'--- NOT WITH RELOAD
'Shell "C:\Program Files\IrfanView\i_view32.exe /slideshow=""D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt"" /fs /silent /one /reloadonloop", vbMinimizedFocus

Shell "C:\Program Files\IrfanView\i_view32.exe /slideshow=""D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt"" /fs /silent /one", vbMinimizedFocus

End

Exit Sub

'---WITH RELOAD
Shell "C:\Program Files\IrfanView\i_view32.exe /slideshow=""D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt"" /fs /silent /one /reloadonloop"



Shell "D:\#\#D\ONE MONTH OF SEPT\i_view32.exe -- IRFAN SLIDESHOW -- FULL DISPLAY -- MAIN.BAT", vbMinimizedNoFocus

End Sub


Private Sub File1_Click()
If LAB2 = True Then Exit Sub

FileName = File1.Path + "\" + File1.List(File1.ListIndex)

If Label6.BackColor = &H80FFFF Then
    Call Label6_Click
    Call RUN_INFRA
    Exit Sub

End If
If Label7.BackColor = &H80FFFF Then
    Call Label7_Click
    Shell "C:\Program Files\Notepad++\notepad++.exe """ + FileName + """"
    
    Exit Sub

End If


Set FS = CreateObject("Scripting.FileSystemObject")
Set F = FS.GetFILE(File1.Path + "\" + File1.List(File1.ListIndex))
F.Copy File1.Path + "\IRFAN SLIDESHOW.txt"

Call RUN_INFRA



End Sub

Public Sub Form_Load()
'FLAGQUART = 30 * 7
File1.Path = "D:\#\#D\ONE MONTH OF SEPT"
File1.Pattern = "IRFAN SLIDESHOW*.*"

FLAGQUART = 400
Label4 = "DATE RECENT" + Str(FLAGQUART) + " DAY"
'RESIZE PROBLEM
On Error Resume Next

Me.Top = 500 + 150
Me.Left = 400

List1.Top = File1.Top
List2.Top = File1.Top

List1.Height = File1.Height
List2.Height = File1.Height

File1.Width = 5500
List1.Width = 1400
List2.Width = 2400

List1.Left = File1.Left + File1.Width + 15
List2.Left = List1.Left + List1.Width + 15

Label1.Left = List2.Left + List2.Width + 15

Label2.Left = List2.Left + List2.Width + 15
Label2.Top = Label1.Top + Label1.Height + 15

Label3.Left = List2.Left + List2.Width + 15
Label3.Top = Label2.Top + Label2.Height + 15

Label4.Left = List2.Left + List2.Width + 15
Label4.Top = Label3.Top + Label3.Height + 15

Label5.AutoSize = False
Label5.Left = List2.Left + List2.Width + 15
Label5.Top = Label4.Top + Label4.Height + 15
Label5.Width = Label4.Width

Label6.AutoSize = False
Label6.Left = List2.Left + List2.Width + 15
Label6.Top = Label5.Top + Label5.Height + 15
Label6.Width = Label5.Width

Label7.AutoSize = False
Label7.Left = List2.Left + List2.Width + 15
Label7.Top = Label6.Top + Label6.Height + 15
Label7.Width = Label6.Width

Label8.WordWrap = False
Label8.AutoSize = True
Label8.Width = Label7.Width
Label8.WordWrap = True
Label8.Left = List2.Left + List2.Width + 15
Label8.Top = Label7.Top + Label7.Height + 15


Me.Width = Label4.Left + Label4.Width + 50
Me.Height = File1.Height + File1.Top + 50

List1.Font = File1.Font
List1.FontSize = File1.FontSize
List1.FontBold = File1.FontBold

List2.Font = File1.Font
List2.FontSize = File1.FontSize
List2.FontBold = File1.FontBold

For R = 0 To File1.ListCount - 1
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set F = FS.GetFILE(File1.Path + "\" + File1.List(R))
    On Error Resume Next
    Err.Clear
    Set D = FS.GetDrive(Chr(R + 67))
    xe = D.VolumeName
    If Err.Number > 0 Then xe = ""
    
    If F.Size > 0 Then
        List1.AddItem Format(F.Size / (1024 ^ 2), "0.00") + "k"
    Else
        List1.AddItem "----"
    End If
    
Next






For R = 0 To File1.ListCount - 1
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set F = FS.GetFILE(File1.Path + "\" + File1.List(R))
    If Len(File1.List(R)) = 21 Then
        X = Mid(File1.List(R), 17, 1)
        
        On Error Resume Next
        Err.Clear
        Set D = FS.GetDrive(X)
        
        xe = D.VolumeName
        If Err.Number > 0 Then xe = ""
        
        List2.AddItem xe
    
        Else
        
        List2.AddItem ""
    
    End If
    
Next

Me.Show
'IR.Show

Form2.SetFocus
'IR.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click()
Me.Hide
Unload Me
End Sub

Private Sub Label2_Click()

LAB2 = True
FILEVAR1 = File1.Path + "\IRFAN SLIDESHOW-COMBINE-.txt"
Open FILEVAR1 For Output As #1

For R = 0 To File1.ListCount - 1
    File1.ListIndex = R + 2
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set F = FS.GetFILE(File1.Path + "\" + File1.List(R))
    On Error Resume Next
    Err.Clear
    
    If F.Size > 0 And Len(File1.List(R)) = 21 Then
    
        Open File1.Path + "\" + File1.List(R) For Input As #2
            L = Input(LOF(2), 2)
            Print #1, L
        Close #2
    
    End If
    
Next

Close #1, #2

Set F = FS.GetFILE(FILEVAR1)
F.Copy File1.Path + "\IRFAN SLIDESHOW.txt"


File1.ListIndex = 0
LAB2 = False
Call Form_Load
Shell "C:\Program Files\IrfanView\i_view32.exe /killmesoftly", vbHide

'--- NOT WITH RELOAD
Shell "C:\Program Files\IrfanView\i_view32.exe /slideshow=""D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt"" /fs /silent /one", vbMinimizedFocus
End Sub

Private Sub Label3_Click()
LAB2 = True
Set FS = CreateObject("Scripting.FileSystemObject")

File1.Pattern = "IRFAN SLIDESHOW-*.*"
File1.Refresh
DoEvents


For R = 0 To File1.ListCount - 1
    If R + 1 < File1.ListCount Then File1.ListIndex = R + 1
    Set F = FS.GetFILE(File1.Path + "\" + File1.List(R))
    If F.Size > 0 And Len(File1.List(R)) = 21 Then
        EX = EX + F.Size
    End If
Next

EX = EX / 4
EZ = 0



FLAGQUART = 1
TRIGX = False



Open File1.Path + "\IRFAN SLIDESHOW-QUART--" + Format(FLAGQUART, "00") + "--.txt" For Output As #1
For R = 0 To File1.ListCount - 1
    If R + 1 < File1.ListCount Then File1.ListIndex = R + 1
    Set F = FS.GetFILE(File1.Path + "\" + File1.List(R))
    
    If F.Size > 0 And Len(File1.List(R)) = 21 Then
    
        Open File1.Path + "\" + File1.List(R) For Input As #2
'            Z = LOF(2)
            'L = INPUT(LOF(2), 2)
            Do
            U = Time
            If OU <> U Then DoEvents
            OU = U
            Line Input #2, L
            EZ = EZ + Len(L)
            Print #1, L
            STR2 = InStrRev(L, "\")
            If EZ > EX And OLSTR <> STR2 Then TRIGX = True
            
            OLSTR = STR2
            
            If EZ > EX And TRIGX = True Then
                TRIGX = False
                FLAGQUART = FLAGQUART + 1
                
                EZ = 0
                Close #1
                Open File1.Path + "\IRFAN SLIDESHOW-QUART--" + Format(FLAGQUART, "00") + "--.txt" For Output As #1

            End If
            Loop Until EOF(2)
        Close #2
    
    End If
    
Next

Close #1, #2

'File1.Pattern = "IRFAN SLIDESHOW*.*"
File1.ListIndex = 1
LAB2 = False
Call Form_Load
Shell "C:\Program Files\IrfanView\i_view32.exe /killmesoftly", vbHide

'--- NOT WITH RELOAD
Shell "C:\Program Files\IrfanView\i_view32.exe /slideshow=""D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt"" /fs /silent /one", vbMinimizedFocus

End Sub

Private Sub Label4_Click()

LAB2 = True
Set FS = CreateObject("Scripting.FileSystemObject")

DoEvents

'
'For R = 0 To File1.ListCount - 1
'    If R + 1 < File1.ListCount Then File1.ListIndex = R + 1
'    Set F = fs.GetFILE(File1.Path + "\" + File1.List(R))
'    If F.Size > 0 And Len(File1.List(R)) = 21 Then
'        EX = EX + F.Size
'    End If
'Next

'EX = EX / 4
'EZ = 0


'AT TOP FORM LOAD

'TRIGX = False


FILEVAR1 = File1.Path + "\IRFAN SLIDESHOW-DATE RECENT -- " + Format(FLAGQUART, "000") + " DAY SET.txt"
FR1 = FreeFile
Open FILEVAR1 For Output As #FR1
For R = 0 To File1.ListCount - 1
    If R + 1 < File1.ListCount Then File1.ListIndex = R
    Set F = FS.GetFILE(File1.Path + "\" + File1.List(R))
    
    If F.Size > 0 And Len(File1.List(R)) = 21 Then
    
        Open File1.Path + "\" + File1.List(R) For Input As #2
            
            
            Do
            U = Time
            If OU <> U Then DoEvents
            OU = U
            
            Line Input #2, L
            If FS.FILEExists(L) Then
                Set FD = FS.GetFILE(L)
                
                XD = FD.DateLastModified
                
                If XD > Now - FLAGQUART Then
                    
                    Print #FR1, L
                    EXCOUNT = EXCOUNT + 1
                    Label5 = EXCOUNT
                End If
            
            End If
            
            Loop Until EOF(2)
        Close #2
    
    End If
    
Next

Close FR1, #2
    
    

Set F = FS.GetFILE(FILEVAR1)
F.Copy File1.Path + "\IRFAN SLIDESHOW.txt"




'File1.Pattern = "IRFAN SLIDESHOW*.*"
File1.ListIndex = 1
LAB2 = False
Call Form_Load
Shell "C:\Program Files\IrfanView\i_view32.exe /killmesoftly", vbHide

'--- NOT WITH RELOAD
Shell "C:\Program Files\IrfanView\i_view32.exe /slideshow=""D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt"" /fs /silent /one", vbMinimizedFocus


End Sub

Private Sub Label6_Click()


Label6.Caption = "SELECT && ADD RECENT - TO GO"
If Label6.BackColor <> &H80FFFF Then
    Label6.BackColor = &H80FFFF
    Exit Sub
End If
If File1.ListIndex = -1 Then Exit Sub

FR1 = FreeFile
FILEVAR1 = File1.Path + "\IRFAN SLIDESHOW.txt"
FR2 = FreeFile
FILEVAR2 = File1.Path + "\IRFAN SLIDESHOW-DATE RECENT -- " + Format(FLAGQUART, "00") + " DAY SET.txt"
FR3 = FreeFile
FILEVAR3 = File1.Path + "\" + File1.List(File1.ListIndex)


FR1 = FreeFile
Open FILEVAR1 For Output As #FR1
FR2 = FreeFile
Open FILEVAR2 For Input As #FR2
FR3 = FreeFile
Open FILEVAR3 For Input As #FR3
    
    Do
        Line Input #FR2, L
        Print #FR1, L
    Loop Until EOF(2)
    
    Do
        Line Input #FR3, L
        Print #FR1, L
    Loop Until EOF(3)
    
Close FR1, FR2, FR3
    
    
Call Form_Load

Shell "C:\Program Files\IrfanView\i_view32.exe /killmesoftly", vbHide

'--- NOT WITH RELOAD
Shell "C:\Program Files\IrfanView\i_view32.exe /slideshow=""D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt"" /fs /silent /one", vbMinimizedFocus


End Sub

Private Sub Label7_Click()
If Label7.BackColor <> &H80FFFF Then
    Label7.BackColor = &H80FFFF
    Exit Sub
End If


End Sub

Private Sub Label8_Click()
FR1 = FreeFile
FILEVAR1 = "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW-AVI.txt"
FR2 = FreeFile
FILEVAR2 = "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW-AVI.txt.TMP"



FR1 = FreeFile
Open FILEVAR1 For Input As #FR1
FR2 = FreeFile
Open FILEVAR2 For Output As #FR2

    
    Do
        Line Input #FR1, L
        If InStr(UCase(L), "PROGRAM F") = 0 Then Print #FR2, L
    Loop Until EOF(1)
    
    
Close FR1, FR2, FR3
    
'D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW-AVI.txt
'C:\Program Files\
End Sub

Private Sub MNU_CREATE_SCRIPT_FILM_AVI_MP4_Click()

NOTEXECUTEINFRAR = True

Shell "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDE FILE LIST BATCH PIPE GENERATOR -- AVI -- .BAT", vbNormalFocus
Shell "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDE FILE LIST BATCH PIPE GENERATOR -- MP4 -- .BAT", vbNormalFocus
Shell "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDE FILE LIST BATCH PIPE GENERATOR -- SETA --.BAT", vbNormalFocus
Shell "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDE FILE LIST BATCH PIPE GENERATOR.BAT", vbNormalFocus
Shell "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDE FILE LIST BATCH PIPE GENERATOR AUTOPIX.BAT", vbNormalFocus
'Shell "", vbNormalFocus
'Shell "", vbNormalFocus
'Shell "", vbNormalFocus

'--- REMOVE FRAME STREAM  NONSE
Call Label8_Click

'--- COMBINE ALL
Call Label2_Click
'--- CUT QUARTERED
Call Label3_Click
'--- DATE RECENT ONE MONTH
Call Label4_Click
'--- REMOVE FRAME STREAM  NONSE
Call Label8_Click


'NOTEXECUTEINFRAR

'---
'Call
'---
'Call

End Sub
