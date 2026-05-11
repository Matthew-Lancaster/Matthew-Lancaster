VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Real Mans Sort..."
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "** FINSIHED HORAY **"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   1065
      TabIndex        =   7
      Top             =   1560
      Width           =   5970
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   3870
      TabIndex        =   6
      Top             =   720
      Width           =   390
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   3870
      TabIndex        =   5
      Top             =   0
      Width           =   390
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Do All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   2040
      TabIndex        =   2
      Top             =   990
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Sen 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send To && Office ToolBar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   3345
      TabIndex        =   3
      Top             =   945
      Visible         =   0   'False
      Width           =   6840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Favorites"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   6555
      TabIndex        =   4
      Top             =   15
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send To"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   4155
      TabIndex        =   1
      Top             =   -30
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Office ToolBar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   -15
      TabIndex        =   0
      Top             =   -75
      Visible         =   0   'False
      Width           =   3900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lovely Code
'I'll be back to use this for more

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()

Set FS = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
For Each Control In Controls
    Control.AutoSize = False
    If Control.Enabled = True Then
    If Control.Width + Control.Left > xx Then xx = Control.Width + Control.Left
    If Control.Height + Control.Top > yy Then yy = Control.Height + Control.Top
    If Control.Width > wx Then wx = Control.Width + 1500
    End If
Next

'pesky control list backwards
xv = 0
For Each Control In Controls
    If Control.Enabled = True Then
    xvv = Control.Height
    xv = xv + xvv
    Control.Width = wx
    Control.Left = 0
    End If
Next
xv = xv - xvv

For Each Control In Controls
    If Control.Enabled = True Then
        Control.Top = xv
        xv = xv - Control.Height
    End If
Next

xx = 0: yy = 0
For Each Control In Controls
    If Control.Enabled = True Then
    If Control.Width + Control.Left > xx Then xx = Control.Width + Control.Left
    If Control.Height + Control.Top > yy Then yy = Control.Height + Control.Top
    End If
Next

On Error GoTo 0

Label8 = "Ready"
Form1.Width = xx + 100
Form1.Height = yy + 400



AT$ = Command$

If AT$ <> "" And Mid$(AT$, 1, 1) = """" Then
AT$ = Mid$(AT$, 2)
AT$ = Mid$(AT$, 1, Len(AT$) - 1)
End If
'Label1.Caption = at$

If AT$ = "" Then AT$ = "E:\00 Office ToolBar" 'AT$ = "E:\01 SendTo\"


If Dir$(AT$) <> "" Then

'MsgBox "Yes File"

AT$ = Mid$(AT$, 1, InStrRev(AT$, "\"))

'MsgBox at$

End If

'If Dir$(at$) = "" Then MsgBox "No Is Folder not file"

'MsgBox (at$)
'End



If AT$ <> "" Then
ScanPath.txtPath.Text = AT$
'txtPath.Text = "C:\01 SendTo\"

On Local Error Resume Next
ChDir AT$
If Err.Number > 0 Then MsgBox ("Error with Dir Existance") ': End
'ScanPath.Show


'SetBackCol
'Label3.BackColor = &HC000&

'LAUGHT TOO HARD
'STORM TROOPER -- DEBBIE DONT LIKE THREAT DRUGS
'MATT HUNTER -- LOOK AT TEXT LOADS FAST THROUGH
'


If Mid(AT, Len(AT), 1) = "\" Then
    AT = Mid(AT, 1, Len(AT) - 1)
End If

Source = AT$ '"E:\01 Office ToolBar"
Call Jeepers(Source)


'Call cmdScan_Click
'Call Jeepers

End If

If AT$ = "" Then
MsgBox "Nothing To Do Auto Sort Folder"
'End
End If
'end



Exit Sub

'Call Jeepers

End Sub

Sub Jeepers(Source)
Dim F1, F2
'ScanPath.Show


ScanPath.cboMask.Text = "*.*"
'ScanPath.cboMask.Text = "*.url"

F1 = Source
F2 = Mid$(Source, 1, 3) + "Temp\" + Mid$(Source, 4)

Label8 = "1 of 3 Move There"
G3 = "There"

Call SendToDestination(F1, F2)

'Offy Bar needs a pause
If F1 = "E:\00 Office ToolBar\" Then
Label8 = "Pause For Offy Bar"
DoEvents
Sleep 400
End If

Label8 = "1 of 3 Cleanse Up After"
G3 = "Cleansed"
DoEvents
Call CleanEmptys(F1)


Label8 = "2 of 3 Move Back"
G3 = "Back"

Call SendToDestination(F2, F1)

Label8 = "3 of 3 Cleanse Up After"
G3 = "Cleansed"
DoEvents
Call CleanEmptys(F2)

Label8 = "** FINSIHED HORAY **"

End Sub

Sub SendToDestination(Source1, Source2)

If Mid(Source1, Len(Source1)) <> "\" Then Source1 = Source1 + "\"
If Mid(Source2, Len(Source2)) <> "\" Then Source2 = Source2 + "\"

On Local Error Resume Next
ChDir Source2
If Err.Number > 0 Then
    Err.Clear
    MkDir Source2
    a = Err.Description
    If Err.Number > 0 Then MsgBox "Error Make First Dir": Stop
End If
On Local Error GoTo 0

ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath.Text = Source1
Call ScanPath.cmdScanDir_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    
    ScanPath.lblCount2 = Str(we)
    
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1) + ScanPath.ListView1.ListItems.Item(we)

    'Make Empty Dir Structure at Destination
    ets1 = Len(ScanPath.txtPath.Text)
    If Mid$(A1$, ets1 + 1) <> "" Then
        c1$ = Source2 + Mid$(A1$, ets1 + 1)
    Else
        c1$ = ""
    End If
    
    On Local Error Resume Next
    Err.Clear
    ChDir c1$
    If Err.Number > 0 And c1$ <> "" Then
        Err.Clear
        MkDir c1$
        If Err.Number > 0 Then MsgBox "Error Make Dir": Stop
    End If
    
Next
    
ScanPath.ListView1.ListItems.Clear
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    
    ScanPath.lblCount2 = Str(we)
    'Get Path
    ets1 = Len(ScanPath.txtPath.Text)
    If Mid$(A1$, ets1 + 1) <> "" Then
        c1$ = Source2 + Mid$(A1$, ets1 + 1)
    Else
        c1$ = Source2
    End If
    
    On Local Error Resume Next
    Err.Clear
    'use of this G$ is the alternate name DOS 8.3 name this will mean it can copy any file even
    'with dodge alpha Chars like Unicode
    Set f = FS.getfile(A1$ + G1$)
    
    
    f.Move c1$
    If Err.Number > 0 Then
        MsgBox ("Error Moving File")
        Stop
    End If
    On Local Error GoTo 0

Next

End Sub

Sub CleanEmptys(Source1)

If Mid(Source1, Len(Source1)) <> "\" Then Source1 = Source1 + "\"


If Source1 = "E:\01 Office ToolBar\" Then Exit Sub

Set f = FS.GetFolder(Source1)
    
'If Total succeeed then left with Empty Folder
If f.Size = 0 Then
    f.Delete True
    Exit Sub
End If

'If Some Left then del folders that are just empty
ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath.Text = Source1
Call ScanPath.cmdScanDir_Click

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1) + ScanPath.ListView1.ListItems.Item(we)
    'b1$ = ListView1.ListItems.Item(we)
    ScanPath.lblCount2 = Str(we)
    
    On Local Error Resume Next
    Err.Clear
    Set f = FS.GetFolder(A1$)
    
    If f.Size = 0 Then
        Err.Clear
        f.Delete True
        If Err.Number > 0 Then MsgBox "Error Delete Dir in Single Mode": Stop
    End If

Next

End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload ScanPath
End Sub

Sub SetBackCol()
    Label1.BackColor = &HFF0000
    Label2.BackColor = &HFF0000
    Label3.BackColor = &HFF0000
    Label4.BackColor = &H800000
    'Label5.BackColor = &H800000
    Label6.BackColor = &HFF0000
    Label7.BackColor = &HFF0000
    Label8.BackColor = &HFF0000
End Sub

Private Sub Label1_Click()
If Label8 <> "Ready" And Label8 <> "** FINSIHED HORAY **" Then Exit Sub
SetBackCol
Label1.BackColor = &HC000&
Source = "E:\01 FAVS"
Call Jeepers(Source)

End Sub

Private Sub Label2_Click()
If Label8 <> "Ready" And Label8 <> "** FINSIHED HORAY **" Then Exit Sub
SetBackCol
Label2.BackColor = &HC000&
 
Source = "C:\01 SendTo"
Call Jeepers(Source)

End Sub

Private Sub Label3_Click()
If Label8 <> "Ready" And Label8 <> "** FINSIHED HORAY **" Then Exit Sub
SetBackCol
Label3.BackColor = &HC000&

Source = "E:\00 Office ToolBar"
Call Jeepers(Source)

End Sub

Private Sub Label4_Click()
If Label8 <> "Ready" And Label8 <> "** FINSIHED HORAY **" Then Exit Sub
SetBackCol
Label4.BackColor = &HC000&

Label3.BackColor = &HC000&
Source = "E:\01 Office ToolBar"
Call Jeepers(Source)

Label2.BackColor = &HC000&
Source = "C:\01 SendTo"
Call Jeepers(Source)

Label1.BackColor = &HC000&
Source = "E:\01 FAVS"
Call Jeepers(Source)

End Sub

Private Sub Label5_Click()
If Label8 <> "Ready" And Label8 <> "** FINSIHED HORAY **" Then Exit Sub
SetBackCol
Label5.BackColor = &HC000&

Label3.BackColor = &HC000&
Source = "E:\01 Office ToolBar"
Call Jeepers(Source)
Label2.BackColor = &HC000&
Source = "C:\01 SendTo"
Call Jeepers(Source)


End Sub
