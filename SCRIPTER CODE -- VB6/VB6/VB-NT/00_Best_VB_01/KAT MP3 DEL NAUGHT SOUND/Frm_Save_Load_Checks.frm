VERSION 5.00
Begin VB.Form Frm_Save_Load_Checks 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer zzCheckTimer 
      Enabled         =   0   'False
      Interval        =   59000
      Left            =   1200
      Top             =   960
   End
End
Attribute VB_Name = "Frm_Save_Load_Checks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Form As Form
Public FORM_ME As Form


'----------------       Latest Version Of Save Chks
'#### REMEBER SWITCHES
'TUE 17 MAY
'--------------------------
'MAKE A MODUAL LIKE THIS
'NAME -- KAT_MP3_RELOAD_VAR
'---------------------
'FOR THIS PUBLIC VAR
'---------------------
'TO USE WITH THIS FORM
'Frm_Save_LOad_Checks
'---------------------

'------------------
Dim OCheckQ
'------------------


'Sub zzLoad_Checks()
'Put This In Declarations
'Dim OCheckQ
'Very Nice Code Will Save all your Check Boxes an Menu Checks and Values you can filter
'some out -- works best as I Know

'if you use menu checks you have to set them yourself on clicks
'If Mnu_VB.Checked = True Then Mnu_VB.Checked = False: Exit Sub
'If Mnu_VB.Checked = False Then Mnu_VB.Checked = True

'3 Parts
' ---
'1 Load
'2 Save
'3 Timer ' This Chks for Changes using XOR Hash
'   Best way I know with Timer Hardly CPU Usage Unfriendly

'zzCheckTimer.Enabled = True
'Dont Have Timer Enabled on Form Load

'Call Timer to Run On Form Unload ----- call zzCheckTimer_Timer
'Then you could set timer slow like 10 Secs - I Do
'-------------------------------
'-------------------------------



Public Sub zzLoad_Checks()

'zzCheckTimer.Enabled = False

Dim TH(), ATD, FR1

'LATER TO USE MULTI FORM


For Each Form In Forms
    If Form.Name = "Form1" Then
        Set FORM_ME = Form
        Exit For
    End If
Next Form


'If FORM_ME.Visible = False Then Exit Sub

ReDim TH(FORM_ME.Controls.Count * 3)

On Error Resume Next
I = 0
FR1 = FreeFile
Open App.Path + "\# DATA\" + GetComputerName + "\0TextData\ChkSettings.txt" For Input As #FR1
Do
    If Not EOF(FR1) Then Line Input #FR1, vv$
    TH(I) = vv$
    I = I + 1
    If Not EOF(FR1) Then Line Input #FR1, vv$
    TH(I) = Val(vv$)
    I = I + 1
    If Not EOF(FR1) Then Line Input #FR1, vv$
    TH(I) = Val(vv$)
    I = I + 1
Loop Until EOF(1)
Close #1
    
tit = I
For Each Control In FORM_ME.Controls
    ATD = 1
    
    ppi = LCase(Control.Name)
    If InStr(ppi, LCase("Combo")) Then ATD = 0
    If InStr(ppi, LCase("Check")) Then ATD = 0
    If InStr(ppi, LCase("Mnu")) Then ATD = 0
    If InStr(ppi, LCase("Menu")) Then ATD = 0
    
    XXD = -1
    For R = 0 To tit
        If Control.Name = TH(R) Then
            XXD = R: Exit For
        End If
    Next
    
    If XXD > 0 And ATD = 0 Then
        On Error Resume Next
        If TH(XXD + 1) <> 0 Then Control.Value = TH(XXD + 1)
        If TH(XXD + 2) <> 0 Then Control.Checked = TH(XXD + 2)
        'Dont Let People Play Around If you Want to Hard Code In Designer To Enable Chk This
        'This Lets Happen Eg If Th(xxd + 2) <> 0
        
        A1 = 0
        If Mid(LCase(Control.Name), 1, 3) <> "mnu" Then
            A1 = Control.Value
        End If
        B1 = 0
        B1 = Control.Checked
        OCheckQ = OCheckQ + Str(A1) + Str(Val(B1))
        On Error GoTo 0
    End If
Next
zzCheckTimer.Enabled = True
End Sub



Sub zzSave_Checks()
Dim A1, B1 As Integer, ATD
On Error Resume Next
'filename=
If FS.FolderExists(App.Path + "\# DATA\" + GetComputerName + "\0TextData") = False Then
    MkDir App.Path + "\# DATA"
    MkDir App.Path + "\# DATA\" + GetComputerName
    MkDir App.Path + "\# DATA\" + GetComputerName + "\0TextData"
End If

Open App.Path + "\# DATA\" + GetComputerName + "\0TextData\ChkSettings.txt" For Output As #1

For Each Control In FORM_ME.Controls
    Err.Clear
    A1 = 0
    A2 = 1 '=ERROR
    ATD = 1
    'If Mid(LCase(Control.Name), 1, 3) = "mnu" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 5) = "timer" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "pic" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "mmco" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "text" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "dir" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "line" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 5) = "label" Then ATD = 0
    If ATD = 1 Then
        A2 = 0
        A1 = Control.Value
        A2 = Err.Number
    End If
    
    Err.Clear
    B1 = 0
    B2 = 1 '=ERROR
    ATD = 1
    If Mid(LCase(Control.Name), 1, 5) = "timer" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "pic" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "mmco" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "text" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 3) = "dir" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 4) = "line" Then ATD = 0
    If Mid(LCase(Control.Name), 1, 5) = "label" Then ATD = 0
    If ATD = 1 Then
        B2 = 0
        B1 = Control.Checked
        B2 = Err.Number
    End If

    If A2 = 0 Or B2 = 0 Then
        Print #1, Control.Name
        Print #1, A1
        Print #1, B1
    End If
Next
Close #1

End Sub



Private Sub Form_Load()

Exit Sub
'NOT WORKING YET
'CONTROL FORM
'GET CONTROL NAMES BUT NOT VALUES


'-----------------------------------------------
Frm_Save_Load_Checks.zzLoad_Checks
'-----------------------------------------------
Frm_Save_Load_Checks.zzCheckTimer.Enabled = True
'-----------------------------------------------

End Sub

Private Sub Form_Unload(Cancel As Integer)

'----------------------
' ABORT USING THIS
' Call zzSave_Checks
'----------------------
' ONLY SAVE IF CHANGES
'----------------------
Call zzCheckTimer_Timer
'----------------------

End Sub

Public Sub zzCheckTimer_Timer()

Dim CHECKQ
On Error Resume Next
'Debug.Print FORM_ME.Name
CHECKQ = ""
For Each Control In FORM_ME.Controls
    CHECK_A1 = 0
    CHECK_A1 = Control.Value
    CHECK_B1 = 0
    CHECK_B1 = Control.Checked
'    Debug.Print Control.Name + Str(CHECK_A1) + Str(Val(CHECK_B1))
    CHECKQ = CHECKQ + Str(CHECK_A1) + Str(Val(CHECK_B1))
Next

'Debug.Print checkq + vbCrLf + OCheckQ

If CHECKQ = OCheckQ Then Exit Sub

OCheckQ = CHECKQ

'Debug.Print OCheckQ

Call Frm_Save_Load_Checks.zzSave_Checks

End Sub



