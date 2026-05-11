VERSION 5.00
Begin VB.Form Frm1_Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Hoggs Menu"
   ClientHeight    =   4860
   ClientLeft      =   48
   ClientTop       =   732
   ClientWidth     =   9960
   Icon            =   "Frm_Process_Hog_Menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   9960
   Visible         =   0   'False
   Begin VB.ListBox List4 
      Height          =   1584
      Left            =   4845
      TabIndex        =   4
      Top             =   585
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox List3 
      Height          =   1776
      ItemData        =   "Frm_Process_Hog_Menu.frx":0442
      Left            =   3045
      List            =   "Frm_Process_Hog_Menu.frx":0444
      TabIndex        =   3
      Top             =   510
      Width           =   1680
   End
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   2760
      Top             =   3120
   End
   Begin VB.Timer KILL_HOGGS 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2205
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1485
      Top             =   3045
   End
   Begin VB.ListBox List2 
      Height          =   1584
      ItemData        =   "Frm_Process_Hog_Menu.frx":0446
      Left            =   1635
      List            =   "Frm_Process_Hog_Menu.frx":0448
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   555
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox List1 
      Height          =   1776
      ItemData        =   "Frm_Process_Hog_Menu.frx":044A
      Left            =   240
      List            =   "Frm_Process_Hog_Menu.frx":044C
      TabIndex        =   0
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PICK AN EXE OFF TO KILL (OR CLOSE (LATER TO CODE IN))"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   165
      Width           =   4560
   End
   Begin VB.Menu MNU_MENU 
      Caption         =   "MENU"
      Begin VB.Menu MNU_NOTEPAD 
         Caption         =   "OPEN NOTEPAD SCRIPT TO HOGG KILL"
      End
      Begin VB.Menu MNU_Process_NirsoftLastExecute 
         Caption         =   "PROCESS NIRSOFT ExecutedProgramsList PASTE DUMP"
      End
   End
   Begin VB.Menu MNU_KILL_HOGGS 
      Caption         =   "KILL HOGGS"
   End
   Begin VB.Menu MNU_VBME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_EXE_COUNT 
      Caption         =   "EXE Count"
   End
   Begin VB.Menu MNU_EXE_COUNT_PROCESS 
      Caption         =   "EXE COUNT"
   End
   Begin VB.Menu MNU_HIDE 
      Caption         =   "HIDE ME"
   End
   Begin VB.Menu MNU_CLOSE 
      Caption         =   "CLOSE"
   End
   Begin VB.Menu MNU_END 
      Caption         =   "END"
   End
End
Attribute VB_Name = "Frm1_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vFile1 = App.Path + "\X SCRIPT FILE PROCESS HOGGS.TXT"

'# -------------------------------------------------------------------
'# --- PROJECT PROCESS HOGG KILLER -----------------------------------
'# --- WHEN YOUR SYSTEM STOPS RESPONDING IN PAGEFILE AND VIRUAL MEMORY
'# --- RUNNING HARD DRIVE CONSTANTLY ACCESSING AND MOUSE STUCK AND
'# ---- TIME STUCK AND CAPS LOCK STUCK
'# --- MATT.LAN@BTINTERNET.COM ---------------------------------------
'# --- 11-12 JAN 2016 MIDNIGHT ---------------------------------------
'# -------------------------------------------------------------------
'# --- REM LINES ARE TRIMMED
'# # = HASH REM LINE ARE ALLOWED
'# --- LINE THAT BEGIN SPACES AND OR EMPTY CRLF ARE ALLOWED
'# --- LINE THAT BEGIN WITH * ALLOWED FOR CODING PUROSE OVER-RIDE EXECPTIONS
'# ### LINE BEGINE TRIPPLE HASH'S ARE ALLOWED NOT TO MODIFY FOR CODE PURPOSE
'# -------------------------------------------------------------------------
'# PASTE THE *EXECUTED-PROGRAMS-LIST* FROM NIRSOFT
'# FOR YOUR MOST RECENT PROGRAMS THT BEEN RUN
'# 1. PASTE THEM IN NOTEPAD - USE MENU OPTION TO FIND
'# 2. PROCESS THEM FOR
'#    1. 1. STRIP THE EXE OUT FROM FIRST TAB CHR(9)
'#       2. KEEP THE # HASHES REMARKS AND SPACES AND *'S
'#    2. REVERSE SORT OLDEST AT END AFTER TEXT FIRST - FOR DUPLICATES
'#       NEWEST AT BEGINING REMAINS INTACT
'#
'# - NOTE -- KEEP YOUR MENU CATAGORYS FOR MOST IMPORTANT TO BATCH KILL
'# - NOTE -- NIRSOFT EXECUTED-PROGRAMS-LIST - DONT KEEP DUPLICATES
'# - ------- BUT BE INFORMED THE PRIORITY OF KEPTED DUPES IS OLDEST
'# -         LAST IN TEXT SCRIPT REMOVED FIRST
'# - ------- THAT IS WHAT YOU WANT WHEN ADD SOME MORE IN BATCH PASTE
'# - ------- OBVIOUSLY WHEN ADD A BATCH MORE YOU MIGHT WANT SOME LEFT OUT
'# - ------- SO I PUT EXECPTIONS IN
'# - NOTE -- ONLY CHECKS FOR DUPES IN EXECUTABLES NOT COMMENTS
'# - ------- A 3 HASH AT BEGINING OF LINE USED FOR HARD CODED CATAGORYS
'# - NOTE -- WHEN EDIT THIS FILE A BACKUP GETS MADE INFINATLY TO DATE
'# -         AND SECOND
'# --------------------------------------------------------------------
'# --------------------------------------------------------------------
'
'
'# -------------------------------------
'### EXCEPTIONS
'# -------------------------------------
'EXE FILES AND PATH
'C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE


'# -------------------------------------
'### DOUBLE CHECK AND ASK EXCEPTIONS
'# -------------------------------------
'# NOTE - PUT A * AT BEGINING OF LINE TO STOP DUPE CHECKER
'# ------ OBVIOUSLY DIY FOR DUPE IN THIS PART
'# -------------------------------------
'
'*C:\WINDOWS\Explorer.EXE

'# -------------------------------------
'
'### BEGIN
'
'# NOTE - TOP PRIORTY BIG MEMORY HOGGS INTO PAGEFILE GRIND TO A HALT
'
'C:\Program Files\Siber Systems\GoodSync\GoodSync.exe
'C:\PROGRAM FILES\FREEFILESYNC\Bin\FREEFILESYNC_XP.EXE
'C:\Program Files\IrfanView\i_view32.exe

'# ----
'
'### KILL FOR THE FUN OF IT LIKE AT BOOT
'### KILLFUN
'C:\WINDOWS\Explorer.EXE

'# ----
'
'### LOAD THESE BACK AFTER KILL
'### KILL_LOAD_BACK
'*C:\WINDOWS\Explorer.EXE

Dim FS

Dim vFile1DATE

Public EXIT_TRUE

Dim TOTALExeCount
Dim OLDWINSTATE
Dim ExeCount
Dim OLDPIDVAR
Dim PROCESS_SCRIPT As String


'PID HAS TO BE -1 -- Process Id Return in PID
'Var - False if Not Exist or PID Remain -1
'----
'USAGE = ----
'PID = -1
'Var = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
'If PID <> -1 Then
'    Call Process_HIGH_PRIORITY_CLASS(PID)
'End If
'USAGE = ----
'----

Private cProcesses As New clsCnProc
Dim PID As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1
Dim vFile1
Dim ScriptVar1


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************


Private Sub Form_Load()


vFile1 = App.Path + "\X SCRIPT FILE PROCESS HOGGS.TXT"

List3.AddItem "Status Window"
List3.AddItem "-------------"
List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- Program Start"
List3.AddItem "-------------------------------------------------------"


Call Load_XScript

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- LOAD SCRIPT"
List3.AddItem "-------------------------------------------------------"


'--------------

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- GRAB FIRST PROCESSES"
List3.AddItem "-------------------------------------------------------"

Call Timer3_Timer

'--------------

OLDPIDVAR = Len(PROCESS_SCRIPT)
MNU_EXE_COUNT_PROCESS.Caption = "Exe Process Count =" + Str(PID)
Frm2_Process_Hog.Label4.Caption = Str(PID)

Timer1.Enabled = True

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- ENABLE TIMER CHECK SCRIPT AT FORM LOAD"
List3.AddItem "-------------------------------------------------------"

Call MNU_Norm_Click

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- POSITION FORM"
List3.AddItem "-------------------------------------------------------"

Load Frm2_Process_Hog

If IsIDE = True Then Me.Show

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- BOTH FORM'S LOADED UP AND RUNNING"
List3.AddItem "-------------------------------------------------------"

End Sub

Private Sub MNU_Norm_Click()
   
    On Error Resume Next
    
    'If FORM_LOAD_FLAG = False Then Me.WindowState = vbNormal
    
    If Me.WindowState <> vbMinimized Then
        Me.Width = Screen.Width - 2000
        Me.Height = Screen.Height - 3500
        
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = ((Screen.Height - Me.Height) / 2) - 900
    End If
    
    'Call Form_Resize
    
    'Call MNU_Norm_Click
    
    Label1.Left = 0
    Label1.Top = 0
    Label1.Width = Me.Width
    
    'End
    List1.Left = 0
    List1.Top = Label1.Height
    List1.Width = Me.Width / 2
    List1.Height = Me.Height - Label1.Height
    
    List3.Left = List1.Width + 20
    List3.Top = Label1.Height
    List3.Width = Me.Width / 2
    List3.Height = Me.Height - Label1.Height
    
    Me.Height = List1.Top + List1.Height + 300 + 480

End Sub


Private Sub Mnu_Center_Click()
    'Me.WindowState = vbMaximized
    If Me.WindowState = vbMaximized Then MsgBox "Not in Windowstate = vbMaximized"
    'Call Form_Resize
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    'Call Form_Resize
    'OLTLWH = 0

End Sub


Private Sub Form_Resize()
If OLDWINSTATE = Me.WindowState Then Exit Sub
OLDWINSTATE = Me.WindowState

Call MNU_Norm_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

'If Me.WindowState <> vbMinimized And EXIT_TRUE = False Then
'    Me.WindowState = vbMinimized
'    'test may have to put back form need reseting
'    'Unload FrmJoy
'    Cancel = True
'    Exit Sub
'End If

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- FORM 2 MENU UNLOAD"
List3.AddItem "PROCESS ID" + Str(PROID) + " -- " + LINESTR2
List3.AddItem "------------------------------------------------------------"
                        

Cancel = False
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_CLOSE_Click()

EXIT_TRUE = True

Dim Form As Form
For Each Form In FORMS
    Unload Form
    Set Form = Nothing
Next Form

End Sub

Private Sub MNU_END_Click()
End
End Sub

Private Sub mnu_exe_count_Click()
' mnu_exe_count
End Sub

Private Sub MNU_HIDE_Click()
Me.Visible = False
End Sub

Public Sub MNU_KILL_HOGGS_Click()

KILL_HOGGS.Enabled = True

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- Demand Button -- Kill_Usage_Hoggs"
List3.AddItem "--------------------------------------------------------------------"

End Sub



Private Sub MNU_NOTEPAD_Click()

If Dir(vFile1) = "" Then
    FR1 = FreeFile
    Open vFile1 For Append As #FR1
    Close #FR1
End If


ShellExecute Me.hwnd, "open", vFile1, vbNullString, vbNullString, conSwNormal


End Sub

Public Sub SET_VAR_FS()

Set FS = CreateObject("Scripting.FileSystemObject")

End Sub

Sub Load_XScript()

Call SET_VAR_FS

Set F = FS.getfile(vFile1)
vFile1DATE2 = F.datelastmodified
Set F = Nothing

If vFile1DATE <> vFile1DATE2 Then
    
    vFile1DATE = vFile1DATE2
    List4.Clear
    
    FR1 = FreeFile
    Open vFile1 For Input As #FR1
    Do
        Line Input #FR1, LINESTR1
        
        List4.AddItem LINESTR1
        
        
    Loop Until EOF(FR1)
    Close #FR1
    
    Call MNU_Process_NirsoftLastExecute_Click
    
End If
    
    
ExeCount = 0
List1.Clear
List2.Clear
    
    
For R = 0 To List4.ListCount - 1
    
    LINESTR1 = List4.List(R)
        
    XGO = 0
    If Mid(LINESTR1, 2, 2) = ":\" Then XGO = 2
    
    If Mid(LINESTR1, 1, 1) = "*" Then XGO = 3
    If Mid(LINESTR1, 1, 1) = "+\" Then XGO = 3
    If Mid(LINESTR1, 1, 1) = "-\" Then XGO = 3
    
    If Mid(LINESTR1, 1, 1) = "#\" Then XGO = 1
    If Trim(LINESTR1) = "" Then XGO = 1
        
    If XGO > 0 Then
    
        List1.AddItem Trim(LINESTR1)
        
        If Mid(LINESTR1, 2, 2) = ":\" Then
            
            MNU_EXE_COUNT.Caption = "Exe Count =" + Str(ExeCount)
            List1.AddItem Trim(LINESTR1)
            'List2.AddItem Trim(LINESTR1)
        'ExeCount = ExeCount + 1
        'MNU_EXE_COUNT.Caption = "Exe Count =" + Str(ExeCount)
        
        End If
        End If
Next

'CHECK MORE - IF EXIST ON DRIVE AND IN MEMORY


TOTALExeCount = ExeCount

End Sub





Private Sub MNU_Process_NirsoftLastExecute_Click()

'CHECK FOR TABS

For R = 0 To List4.ListCount - 1
    LINESTR1 = List4.List(R)
    If Mid(LINESTR1, 2, 2) = ":\" And InStr(LINESTR1, Chr(9)) > 0 Then
    
        List4.List(R) = Mid(LINESTR1, 1, InStr(LINESTR1, Chr(9)) - 1)
        CHANGEMADE1 = True
        CHANGEMADE = True
    End If
Next

'CHECK FOR DUPES
'ONLY RAW PATHS NOT SPECIAL FUNCTION PRE ID'ED

'YES LIKE :\
'NOT LIKE + - * # (SPACES LINES)

CHUNK = "&"

For R = List4.ListCount - 1 To 0 Step -1
    If Mid(List4.List(R), 2, 2) = ":\" Then
        CHUNK = CHUNK + Chr(255) + Format(R, "0000") + LCase(List4.List(R))
    End If
Next

For R = List4.ListCount - 1 To 0 Step -1
    If Mid(List4.List(R), 2, 2) = ":\" Then
        VAR1 = 1
        Do
        VAR1 = InStr(VAR1, CHUNK, LCase(List4.List(R)))
            If VAR1 > 0 Then
                IndexVAR = Val(Mid(CHUNK, VAR1 - 4, 4))
                If IndexVAR <> R Then
                    List4.RemoveItem (R)
                    CHANGEMADE2 = True
                    CHANGEMADE = True
                End If
                VAR1 = VAR1 + 7
            
            End If
        Loop Until VAR1 = 0
    End If
Next


If CHANGEMADE = False Then Exit Sub

If CHANGEMADE1 = True Then
    List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- CHANGE MADE TO XSCRIPT OF *TABS* DETECTED TRIMMED"
    List3.AddItem "------------------------------------------------------------"
End If

If CHANGEMADE2 = True Then
    List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- CHANGE MADE TO XSCRIPT OF *DUPES* DETECTED REMOVED"
    List3.AddItem "------------------------------------------------------------"
End If


Dim vFile2

'vFile1 = App.Path + "\X SCRIPT FILE PROCESS HOGGS.TXT"
vFile2 = App.Path + "\X SCRIPT FILE PROCESS HOGGS - " + Format(Now, "YYYY-MM-DD HH-MM-SS") + ".TXT"

On Error Resume Next

Name vFile1 As vFile2

If Err.Number > 0 Then
    MSGTEXT = "Error Rename a Backup for Changes to " + vbCrLf + vFile1 + vbCrLf + vFile2 + vbCrLf + Err.Description + vbCrLf + "Abort Continue to Save -- Do Again"
    MsgBox MSGTEXT
    GoTo END1
End If

Err.Clear

FR1 = FreeFile
Open vFile1 For Output As #FR1

If Err.Number > 0 Then
    MSGTEXT = "Error Open File for Changes to " + vbCrLf + vFile1 + vbCrLf + Err.Description + vbCrLf + "Abort Continue to Save -- Do Again"
    MsgBox MSGTEXT
    GoTo END1
End If
    
Err.Clear
    
For R2 = 0 To List4.ListCount - 1
    Print #FR1, List4.List(R2)
Next

If Err.Number > 0 Then
    MSGTEXT = "Error During Save File Data Changes to " + vbCrLf + vFile1 + vbCrLf + Err.Description + vbCrLf + "Check it Over and Maybe Do Again"
    MsgBox MSGTEXT
    GoTo END1
End If

Err.Clear

Close #FR1

If Err.Number > 0 Then
    MSGTEXT = "Error During Close File of Data Changes to " + vbCrLf + vFile1 + vbCrLf + Err.Description + vbCrLf + "Check it Over and Maybe Do Again"
    MsgBox MSGTEXT
    GoTo END1
End If

                        
List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- CHANGE MADE TO XSCRIPT - WRITE FILE BACK *SUCCESS*"
List3.AddItem "------------------------------------------------------------"
                        
Exit Sub

END1:

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- CHANGE MADE TO XSCRIPT - " + MSGTEXT
List3.AddItem "------------------------------------------------------------"

End Sub

Private Sub List1_Click()

'CALLED CLICK TO KILL

Dim LINESTR1 As String

STARTLIST1COUNT = List1.ListCount

'For ScriptVar1 = 0 To List1.ListCount - 1
    
    DoEvents

    LINESTR1 = List1.List(ScriptVar1)
    
    If InStr(LINESTR1, "###") > 0 Then XGO = 0
    If InStr(LINESTR1, "### BEGIN") > 0 Then XGO = 1
    If InStr(LINESTR1, "### KILLFUN") > 0 Then XGO = 1
    
    If XGO = 1 Or 1 = 1 Then
        
        
        If Mid(LINESTR1, 2, 2) = ":\" Or Mid(LINESTR1, 3, 2) = ":\" Then
        
            'MNU_EXE_COUNT.Caption = "Exe Count =" + Str(ExeCount) + " Checking for Existing" + Str(ScriptVar1) + " of" + Str(STARTLIST1COUNT)
            
            If Mid(LINESTR1, 1, 1) = "*" Then LINESTR1 = Mid(LINESTR1, 2)
            If Mid(LINESTR1, 1, 1) = "+" Then LINESTR1 = Mid(LINESTR1, 2)
            If Mid(LINESTR1, 1, 1) = "-" Then LINESTR1 = Mid(LINESTR1, 2)
            
            LINESTR2 = LCase(LINESTR1)
            LINESTR1 = LCase(Mid(LINESTR1, InStrRev(LINESTR1, "\") + 1))
            
            If InStr(PROCESS_SCRIPT, LINESTR1) > 0 Then
                
                If InStr(LCase(List1.List(List1.ListIndex)), LINESTR1) > 0 Then
                
                    ISPOINTER = 1
                    Do
                        PROID = InStr(ISPOINTER, PROCESS_SCRIPT, LINESTR1)
                        If PROID = 0 Then Exit Do
                        PROID2 = InStrRev(PROCESS_SCRIPT, vbCrLf, PROID)
                        PROID = InStr(PROID2, PROCESS_SCRIPT, "-")
                        PROID = Val(Mid(PROCESS_SCRIPT, PROID2 + 3, PROID - PROID2 - 3))
                        
                        ISPOINTER = PROID2 + Len(LINESTR2)
                        
                        'Call Process_HIGH_PRIORITY_CLASS(PROID)
                        
                        MsgBox "KILL PROCESS " + vbCrLf + "PROCESS ID" + Str(PROID) + vbCrLf + LINESTR2
                        
                        List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- KILL ONE PROCESS BY HITT SCRIPT BOX"
                        List3.AddItem "PROCESS ID" + Str(PROID) + " -- " + LINESTR2
                        List3.AddItem "------------------------------------------------------------"
                        
                    Loop Until 1 = 2
                End If

            End If
        End If
    End If
'Next

End Sub


Private Sub MNU_VBME_Click()

List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- LOAD VISUAL BASIC TO PROGRAM"
List3.AddItem App.Path + "\" + App.EXEName + ".vbp"
List3.AddItem "------------------------------------------------------------"

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus

EXIT_TRUE = True
Unload Me

End Sub

Private Sub Timer1_Timer()
'CALLED FROM FORM LOAD AND PROCEES COUNT CHANGE
'REFRESH
'Call Avon Calling
'Shell Out Splash the Cash

'REMOVE ANY EXE THAT ARE NOT THERE EXIST


ExeCount = TOTALExeCount
ExeCount = 0

Dim LINESTR1 As String

STARTLIST1COUNT = List1.ListCount

For ScriptVar1 = List1.ListCount - 1 To 0 Step -1

    DoEvents
    
    LINESTR1 = List1.List(ScriptVar1)
    
    If Mid(LINESTR1, 2, 2) = ":\" Or Mid(LINESTR1, 3, 2) = ":\" Then
    
        MNU_EXE_COUNT.Caption = "Exe Count =" + Str(ExeCount) + " Checking for Existing" + Str(ScriptVar1) + " of" + Str(STARTLIST1COUNT)
        
        If Mid(LINESTR1, 1, 1) = "*" Then
            LINESTR1 = Mid(LINESTR1, 2)
        End If
        If Mid(LINESTR1, 1, 1) = "+" Then
            LINESTR1 = Mid(LINESTR1, 2)
        End If
        If Mid(LINESTR1, 1, 1) = "-" Then
            LINESTR1 = Mid(LINESTR1, 2)
        End If
        
        LINESTR1 = LCase(Mid(LINESTR1, InStrRev(LINESTR1, "\") + 1))
        
        If InStr(PROCESS_SCRIPT, LINESTR1) = 0 Then
            List1.RemoveItem (ScriptVar1)
            
            'Call Process_HIGH_PRIORITY_CLASS(PID)
            
            'ExeCount = ExeCount - 1
        
        Else
            
            
            If InStr(COMPvAR, LINESTR1) = 0 Then
            
                COMPvAR = COMPvAR + LINESTR1
            
                ExeCount = ExeCount + 1
        
            End If
        
        End If
    End If
Next




MNU_EXE_COUNT.Caption = "Exe Count =" + Str(ExeCount) + " RUN Existing"

Timer1.Enabled = False

'AFTER WE FINSISH OUR FILTER
'ANOTHER QUICK ONE
'REMOVE SOME DOUBLE LINE SPACES
For R = List1.ListCount - 1 To 0 Step -1
    If Trim(List1.List(R)) <> "" Then FLAG = 0
    If Trim(List1.List(R)) = "" Then FLAG = FLAG + 1

    If FLAG >= 2 Then
        List1.RemoveItem (R)
    
    End If
Next

End Sub


Public Sub KILL_HOGGS_Timer()

'CALLED FROM -- MNU_KILL_HOGGS_Click
'AND DETECT GRAINGING TO A HALT

Dim LINESTR1 As String
Dim PROID As Long, PROID2

For ScriptVar1 = 0 To List1.ListCount - 1
    
    DoEvents

    LINESTR1 = List1.List(ScriptVar1)
    
    If InStr(LINESTR1, "###") > 0 Then
        XGO = 0
    End If
    If InStr(LINESTR1, "### BEGIN") > 0 Then XGO = 1
    If InStr(LINESTR1, "### KILLFUN") > 0 Then XGO = 1
    
    If XGO = 1 Then
        
        
        If Mid(LINESTR1, 2, 2) = ":\" Or Mid(LINESTR1, 3, 2) = ":\" Then
        
            If Mid(LINESTR1, 1, 1) = "*" Then LINESTR1 = Mid(LINESTR1, 2)
            If Mid(LINESTR1, 1, 1) = "+" Then LINESTR1 = Mid(LINESTR1, 2)
            If Mid(LINESTR1, 1, 1) = "-" Then LINESTR1 = Mid(LINESTR1, 2)
            
            LINESTR2 = LINESTR1
            LINESTR1 = LCase(Mid(LINESTR1, InStrRev(LINESTR1, "\") + 1))
            
            If InStr(PROCESS_SCRIPT, LINESTR1) > 0 Then
                
                'MSGB = MSGB + "KILL PROCESS " + vbCrLf + "PROCESS ID" + Str(PROID) + vbCrLf + LINESTR2 + vbCrLf

                ISPOINTER = 1
                Do
                    PROID = InStr(ISPOINTER, PROCESS_SCRIPT, LINESTR1)
                    If PROID = 0 Then Exit Do
                    PROID2 = InStrRev(PROCESS_SCRIPT, vbCrLf, PROID)
                    PROID = InStr(PROID2, PROCESS_SCRIPT, "-")
                    PROID = Val(Mid(PROCESS_SCRIPT, PROID2 + 3, PROID - PROID2 - 3))
                    
                    ISPOINTER = PROID2 + Len(LINESTR2)
                    
                    'Call Process_HIGH_PRIORITY_CLASS(PROID)
                    
                    List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- KILL HOGGS "
                    List3.AddItem "PROCESS ID" + Str(PROID) + " -- " + LINESTR2
                    List3.AddItem "------------------------------------------------------------"
                    
                Loop Until 1 = 2

            End If
        End If
    End If
Next

'MsgBox "ALL THESE PROCESS KILLED " + Format(Now, "DD-MMM-YYYY HH:MM:SS") + vbCrLf + MSGB

KILL_HOGGS.Enabled = False


End Sub

Private Sub Timer3_Timer()

'CALLED FROM FORM_LOAD
'AND CONSTANT TIMER

Var = cProcesses.GetEXEID_WILDCARD_PROCESS_SCRIPT(PID, PROCESS_SCRIPT)
PROCESS_SCRIPT = LCase(PROCESS_SCRIPT)

If OLDPIDVAR = 0 Then FLAGFIRSTRUN = True

If OLDPIDVAR = Len(PROCESS_SCRIPT) Then Exit Sub

OLDPIDVAR = Len(PROCESS_SCRIPT)
    
MNU_EXE_COUNT_PROCESS.Caption = "Exe Process Count =" + Str(PID)
Frm2_Process_Hog.Label4.Caption = Trim(Str(PID))

If FLAGFIRSTRUN = True Then
    List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- Process Count 1ST LOAD =" + Str(PID)
    List3.AddItem "------------------------------------------------------------"
Else
    List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- Process Count Changed =" + Str(PID)
    List3.AddItem "------------------------------------------------------------"
End If

Call Load_XScript

'REFRESH
Timer1.Enabled = True

Call SET_HIGH_LOW_PROCESSES

End Sub


Sub SET_HIGH_LOW_PROCESSES()

'CALLED FROM -- Timer3_Timer

Dim LINESTR1 As String

Dim PROID As Long, PROID2

For ScriptVar1 = 0 To List1.ListCount - 1
    
    DoEvents

    LINESTR1 = List1.List(ScriptVar1)
    
    If InStr(LINESTR1, "###") > 0 And XGO = 1 Then XGO = 0
    If InStr(LINESTR1, "### HIGH_LOW_PRIORITY") > 0 Then XGO = 1
    
    If XGO = 1 Then
        
        If Mid(LINESTR1, 2, 2) = ":\" Or Mid(LINESTR1, 3, 2) = ":\" Then
        
            If Mid(LINESTR1, 1, 1) = "*" Then LINESTR1 = Mid(LINESTR1, 2)
            If Mid(LINESTR1, 1, 1) = "+" Then LINESTR1 = Mid(LINESTR1, 2): HIGH = "HIGH"
            If Mid(LINESTR1, 1, 1) = "-" Then LINESTR1 = Mid(LINESTR1, 2): HIGH = "LOW"
            
                LINESTR2 = LINESTR1
                LINESTR1 = LCase(Mid(LINESTR1, InStrRev(LINESTR1, "\") + 1))
                
                If InStr(PROCESS_SCRIPT, LINESTR1) > 0 Then
                    ISPOINTER = 1
                    Do
                        PROID = InStr(ISPOINTER, PROCESS_SCRIPT, LINESTR1)
                        If PROID = 0 Then Exit Do
                        PROID2 = InStrRev(PROCESS_SCRIPT, vbCrLf, PROID)
                        PROID = InStr(PROID2, PROCESS_SCRIPT, "-")
                        PROID = Val(Mid(PROCESS_SCRIPT, PROID2 + 3, PROID - PROID2 - 3))
                        
                        ISPOINTER = PROID2 + Len(LINESTR2)
                        
                        'If LCase(XX) <> LINESTR2 Then
                        
                        'List1.List (ScriptVar1)
                        
                        If HIGH = "HIGH" Then
                            Call Process_HIGH_PRIORITY_CLASS(PROID)
                        Else
                            Call Process_LOW_PRIORITY_CLASS(PROID)
                        End If
                        
                        List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- Process Priority Set " + HIGH
                        List3.AddItem "PROCESS ID" + Str(PROID) + " -- " + LINESTR2
                        List3.AddItem "------------------------------------------------------------"
                        
                        'MsgBox "Process_HIGH_PRIORITY_CLASS " + vbCrLf + "PROCESS ID" + Str(PROID) + vbCrLf + LINESTR2
                        
                        'XX = LINESTR2
                    Loop Until 1 = 2
    
                'End If
            End If
        End If
    End If

Next



End Sub
