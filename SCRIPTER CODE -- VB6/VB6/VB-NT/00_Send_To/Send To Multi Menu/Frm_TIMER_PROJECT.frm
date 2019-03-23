VERSION 5.00
Begin VB.Form Frm_TIMER_PROJECT 
   BackColor       =   &H00800000&
   Caption         =   "Form2"
   ClientHeight    =   9036
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   10800
   LinkTopic       =   "Form2"
   ScaleHeight     =   9036
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4908
      Left            =   24
      TabIndex        =   2
      Top             =   2868
      Width           =   10728
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2340
      Top             =   1968
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Check Mate -- Wanna Sexy Date - Be One Second Two Second -- Matt.Lan@btinternet.Com -- Sat Sun 23 July 2kSixteen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   48
      TabIndex        =   3
      Top             =   24
      Width           =   10704
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Label2 FOLDER+FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1272
      Left            =   36
      TabIndex        =   1
      Top             =   1584
      Width           =   10716
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Label1 FOLDER+FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   36
      TabIndex        =   0
      Top             =   348
      Width           =   10716
      WordWrap        =   -1  'True
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB"
   End
   Begin VB.Menu Mnu_VB_Folder 
      Caption         =   "VB Folder"
   End
   Begin VB.Menu Mnu_Back_Form 
      Caption         =   "Back Form"
      Begin VB.Menu MNU_MAIN_FORM 
         Caption         =   "Back_To_--_Multi_Send_To_Menu_Main_Form"
      End
   End
   Begin VB.Menu MNU_PROCESS_OUTPUT 
      Caption         =   "Process_&&_Output"
   End
   Begin VB.Menu MNU_OPEN_INPUT_TEXT 
      Caption         =   "Open_Input_Text"
   End
   Begin VB.Menu MNU_OPEN_OUTPUT_TEXT 
      Caption         =   "Open_Output_Text"
   End
   Begin VB.Menu MNU_EXPLORER_FOLDER 
      Caption         =   "Explorer_Folder"
   End
   Begin VB.Menu MNU_CLIP_BOARD_OUTPUT 
      Caption         =   "Clipboard_Output"
   End
   Begin VB.Menu MNU_LAST_PROJECT 
      Caption         =   "Use_Last_Project Not Inc App.Path"
   End
   Begin VB.Menu MNU_USE_APP_PATH_PROJECT 
      Caption         =   "Use_App.Path_Project"
   End
   Begin VB.Menu MNU_LAST_PROJECT_HISTORY 
      Caption         =   "Open_Last_Project_History"
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "Ende_Exit"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "Frm_TIMER_PROJECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ID11, ID12, ID13, ID14

Dim F11, F12, F13, F14
Dim F41, F42, F43, F44

Dim F21 As Long
Dim F22 As Long
Dim F23 As Long

Dim InPut_Date1()
Dim Test__Date1()

Dim IS_PROJECT_DONE_VAR
Dim TOTAL_DATA_COUNT

Dim TIME__RESULT_____1()
Dim TIME__RESULT_MOD_2()
Dim TIME_TOTAL_______1()
Dim TIME_TOTAL___MOD_2()

Dim OUTPUT_RESULT_STRING
Dim FS

Dim LABEL_MODIFY_ALLOW_CHANGE

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Sub FORM_LOAD()

    Set FS = CreateObject("Scripting.FileSystemObject")
    
    Me.Caption = "Time Management -- Multi Schedual -- Project Time Cost -- Matt.Lan@btinternet.com"
    
    'Form1.Timer2.Enabled = False
    
    'Call Form1.GET_PATH_OR_FILE_PATH("E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT")
    'Call Form1.GET_PATH_OR_FILE_PATH(App.Path)
    'Call Form1.GET_PATH_OR_FILE_PATH("")
    
    'Call PROCESS_LOAD_DATA
    
    If Dir(App.Path + "\DATA_LAST_PROJECT.TXT") = "" Then MNU_LAST_PROJECT.Enabled = False
    If Dir(App.Path + "\DATA_LAST_PROJECT_LOGGER.TXT") = "" Then MNU_LAST_PROJECT_HISTORY.Enabled = False
    
    If IsIDE = True Then Mnu_VB_ME.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Me.Hide
    'Form1.Show
    'Form1.WindowState = vbNormal

    Unload Me
    Unload Form1


End Sub

Private Sub Label1_Change()
    
    If LABEL_MODIFY_ALLOW_CHANGE = True Then
        LABEL_MODIFY_ALLOW_CHANGE = False
        Exit Sub
    End If

    
    If InStr("*" + Frm_TIMER_PROJECT.Label1, "None Path/file Found") > 0 Then
        Label2.Caption = "Clipboard See Above"
    End If
    
    If Dir(Frm_TIMER_PROJECT.Label1) = "" Then
        Label2.Caption = "Project Timer Don't Exist -- Click On Here To Make One or Use Selection Menu"
        MNU_PROCESS_OUTPUT.Enabled = False
        MNU_OPEN_INPUT_TEXT.Enabled = False
        MNU_OPEN_INPUT_TEXT.Enabled = False
        MNU_OPEN_OUTPUT_TEXT.Enabled = False
        'MNU_EXPLORER_FOLDER.Enabled = False
        MNU_CLIP_BOARD_OUTPUT.Enabled = False
    End If
        
    
    
    
    'STRIP TO FOLDER PATH
    I4 = Frm_TIMER_PROJECT.Label1.Caption
    I4 = Mid(I4, 1, InStrRev(I4, "\"))
    If FS.FolderExists(I4) = True Then
        MNU_EXPLORER_FOLDER.Enabled = True
    End If
        
    DONE_FLAG = False
    If InStr(Frm_TIMER_PROJECT.Label1, "\PROJECT_TIMER") = 0 Then
        If FS.FolderExists(Frm_TIMER_PROJECT.Label1.Caption) = True Then
            
            I4 = Frm_TIMER_PROJECT.Label1.Caption
            If Right(I4, 1) <> "\" Then
                I4 = I4 + "\PROJECT_TIMER.TXT"
                DONE_FLAG = True
                LABEL_MODIFY_ALLOW_CHANGE = True
                Frm_TIMER_PROJECT.Label1.Caption = I4
            End If

        End If
        If DONE_FLAG = False And FS.FileExists(Frm_TIMER_PROJECT.Label1) = True Then
            I4 = Frm_TIMER_PROJECT.Label1.Caption
            I4 = Mid(I4, 1, InStrRev(I4, "\"))
            LABEL_MODIFY_ALLOW_CHANGE = True
            Frm_TIMER_PROJECT.Label1.Caption = I4
        End If
    End If
        
    
    If FS.FileExists(Frm_TIMER_PROJECT.Label1) = True Then
        MNU_PROCESS_OUTPUT.Enabled = True
        MNU_OPEN_INPUT_TEXT.Enabled = True
        MNU_OPEN_INPUT_TEXT.Enabled = True
        MNU_OPEN_OUTPUT_TEXT.Enabled = True
        MNU_EXPLORER_FOLDER.Enabled = True
        MNU_CLIP_BOARD_OUTPUT.Enabled = True
        
        Label1.AutoSize = True
        Label1.Refresh
        Label1.AutoSize = False
        Label1.Width = Me.Width - 300
        Label1.Height = Label1.Height + 20
        Label2.Top = Label1.Height + Label1.Top
        
        Label2 = "Project Ready"
        Label2.Height = 400
        'List1.Top = Label2.Height + Label2.Top + 20
        Label2.Visible = False
        List1.Top = Label2.Top + 20
        List1.Height = 5000 + 2200
        
        Me.Height = List1.Height + List1.Top + 1800
        
        Call MNU_PROCESS_OUTPUT_Click
    End If

End Sub

Private Sub Label2_Click()

'If Dir(Frm_TIMER_PROJECT.Label1) <> "" Then Exit Sub

If InStr("*" + Frm_TIMER_PROJECT.Label1, "None Path/file Found") > 0 Then
    Exit Sub
End If

'WRITE TEMPLATE FILE ---------------------------------------------


FR1 = FreeFile
Open Frm_TIMER_PROJECT.Label1 For Output As #FR1

    Print #FR1, "# MATT.LAN@BTINTERNET.COM"
    Print #FR1, "# SAT SUN -- 24 JULY"
    Print #FR1, "#     MON -- 08 AUGUST TIDY UP"
    Print #FR1, "# ------------------"
    Print #FR1, "# HI HERE IS THE LAY"
    Print #FR1, "#"
    Print #FR1, "# HASH ARE REM AND EMPTY LINE SPACE"
    Print #FR1, "# A lOT OF DATE FORMAT ARE PICKED UP WITH DATEVAULE"
    Print #FR1, "#"
    Print #FR1, "# SET A VARIABLE BIT AT TOP LIKE HERE WITH REM -- USE ONE LINE YES OT NO"
    Print #FR1, "#"
    Print #FR1, "# ------------------"
    Print #FR1, "PROJECT CURRENTLY ACTIVE NOW = NOT"
    Print #FR1, "# ------------------"
    Print #FR1, "# PROJECT CURRENTLY ACTIVE NOW = YES"
    Print #FR1, "# PROJECT CURRENTLY ACTIVE NOW = NOT"
    Print #FR1, "#"
    Print #FR1, "# WITH THIS VAR FLAG PROJECT ACTIVE WILL TAKE A LAST START TIME AND REPLACE LAST DATE VALUE WITH NOW TIME"
    Print #FR1, "#"
    Print #FR1, "# THIS IS EXAMPLE OF WHAT I USE"
    Print #FR1, "#"
    Print #FR1, "# 22-Jul-2015 12:09:37"
    Print #FR1, "# 23-Jul-2015 00:38:31"
    Print #FR1, "#"
    Print #FR1, "# SLOT ALL THESE DATE TOGETHER AT AN ORDER BEGIN LINE WITH REM"
    Print #FR1, "# FOR LIKE SESSION 1.. ONE 2 3 4 5 -- ACCUMALATOR"
    Print #FR1, "# AND THE LAST ONE CAN SET IN DOUBLE BLOCK WITH BEGIN END "
    Print #FR1, "# BUT END IGNOR IS PROJECT USE ACTIVE NOW"
    Print #FR1, "# HERE IS ONE TO GET GOING"
    Print #FR1, "#"
    Print #FR1, "# WHEN CODE RUN ONCE AND RESULT OUTPUT WILL BE CLIPBOARD IF WANT AND OR ALWAYS FILE NEARBY SIMULAR NAME"
    Print #FR1, "#"
    Print #FR1, "# IN ORDER TO DO BETTER DATE ERROR CHECKING AT PROCESS LOAD FILE SEPERATE DATE AND TIME WITH A SPACE AND NOT OTHER SPACE"
    Print #FR1, "#"
    Print #FR1, "# Ability to Use a Day 3 Char of Week Is Avail Keep Dash Between Will be Striped Out for Processor"
    Print #FR1, "# HAS TO BE STRIPPED OUT FOR PROCESSOR AS DATEVALUE DONT CALC A DATE WITH A DAY OF WEEK"
    Print #FR1, "#"
    Print #FR1, "# AS OF MON AUG 15 -- ABLE TO CREATE OWN PART FILENAME PROJECT OPEN IN FILE NOT OUTPUT YET"
    Print #FR1, "# FRONT OF FILENAME MUST BE "
    Print #FR1, "# \PROJECT_TIMER -- AND EXTENSION IS NOT SEARCHED MAKE EASIER CODER AJUST"
    Print #FR1, "#"
    Print #FR1, "# ------------------"
    Print #FR1, Format(Int(Now), "DDD-DD-MMM-YYYY HH:MM:SS")
    Print #FR1, Format(Now, "DDD-DD-MMM-YYYY HH:MM:SS")
Close #FR1

Frm_TIMER_PROJECT.Label2 = "Projet Ready"

'Shell "EXPLORER /E, " + Form1.Label_2_FOLDER, vbNormalFocus

'Call MNU_OPEN_INPUT_TEXT_Click

Call PROCESS_LOAD_DATA
Call TEST_ROUTINE
End Sub

Private Sub MNU_CLIP_BOARD_OUTPUT_Click()

'I = MsgBox("MSGBOX -- YES IF YOU WANT THIS ON CLIPBOARD" + vbCrLf + I1, vbYesNo)
'If I = vbYes Then

If OUTPUT_RESULT_STRING <> "" Then
    Clipboard.Clear
    Sleep 200
    Clipboard.SetText OUTPUT_RESULT_STRING
End If
Beep

'End If


'End

End Sub

Private Sub MNU_EXIT_Click()
'Unload Form1
Unload Me
Unload Form1
End Sub

Private Sub MNU_EXPLORER_FOLDER_Click()

If Dir(Frm_TIMER_PROJECT.Label1) <> "" Then
    Shell "EXPLORER.EXE /select, " + Frm_TIMER_PROJECT.Label1 + ", vbNormalFocus"
    Else
    Shell "EXPLORER /E, " + Form1.Label_2_FOLDER, vbNormalFocus
End If

End Sub

Private Sub MNU_LAST_PROJECT_Click()
    
    If Dir(App.Path + "\DATA_LAST_PROJECT.TXT") = "" Then Beep: MNU_LAST_PROJECT.Enabled = False: Exit Sub
    
    FR1 = FreeFile
    Open App.Path + "\DATA_LAST_PROJECT.TXT" For Input As #FR1
        Line Input #FR1, LINE2
        'LINE2 = Mid(LINE2, 20)
    Close #FR1
    
    'Kill App.Path + "\DATA_LAST_PROJECT.TXT"
    
    Form1.Timer2.Enabled = False
    
    Label1.Caption = LINE2
    
    'MAYBE WANTING
    'Call Form1.GET_PATH_OR_FILE_PATH(LINE2)
    

End Sub

Private Sub MNU_LAST_PROJECT_HISTORY_Click()
    
    If Dir(App.Path + "\DATA_LAST_PROJECT_LOGGER.TXT") = "" Then Beep: MNU_LAST_PROJECT_HISTORY.Enabled = False: Exit Sub
    
    Shell "NOTEPAD " + App.Path + "\DATA_LAST_PROJECT_LOGGER.TXT", vbNormalFocus

End Sub

Private Sub MNU_MAIN_FORM_Click()
    'Unload Me
    Me.Hide
    Form1.Show
    Form1.WindowState = vbNormal
End Sub

Private Sub MNU_OPEN_INPUT_TEXT_Click()

'Shell "EXPLORER " + Frm_TIMER_PROJECT.Label1, vbNormalFocus
If Dir("") <> "" Then
  
End If

Shell "NOTEPAD " + Frm_TIMER_PROJECT.Label1, vbNormalFocus

End Sub

Sub PROCESS_LOAD_DATA()

    COUNTI = 1
    LC = 1
    ReDim InPut_Date1(2000)
    ReDim Test__Date1(2000)

    'GOT EVENT WRONG HERE SLOPPY START UP
    'WORK AROUND
    If Frm_TIMER_PROJECT.Label1 = "Label1 FOLDER+FILE" Then
        Exit Sub
    End If
    
    'NOT DO MUCH HERE
    If InStr("*" + Frm_TIMER_PROJECT.Label1, "None Path/file Found") > 0 Then
        Label2.Caption = "Clipboard See Above"
        Exit Sub
    End If
    
    'Frm_TIMER_PROJECT.Label1 = Form1.Label_2_FOLDER.Caption + "\PROJECT_TIMER.TXT"

    MNU_PROCESS_OUTPUT.Enabled = True
    MNU_OPEN_INPUT_TEXT.Enabled = True
    MNU_OPEN_INPUT_TEXT.Enabled = True
    MNU_OPEN_OUTPUT_TEXT.Enabled = True
    MNU_EXPLORER_FOLDER.Enabled = True
    MNU_CLIP_BOARD_OUTPUT.Enabled = True
    
    MNU_LAST_PROJECT.Enabled = True
    MNU_LAST_PROJECT_HISTORY.Enabled = True
    
    If InStr(Label1.Caption, App.Path + "\PROJECT_TIMER") > 0 Then
        FR1 = FreeFile
        Open App.Path + "\DATA_LAST_PROJECT.TXT" For Output As #FR1
            Print #FR1, Label1.Caption
        Close #FR1
    End If
    
    FR1 = FreeFile
    Open App.Path + "\DATA_LAST_PROJECT_LOGGER.TXT" For Append As #FR1
        Print #FR1, Format(Now, "YYYY-MM-DD HH:MM:SS") + " -- " + Label1.Caption
    Close #FR1


    'MAIN READ FILE ------------------------------------
    
    If Dir(Frm_TIMER_PROJECT.Label1.Caption) <> "" Then

        IS_PROJECT_DONE_VAR = False
        FR1 = FreeFile
        i = Frm_TIMER_PROJECT.Label1
        FILEX = i
        
        Open FILEX For Input As #FR1
            Do
                Line Input #FR1, line1
                LC = LC + 1
                FLAG_DONE = False
                If Trim(line1) = "PROJECT CURRENTLY ACTIVE NOW = YES" Then
                    '-------------
                    IS_PROJECT_DONE_VAR = False
                    FLAG_DONE = True
                End If
                If Trim(line1) = "PROJECT CURRENTLY ACTIVE NOW = NO" Then
                    '-------------
                    IS_PROJECT_DONE_VAR = True
                    FLAG_DONE = True ' VALUE FOR NOT OBTAIN ANY LINE FURTHER CHECK
                End If
                If Trim(line1) = "PROJECT CURRENTLY ACTIVE NOW = NOT" Then
                    '-------------
                    IS_PROJECT_DONE_VAR = True
                    FLAG_DONE = True ' VALUE FOR NOT OBTAIN ANY LINE FURTHER CHECK
                End If
                
                If Trim(line1) <> "" And Mid(line1, 1, 1) <> "#" And FLAG_DONE = False Then
                    '-------------
                    If Test__Date1(COUNTI) <> "" Then COUNTI = COUNTI + 1
                    On Error Resume Next
                    ERRDATETIME1 = False
                    ERRDATETIME2 = False
                    
                    'strip the day of week out
                    For rday = 1 To 7
                        tday = UCase(Format(Now + rday, "DDD") + "-") 'NEW FORMAT WITH DASH
                        line1 = Replace(UCase(line1), tday, "")
                        tday = UCase(Format(Now + rday, "DDD") + " ") 'OLD FORMAT WITH SPACE
                        line1 = Replace(UCase(line1), tday, "")
                    Next
                    
                    If InPut_Date1(COUNTI) = "" Then
                        'SET FIRST LINE OF BEGIN TIME
                        InPut_Date1(COUNTI) = line1
                    Else
                        'SET SECOND LINE OF END TIME
                        Test__Date1(COUNTI) = line1
                    End If
                    
                    'JOB
                    'IF Date Test OKAY REPLACE DATE FORMAT WITH OWN CONSISTANT
                    
                    'ERROR CHECKING
                    Err.Clear
                    If TimeValue(line1) = 0 Then ERRDATETIME1 = True
                    If Err.Number > 0 Then ERRDATETIME2 = True
                    LINE1DATE = line1
                    
                    If ERRTIME1 = True Or ERRTIME2 = True Then
                        MBI = "READING THE INPUT SCRIPT AND A ERROR WITH ON LOAD DATE TIME LINE"
                        MBI = MBI + vbCrLf + vbCrLf
                        MBI = MBI + "DATE LINE COUNT -- LINE =" + Str(LC) + "DATE PAIR COUNT =" + Str(COUNTI)
                        MBI = MBI + "DATE LINE STRING ------ = " + line1
                        MBI = MBI + vbCrLf + vbCrLf
                        MBI = MBI + vbCrLf + vbCrLf
                        MBI = MBI + "CHECK ERROR DATE = "
                        If ERRDATETIME1 = True Then MBI = MBI + "ERROR -- IS WITH DATEVALUE IS RETURN NOTHING"
                        If ERRDATETIME2 = True Then MBI = MBI + "ERROR -- IS WITH DATEVAULE RETURNED ERROR -- " + Err.Description
                        MBI = MBI + vbCrLf + vbCrLf
                        MBI = MBI + vbCrLf + vbCrLf
                        Err.Clear
                        MBI = MBI + "CHECK DATE FORMAT INTEGRITY"
                        If InStr(line1, " ") = 0 Then
                            MBI = MBI + "CHECK DATE HAS A MIDDLE SPACE CHAR = FALSE " + "-- TOTAL LENGHT STRING =" + Str(Len(line1))
                        Else
                            MBI = MBI + "CHECK DATE HAS A MIDDLE SPACE CHAR = IS AT POSITION" + Str(InStr(line1, " ")) + "-- TOTAL LENGHT STRING =" + Str(Len(line1))
                        End If
                        MBI = MBI + vbCrLf + vbCrLf
                        CHECK_ERR_DATE = Mid(line1, 1, InStr(line1, " "))
                        CHECK_ERR_TIME = Mid(line1, InStr(line1, " "))
                        '----------------------------------------
                        MBI = MBI + "CHECK DATE FORMAT INTEGRITY"
                        If DateValue(CHECK_ERR_DATE) = 0 Then
                            MBI = MBI + "CHECK DATE = DATE IS PROBLEM WITH VAULE NOTHING"
                            MBI = MBI + vbCrLf + vbCrLf
                        End If
                        Err.Clear
                        If DateValue(CHECK_ERR_DATE) = 0 Then
                            MBI = MBI + "CHECK DATE = DATE IS PROBLEM RETURN ERROR -- " + Err.Description
                            MBI = MBI + vbCrLf + vbCrLf
                        End If
                        '----------------------------------------
                        MBI = MBI + vbCrLf + vbCrLf
                        MBI = MBI + "CHECK TIME FORMAT INTEGRITY"
                        If TimeValue(CHECK_ERR_TIME) = 0 Then
                            MBI = MBI + "CHECK TIME = TIME IS PROBLEM WITH VAULE NOTHING"
                            MBI = MBI + vbCrLf + vbCrLf
                        End If
                        Err.Clear
                        If DateValue(CHECK_ERR_TIME) = 0 Then
                            MBI = MBI + "CHECK TIME = TIME IS PROBLEM RETURN ERROR -- " + Err.Description
                            MBI = MBI + vbCrLf + vbCrLf
                        End If
                        
                        MsgBox MBI, vbMsgBoxSetForeground
                    
                    End If
                End If
            Loop Until EOF(1)
        Close #FR1
    End If
    
    If Dir(Frm_TIMER_PROJECT.Label1) = "" Then Exit Sub
    
    TOTAL_DATA_COUNT = COUNTI

    Call TEST_ROUTINE
End Sub

Private Sub MNU_PROCESS_OUTPUT_Click()
    
    Call PROCESS_LOAD_DATA
    Call TEST_ROUTINE

End Sub

Private Sub MNU_USE_APP_PATH_PROJECT_Click()
    
    If Dir(App.Path + "\DATA_LAST_PROJECT.TXT") = "" Then Beep: MNU_LAST_PROJECT.Enabled = False: Exit Sub
    
    Form1.Timer2.Enabled = False
    
    File = Dir(App.Path + "\PROJECT_TIMER*")
    Label1.Caption = App.Path + "\" + File '"\PROJECT_TIMER.TXT"

End Sub

Private Sub MNU_VB_ME_Click()

Call Form1.MNU_VB_ME_Click

End Sub

Private Sub MNU_VB_FOLDER_Click()

Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
Unload Me

End Sub

Private Sub Timer1_Timer()

'Label1 = Form1.Label1



End Sub


Sub DECLARE_VAR()

'THIS SUB IS FOR DECLARE VAR COMMON
'PUT THESE VAR IN THARE DELACE AT THE TOP OF CODE
'SO COMMON BETWEEN SUB AND FUNCTION
Exit Sub
Dim F1, F2, F3, F4
Dim F41, F42, F43, F44


End Sub

Sub TEST_ROUTINE()

ReDim TIME__RESULT_____1(4, 2000)
ReDim TIME__RESULT_MOD_2(4, 2000)
ReDim TIME_TOTAL_______1(4, 2000)
ReDim TIME_TOTAL___MOD_2(4, 2000)


' NOT HARDCODED DESIGN STAGE --SET FROM SCRIPT FILE NOW --------
'-------------------- SET FROM SCRIPT FILE NOW
'TOTAL_DATA_COUNT = 1
i = TOTAL_DATA_COUNT
'---------------------------------------SET FROM SCRIPT FILE NOW
'InPut_Date1(I) = "22-Jul-2016 12:09:37"
'Test__Date1(I) = "23-Jul-2016 00:38:31"
''--------------------------------------
'TOTAL_DATA_COUNT = TOTAL_DATA_COUNT + 1
'I = TOTAL_DATA_COUNT
'InPut_Date1(I) = "23-Jul-2016 08:20:19"
'Test__Date1(I) = "23-Jul-2016 11:05:40"
'
'TOTAL_DATA_COUNT = TOTAL_DATA_COUNT + 1
'I = TOTAL_DATA_COUNT
'InPut_Date1(I) = "24-Jul-2016 10:00:00"

'---------------------------SET FROM SCRIPT FILE NOW
'----------------------------------------------------------------------
' IS PROJECT CURRENTLY WORKING ON
' AND THEN GET A REAL TIME UPDATE LENGHT
' SET START TIME FOR CURRENT DAY
' OR PLANNING TO DO TIME
'----------------------------------------------------------------------
'---------------------------SET FROM SCRIPT FILE NOW
'IS_PROJECT_DONE_VAR = FASLE
'----------------------------------------------------------------------

If IS_PROJECT_DONE_VAR = False Then
    Test__Date1(i) = Format(Now, "DD-MMM-YYYY HH:MM:SS")
Else
    'TOTAL_DATA_COUNT = TOTAL_DATA_COUNT - 1
End If

DAY_USE = ""
For R = 1 To TOTAL_DATA_COUNT
    
    InPut_Date2 = DateValue(InPut_Date1(R)) + TimeValue(InPut_Date1(R))
    Test__Date2 = DateValue(Test__Date1(R)) + TimeValue(Test__Date1(R))
    I1 = InPut_Date2
    I2 = Test__Date2
    
    Call FindTimeInfo_PROJECT(I1, I2)
    
    '--------------------------------------------
    '--------------------------------------------
    For R2 = 0 To F11
        i = I1 + R2
        DAY_USE = Format(i, "DDD") + " " + Format(Day(i), "00") + " " + Format(i, "MMM") + " Shift" + Str(R) + vbCrLf
        If InStr(DAY_USE_2, DAY_USE) = 0 Then
            DAY_USE_2 = DAY_USE_2 + "#" + Trim(Str(DAY_USE_3 + 1)) + ".. " + DAY_USE
            DAY_USE_3 = DAY_USE_3 + 1
        End If
    Next
    '--------------------------------------------
    'THIS TWO MORE LIKLY TO BE ONLY ONE DAY EXTRA
    'COVER IT ALL
    '--------------------------------------------
    If Day(I2) <> Day(i) Then
        i = I2
        DAY_USE = Format(i, "DDD") + Str(Day(i)) + " " + Format(i, "MMM") + " Shift" + Str(R) + vbCrLf
        If InStr(DAY_USE_2, DAY_USE) = 0 Then
            DAY_USE_2 = DAY_USE_2 + "#" + Trim(Str(DAY_USE_3 + 1)) + ".. " + DAY_USE
            DAY_USE_3 = DAY_USE_3 + 1
        End If
    
    End If
    '--------------------------------------------
    'THIS MAY NOT EVER RUN
    If IS_PROJECT_DONE_VAR = False Then i = Now
    If Day(I2) <> Day(i) And R = TOTAL_DATA_COUNT Then
        i = Now
        DAY_USE = Format(i, "DDD") + Str(Day(i)) + " " + Format(i, "MMM") + " Shift" + Str(R) + vbCrLf
        If InStr(DAY_USE_2, DAY_USE) = 0 Then
            DAY_USE_2 = DAY_USE_2 + "#" + Trim(Str(DAY_USE_3 + 1)) + ".. " + DAY_USE
            DAY_USE_3 = DAY_USE_3 + 1
        End If
    End If
    '--------------------------------------------
    '--------------------------------------------
   
    TIME__RESULT_____1(1, R) = F11 'DAY
    TIME__RESULT_____1(2, R) = F12 'HOUR
    TIME__RESULT_____1(3, R) = F13 'MIN
    TIME__RESULT_____1(4, R) = F14 'SEC
    
    F21 = F11         'DAY
    F22 = F12 Mod 24  'HOUR
    F23 = F13 Mod 60  'MIN
    F24 = F14 Mod 60  'SEC
    
    TIME__RESULT_MOD_2(1, R) = F21 'DAY
    TIME__RESULT_MOD_2(2, R) = F22 'HOUR
    TIME__RESULT_MOD_2(3, R) = F23 'MIN
    TIME__RESULT_MOD_2(4, R) = F24 'SEC
    
    'RUN INDIVIUAL TOTAL HISTORY IF WANT LATER
    TIME_TOTAL_______1(1, R) = TIME_TOTAL_______1(1, R - 1) + F11 'DAY
    TIME_TOTAL_______1(2, R) = TIME_TOTAL_______1(2, R - 1) + F12 'HOUR
    TIME_TOTAL_______1(3, R) = TIME_TOTAL_______1(3, R - 1) + F13 'MIN
    TIME_TOTAL_______1(4, R) = TIME_TOTAL_______1(4, R - 1) + F14 'SEC

    F41 = TIME_TOTAL_______1(1, R)         'DAY
    F42 = TIME_TOTAL_______1(2, R) Mod 24  'HOUR
    F43 = TIME_TOTAL_______1(3, R) Mod 60  'MIN
    F44 = TIME_TOTAL_______1(4, R) Mod 60  'SEC

    TIME_TOTAL___MOD_2(1, R) = F41  'DAY
    TIME_TOTAL___MOD_2(2, R) = F42 'HOUR
    TIME_TOTAL___MOD_2(3, R) = F43 'MIN
    TIME_TOTAL___MOD_2(4, R) = F44 'SEC

Next



RESULT_MESSAGE = "PROJECT TIME COUNTER RESULT SCRIPT -- TIME MACHINE" + vbCrLf
I1 = RESULT_MESSAGE
I1 = I1 + "TOTAL PROJECT SESSION COUNT SHIFT'S SPENT =" + Str(TOTAL_DATA_COUNT)

If IS_PROJECT_DONE_VAR = True Then
    I1 = I1 + "PROJECT IS DONE FINISHED" + vbCrLf
Else
    I1 = I1 + " IS CURRENTLY ACTIVE NOW" + vbCrLf
    'I1 = I1 + "PROJECT CURRENTLY ACTIVE NOW" + vbCrLf
End If

'I1 = I1 + "--------------------------------------------" + vbCrLf
'DAY_USE_2 = "DAY SCRIPT OF INCLUDED PROJECT ACTIVE DAY COUNT =" + Str(DAY_USE_3) + vbCrLf + DAY_USE_2
I1 = I1 + DAY_USE_2 ' ALREADY GOT VBCRLF
I1 = I1 + "------------------------------------------------------------------" + vbCrLf


I1 = I1 + "TOTAL TIME SPENT EACH SESSION ---- & ---- MOD DIVEDE -------------" + vbCrLf
I1 = I1 + "------------------------------------------------------------------" + vbCrLf
'I1 = I1 + vbCrLf

For RESULT_OUT = 1 To TOTAL_DATA_COUNT
    I2 = RESULT_OUT
    'work the gap time -- longest
    If I2 > 1 Then
        GAP_T1 = Test__Date1(I2 - 1)
        GAP_T2 = InPut_Date1(I2)
        gap_t_day = DATEDIFF_DAY_BY_COUNT_SECONDS(DateDiff("S", GAP_T1, GAP_T2))
        If gap_t_day > ogap_t_day Then ogap_t_day = gap_t_day
    End If
Next
Longest_Gap = String(Len(Str(ogap_t_day)) - 1, "0")


For RESULT_OUT = 1 To TOTAL_DATA_COUNT
    I2 = RESULT_OUT
    'work the gap time
    If I2 > 1 Then
        GAP_T1 = Test__Date1(I2 - 1)
        GAP_T2 = InPut_Date1(I2)
    
        gap_t_day = DATEDIFF_DAY_BY_COUNT_SECONDS(DateDiff("S", GAP_T1, GAP_T2))
        gap_t_hour = DateDiff("h", GAP_T1, GAP_T2) Mod 24
        gap_t_min = DateDiff("n", GAP_T1, GAP_T2) Mod 60
        gap_t = " - Gap Time(" + Format(gap_t_day, Longest_Gap) + " Day)(" + Format(gap_t_hour, "00") + " Hour)(" + Format(gap_t_min, "00") + " Min) After " + Format(Test__Date1(I2 - 1), "DD MMM")
    End If
    
    'DASH PAIR 01 OF 02
    'INFO_SESSION_NUMBER = "SESSION -- " + Format((I2), "00") + " of -- " + Format((1), "00") + " to " + Format((TOTAL_DATA_COUNT), "00") + " -------- " + vbCrLf
    INFO_SESSION_NUMBER = "SESSION -- " + Format((I2), "00") + " of " + Format((TOTAL_DATA_COUNT), "00") + " ----------------- " + vbCrLf
    I1 = I1 + INFO_SESSION_NUMBER
    'DASH PAIR 02 OF 02
    
    
    I1 = I1 + "-------------------------------------" + vbCrLf
'    I1 = I1 + "BEGIN ---- " + InPut_Date1(I2) + " ------------" + vbCrLf
'    I1 = I1 + "END ------ " + Test__Date1(I2) + " ------------" + vbCrLf
    '-----------------------------
    'Gap Time Used Here with Input Date
    '-----------------------------
    I1 = I1 + "BEGIN ---- " + Format(InPut_Date1(I2), "DDD ") + InPut_Date1(I2) + gap_t + vbCrLf
    I1 = I1 + "END ------ " + Format(Test__Date1(I2), "DDD ") + Test__Date1(I2) + vbCrLf
    I1 = I1 + "------------------------------------------------------------------" + vbCrLf

    '--------------
    '4 COLUMN OF DATA -- DAY HOUR MIN SEC
    For R2 = 1 To 4
        I1 = I1 + FORMATTING(R2, Format(TIME__RESULT_____1(R2, I2), "00"))
    Next
    I1 = I1 + vbCrLf
    'I1 = I1 + "--------------------------------------------------------------------" + vbCrLf
    
    '--------------
    '4 COLUMN OF DATA -- DAY HOUR MIN SEC
    For R2 = 1 To 4
        I1 = I1 + FORMATTING(R2, Format(TIME__RESULT_MOD_2(R2, I2), "00"))
    Next
    I1 = I1 + vbCrLf
    If RESULT_OUT <> TOTAL_DATA_COUNT Then
        I1 = I1 + "------------------------------------------------------------------" + vbCrLf
    End If
Next

'----------------------------------
'STRING MANNIPULATE IS LIKE SURGOEN
'MAKE A SPLIT AND PUT SOMETHING IN
'DON'T LEAVE A TOOL BEHIND
'----------------------------------

i = TOTAL_DATA_COUNT
'------------------------------------------------------------------------------------------
'I1 = I1 + vbCrLf
I1 = I1 + "==================================================================" + vbCrLf
I1 = I1 + "TOTAL TIME SPENT ALL" + Str(i) + " SESSION ---- & ---- MOD DIVEDE ------------" + vbCrLf
I1 = I1 + "------------------------------------------------------------------" + vbCrLf
'I1 = I1 + vbCrLf

xspace = String(Str(TOTAL_DATA_COUNT) - 1, "0")
For RESULT_OUT = 1 To TOTAL_DATA_COUNT
    I2 = RESULT_OUT
    
    session = "Sess " + Format(I2, xspace) + "-" + Format(i, xspace) + " "
    session = session + Mid(Format(InPut_Date1(I2), "DDD"), 1, 2) + "-" + Mid(Format(Test__Date1(I2), "DDD"), 1, 2) + " "
    session = session + Format(InPut_Date1(I2), "DD") + "-"
    session = session + Mid(Format(InPut_Date1(I2), "MMM"), 1, 3) + " "
    I1 = I1 + session
    
    For R2 = 1 To 4
        I1 = I1 + FORMATTING(R2, Format(TIME_TOTAL_______1(R2, I2), "00"))
    Next
    I1 = I1 + vbCrLf
    I1 = I1 + session
    For R2 = 1 To 4
        I1 = I1 + FORMATTING(R2, Format(TIME_TOTAL___MOD_2(R2, I2), "00"))
    Next
    I1 = I1 + vbCrLf
    I1 = I1 + "------------------------------------------------------------------" + vbCrLf

Next

'MAIN WRITE OUTPUT RESULT FILE ----------------------------------

FR1 = FreeFile
i = Frm_TIMER_PROJECT.Label1
FILEX = Mid(i, 1, InStrRev(i, "\")) + "PROJECT_TIMER_OUTPUT.TXT"
Open FILEX For Output As #FR1
    Print #FR1, I1
Close #FR1


List1.Clear
FR1 = FreeFile
Open FILEX For Input As #FR1
    Do
        Line Input #FR1, line1
        List1.AddItem line1
    Loop Until EOF(FR1)
Close #FR1
    
OUTPUT_RESULT_STRING = I1
    
    
End Sub

Function FORMATTING(ID2, STRING1)
    
    If ID11 < 2 Then ID11 = 2
    If ID12 < 3 Then ID12 = 3
    If ID13 < 4 Then ID13 = 4
    If ID14 < 5 + 1 Then ID14 = 5 + 1
    
    If ID2 = 1 Then
        STRING2 = "DAY "
        If Len(STRING1) > ID11 Then ID11 = Len(STRING1)
        ID1 = ID11
    End If
    If ID2 = 2 Then
        STRING2 = "HOUR "
        If Len(STRING1) > ID12 Then ID12 = Len(STRING1)
        ID1 = ID12
    End If
    If ID2 = 3 Then
        STRING2 = "MINUTE "
        If Len(STRING1) > ID13 Then ID13 = Len(STRING1)
        ID1 = ID13
    End If
    If ID2 = 4 Then
        STRING2 = "SECOND "
        If Len(STRING1) > ID14 Then ID14 = Len(STRING1)
        ID1 = ID14
    End If
    NUMERIC_VAR = Space(ID1)
    RSet NUMERIC_VAR = STRING1
    i = "(" + STRING2 + NUMERIC_VAR + ") -- "
    FORMATTING = i

End Function


Sub FindTimeInfo_PROJECT(InPutDate1, TestDate1)

'---------------------------------------------------------------
'DON'T MATTER IF -- BACK TO FRONT -- AFTER BEFORE -- BACK FUTURE
'BECAUSE -- ABS FUNCTION
'---------------------------------------------------------------
'LIKE LIKE LIKE FIND AND FIND NEXT
'---------------------------------------------------------------

InPutDate = DateValue(InPutDate1) + TimeValue(InPutDate1)
TestDate = DateValue(TestDate1) + TimeValue(TestDate1)

'F11 = Abs(DateDiff("d", InPutDate, TestDate))
'DONT USE THIS DAY BECAUSE CROSS ON BOUNDRY OF A DAY MIDNIGHT
'DO COUNT BY SECOND ACCURATE FOR QUICKER
F11 = Abs(DateDiff("S", InPutDate, TestDate))
F11 = DATEDIFF_DAY_BY_COUNT_SECONDS(F11)


F12 = Abs(DateDiff("h", InPutDate, TestDate))
F13 = Abs(DateDiff("n", InPutDate, TestDate))
F14 = Abs(DateDiff("s", InPutDate, TestDate))

F21 = F11 'Day
F22 = F12 Mod 24
F23 = F13 Mod 60
F24 = F14 Mod 60


Exit Sub

'OUTPUT FORMAT RESULT EXAMPLE
'EX0 = Format$(TimeSerial(0, F3, F4), "H:MM:SS")
'EX1 = Format$(TimeSerial(0, F3, F4), "hH:MM:SS")
'EX2 = Trim(Str(Val(Mid$(EX1, 1, 2)))) + "h" + Str(Val(Mid$(EX1, 4, 2))) + "m" + Str(Val(Mid$(EX1, 7, 2))) + "s"
'EX3 = Trim(Str(Val(Mid$(EX1, 1, 2)))) + "h" + Str(Val(Mid$(EX1, 4, 2))) + "m"
'
'DexFORMAT1 = Trim(Str(F1)) + "d " + EX0
'DexFORMAT2 = Trim(Str(F1)) + "d " + EX2
'DexFORMAT3 = Trim(Str(F1)) + "d " + EX3

End Sub


Function DATEDIFF_DAY_BY_COUNT_SECONDS(F11)

'load the seconds since date
'AND CODE WILL SUBTRACT DAY OF SECOND UNTIL COUNT OF DAY IS LEFT
'THAT WAY THE NORMAL DATEDIFF WON'T BE CHANGE AT MIDNIGHT 1 DAY

'---- 86400 ' -- TOTAL SECONDS IN A DAY
DATEDIFF_DAY_BY_COUNT_SECONDS = 0
Z = 0
Do
    If F11 >= 86400 Then Z = Z + 1: F11 = F11 - 86400
Loop Until F11 < 86400
F11 = Z
DATEDIFF_DAY_BY_COUNT_SECONDS = F11

End Function




'----------------------------------
'ALL BELOW HERE ARE NOT PROJECT USE
'OLDER CODE
'BUILT HERE UP FROM
'----------------------------------
'----------------------------------
'----------------------------------
'----------------------------------
'----------------------------------


Function FindTimeInfo_ADD_UP(InPutDate, TestDate)

'------------------------------------------------------
'THIS CODE IS BASICLY A COPY PASTE LAZY REPEAT OF OTHER
'WITH ADDING
'------------------------------------------------------
'LIKE LIKE LIKE FIND AND FIND NEXT
'------------------------------------------------------

'------------------------------------------------------
'InPutDate = DateValue("01-12-2008")
'TestDate = DateValue("31-05-2009")
'------------------------------------------------------

DayCountT = DateDiff("d", InPutDate, TestDate)

Year5 = 0
For R5 = Year(InPutDate) + 1 To Year(TestDate) + 1
    If DateSerial(R5, Month(InPutDate), Day(InPutDate)) < Now Then Year5 = Year5 + 1
Next
For R5 = 1 To -2 Step -1
    XX = DateSerial(Year(TestDate), Month(TestDate) + R5, Day(InPutDate))
    MonthStart = XX
    If XX <= TestDate Then Exit For
Next
'Month5 = 0
'Month7 = 0
For R5 = 1 To 10000
    XX = DateSerial(Year(InPutDate), Month(InPutDate) + R5, Day(InPutDate))
    If Year(XX) <> oxx Then Month7 = 0
    oxx = Year(XX)
    If XX <= TestDate Then Month5 = Month5 + 1: Month7 = Month7 + 1
    If XX >= TestDate Then Exit For
Next

For R5 = Year(TestDate) To 1 Step -1
    If DateSerial(R5, Month(InPutDate), Day(InPutDate)) < Now Then
        DayCount1 = DateDiff("d", DateSerial(R5, Month(InPutDate), Day(InPutDate)), TestDate)
        DayCountMonth = DateDiff("d", MonthStart, TestDate)
        WholeYear1 = DateDiff("d", DateSerial(R5, Month(InPutDate), Day(InPutDate)), DateSerial(R5 + 1, Month(InPutDate), Day(InPutDate)))
        'Month5 = DateDiff("m", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        WeeksSinceYear = DateDiff("w", DateSerial(R5, Month(InPutDate), Day(InPutDate)), TestDate) - 1
        WeeksSinceInput = Int((DateDiff("D", InPutDate, TestDate) - 1) / 7)
        WeeksPlusDays = (DateDiff("d", InPutDate, TestDate) - 1) - (WeeksSinceInput * 7)
        
        ResultYearDate = Format$(Year5 + DayCount1 / WholeYear1, ".000")
        Exit For
    End If
Next


DexFORMAT1 = 0
DexFORMAT2 = 0
DexFORMAT3 = 0

F1 = 0: F2 = 0: F3 = 0: F4 = 0

'DO SECONDS IF MASSIVE BUT ONLY TO THE LAST REMAINING DAY

'tMm = DateDiff("s", Int(TestDate), TestDate)
'F1 = Int((tMm / 60 ^ 2) / 24)
'If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)
'F2 = Int((tMm / 60 ^ 2))
'F3 = Int(tMm / 60)
'F4 = tMm Mod 60

'DONT DO SECONDS IF MASSIVE
If Abs(DateDiff("D", InPutDate, TestDate)) > 366 Then Exit Function
'------------------------------------------------------------------

tMm = DateDiff("s", InPutDate, TestDate)
F1 = Int((tMm / 60 ^ 2) / 24)
If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)
F2 = Int((tMm / 60 ^ 2))
F3 = Int(tMm / 60)
F4 = tMm Mod 60

'----------------------------
'THIS MAY NOT BE CORRECT TEST
'If F1 < 0 Then F1 = 0

'If F1 < 0 Then Stop
'----------------------------

'EXAMPLE
EX0 = Format$(TimeSerial(0, F3, F4), "H:MM:SS")
EX1 = Format$(TimeSerial(0, F3, F4), "hH:MM:SS")
EX2 = Trim(Str(Val(Mid$(EX1, 1, 2)))) + "h" + Str(Val(Mid$(EX1, 4, 2))) + "m" + Str(Val(Mid$(EX1, 7, 2))) + "s"
EX3 = Trim(Str(Val(Mid$(EX1, 1, 2)))) + "h" + Str(Val(Mid$(EX1, 4, 2))) + "m"

DexFORMAT1 = Trim(Str(F1)) + "d " + EX0
DexFORMAT2 = Trim(Str(F1)) + "d " + EX2
DexFORMAT3 = Trim(Str(F1)) + "d " + EX3

If F1 = 1 Then
    DexFORMAT1 = Trim(Str(F1)) + "d " + EX0
    DexFORMAT2 = Trim(Str(F1)) + "d " + EX2
    DexFORMAT3 = Trim(Str(F1)) + "d " + EX3
End If

If F1 = 0 Then
    DexFORMAT1 = EX0
    DexFORMAT2 = EX2
    DexFORMAT3 = EX3
End If

End Function


Function FindTimeInfo(InPutDate, TestDate)

'InPutDate = DateValue("01-12-2008")
'TestDate = DateValue("31-05-2009")

DayCountT = DateDiff("d", InPutDate, TestDate)

Year5 = 0
For R5 = Year(InPutDate) + 1 To Year(TestDate) + 1
    If DateSerial(R5, Month(InPutDate), Day(InPutDate)) < Now Then Year5 = Year5 + 1
Next
For R5 = 1 To -2 Step -1
    XX = DateSerial(Year(TestDate), Month(TestDate) + R5, Day(InPutDate))
    MonthStart = XX
    If XX <= TestDate Then Exit For
Next
Month5 = 0
Month7 = 0
For R5 = 1 To 10000
    XX = DateSerial(Year(InPutDate), Month(InPutDate) + R5, Day(InPutDate))
    If Year(XX) <> oxx Then Month7 = 0
    oxx = Year(XX)
    If XX <= TestDate Then Month5 = Month5 + 1: Month7 = Month7 + 1
    If XX >= TestDate Then Exit For
Next

For R5 = Year(TestDate) To 1 Step -1
    If DateSerial(R5, Month(InPutDate), Day(InPutDate)) < Now Then
        DayCount1 = DateDiff("d", DateSerial(R5, Month(InPutDate), Day(InPutDate)), TestDate)
        DayCountMonth = DateDiff("d", MonthStart, TestDate)
        WholeYear1 = DateDiff("d", DateSerial(R5, Month(InPutDate), Day(InPutDate)), DateSerial(R5 + 1, Month(InPutDate), Day(InPutDate)))
        'Month5 = DateDiff("m", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        WeeksSinceYear = DateDiff("w", DateSerial(R5, Month(InPutDate), Day(InPutDate)), TestDate) - 1
        WeeksSinceInput = Int((DateDiff("D", InPutDate, TestDate) - 1) / 7)
        WeeksPlusDays = (DateDiff("d", InPutDate, TestDate) - 1) - (WeeksSinceInput * 7)
                
        ResultYearDate = Format$(Year5 + DayCount1 / WholeYear1, ".000")
        Exit For
    End If
Next

DexFORMAT1 = 0
DexFORMAT2 = 0
DexFORMAT3 = 0

F1 = 0: F2 = 0: F3 = 0: F4 = 0

'DO SECONDS IF MASSIVE BUT ONLY TO THE LAST REMAINING DAY

'tMm = DateDiff("s", Int(TestDate), TestDate)
'F1 = Int((tMm / 60 ^ 2) / 24)
'If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)
'F2 = Int((tMm / 60 ^ 2))
'F3 = Int(tMm / 60)
'F4 = tMm Mod 60

'DONT DO SECONDS IF MASSIVE
If Abs(DateDiff("D", InPutDate, TestDate)) > 366 Then Exit Function

tMm = DateDiff("s", InPutDate, TestDate)
F1 = Int((tMm / 60 ^ 2) / 24)
If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)
F2 = Int((tMm / 60 ^ 2))
F3 = Int(tMm / 60)
F4 = tMm Mod 60

'THIS MAY NOT BE CORRECT TEST
'If F1 < 0 Then F1 = 0

'If F1 < 0 Then Stop

'EXAMPLE
ex = Format$(TimeSerial(0, F3, F4), "H:MM:SS")
EX1 = Format$(TimeSerial(0, F3, F4), "hH:MM:SS")
EX2 = Trim(Str(Val(Mid$(EX1, 1, 2)))) + "h" + Str(Val(Mid$(EX1, 4, 2))) + "m" + Str(Val(Mid$(EX1, 7, 2))) + "s"
EX3 = Trim(Str(Val(Mid$(EX1, 1, 2)))) + "h" + Str(Val(Mid$(EX1, 4, 2))) + "m"

DexFORMAT1 = Trim(Str(F1)) + "d " + ex
DexFORMAT2 = Trim(Str(F1)) + "d " + EX2
DexFORMAT3 = Trim(Str(F1)) + "d " + EX3

If F1 = 1 Then
    DexFORMAT1 = Trim(Str(F1)) + "d " + ex
    DexFORMAT2 = Trim(Str(F1)) + "d " + EX2
    DexFORMAT3 = Trim(Str(F1)) + "d " + EX3
End If

If F1 = 0 Then
    DexFORMAT1 = ex
    DexFORMAT2 = EX2
    DexFORMAT3 = EX3
End If

End Function


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
  'Test = False
End Function
'***********************************************

'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

Timer_1_MINUTE.Interval = 60000

Call VB_PROJECT_CHECKDATE

End Sub
'-------------------------------------------------
'-------------------------------------------------
Sub VB_PROJECT_CHECKDATE()

'TOP DELCLARE -------------
'DIM XVB_DATE
'--------------------------

If Mid(App.Path, 1, 1) <> "D" Then Exit Sub

PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

'---------------------------------------------------
'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
'---------------------------------------------------
'MINOR WORK AROUND CREATE PATH
'---------------------------------------------------
If Dir(PATH_FILE_NAME2) = "" Then
    On Error Resume Next
        MkDir Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        FS.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If

Set F = FS.GetFile(PATH_FILE_NAME2)

VB_DATE = F.DateLastModified

'--------
'01 OF 02
'----------------------------------
'WRITE INFO ABOUT DATE CHANGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    
    '-----------------------------------
    'RUN A VBS TO COPY OVER AND RELOADER
    '-----------------------------------
    
'    FR1 = FreeFile
'    Open App.Path + "\RELOAD AND COPY.VBS" For Output As #FR1


    '----------------------------------
    'Mon 10 April 2017 16:30:50--------
    '----------------------------------
    'PROJECT REFERANCE ----------------
    'wshom.ocx
    'IF HARD FIND DO BROWSER
    'IN SYSTEM32 AND SYSWOW64
    'DIR wshom.ocx /S
    '----------------------------------
    'WINDOWS SCRIPT HOST OBJECT MODEL
    'AFTER FIND ALSO HAVE TO SELECT HER
    '----------------------------------
    '----------------------------------
    'Mon 10 April 2017 15:45:11--------
    '----
    'VISUAL BASIC wshom.ocx NOT WORKING - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=VISUAL+BASIC+wshom.ocx+NOT+WORKING&oq=VIUAL+BASIC+wshom.ocx+NOT+WORKING&gs_l=serp.3..30i10k1.1533.3987.0.4658.12.12.0.0.0.0.127.888.11j1.12.0....0...1c.1.64.serp..0.12.886...33i160k1j33i21k1.fYCPCpVPeVk
    '--------
    'WScript reference in VB
    'http://forums.codeguru.com/showthread.php?30458-WScript-reference-in-VB
    '----
    'Set objShell = WScript.CreateObject("WScript.Shell") - ------Not WORKER
    '----------------------------------
    '----------------------------------
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE

End Sub
