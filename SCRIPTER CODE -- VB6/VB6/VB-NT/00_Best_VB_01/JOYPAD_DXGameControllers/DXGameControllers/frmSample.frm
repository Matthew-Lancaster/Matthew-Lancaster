VERSION 5.00
Begin VB.Form frmDXGAMEJOY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JoyPad_DXGameControllers"
   ClientHeight    =   3345
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdFoc 
      Caption         =   "Ignore Background Events"
      Height          =   375
      Left            =   4980
      TabIndex        =   3
      Top             =   2880
      Width           =   2355
   End
   Begin VB.ListBox lboEvt 
      Height          =   2205
      ItemData        =   "frmSample.frx":0000
      Left            =   4980
      List            =   "frmSample.frx":0002
      TabIndex        =   2
      Top             =   60
      Width           =   3735
   End
   Begin VB.CommandButton cmdScn 
      Caption         =   "Rescan For Devices"
      Height          =   375
      Left            =   132
      TabIndex        =   1
      Top             =   2856
      Width           =   2355
   End
   Begin VB.ListBox lboLst 
      Height          =   2205
      ItemData        =   "frmSample.frx":0004
      Left            =   120
      List            =   "frmSample.frx":0006
      TabIndex        =   0
      Top             =   60
      Width           =   4755
   End
   Begin VB.Menu MNU_JOYPAD_SOFTWARE_STATUS 
      Caption         =   "JOY_SOFTWARE_STATUS_NOT_OKAY"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmDXGAMEJOY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public JOY_CONTROLLER
Public JOYPAD_DX_LOADED
Dim PRESSED_YET, JOY_RUN_ONCE
Dim A2

Dim ElementName_VAR


' IN TIDAL PROGRAM THIS FORM IS CALLED
' frmDXGAMEJOY

'NOTE BY MATTHEW
'
'------------------------------------------------
'Mon 09-May-2016 11:51:33 a
'D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\DXGameControllers\DXGameControllers\dx8vb.dll
'D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\DXGameControllers\DXGameControllers\dx8vb.txt
'------------------------------------------------
'
'THIS FILE YOU COPY INTO THE VB APPLICATION PROJECT AREA
'FOR THE JOYPAD
'--------------
'
'dx8vb.dll activex component can't create object 429
'
'THE PROBLEM WITH THIS .DLL
'AND FIX LIKE THIS
'regsvr32 C:\Windows\system32\dx8vb.dll
'OR COPY
'COPY D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\dx8vb.dll C:\Windows\system32\dx8vb.dll
'
'SEEM NOT GOOD ENOUGH TO PUT .DLL IN PATH WHERE EXE APPLICATION CODE IS
'MAYBE IF REGISTER IT THERE
'
'https://www.google.co.uk/search?num=100&rlz=1C1SKPL_enGB417&q=dx8vb.dll+activex+component+can%27t+create+object+429&oq=dx8vb.dll+activex+component+can%27t+create+object+429&gs_l=serp.12...3958.5360.0.8774.2.2.0.0.0.0.242.327.1j0j1.2.0....0...1c.1.64.serp..0.1.236...30i10.zVuh2klQo68
'
'http://forums.pcsx2.net/Thread-TwinPad-v0-9-25?page=5
'
'----
'THE ---DX8VB.DLL
'MIGHT OF BEEN INSTALLED BY LOGITECH JOY PAD SOFTWARE
'AND IT IS NOW AT MOMENT -- NOT SURE KNOW IF WANTED THE
'ON MY COMPUTER IT IS HERE
'D:\#0 1 INSTALLATIONS\Installation programs\# COMMS JOYPAD BLUETOOTH\JOY PAD


'
'----
'Download DirectX 9.0c End-User Runtime from Official Microsoft Download Center
'https://www.microsoft.com/en-in/download/confirmation.aspx?id=34429
'----
'
'




' Name:    frmSample
' Author:  Michael A Seel
' Date:    Feb 2004

Option Explicit

Dim JOY_CONTROLLER_F710_PORT
Dim JOY_CONTROLLER_RumblePad_2_PORT

Public FIRST_PASS

Implements DirectXEvent8  'Allow form to receive device event notification

Private WithEvents mLst As cControllerList  'Our controller list
Attribute mLst.VB_VarHelpID = -1
Private mFoc As Boolean  'True = Allow background events (needed only for this sample)

'---------------------------------------
'---------------------------------------
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'---------------------------------------
'---------------------------------------
Private Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
'---------------------------------------
Private Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
'---------------------------------------
'---------------------------------------



Private Sub cmdFoc_Click()
   Dim i As Integer
   mFoc = Not mFoc
   If mFoc Then
      cmdFoc.Caption = "Allow Background Events"
   Else
      cmdFoc.Caption = "Ignore Background Events"
   End If
   For i = 1 To mLst.Count
      mLst.Item(i).FocusEventsOnly = mFoc
   Next
End Sub

Private Sub cmdScn_Click()
   lboLst.Clear
   lboEvt.Clear
   Me.Refresh
   mLst.RescanForDevices  'Rescan
   RefreshList
End Sub

Private Sub DirectXEvent8_DXCallback(ByVal EventID As Long)
'EventID is specific to each device if DirectXEvent8 is implemented on the "EventForm"
'If it is not implemented, all devices will have an EventID of '0', and this procedure
'is never called.
   mLst.CheckForEvents EventID  'Check for events on device with EventID
End Sub

Private Sub Form_Load()
   
    Me.Show
   
    Dim RE
    Dim MSGB2, MSGBOX_QUESTION, Result, RESULT_VAR
   
    frmDXGAMEJOY.JOYPAD_DX_LOADED = True
   
'   If Dir("C:\Program Files\Logitech\Gaming Software\LWEMon.exe") = "" Then
   
    DoEvents
    Set mLst = New cControllerList  'Create an instance of the controller list.
    'MsgBox "ME 0" + str(Err.Number)
    
    
    'regsvr32 C:\Windows\system32\dx8vb.dll
        
    
    Dim ERR_FILENAME_1, ERR_FILENAME_2
    Dim GIVE_A_GO
    
    ERR_FILENAME_1 = False
    ERR_FILENAME_2 = True
    If Dir("C:\Program Files\Logitech\Gaming Software\LWEMon.exe") = "" Then
        ERR_FILENAME_1 = True
    End If
    ERR_FILENAME_2 = False
    
'    Dim FS, F, F1, FC, S
'    Set FS = CreateObject("SCRIPTING.FILESYSTEMOBJECT")
'    Set F = FS.GetFolder("C:\Windows\System32")
'    Set FC = F.Files
'    For Each F1 In FC
'        If F1.Name = "dx8vb.dll" Then Stop: Exit For
'    Next
    
    
'    If FS.FileExists("C:\Windows\System32\dx8vb.dll") = True Then
'        ERR_FILENAME_2 = False
'    End If
'
'    If Dir("C:\Windows\System32\dx8vb.dll") <> "" Then
'        ERR_FILENAME_2 = False
'    End If
'    If Dir("C:\Windows\SysWOW64\dx8vb.dll") <> "" And ERR_FILENAME_2 = True Then
'        ERR_FILENAME_2 = False
'    End If
    
    
'    If IsIDE = True Or 1 = 1 Then ERR_FILENAME_2 = False
'
'    'GET AN EXE LOAD PROBLEM IF RUN THIS WAY
'    'OKAY IN IDE BUT NOT EXE __ DRIVER MUST EXIST ____
    If ERR_FILENAME_2 = False Then
        On Error Resume Next
        Err.Clear
        mLst.Initialize Me, False  'Initialize the list using this form as the owner.
        '---------------------------------
        cmdScn_Click
'        For RE = 0 To lboLst.ListCount
'            If InStr(lboLst.List(RE), "Controller (Wireless Gamepad F710)") > 0 Then
'                JOY_CONTROLLER = "Controller (Wireless Gamepad F710)"
'                Exit For
'            End If
'        Next
        '---------------------------------
    End If
    
    'MsgBox "ME 1"
    
    'IF ERROR HERE MOST LIKELY NOT INSTALLED THE LOGITECK SOFTWARE
    '---------------------------------------------------------------
    'REQUIRING DETECT TWO TYPE OF CONTROLLER
    'AS NEW MODERN ONE HAS SWAP BUTTON AROUND OF 1 TO 4 ON THE RIGHT
    '---------------------------------------------------------------
    'ERR.NUMBER = 91
    'Watch :   : Err.Description : "Object variable or With block variable not set" : String : frmDXGAMEJOY.Form_Load
    'Debug.Print Err.Description
    'NOT UNDERSTAND ERROR _ BUT WORK OKAY FOR GET JOYPAD
    'DISCARD IT USER
    '---------------
    
    'Err.Clear
    'Call RefreshList
    
    'MIGHT BE A CRASH PROBLEM IF JOYPAD SOFTWARE NOT INSTALLER CORRECT AS ABOVE ERROR 91
    'AND MAYBE HAS THE USE OF JOY PAD WITH SEMI DRIVER INSTALL REGISTER TYPE THING
    'BUT HAS A EXE COMPILED ERROR FREEZER
    'ONLY WAY IS IN IDE
    '------------------------------------
    
    
    
    If Err.Number = 0 Then 'And ERR_FILENAME_2 = False Then
        MNU_JOYPAD_SOFTWARE_STATUS.Visible = False
        RefreshList  'Update our controller display
        Exit Sub
    End If
    
    'MsgBox "ME 2"
    
    '-----------
    'RESULT ----
    'RefreshList
    'DEBUG.PRINT SUB ROUTINE ABOVE __
    'VERSION 1ST ONE __ Logitech Cordless RumblePad 2
    '1....Logitech Cordless RumblePad 2
    '---------------------------------------
    'I DID HAVE TWO OF THESE Logitech Cordless RumblePad 2
    'BUTTER LOST THE REMOTE FOR ONE WHILE MOVING OUT
    'DONE BY ANOTHER
    'AND ANY HOW NOT WORKER WITH THEM AS TWO REMOTE RECEIVER ARE SAME FREQENCY
    'VERSION 2__ ______ Controller (Wireless Gamepad F710)
    '2....Controller (Wireless Gamepad F710)
    '---------------------------------------
    
    '--------------------------------
    'READ HERE FOR A PROBLEM SORT_ER
    '2017
    'D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\DXGameControllers\dxwebsetup -- 2012.txt
    '--------------------------------
    
    '-----------------------------------------------------------------------
    'THE LESS SOPHISTICATED CONTROL SOFTWARE WHICH NOT LIKE THE DIRECTX ONE
    'IS UNABLE TO DETECT THE MODERN JOYPAD F710
    'AND MOST SOFTWARE WRITTEN FOR THAT ONE
    'MAYBE EAYS CONVERT TO THE DIRECT X
    'BUT FELT A PROBLEM BEFORE WITH EVERY INFINATE DETAIL
    'AND ALSO THAT CODE APPEAR TO BE MISSING
    '-----------------------------------------------------------------------
    'THE JOYPAD LESS VERSION BE NOT DIRECT X
    'IS NOT SHOW ANYTHING WHILE NOT IN ISIDE MODE
    'TO SAVE ON PROCESS WHILE EXE
    '-----------------------------------------------------------------------
    'YES EASY TO USER TWO CONTROLER AT ONCE THING WHILE RADIO SIGNAL OKAY NOT COMPARE SAME
    'EACH EVENT HAS CORSPONDING EVENT TO JOYPAD NUMBER_ER
    '-----------------------------------------------------------------------
    'TO WRITE INTERESTING SOFTWARE AGAIN LATER
    '-----------------------------------------------------------------------

    
    If Err.Number > 0 Or ERR_FILENAME_2 = True Then
        Dim FileName, A1, FR, A_JOY_2
        FileName = "Error With Joypad Driver Message Display Counter.txt"
        
        If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--", vbDirectory) = "" Then
            MkDir App.Path + "\00_Text_Data"
            MkDir App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--"
        End If
        
        If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\" + FileName) <> "" Then
            On Error Resume Next
            FR = FreeFile
            Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\" + FileName For Input As #FR
                Line Input #FR, A1
            Close #FR
        End If
            
        'If Err.Number = 0 Then
        On Error GoTo 0
            
        A_JOY_2 = Val(A1) + 1
        If A_JOY_2 > 10 Then A_JOY_2 = 1
        
        If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--", vbDirectory) = "" Then
            MkDir App.Path + "\00_Text_Data"
            MkDir App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--"
        End If
        FR = FreeFile
        Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\" + FileName For Output As #FR
            Print #FR, A_JOY_2
        Close #FR
    
        MSGB2 = ""
        MSGB2 = MSGB2 + "" ' + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "LOGITECK JOYPAD SOFTWARE ----" + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "ERROR DESCRIPTION = " + Err.Description + vbCrLf
        MSGB2 = MSGB2 + "ERROR NUMBER      =" + Str(Err.Number) + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "IF ERROR HERE MOST LIKELY NOT INSTALLED PROPER THE" + vbCrLf
        MSGB2 = MSGB2 + "LOGITECH JOYPAD SOFTWARE" + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "MY INSTALLER IS HERE" + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "D:\#0 1 INSTALLATIONS\Installation programs\# COMMS JOYPAD BLUETOOTH\JOY PAD\Demo32.exe" + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "MIGHT REQUIRE INSTALLER RUN TWICE WITH" + vbCrLf
        MSGB2 = MSGB2 + "REMOVE OPTION QUESTION FIRST RUN HITT IT BUT MISSED INFO ABOUT WHAT IT WAS" + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "WHEN INSTALLED JOYPAD SOFTWARE LOGITECH -- THE CONTROL PANEL AND OPTIONS IS HERE" + vbCrLf
        MSGB2 = MSGB2 + "C:\Program Files\Logitech\Gaming Software\LWEMon.exe" + vbCrLf
        
        If Dir("C:\Program Files\Logitech\Gaming Software\LWEMon.exe") <> "" Then RESULT_VAR = "TRUE ****" Else RESULT_VAR = "FALSE XXXX"
        MSGB2 = MSGB2 + "DECTION RESULT EXIST -- LWEMon.exe :- " + RESULT_VAR + vbCrLf + vbCrLf
        
        MSGB2 = MSGB2 + "AND THIS INSTALLED AND REGISTERED" + vbCrLf
        MSGB2 = MSGB2 + "regsvr32 C:\Windows\system32\dx8vb.dll" + vbCrLf
        
        If Dir("C:\Windows\System32\dx8vb.dll") <> "" Then RESULT_VAR = "TRUE ****" Else RESULT_VAR = "FALSE XXXX"
        MSGB2 = MSGB2 + "DECTION RESULT EXIST -- \Windows\System32\dx8vb.dll :- " + RESULT_VAR + vbCrLf
        
        If Dir("C:\Windows\SysWOW64\dx8vb.dll") <> "" Then RESULT_VAR = "TRUE *****" Else RESULT_VAR = "FALSE XXXX"
        MSGB2 = MSGB2 + "DECTION RESULT EXIST -- \Windows\SysWOW64\dx8vb.dll :- " + RESULT_VAR + vbCrLf + vbCrLf
        
        
        
'        MSGB2 = MSGB2 + "THIS DOWNLOAD HERE DirectX 9.0c End-User Runtime" + vbCrLf
'        MSGB2 = MSGB2 + "https://www.microsoft.com/en-in/download/confirmation.aspx?id=34429" + vbCrLf + vbCrLf
'
'        MSGB2 = MSGB2 + "THIS INSTALL IS 64 BIT IN \Program Files" + vbCrLf + vbCrLf
        
        
        MSGB2 = MSGB2 + "FOLLOW SOME OF HERE" + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\DXGameControllers\dxwebsetup -- 2012.txt" + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "LOADING NOTEAD" + vbCrLf + vbCrLf
        MSGB2 = MSGB2 + "MESSAGE HERE SET TO ONLY SHOW ONCE PER EACH EXE RUNNING PER 10 TIMES"
        
        
        
        '----------------------------
        MSGBOX_QUESTION = vbOK
        '----------------------------
        MSGB2 = MSGB2
        
        '------------------------------------------------------
        'NOT ABLE TO USE HERE THE CANCLE BUTTON IS ALWAYS ADDED
        '------------------------------------------------------
        'MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
        '------------------------------------------------------
        
        MNU_JOYPAD_SOFTWARE_STATUS.Visible = True
        If A2 = 1 Or FIRST_PASS = True Then
            Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
            If IsIDE = True And Result = vbCancel Then Stop
            Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
            '----------------------------
        
            'Shell "NOTEPAD D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\DXGameControllers\dxwebsetup -- 2012.txt", vbMaximizedFocus
        
            'A = A
            Dim WSHShell
            Set WSHShell = CreateObject("WScript.Shell")
    
            WSHShell.Run """" + App.Path + "\0 ROOT INFO\dxwebsetup -- 2012.txt" + """"
            Set WSHShell = Nothing

        
        End If
        FIRST_PASS = True
        
        
    End If
    On Error Resume Next
    If Err.Number = 0 And ERR_FILENAME_2 = False Then
        RefreshList  'Update our controller display
    End If
    
'    Call cmdScn_Click
    
End Sub

Function MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, MSGBOX_RESULT, MSGBOX_QUESTION)
Dim Result

'------------------------------
'REM OUT -- If Want to Stop in Subrountine Function
'If RESULT = vbCancel Then Stop
'------------------------------
MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN = MSGBOX_RESULT
If MSGBOX_RESULT = vbCancel Then
    'DON'T ASK AGAIN IF IS ONLY A OKAY QUESTION -------------------------
    If MSGBOX_QUESTION = vbOKCancel Then MSGBOX_QUESTION = vbOK: Exit Function
    If MSGBOX_QUESTION = vbYesNoCancel Then MSGBOX_QUESTION = vbYesNo
    Result = MsgBox("QUESTIONED ANSWERED - BY STOP -- HOW DO YOU WANT THIS QUESTION NOW " + vbCrLf + vbCrLf + MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
    MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN = Result
End If


End Function

Function MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)

MSGBOX_QUESTION_CONVERT_ISIDE = MSGBOX_QUESTION
If IsIDE = True Then
    If MSGBOX_QUESTION = vbOK Then MSGBOX_QUESTION = vbOKCancel
    If MSGBOX_QUESTION = vbYesNo Then MSGBOX_QUESTION = vbYesNoCancel
    MSGBOX_QUESTION_CONVERT_ISIDE = MSGBOX_QUESTION
End If

End Function


Private Sub Form_Unload(Cancel As Integer)

'If ATidalX.ALL_FORM_REQUEST_TO_END = False Then
'    Cancel = True: Me.Hide: Exit Sub
'End If

On Error Resume Next
mLst.Terminate  'Clean up the controller list
Set mLst = Nothing  'Discard the controller list
End Sub

Private Sub RefreshList()
    Dim i As Integer
    lboLst.Clear
    For i = 1 To mLst.Count  'Loop through each controller...
         With mLst.Item(i)
             lboLst.AddItem CStr(i) & ": " & .Name & " " & .GUID  '...and list some info
             Debug.Print .Name
             'VERSION 1ST ONE __ Logitech Cordless RumblePad 2
             'VERSION 2__ ______ Controller (Wireless Gamepad F710)
            If .Name = "Controller (Wireless Gamepad F710)" Then
                JOY_CONTROLLER_F710_PORT = i
            End If
            If .Name = "Logitech Cordless RumblePad 2" Then
                JOY_CONTROLLER_RumblePad_2_PORT = i
            End If
        End With
    Next


    For i = 0 To lboLst.ListCount
        If InStr(lboLst.List(i), "Controller (Wireless Gamepad F710)") > 0 Then
            JOY_CONTROLLER = "Controller (Wireless Gamepad F710)"
            Exit For
        End If
    Next

End Sub

Private Sub mLst_ElementChange(ByVal index As Integer, ByVal Element As ControllerElement, ByVal Pressed As Boolean)
    
    ' Index: The number of the controller
    ' Element: Which element changed
    ' Pressed: The state of the element
    
    Dim JOY_CONTROLLER_VAR
    
    If Pressed Then
        ElementName_VAR = ElementName(Element)
        lboEvt.AddItem "Controller " & CStr(index) & ": Pressed " & ElementName_VAR
    Else
        ElementName_VAR = ElementName(Element)
        lboEvt.AddItem "Controller " & CStr(index) & ": Released " & ElementName_VAR
    End If
    lboEvt.ListIndex = lboEvt.NewIndex
    lboEvt.TopIndex = lboEvt.NewIndex
    If lboEvt.ListCount > 20 Then lboEvt.RemoveItem (0)

    JOY_CONTROLLER = ""
    JOY_CONTROLLER_VAR = lboLst.List(index - 1)
    If InStr(JOY_CONTROLLER_VAR, "Controller (Wireless Gamepad F710)") > 0 Then
        JOY_CONTROLLER = "Controller (Wireless Gamepad F710)"
    End If
    If InStr(JOY_CONTROLLER_VAR, "Logitech Cordless RumblePad 2") > 0 Then
        JOY_CONTROLLER = "Logitech Cordless RumblePad 2"
    End If
    
    
    
    
    '--------------------------
    'TALK DATE
    '--------------------------
'    If JOY_CONTROLLER = "" Or JOY_CONTROLLER <> "" Then
'        If ElementName_VAR = "" Then
'            If Pressed = True Then
'                'ATidalX.Label7_Click
'            End If
'        End If
'    End If
    
    '--------------------------
    'TALK TIME
    '--------------------------
'    If JOY_CONTROLLER = "" Then
'        If ElementName_VAR = "" Then
'            If Pressed = True Then
'                'ATidalX.LABEL_TIME_Click
'            End If
'        End If
'    End If
    
    
    '--------------------------
    'TEXT TO SPEECH __ ME
    '--------------------------
    'D:\0 00 MUSIC ---\01 SOUND EFFECT & TECHNO SAMPLES\Smalls My Audio Recording Sounds\180419_0012_2018-04-19--10-50__ME__WORD_TEXT_TO_SPEECH_SOUND_EFFECT__WWW.TEXTOSPEECH.INFO.MP3
    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
        If ElementName_VAR = "Up" Then
            If Pressed = True Then
'                ATidalX.ME_PLAY_TEXT_TO_SPEECH_TIMER.Enabled = True
            End If
            If Pressed = False Then
'                ATidalX.ME_PLAY_TEXT_TO_SPEECH_TIMER.Enabled = False
            End If
        End If
    End If
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Up" Then
            If Pressed = True Then
'                ATidalX.ME_PLAY_TEXT_TO_SPEECH_TIMER.Enabled = True
            End If
            If Pressed = False Then
'                ATidalX.ME_PLAY_TEXT_TO_SPEECH_TIMER.Enabled = False
            End If
        End If
    End If
    
    '--------------------------
    'TALK TIME
    '--------------------------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Left" Then
            If Pressed = True Then
'                ATidalX.MNU_SPEECH_OFF.Checked = False
'                Call ATidalX.SPEECH_OFF_SETTER

'                Call ATidalX.FORCESAYTIME
            End If
        End If
    End If
    
    '--------------------------
    'TALK DATE
    '--------------------------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Right" Then
            If Pressed = True Then
                'SAY DATE
'                ATidalX.MNU_SPEECH_OFF.Checked = False
'                Call ATidalX.SPEECH_OFF_SETTER
'                Call ATidalX.TIMER_SECONDS20_AFTER_HOUR_TIMER_Timer
            End If
        End If
    End If
    
    
    '--------------------------
    'F-FORWARD
    '--------------------------
    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
        If ElementName_VAR = "Button 8" And Pressed = True Then
'            MOTION_F_FORWARD_TRIGGER = True
        End If
        If ElementName_VAR = "Button 8" And Pressed = False Then
'            MOTION_F_FORWARD_TRIGGER = False
        End If
    End If
    '--------------------------
    'REWIND
    '--------------------------
    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
        If ElementName_VAR = "Button 7" And Pressed = True Then
'            MOTION_REWIND_TRIGGER = True
        End If
        If ElementName_VAR = "Button 7" And Pressed = False Then
'            MOTION_REWIND_TRIGGER = False
        End If
    End If
    
    '--------------------------
    'F-FORWARD
    '--------------------------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "In" And Pressed = True Then
'            MOTION_F_FORWARD_TRIGGER = True
        End If
        If ElementName_VAR = "In" And Pressed = False Then
'            MOTION_F_FORWARD_TRIGGER = False
        End If
    End If
    '--------------------------
    'REWIND
    '--------------------------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Out" And Pressed = True Then
'            MOTION_REWIND_TRIGGER = True
        End If
        If ElementName_VAR = "Out" And Pressed = False Then
'            MOTION_REWIND_TRIGGER = False
        End If
    End If

    
    
    
    '--------------------------
    'VOLUME UP
    '--------------------------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Button 6" And Pressed = True Then
'            JOY_VOLUME_UP_TRIGGER = True
        End If
        If ElementName_VAR = "Button 6" And Pressed = False Then
'            JOY_VOLUME_UP_TRIGGER = False
        End If
    End If
    '--------------------------
    'VOLUME DOWN
    '--------------------------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Button 5" And Pressed = True Then
'            JOY_VOLUME_DOWN_TRIGGER = True
        End If
        If ElementName_VAR = "Button 5" And Pressed = False Then
'            JOY_VOLUME_DOWN_TRIGGER = False
        End If
    End If

    '--------------------------
    'VOLUME UP
    '--------------------------
    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
        If ElementName_VAR = "Button 6" And Pressed = True Then
'            JOY_VOLUME_UP_TRIGGER = True
        End If
        If ElementName_VAR = "Button 6" And Pressed = False Then
'            JOY_VOLUME_UP_TRIGGER = False
        End If
    End If
    '--------------------------
    'VOLUME DOWN
    '--------------------------
    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
        If ElementName_VAR = "Button 5" And Pressed = True Then
'            JOY_VOLUME_DOWN_TRIGGER = True
        End If
        If ElementName_VAR = "Button 5" And Pressed = False Then
'            JOY_VOLUME_DOWN_TRIGGER = False
        End If
    End If


    '--------------------------
    'Next Track
    '--------------------------
    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
        If ElementName_VAR = "Button 2" Then
'            If Pressed = True Then
'                If JOY_NEXT_TRACK_TRIGGER = False Then
'                    JOY_NEXT_TRACK_TRIGGER = True
'                    Call WinAmpCommands(3)
'                End If
'            End If
'            If Pressed = False Then
'                JOY_NEXT_TRACK_TRIGGER = False
'            End If
        End If
    End If
    
    '--------------------------
    'Previous Track
    '--------------------------
    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
        If ElementName_VAR = "Button 4" Then
'            If Pressed = True Then
'                If JOY_PREVIOUS_TRACK_TRIGGER = False Then
'                    JOY_PREVIOUS_TRACK_TRIGGER = True
'                    Call WinAmpCommands(6)
'                End If
'            End If
'            If Pressed = False Then
'                JOY_PREVIOUS_TRACK_TRIGGER = False
'            End If
        End If
    End If
    
    
    
    
    
    '--------------------------
    'Next Track F710
    '--------------------------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Button 1" Then
'            If Pressed = True Then
'                If JOY_NEXT_TRACK_TRIGGER = False Then
'                    JOY_NEXT_TRACK_TRIGGER = True
'                    Call WinAmpCommands(3)
'                End If
'            End If
'            If Pressed = False Then
'                JOY_NEXT_TRACK_TRIGGER = False
'            End If
        End If
    End If
    
    '--------------------------
    'Previous Track
    '--------------------------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Button 4" Then
'            If Pressed = True Then
'                If JOY_PREVIOUS_TRACK_TRIGGER = False Then
'                    JOY_PREVIOUS_TRACK_TRIGGER = True
'                    Call WinAmpCommands(6)
'                End If
'            End If
'            If Pressed = False Then
'                JOY_PREVIOUS_TRACK_TRIGGER = False
'            End If
        End If
    End If
    
    
    
    '----------
    'PAUSE R 2
    '----------
    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
        If ElementName_VAR = "Button 3" Then
'            If Pressed = True Then
'                If JOY_PAUSE_TRIGGER = False Then
'                    JOY_PAUSE_TRIGGER = True
'                    Call WinAmpCommands(2)
'                End If
'            End If
'            If Pressed = False Then
'                JOY_PAUSE_TRIGGER = False
'            End If
        End If
    End If

    '----------
    'PAUSE F710
    '----------
    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
        If ElementName_VAR = "Button 2" Then
'            If Pressed = True Then
'                If JOY_PAUSE_TRIGGER = False Then
'                   JOY_PAUSE_TRIGGER = True
'                    Call WinAmpCommands(2)
'                End If
'            End If
'            If Pressed = False Then
'                JOY_PAUSE_TRIGGER = False
'            End If
        End If
    End If

'    Call FrmJoy.Timer1_Timer

End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
  'TEST
  'IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

'***********************************************
'***********************************************


