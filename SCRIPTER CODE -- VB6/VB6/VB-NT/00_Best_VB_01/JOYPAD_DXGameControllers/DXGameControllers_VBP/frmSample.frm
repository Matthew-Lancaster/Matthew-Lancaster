VERSION 5.00
Begin VB.Form frmDXGAMEJOY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DXGameControllers"
   ClientHeight    =   3336
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8784
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3336
   ScaleWidth      =   8784
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
      Height          =   2544
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
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   2355
   End
   Begin VB.ListBox lboLst 
      Height          =   2544
      ItemData        =   "frmSample.frx":0004
      Left            =   120
      List            =   "frmSample.frx":0006
      TabIndex        =   0
      Top             =   60
      Width           =   4755
   End
End
Attribute VB_Name = "frmDXGAMEJOY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
   
    Dim MSGB2, MSGBOX_QUESTION, Result, RESULT_VAR
   
'   If Dir("C:\Program Files\Logitech\Gaming Software\LWEMon.exe") = "" Then
   
    DoEvents
    Set mLst = New cControllerList  'Create an instance of the controller list.
    On Error Resume Next
    Err.Clear
    mLst.Initialize Me, False  'Initialize the list using this form as the owner.
    
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
    
'    Err.Clear
    Call RefreshList
    
    If Err.Number = 0 Then
        'ATidalX.MNU_JOYPAD_SOFTWARE_STATUS.Visible = False
        RefreshList  'Update our controller display
        Exit Sub
    End If
    
    If Err.Number > 0 Then
        Dim FileName, A1, A2, FR
        FileName = "Error With Joypad Driver Message Display Counter.txt"
        If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\" + FileName) <> "" Then
            On Error Resume Next
            FR = FreeFile
            Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\" + FileName For Input As #FR
                Line Input #FR, A1
            Close #FR
            If Err.Number = 0 Then
                On Error GoTo 0
                
                A2 = Val(A1) + 1
                If A2 > 10 Then
                    A2 = 1
                    FR = FreeFile
                    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\" + FileName For Output As #FR
                        Print #FR, "1"
                    Close #FR
                End If
            End If
        End If
    
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
        MSGB2 = MSGB2 + "THIS DOWNLOAD HERE DirectX 9.0c End-User Runtime" + vbCrLf
        MSGB2 = MSGB2 + "https://www.microsoft.com/en-in/download/confirmation.aspx?id=34429" + vbCrLf + vbCrLf
        
        MSGB2 = MSGB2 + "THIS INSTALL IS 64 BIT IN \Program Files" + vbCrLf + vbCrLf
        
        '----------------------------
        MSGBOX_QUESTION = vbOK
        '----------------------------
        MSGB2 = MSGB2
        
        '------------------------------------------------------
        'NOT ABLE TO USE HERE THE CANCLE BUTTON IS ALWAYS ADDED
        '------------------------------------------------------
        'MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
        '------------------------------------------------------
        
'        ATidalX.MNU_JOYPAD_SOFTWARE_STATUS.Visible = True
        If A2 = 1 Or FIRST_PASS = True Then
            Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
'            If IsIDE = True And Result = vbCancel Then Stop
            Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
            '----------------------------
        End If
        FIRST_PASS = True
    End If
    RefreshList  'Update our controller display
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
   mLst.Terminate  'Clean up the controller list
   Set mLst = Nothing  'Discard the controller list
End Sub

Private Sub RefreshList()
   Dim i As Integer
   lboLst.Clear
   For i = 1 To mLst.Count  'Loop through each controller...
      With mLst.Item(i)
         lboLst.AddItem CStr(i) & ": " & .Name & " " & .GUID  '...and list some info
      End With
   Next
End Sub

Private Sub mLst_ElementChange(ByVal index As Integer, ByVal Element As ControllerElement, ByVal Pressed As Boolean)
' Index: The number of the controller
' Element: Which element changed
' Pressed: The state of the element
   If Pressed Then
      lboEvt.AddItem "Controller " & CStr(index) & ": Pressed " & ElementName(Element)
   Else
      lboEvt.AddItem "Controller " & CStr(index) & ": Released " & ElementName(Element)
   End If
   lboEvt.ListIndex = lboEvt.NewIndex
   lboEvt.TopIndex = lboEvt.NewIndex
End Sub
