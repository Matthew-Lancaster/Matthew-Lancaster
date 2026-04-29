VERSION 5.00
Begin VB.Form FrmRUN_PROCESS_RELOAD 
   Caption         =   "Form1"
   ClientHeight    =   3912
   ClientLeft      =   48
   ClientTop       =   696
   ClientWidth     =   12864
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3912
   ScaleWidth      =   12864
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9504
      Top             =   3132
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ELITE SPY MAINTAIN ____ RUNNER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   28.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2652
      Left            =   72
      TabIndex        =   0
      Top             =   204
      Width           =   15336
   End
   Begin VB.Menu MNU_MAIN_FORM 
      Caption         =   "BRING MAIN FRM _ELITESPY_ BACK SAME AS CLOSE_ IT HAPPENING"
   End
End
Attribute VB_Name = "FrmRUN_PROCESS_RELOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim O_HEIGHT, O_WIDTH
Dim MOUSE_DOWN
Dim COUNT_DOWN
Dim ENDER_HAPPEN
Dim RESIZE_ALLOW
Public cProcesses As New clsCnProc

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long



Private Sub Form_Activate()
    If APP_NAME_RELOAD_IT_ER____ = "" Then Unload Me
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
'Unload Me
Exit Sub
Dim LET_PASS
'If GetComputerName <> "7-ASUS-GL522VW" Then Unload Me: Exit Sub
If GetComputerName = "7-ASUS-GL522VW" And GetUserName <> "MATT 01" Then
    LET_PASS = False
    If InStr(UCase(APP_NAME_RELOAD_IT_ER_EXE), UCase("Explorer.exe")) > 0 Then LET_PASS = True
    If InStr(UCase(APP_NAME_RELOAD_IT_ER_EXE), UCase("ProcessLasso.exe")) > 0 Then LET_PASS = True
    If InStr(UCase(APP_NAME_RELOAD_IT_ER_EXE), UCase("GoodSync-v10.exe")) > 0 Then LET_PASS = True
    If LET_PASS = False Then Exit Sub
End If



Dim PID As Long
Dim APP_NAME_EXE_PASS_FOR_CALL_BACK As String

PID = -1
'i = GetEXEID_WILDCARD_PROCESS_SCRIPT(PID, APP_NAME_RELOAD_IT_ER_EXE)
APP_NAME_EXE_PASS_FOR_CALL_BACK = APP_NAME_RELOAD_IT_ER_EXE
i = cProcesses.GetEXEID(PID, APP_NAME_EXE_PASS_FOR_CALL_BACK)

If PID > 0 Then
    '-------------------------------------
    'TAKE A DELAY INSTANCE TO GET RUNNER
    'ABORT REPEAT ATTEMPT FROM FIND_WINDOW
    '-------------------------------------
    
    If UCase(APP_NAME_EXE_PASS_FOR_CALL_BACK) <> "EXPLORER.EXE" Then
        Timer1.Enabled = False
        ENDER_HAPPEN = True
        'Unload Me
        APP_NAME_RELOAD_IT_ER____ = ""
        DoEvents
        GoTo ENDER_SUB
    End If
End If

If APP_NAME_RELOAD_IT_ER____ = "" Then GoTo ENDER_SUB

frmMain.WindowState = vbMinimized

COUNT_DOWN = 40
Call Timer1_Timer
RESIZE_ALLOW = False

XW = Me.height
YH = Me.width
Debug.Print XW '4300
Debug.Print YH '15000
XW = 4300
YH = 15000

'FrmRUN_PROCESS_RELOAD
'HERE TOP PROJECT EXPLORER
'-------------------------
Me.height = XW
Me.width = YH
Label1.width = Me.width
Label1.height = Me.height
Label1.Left = 0
Label1.Top = 0
Label1.FontSize = 17

frmMain.WindowState = vbMinimized

ReDim Preserve BLOCK_RUN_1(20)
ReDim Preserve BLOCK_RUN_2(20)
ReDim Preserve BLOCK_RUN_3(20)

Me.Show
Me.SetFocus
RESIZE_ALLOW = True

Timer1.Enabled = True

ENDER_SUB:

'-----------------------------------------------
'If APP_NAME_RELOAD_IT_ER____ = "" Then Unload Me
'LOOK TO Form_Activate() FOR NEXT PART
'-----------------------------------------------
End Sub


Private Sub Form_Resize()

'If RESIZE_ALLOW = True Then COUNT_DOWN = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel = True
'Exit Sub

If ENDER_HAPPEN = True Then Exit Sub

VAR_ARRAY = VAR_ARRAY + 1
BLOCK_RUN_1(VAR_ARRAY) = Now + TimeSerial(0, 10, 0)
BLOCK_RUN_2(VAR_ARRAY) = APP_NAME_RELOAD_IT_ER_EXE
BLOCK_RUN_3(VAR_ARRAY) = False
Timer1.Enabled = False

APP_NAME_RELOAD_IT_ER____ = ""
Unload Me

End Sub

Private Sub Label1_Click()

Timer1.Enabled = False

APP_NAME_RELOAD_IT_ER____ = ""

Unload Me
End

End Sub

Private Sub Label1_DblClick()

VAR_ARRAY = VAR_ARRAY + 1
BLOCK_RUN_1(VAR_ARRAY) = Now + TimeSerial(0, 10, 0)
BLOCK_RUN_2(VAR_ARRAY) = APP_NAME_RELOAD_IT_ER_EXE
BLOCK_RUN_3(VAR_ARRAY) = False
Timer1.Enabled = False
APP_NAME_RELOAD_IT_ER____ = ""
Unload Me

End Sub

Private Sub MNU_MAIN_FORM_Click()

Unload Me

frmMain.WindowState = vbNormal

End Sub

Private Sub Timer1_Timer()

If APP_NAME_RELOAD_IT_ER____ = "" Then GoTo End_E
If Dir(APP_NAME_RELOAD_IT_ER_EXE) = "" Then GoTo End_E


On Error Resume Next
Me.SetFocus
On Error GoTo 0

If O_HEIGHT > 0 And O_WIDTH > 0 Then
    If O_HEIGHT <> Me.Top Or O_WIDTH <> Me.width Then COUNT_DOWN = 0
End If
O_HEIGHT = Me.Top
O_WIDTH = Me.width

COUNT_DOWN = COUNT_DOWN - 1

APP_NAME = "GOODSYNC DESKTOP ____"
APP_NAME = "PROCESS LASSO ____"

'DUPLICATE INFO __ A=A B=B
'-------------------------
APP_NAME_RELOAD_IT_ER____ = APP_NAME_RELOAD_IT_ER____
APP_NAME_RELOAD_IT_ER_EXE = APP_NAME_RELOAD_IT_ER_EXE

Label1.Caption = Trim(Str(COUNT_DOWN)) + " " + APP_NAME_RELOAD_IT_ER____ + vbCrLf + vbCrLf + APP_NAME_RELOAD_IT_ER____ + vbCrLf + "IS RUNNER" + vbCrLf + "HITT BUTTON HERE TO EXIT" + vbCrLf + "DOUBLE BUTTON TO TIMER" + vbCrLf + "AND CLOSE TO SHORT TIMER" + vbCrLf + "AND MOVE THE WINDOW GIVE IT A SHAKE" + vbCrLf + "OR RESIZE MINIMIZE BUTTON TO RUN NOW"
'" ELITE SPY __ MAINTAIN DO_ER"
Me.Caption = Replace(Replace(Label1.Caption, vbCrLf, " "), "__", "_")

RESIZE_ALLOW = True

If COUNT_DOWN > 0 Then Exit Sub

'APP_NAME_RELOAD_IT_ER_EXE = "C:\Program Files\Process Lasso\ProcessLasso.exe"
On Error Resume Next
Shell APP_NAME_RELOAD_IT_ER_EXE, vbMinimizedNoFocus
On Error GoTo 0


End_E:
Timer1.Enabled = False
ENDER_HAPPEN = True
APP_NAME_RELOAD_IT_ER____ = ""
Unload Me

End Sub





Public Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        
        '----------------------------------------------------------------------------
        'ROUTINE TAKEN FROM
        '----------------------------------------------------------------------------
        'SEND_TO_SCRIPT_IRFAR - Microsoft Visual Basic [design] - [mdlFileSys (Code)]
        'D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAR\#0 Send To Text List of Files and Sub Folders IRFAN.exe
        '----------------------------------------------------------------------------
        'MODIFIED A BIT DIR COMMAND REPLACE MORE COMPLEX WAY
        '----------------------------------------------------------------------------
        
        'If Not FolderExists(Left$(sPath, nPos - 1)) Then
        
        If Dir((Left$(sPath, nPos - 1)), vbDirectory) = "" Then
            MkDir Left$(sPath, nPos - 1)
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    'If Not FolderExists(sPath) Then MkDir sPath
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
    CreateFolderTree = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function

'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function



