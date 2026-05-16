VERSION 5.00
Begin VB.Form frmMain2 
   Caption         =   "Is This EXE Running?"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3465
      Top             =   1845
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1793
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtEXEName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "explorer.exe"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      Caption         =   $"frmMain.frx":0000
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblEXEName 
      AutoSize        =   -1  'True
      Caption         =   "Filename of EXE To Check:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1965
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XXD1, XXH1, XXH2, XXH3, OTrig, RunOnce
Dim Active, OActive, Huge2, ActiveX, OHuge2
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
    '#### Functions/Consts used for GetHWnd() (part of Convert())
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long  'MODULE 1141
'Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'Public Const HWND_TOPMOST = -1
'Public Const HWND_NOTOPMOST = -2
'Public Const SWP_NOACTIVATE = &H10
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOSIZE = &H1
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDPREV = 3
Private Const GWL_WNDPROC = (-4)
Private Const IDANI_OPEN = &H1
Private Const IDANI_CLOSE = &H2
Private Const IDANI_CAPTION = &H3
Private Const WM_USER = &H400

Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long


'#####################################################################################
'#  Example program code for clsCnProc usage
'#      By: Nick Campbeln
'#
'#      Revision History:
'#          1.0 (Aug 28, 2002):
'#              Initial Release
'#
'#      Copyright © 2002 Nick Campbeln (opensource@nick.campbeln.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  NOTE: WinNT4 support requires that "PSAPI.dll" be present in the target WinNT4 system's path (current directory, system directory, etc). This DLL is not installed by default on WinNT4, so be advised. It is provided in the PSC.com zip as 'PSAPI-dll', simply rename to 'PSAPI.dll'.
'#
'#  Please remember to vote on PSC.com if you like this code!
'#  Code URL: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=38425&lngWId=1
'#####################################################################################


Private cProcesses As New clsCnProc


'#########################################################
'# cmdClose code to shutdown the program
'#########################################################
Private Sub cmdClose_Click()
    Call Unload(Me)
End Sub


'#########################################################
'# cmdClose code to determine if the EXEName in txtEXEName is running
'#########################################################
Private Sub cmdTest_Click()
        
        '#### If the EXE is running, let the user know
    If (IsThisEXE(txtEXEName.Text)) Then
        Call MsgBox("'" & txtEXEName.Text & "' is currently running on the local system.", vbOKOnly + vbInformation, "'" & txtEXEName.Text & "' Is Running!")

        '#### Else the EXE is not running
    Else
        Call MsgBox("'" & txtEXEName.Text & "' is not running on the local system.", vbOKOnly + vbCritical, "'" & txtEXEName.Text & "' Is Not Running.")
    End If
End Sub


'#########################################################
'# Tests For a Running EXE (referenced by it's passed filename)
'#########################################################
Public Function IsThisEXE(ByVal sRunningEXE As String) As Boolean
    Dim lProcessID As Long

        '#### Default the return value
    IsThisEXE = False

        '#### If we're able to successfully open the processes
    If (cProcesses.OpenProcesses()) Then
            '#### UCase sRunningEXE
        sRunningEXE = UCase(sRunningEXE)

            '#### Do while we still have processes
        Do
                '#### If the EXE names match
            If (InStr(1, UCase(cProcesses.EXE), sRunningEXE, vbBinaryCompare) > 0) Then
                IsThisEXE = True
                Exit Do
            End If
        Loop While (cProcesses.NextProcess())
    End If

        '#### Close the class
    Call cProcesses.CloseProcesses
End Function

Private Sub Form_Load()
        
    frmMain.Show
        
'    frmMain.PDHQuery.Class_Terminate
'    Set PDHQuery = Nothing

End Sub

Private Sub Timer1_Timer()
Dim Trig, Trig2
XXD1 = FindWindow(vbNullString, "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-]")
If XXD1 > 0 Then Trig2 = 1

'If XXD1 <> XXH1 And XXD1 > 0 Then
'    XXH1 = XXD1
'End If

XXD1 = FindWindow("Winamp v1.x", vbNullString)
If XXD1 > 0 Then Trig2 = Trig2 + 2

'If XXD1 <> XXH2 And XXD1 > 0 Then
'    XXH2 = XXD1
'End If

Trig2 = GetForegroundWindow
If OTrig <> Trig2 Then
'    MUSTUNLOAD = False
    Call frmMain.Reload
'    Unload frmMain
'   If Trig2 > 0 Then DoEvents: DoEvents: DoEvents: frmMain.Show
    OTrig = Trig2
'    MUSTUNLOAD = True
End If



End Sub





Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function



Function FindWinPart() As Long

Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String, i, Huge
'ReDim Preserve Huge(30)

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
'test_hwnd = GetDesktopWindow
Do While test_hwnd <> 0
    
'        ash$ = GetWindowTitle(test_hwnd)
        ash$ = GetWindowClass(test_hwnd)
        If ash$ <> "" And InStr(ash$, "Winamp v1.x") > 0 Then
            Huge = test_hwnd
        End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

FindWinPart = False
If Huge > 0 Then FindWinPart = Huge

End Function



Function GetActiveWindowClass(ReturnParent As Long) As String
    GetActiveWindowClass = GetWindowClass(ReturnParent)
End Function

Function GetWindowClass(ByVal hwnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hwnd, sText, 255)
    sText = Left$(sText, Ret)
   GetWindowClass = sText
End Function


