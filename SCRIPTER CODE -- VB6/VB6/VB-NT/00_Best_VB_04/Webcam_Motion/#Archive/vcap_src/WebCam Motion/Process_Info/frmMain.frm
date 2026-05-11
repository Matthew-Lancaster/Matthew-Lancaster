VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Is This EXE Running?"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
