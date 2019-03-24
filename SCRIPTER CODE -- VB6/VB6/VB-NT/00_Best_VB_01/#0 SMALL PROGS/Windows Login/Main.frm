VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Winlog"
   ClientHeight    =   765
   ClientLeft      =   4815
   ClientTop       =   3120
   ClientWidth     =   2070
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' This is where the app quits. Before it quit it must write the logof data
    ' to the logfile.
    
    Call LogOf
    
    ' Now unregister the app as a service
    Call UnMakeMeService
End Sub

Private Sub Timer1_Timer()

    ' This sub tests for the existence of e-winlog.dat in system dir.
    ' If it exists the app deletes it before terminating itself.
    ' The e-winlog.dat file is made by starting the app with the
    ' /q argument. Its the way to kill a already running instance of the app.
    
    ' It also tests for the existence of e-winlogFileInfo.dat in system dir.
    ' If it exists the app displays a messagebox about where the logfile is
    ' beeing made (if app was started with the /r argument).
    
    ' Then it tests for the existence og e-winlogViewLog.dat in system dir.
    ' If it exists the app will launch Notepad with the logfile.
    
    Dim test As Boolean
    
    ' test existance of e-winlog.dat in system dir.
    test = FileExists(WinSysDir & "e-winlog.dat")
    If test = True Then    ' file does exist
        Kill WinSysDir & "e-winlog.dat"
        End     ' Quit
    End If
    
    ' test existance of e-winlogFileInfo.dat in system dir.
    test = FileExists(WinSysDir & "e-winlogFileInfo.dat")
    If test = True Then    ' file does exist
        Kill WinSysDir & "e-winlogFileInfo.dat"
        MsgBox "Logfile for current " & App.ProductName & " session: " & LogfileDir & logFile, vbInformation
    End If

    ' test existance of e-winlogViewLog.dat in system dir.
    test = FileExists(WinSysDir & "e-winlogViewLog.dat")
    If test = True Then    ' file does exist
        Kill WinSysDir & "e-winlogViewLog.dat"
        Dim RetVal
        RetVal = Shell(WindDir & "NOTEPAD.EXE " & LogfileDir & logFile, 1)
    End If

End Sub
