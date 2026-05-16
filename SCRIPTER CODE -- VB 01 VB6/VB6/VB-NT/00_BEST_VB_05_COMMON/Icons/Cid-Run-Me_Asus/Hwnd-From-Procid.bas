Attribute VB_Name = "Module2"
Option Explicit

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Declare Function GetWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" _
    (ByVal hWnd As Long) As Long

Public lProcessID As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
        As Long

Public w34 As Long





Public Function hWndFromProcIDPlayList(ProcID As Long) As Long
    
Dim lhTmp           As Long
Dim lhTmp2          As Long
Dim lRetVal         As Long
Dim lCurrentProcID  As Long
Dim lpClassName As String
Dim sText As String
Dim Ret As Long


'sText = Space(255)
'Ret = GetClassName(winamp22handle2, sText, 255)
'sText = Left$(sText, Ret)
                
                
                
On Error GoTo errhWndFromProcID

Dim ash$

Dim Clex



'lhTmp = FindWindow(vbNullString, "DDE Server Window")

lhTmp = GetDesktopWindow()

lhTmp = GetWindow(lhTmp, GW_CHILD)
Do While lhTmp <> 0
    lRetVal = GetWindowThreadProcessId(lhTmp, lCurrentProcID)
    If lCurrentProcID = ProcID Then
        hWndFromProcIDPlayList = lhTmp
        ash$ = GetWindowTitle(lhTmp)
        If InStr(ash$, "Winamp Playlist Editor") Then
            sText = Space(255)
            Ret = GetClassName(lhTmp, sText, 255)
            sText = Left$(sText, Ret)
            If sText = "Winamp PE" Then
                Clex = 1: Exit Do
            End If
        End If
    End If
    lhTmp = GetWindow(lhTmp, GW_HWNDNEXT)
Loop

'w34 = w34 + 1: MsgBox (str(w34) + vbCr + str(lhTmp) + vbCr + str(ProcID) + vbCr + str(hWndFromProcID))

If Clex = 0 Then hWndFromProcIDPlayList = -2


Exit Function
errhWndFromProcID:
    hWndFromProcIDPlayList = 0
End Function



Public Function hWndFromProcID(ProcID As Long) As Long
    
Dim lhTmp           As Long
Dim lhTmp2          As Long
Dim lRetVal         As Long
Dim lCurrentProcID  As Long
Dim lpClassName As String
Dim sText As String
Dim Ret As Long


'sText = Space(255)
'Ret = GetClassName(winamp22handle2, sText, 255)
'sText = Left$(sText, Ret)
                
                
                
On Error GoTo errhWndFromProcID

Dim ash$

Dim Clex



'lhTmp = FindWindow(vbNullString, "DDE Server Window")

lhTmp = GetDesktopWindow()

lhTmp = GetWindow(lhTmp, GW_CHILD)
Do While lhTmp <> 0
    lRetVal = GetWindowThreadProcessId(lhTmp, lCurrentProcID)
    If lCurrentProcID = ProcID Then
        hWndFromProcID = lhTmp
        ash$ = GetWindowTitle(lhTmp)
        If InStr(ash$, "- Winamp") Or InStr(ash$, "Winamp 5.094") Then
            sText = Space(255)
            Ret = GetClassName(lhTmp, sText, 255)
            sText = Left$(sText, Ret)
            If sText = "Winamp v1.x" Then
                Clex = 1: Exit Do
            End If
        End If
    End If
    lhTmp = GetWindow(lhTmp, GW_HWNDNEXT)
Loop

'w34 = w34 + 1: MsgBox (str(w34) + vbCr + str(lhTmp) + vbCr + str(ProcID) + vbCr + str(hWndFromProcID))

If Clex = 0 Then hWndFromProcID = -2


Exit Function
errhWndFromProcID:
    hWndFromProcID = 0
End Function


