Attribute VB_Name = "modFrmOps"
'This module is for basic api declares and some form effects
'Author: Matt Gillmore (SCO_STINKS@Yahoo.com)

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal _
    hdc As Long) As Long
Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal _
    crColor As Long) As Long
Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal lngX As Long, ByVal lngY As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal lngXSrc As Long, ByVal lngYSrc As Long, _
    ByVal dwRop As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long


Type POINTAPI
        X As Long
        Y As Long
End Type

Public Sub ExplodeForm(f As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, Cx%, Cy%
Dim TheScreen As Long
Dim Brush As Long
    
    GetWindowRect f.hWnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    For i = 1 To Movement
        Cx = formWidth * (i / Movement)
        Cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - Cx) / 2
        Y = myRect.Top + (formHeight - Cy) / 2
        Rectangle TheScreen, X, Y, X + Cx, Y + Cy
        DoEvents
    Next i
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
End Sub
Public Sub ImplodeForm(f As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, Cx%, Cy%
Dim TheScreen As Long
Dim Brush As Long
    
    GetWindowRect f.hWnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    For i = Movement To 1 Step -1
        Cx = formWidth * (i / Movement)
        Cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - Cx) / 2
        Y = myRect.Top + (formHeight - Cy) / 2
        Rectangle TheScreen, X, Y, X + Cx, Y + Cy
        DoEvents
    Next i
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
End Sub
Function GetNumTracks%()
Dim s As String * 30
    
    GetNumTracks = "0"
    mciSendString "status cd number of tracks wait", s, Len(s), 0
    
    If Left(s, 1) <> vbNullChar Then
        
        GetNumTracks = CInt(Mid$(s, 1, 2))

    End If

End Function

Function GetCurTrack() As Integer

Dim lRet As Integer
Dim c1(2) As POINTAPI
Dim sPos As String * 30
Dim sTrack As Integer
Dim sMin As Integer
Dim sSec As Integer

    lRet = mciSendString("status cd position", sPos, Len(sPos), 0)
    If lRet = 0 Then
        GetCurTrack = CInt(Mid$(sPos, 1, 2))
    End If
    
End Function

