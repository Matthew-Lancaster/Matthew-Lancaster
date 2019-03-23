Attribute VB_Name = "MemCap"
'*
'* Author: E. J. Bantz Jr.
'* Copyright: None, use and distribute freely ...
'* E-Mail: ejbantz@usa.net
'* Web: http://www.inlink.com/~ejbantz

'// ------------------------------------------------------------------
'//  Windows API Constants / Types / Declarations
'// ------------------------------------------------------------------
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = 1
Public Const SWP_NOZORDER = &H4
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SM_CYCAPTION = 4
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const WS_EX_TRANSPARENT = &H20&
Public Const GWL_STYLE = (-16)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


'// Memory manipulation
Declare Function lStrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Declare Function lStrCpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As Any, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (ByVal hpvDest As Long, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub hmemcpy Lib "kernel32" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
    
'// Window manipulation
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public lwndC As Long       ' Handle to the Capture Windows

Function MyFrameCallback(ByVal lwnd As Long, ByVal lpVHdr As Long) As Long

    Debug.Print "FrameCallBack"
    
    Dim VideoHeader As VIDEOHDR
    Dim VideoData() As Byte
    
    '//Fill VideoHeader with data at lpVHdr
    RtlMoveMemory VarPtr(VideoHeader), lpVHdr, Len(VideoHeader)
    
    '// Make room for data
    ReDim VideoData(VideoHeader.dwBytesUsed)
    
    '//Copy data into the array
    RtlMoveMemory VarPtr(VideoData(0)), VideoHeader.lpData, VideoHeader.dwBytesUsed

    Debug.Print VideoHeader.dwBytesUsed
    Debug.Print VideoData
    
End Function

Function MyYieldCallback(lwnd As Long) As Long

    Debug.Print "Yield"

End Function

Function MyErrorCallback(ByVal lwnd As Long, ByVal iID As Long, ByVal ipstrStatusText As Long) As Long
    
    If iID = 0 Then Exit Function
    
    Dim sStatusText As String
    Dim usStatusText As String
    
    'Convert the Pointer to a real VB String
    sStatusText = String$(255, 0)                                      '// Make room for message
    lStrCpy StrPtr(sStatusText), ipstrStatusText                       '// Copy message into String
    sStatusText = Left$(sStatusText, InStr(sStatusText, Chr$(0)) - 1)  '// Only look at left of null
    usStatusText = StrConv(sStatusText, vbUnicode)                     '// Convert Unicode
            
    LogError usStatusText, iID

End Function

Function MyStatusCallback(ByVal lwnd As Long, ByVal iID As Long, ByVal ipstrStatusText As Long) As Long

    If iID = 0 Then Exit Function
   
    Dim sStatusText As String
    Dim usStatusText As String
    
    '// Convert the Pointer to a real VB String
    sStatusText = String$(255, 0)                                      '// Make room for message
    lStrCpy StrPtr(sStatusText), ipstrStatusText                       '// Copy message into String
    sStatusText = Left$(sStatusText, InStr(sStatusText, Chr$(0)) - 1)  '// Only look at left of null
    usStatusText = StrConv(sStatusText, vbUnicode)                     '// Convert Unicode
    
    frmMain.StatusBar.SimpleText = usStatusText
    Debug.Print "Status: ", usStatusText, iID

    Select Case iID '

    
    End Select


End Function

Sub ResizeCaptureWindow(ByVal lwnd As Long)

    Dim CAPSTATUS As CAPSTATUS
    Dim lCaptionHeight As Long
    Dim lX_Border As Long
    Dim lY_Border As Long
    
    
    lCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
    lX_Border = GetSystemMetrics(SM_CXFRAME)
    lY_Border = GetSystemMetrics(SM_CYFRAME)
    
    '// Get the capture window attributes .. width and height
    If capGetStatus(lwnd, VarPtr(CAPSTATUS), Len(CAPSTATUS)) Then
        
        '// Resize the capture window to the capture sizes
        SetWindowPos lwnd, HWND_BOTTOM, 0, 0, _
                           CAPSTATUS.uiImageWidth + (lX_Border * 2), _
                           CAPSTATUS.uiImageHeight + lCaptionHeight + (lY_Border * 2), _
                           SWP_NOMOVE Or SWP_NOZORDER
         'we = Me.hWnd
        frmMain.Width = (CAPSTATUS.uiImageWidth + (lX_Border * 2)) * Screen.TwipsPerPixelX + 150
        frmMain.Height = (CAPSTATUS.uiImageHeight + lCaptionHeight + (lY_Border * 2)) * Screen.TwipsPerPixelY + 1000
    End If

    Debug.Print "Resize Window."

End Sub

Function MyVideoStreamCallback(lwnd As Long, lpVHdr As Long) As Long

    Beep  '// Replace this with your code!
  
End Function

Function MyWaveStreamCallback(lwnd As Long, lpVHdr As Long) As Long

    Debug.Print "WaveStream"

End Function

Sub LogError(txtError As String, lID As Long)

    frmMain.StatusBar.SimpleText = txtError
    Debug.Print "Error: ", txtError, lID
 
End Sub
