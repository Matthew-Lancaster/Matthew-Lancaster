VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Public WithEvents Mouse As clsMouse


Option Explicit

Public Event PositionChanged()
Public Event SytemClick(ByVal Button As MouseButtonConstants)

Private Declare Function GetAsyncKeyState Lib "user32" _
        (ByVal vKey As Long) As Integer
        
Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
             
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal _
        wCount As Long) As Long
        
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
                          
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
        ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, _
        ByVal dwExtraInfo As Long)
 
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) _
        As Long
                          
Private Declare Function LoadCursorFromFile Lib "user32" Alias _
        "LoadCursorFromFileA" (ByVal lpFileName As String) As Long

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
        As Long
                                                  
Private Declare Function SystemParametersInfo Lib "user32" Alias _
        "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam _
        As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As _
        Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex _
        As Long) As Long

Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap _
        As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As _
        Long) As Long

Private Declare Function GetCapture Lib "user32" () As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function BlockInput Lib "user32" (ByVal fBlock As _
        Long) As Boolean


Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor _
        As Any) As Long
  
Private Declare Function SetSystemCursor Lib "user32" (ByVal hCur As Long, _
        ByVal id As Long) As Long
  
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" _
        (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Const SM_CMOUSEBUTTONS = 43
Private Const SM_MOUSEWHEELPRESENT = 75
Private Const SM_MOUSEPRESENT = 19
Private Const SM_SWAPBUTTON = 23

Private Const SPI_GETMOUSE = 3
Private Const SPI_GETMOUSEKEYS = 54
Private Const SPI_GETMOUSETRAILS = 94
Private Const SPI_GETMOUSEHOVERTIME = 102
Private Const SPI_GETMOUSEHOVERHEIGHT = 100
Private Const SPI_GETMOUSEHOVERWIDTH = 98

Private Const SPI_SETMOUSESPEED = 113
Private Const SPI_GETMOUSESPEED = 112
        
Private Const GCLNG_CURSOR = -12
        
Private Const VK_RBUTTON = &H2
Private Const VK_MBUTTON = &H4
Private Const VK_LBUTTON = &H1
        
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Private Const OCR_NORMAL = 32512

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type CURSORTYPE
    hWnd As Long
    hCursor As Long
    Classname As String
End Type

Private m_Key As String
Private m_WatchPosition As Boolean
Private m_WatchSystemClicks As Boolean
Private m_Position As New clsMousePosition
Private m_Clip As New clsMouseClip
Private m_Visible As Boolean
Private m_Blocked As Boolean

Private hNewCursor As Long
Private hOldCursor() As CURSORTYPE
Private hNewSysCursor As Long
Private hOldSysCursor As Long

Private AniParent As Form
Private AniCursor As avbLocationConstants
Private ClassNames As String

Public Enum avbLocationConstants
    avbGlobal = 1
    avbLocal = 2
End Enum

Private Sub Class_Initialize()
    Call modMouse.AddObject(Me, m_Key)
    m_Visible = True
End Sub

Private Sub Class_Terminate()
    If Not m_Visible Then Me.Visible = True
    If m_Blocked Then Me.Block = False
    
    If AniCursor = avbLocal Then
        Call DeleteAniCursor
    ElseIf AniCursor = avbGlobal Then
 '       Call DeleteAniCursorGlobal
    End If
    
    Call modMouse.RemoveObject(m_Key)
    
    Set m_Position = Nothing
    Set m_Clip = Nothing
End Sub

Public Property Get ButtonsSwapped() As Boolean
    ButtonsSwapped = (GetSystemMetrics(SM_SWAPBUTTON) <> 0)
End Property

Public Property Let ButtonsSwapped(Value As Boolean)
    Call SwapMouseButton(IIf(Value, 1, 0))
End Property

Public Property Get Position() As clsMousePosition
    Set Position = m_Position
End Property

Public Property Get Clip() As clsMouseClip
    Set Clip = m_Clip
End Property

Public Property Let Visible(Value As Boolean)
    If Value <> m_Visible Then
        Call ShowCursor(IIf(Value, 1, 0))
        m_Visible = Value
    End If
End Property

Public Property Get Visible() As Boolean
     Visible = m_Visible
End Property

Public Property Get WheelExists() As Long
    WheelExists = (GetSystemMetrics(SM_MOUSEWHEELPRESENT) <> 0)
End Property

Public Property Get Exists() As Long
    Exists = (GetSystemMetrics(SM_MOUSEPRESENT) <> 0)
End Property

Public Property Get Buttons() As Long
    Buttons = GetSystemMetrics(SM_CMOUSEBUTTONS)
End Property

Public Property Get Capture() As Long
    Capture = GetCapture
End Property

Public Property Let Capture(hWnd As Long)
    Call SetCapture(hWnd)
End Property

Public Sub CaptureRelease()
    Call ReleaseCapture
End Sub

Public Property Get Speed() As Long
    Dim Result As Long
        
        Call SystemParametersInfo(SPI_GETMOUSESPEED, 0, Result, 0)
        Speed = Result
End Property

Public Property Let Speed(Value As Long)
    Dim V As Long
    
        V = Value
        If V < 1 Then V = 1
        If V > 20 Then V = 20
        
        Call SystemParametersInfo(SPI_SETMOUSESPEED, 0&, ByVal V, 0&)
End Property

Public Property Get DoubleClickTime() As Long
    DoubleClickTime = GetDoubleClickTime
End Property

Public Property Let DoubleClickTime(ms As Long)
    Dim V As Long
    
        V = ms
        If V < 1 Then V = 1
        If V > 2000 Then V = 2000
        Call SetDoubleClickTime(V)
End Property

Public Property Let WatchPosition(n_WatchPosition As Boolean)
    m_WatchPosition = n_WatchPosition
End Property

Public Property Get WatchPosition() As Boolean
    WatchPosition = m_WatchPosition
End Property

Public Property Let WatchSystemClicks(n_WatchSystemClicks As Boolean)
    m_WatchSystemClicks = n_WatchSystemClicks
End Property

Public Property Get WatchSystemClicks() As Boolean
    WatchSystemClicks = m_WatchSystemClicks
End Property

Public Property Let Block(Value As Boolean)
    Call BlockInput(IIf(Value, 1, 0))
    m_Blocked = Value
End Property

Public Property Get Block() As Boolean
    Block = m_Blocked
End Property

Public Sub Click(Optional MouseButton As MouseButtonConstants = vbLeftButton)
    If (MouseButton = vbLeftButton) Then
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
    ElseIf (MouseButton = vbMiddleButton) Then
        Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_MIDDLEUP, 0&, 0&, 0&, 0&)
    ElseIf (MouseButton = vbRightButton) Then
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0&, 0&, 0&, 0&)
    End If
End Sub

Public Sub Animate(ByRef Parent As Form, ByRef Path As String, Optional ByRef Location As avbLocationConstants = avbLocal)
    If Location = avbLocal Then
        If AniCursor = avbGlobal Then Call DeleteAniCursorGlobal
        
        Call DeleteAniCursor
        Call CreateAniCursor(Parent, Path)
    ElseIf Location = avbGlobal Then
        If AniCursor = avbLocal Then Call DeleteAniCursor
        
        Call CreateAniCursorGlobal(Path)
    End If
End Sub

Private Sub CreateAniCursorGlobal(Path As String)
    If (AniCursor = 0) And (Dir$(Path) <> "") And (Path <> "") Then
        'hOldSysCursor = CopyIcon(GetCursor())

        
        hOldSysCursor = hOldSysCursor2
        hNewSysCursor = LoadCursorFromFile(Path)
        Call SetSystemCursor(hNewSysCursor, OCR_NORMAL)
        AniCursor = avbGlobal
    End If
End Sub

Private Sub DeleteAniCursorGlobal()
    Dim hCursor As Long
    
    If AniCursor = avbGlobal Then
        hOldSysCursor = hOldSysCursor2
        Call SetSystemCursor(hOldSysCursor, OCR_NORMAL)
        Call DestroyCursor(hNewSysCursor)
        Call DestroyCursor(hOldSysCursor)
        AniCursor = 0
    End If
End Sub

Private Sub CreateAniCursor(Parent As Form, Path As String)
    Dim Obj As Object
    
        If (AniCursor = 0) And (Dir$(Path) <> "") And (Path <> "") Then
            ReDim hOldCursor(1)
            Set AniParent = Parent
            hNewCursor = LoadCursorFromFile(Path)
            
            ClassNames = Chr$(0)
            For Each Obj In AniParent.Controls
                If HashWnd(Obj) Then
                    If InStr(ClassNames, Chr$(0) & TypeName(Obj) & Chr$(0)) = 0 Then
                        
                        With hOldCursor(UBound(hOldCursor))
                            .hWnd = Obj.hWnd
                            .hCursor = SetClassLong(Obj.hWnd, GCLNG_CURSOR, hNewCursor)
                            .Classname = TypeName(Obj)
                            ClassNames = ClassNames & .Classname & Chr$(0)
                        End With
                        
                        ReDim Preserve hOldCursor(0 To UBound(hOldCursor) + 1)
                    End If
                End If
            Next Obj
            
            With hOldCursor(0)
                .hWnd = AniParent.hWnd
                .hCursor = SetClassLong(AniParent.hWnd, GCLNG_CURSOR, hNewCursor)
                .Classname = TypeName(AniParent)
                ClassNames = ClassNames & .Classname & Chr$(0)
            End With
            
            AniCursor = avbLocal
        End If
End Sub

Private Sub DeleteAniCursor()
    Dim Obj As Object, z As Long
    
        If GetUBound(hOldCursor) And (AniCursor = avbLocal) Then
            
            Call SetClassLong(hOldCursor(0).hWnd, GCLNG_CURSOR, hOldCursor(0).hCursor)
            ClassNames = Replace(ClassNames, Chr$(0) & hOldCursor(0).Classname & Chr$(0), Chr$(0))
            
            For Each Obj In AniParent.Controls
                If HashWnd(Obj) Then
                    
                    If InStr(ClassNames, Chr$(0) & TypeName(Obj) & Chr$(0)) > 0 Then
                        For z = 0 To UBound(hOldCursor) - 1
                            If hOldCursor(z).hWnd = Obj.hWnd Then
                                Call SetClassLong(Obj.hWnd, GCLNG_CURSOR, hOldCursor(z).hCursor)
                                ClassNames = Replace(ClassNames, Chr$(0) & hOldCursor(z).Classname & Chr$(0), Chr$(0))
                                
                                Exit For
                            End If
                        Next z
                        
                    End If
                End If
            Next Obj
            
            Call DestroyCursor(hNewCursor)
            Set AniParent = Nothing
        End If
        
        AniCursor = 0
End Sub

Private Function GetUBound(Field() As CURSORTYPE) As Long
    On Error Resume Next
    GetUBound = UBound(Field)
End Function

Private Function HashWnd(Obj As Object) As Boolean
    On Error GoTo Err
    If Obj.hWnd Then HashWnd = True
Err:
End Function

Public Sub Deactivate()
    Shell "Rundll32.exe mouse,disable"
End Sub

Private Function CompKey(KCode As Long) As Boolean
    Dim Result As Integer
    
        Result = GetAsyncKeyState(KCode)
        
        If Result = -32767 Then
            CompKey = True
        Else
            CompKey = False
        End If
End Function

Public Property Let TimerEvent(Dummmy As Boolean)
     Dim CurPos As POINTAPI, Value As MouseButtonConstants
     Static First As Boolean, mx As Long, my As Long
     
        If m_WatchPosition Then
            Call GetCursorPos(CurPos)
            
            If First Then
                If CurPos.x <> mx Or CurPos.y <> my Then
                    RaiseEvent PositionChanged
                End If
            End If
            
            mx = CurPos.x
            my = CurPos.y
        End If
        
        If m_WatchSystemClicks Then
            If CompKey(VK_LBUTTON) Then Value = vbLeftButton
            If CompKey(VK_RBUTTON) Then Value = Value Or vbRightButton
            If CompKey(VK_MBUTTON) Then Value = Value Or vbMiddleButton
            
            If Value <> 0 Then
                RaiseEvent SytemClick(Value)
            End If
        End If
        
        First = True
End Property
