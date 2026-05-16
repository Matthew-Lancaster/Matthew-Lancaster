VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00A4A3A5&
   BorderStyle     =   0  'None
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox VUBox 
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   480
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   14
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1560
      TabIndex        =   12
      Top             =   240
      Width           =   3495
      Begin VB.Label lblstatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   540
         TabIndex        =   13
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      Height          =   2985
      ItemData        =   "mixer.frx":0000
      Left            =   4080
      List            =   "mixer.frx":0002
      TabIndex        =   11
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "X"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "STOP"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdVolume 
      Caption         =   "Vol"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdEject 
      Caption         =   "EJ"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.PictureBox FFTPanel 
      BackColor       =   &H00000000&
      DrawStyle       =   5  'Transparent
      FillStyle       =   0  'Solid
      ForeColor       =   &H000080FF&
      Height          =   3825
      Left            =   120
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   251
      TabIndex        =   2
      Top             =   1680
      Width           =   3825
   End
   Begin VB.PictureBox Scope 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   0
      Left            =   5280
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   1
      Top             =   5160
      Width           =   1080
   End
   Begin VB.PictureBox Scope 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H000080FF&
      Height          =   360
      Index           =   1
      Left            =   4080
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   0
      Top             =   5160
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'Matt Player
'******************************************************
'Thank you for Downloading MattPlayer.  This program has been
'created as an experiment with FFT (Fast Fourier Transforms) to
'visualize sound.  This application is a little rough around the
'edges but is fun to use and watch.
'NOTE:  The visualization will not work unless either the Wave OUT Mixer, or CDPlayer
'is selected in the volume control under the Recording Control.
'Please email me any bug fixes or enhancements at sco_stinks@yahoo.com
'
'Enjoy
'Matt Gillmore

'Kudos to Murphy McCauley for his FFT algorithm


Option Explicit

Private Type BGRQuad
  b As Byte
  g As Byte
  R As Byte
  Empty As Byte
End Type
Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, D@) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'Mixer-Vars
Dim hMixer&, SrcArr(20) As MIXERCONTROL, DstArr(20) As MIXERCONTROL
'WaveIn-Vars
Dim DevHandle&, RBuf(29) As WavBuf, RIdx&, FNr&, FSize&
Dim WithEvents xt As XTimer, pArr() As BGRQuad, MaxL&, MaxR&
Attribute xt.VB_VarHelpID = -1
Dim status As Boolean
Dim Num_Tracks As Long
Dim curTrack As Integer
Dim Cd_Open As Boolean
Dim gX As Long
Dim gY As Long
Dim offset As Integer

Private Sub cmdEject_Click()
Dim lRet As Long
    If Not Cd_Open Then
        lRet = mciSendString("set cd door open", 0, 0, 0)
        Cd_Open = True
    Else
        lRet = mciSendString("set cd door closed", 0, 0, 0)
        Cd_Open = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
Dim lRet As Long
    curTrack = GetCurTrack()
    curTrack = curTrack + 1
    If curTrack <> GetNumTracks Then
        
        If CheckIfPlaying = 1 Then
            mciSendString "play cd from " & CStr(curTrack), 0, 0, 0
        Else
            mciSendString "seek cd to " & CStr(curTrack), 0, 0, 0
        End If
    
    End If
End Sub

Private Sub cmdPlay_Click()
Dim lRet As Long
Dim sPos As String * 30
Dim nCurrentTrack As Integer
    
    lRet = mciSendString("play cd", 0&, 0, 0)
    nCurrentTrack = 1
    lRet = mciSendString("play cd from" & Str(nCurrentTrack), 0&, 0, 0)

End Sub



Private Sub cmdPrev_Click()
Dim lRet As Long
    If curTrack = 1 Then
        '
    Else
        curTrack = curTrack - 1
        If CheckIfPlaying = 1 Then
            lRet = mciSendString("play cd from" & Str(curTrack), 0&, 0, 0)
        Else
            mciSendString "seek cd to " & CStr(curTrack), 0, 0, 0
        End If
    End If
    
End Sub

Private Sub cmdStop_Click()
    Dim lRet As Long
    
    lRet = mciSendString("stop cd wait", 0&, 0, 0)

    DoEvents

End Sub

Private Sub cmdVolume_Click()
    Shell "sndvol32.exe", vbNormalFocus
End Sub



Public Sub SetTrackList(lnumfiles As Long)
Dim i As Integer
    List1.Clear
    For i = 1 To lnumfiles Step 1
        Form1.List1.AddItem "Track " & i
    Next

End Sub
Private Sub Form_Load()
Dim i&, mxl As MIXERLINE, MCaps As MIXERCAPS
    
    If Not InitDevices Then Unload Me: Exit Sub
    
    If mixerOpen(hMixer, 0, 0, 0, 0) Then Exit Sub
    
    Randomize
  
    mixerGetDevCaps hMixer, MCaps, Len(MCaps)
    For i = 0 To MCaps.cDestinations - 1
        mxl.cbStruct = Len(mxl)
        mxl.dwDestination = i
        mixerGetLineInfo hMixer, mxl, 0 'by DestinationLine
    Next i
  

    i = mciSendString("open cdaudio alias cd wait shareable", 0&, 0, 0)
    i = mciSendString("set cd time format tmsf", 0&, 0, 0)

    SetWindowRgn Form1.hWnd, CreateRoundRectRgn(0, 0, Form1.ScaleWidth, Form1.ScaleHeight, 150, 150), True
    SetWindowRgn List1.hWnd, CreateRoundRectRgn(0, 0, List1.Width, List1.Height, 15, 15), True
    SetWindowRgn Frame1.hWnd, CreateRoundRectRgn(0, 0, Frame1.Width, Frame1.Height, 15, 15), True
    Num_Tracks = GetNumTracks
    If Num_Tracks <> 0 Then
        SetTrackList Num_Tracks
    End If
    Set xt = New XTimer
    xt.Interval = 1
    
'    ExplodeForm Form1, 500
    StartListening
    curTrack = 1
    Cd_Open = False

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    status = True
    gX = X
    gY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lXdiff As Long
Dim LYdiff As Long

    lXdiff = gX - X
    LYdiff = gY - Y
    If status = True Then

        Form1.Move Form1.Left - lXdiff, Form1.Top - LYdiff
    End If
End Sub

Private Function FillControlInfos(mxl As MIXERLINE, Arr() As MIXERCONTROL, L As ListBox)
Dim i&, mxlc As MIXERLINECONTROLS, hMem&
    mxlc.cbStruct = Len(mxlc)
    mxlc.dwLineID = mxl.dwLineID
    mxlc.cControls = mxl.cControls
    mxlc.cbmxctrl = Len(Arr(0))
    hMem = GlobalAlloc(&H40, Len(Arr(0)) * mxlc.cControls)
    mxlc.pamxctrl = GlobalLock(hMem)
    L.Clear
    If mixerGetLineControls(hMixer, mxlc, 0) Then Exit Function
    
    mixerGetLineControls hMixer, mxlc, 0 'again, because mxlc.cControls is now exact
    
    For i = 0 To mxl.cControls - 1
        CopyStructFromPtr Arr(i), mxlc.pamxctrl + i * mxlc.cbmxctrl, Len(Arr(i))
        L.AddItem Left(Arr(i).szShortName, InStr(Arr(i).szShortName, Chr(0)) - 1)
    Next i
    
    If L.ListCount Then L.ListIndex = 0
    GlobalFree hMem
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    status = False
End Sub




Private Function SetValue(mxctl As MIXERCONTROL, ByVal Value&) As Boolean
Dim MCD As MIXERCONTROLDETAILS, ValArr&(50)
    
    MCD.cbStruct = Len(MCD)
    MCD.dwControlID = mxctl.dwControlID
    
    If mxctl.fdwControl And 2 Then
        MCD.item = mxctl.cMultipleItems
        ValArr(Value) = 1
    Else 'Default
        ValArr(0) = Value
    End If
    
    MCD.cbDetails = Len(ValArr(0))
    MCD.paDetails = VarPtr(ValArr(0))
    MCD.cChannels = 1
    
    If mixerSetControlDetails(hMixer, MCD, 0) Then Exit Function
    SetValue = True
End Function
 
Private Function GetValue&(mxctl As MIXERCONTROL)
Dim MCD As MIXERCONTROLDETAILS, ValArr&(50), i&
    MCD.cbStruct = Len(MCD)
    MCD.dwControlID = mxctl.dwControlID
    If mxctl.fdwControl And 2 Then MCD.item = mxctl.cMultipleItems
    
    MCD.cbDetails = Len(ValArr(0))
    MCD.paDetails = VarPtr(ValArr(0))
    MCD.cChannels = 1
    If mixerGetControlDetails(hMixer, MCD, 0) Then Exit Function
    GetValue = ValArr(0)
    If mxctl.fdwControl And 2 Then
        For i = 0 To MCD.item - 1
            If ValArr(i) Then Exit For
        Next i
        GetValue = i
    End If
End Function

'*******WaveIn-Example follows
Private Function InitDevices()
Dim Caps As WaveInCaps, i&
    DevicesBox.Clear
    For i = 0 To waveInGetNumDevs - 1
        waveInGetDevCaps i, Caps, Len(Caps)
        If Caps.Formats And WAVE_FORMAT_4S16 Then DevicesBox.AddItem StrConv(Caps.ProductName, vbUnicode)
    Next
    If DevicesBox.ListCount Then InitDevices = True: DevicesBox.ListIndex = 0
End Function

Private Sub StartListening()
Static WaveFormat As WaveFormatEx
On Error Resume Next
'    FNr = FreeFile
'    If Dir(App.Path & "Test.Wav") <> "" Then Kill App.Path & "Test.Wav"
'    Open App.Path & "Test.Wav" For Binary As FNr
'    FSize = 0
'    If Err Then MsgBox "Cannot open File!"
    
    With WaveFormat
      .FormatTag = WAVE_FORMAT_PCM
      .Channels = 2
      .SamplesPerSec = 44100
      .BitsPerSample = 16
      .BlockAlign = (.Channels * .BitsPerSample) \ 8
      .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
    End With
    
    waveInOpen DevHandle, DevicesBox.ListIndex, WaveFormat, 0, 0, 0
    If DevHandle = 0 Then Exit Sub
    
    waveInStart DevHandle
    xt.Enabled = True
End Sub

Private Sub StopButton_Click()
    DoStop
    Close FNr
End Sub

Private Sub DoStop()
Dim i&
    xt.Enabled = False
        If DevHandle Then
        For i = 0 To UBound(RBuf)
            waveInUnprepareHeader DevHandle, RBuf(i).Hdr, Len(RBuf(i).Hdr)
        Next i
        waveInReset DevHandle
        waveInClose DevHandle
        DevHandle = 0
    End If
End Sub

Private Sub DrawOsz(Arr() As Stereo)
Dim X&, Y&, dx&, dy&, dy2&, dc0&, dc1&

    dx = Scope(0).ScaleWidth: dy = Scope(0).ScaleHeight
    dy2 = dy \ 2
    dc0 = Scope(0).hdc: dc1 = Scope(1).hdc
    
    Scope(0).ForeColor = Scope(0).BackColor: Scope(1).ForeColor = Scope(1).BackColor
    Rectangle dc0, 0, 0, dx, dy: Rectangle dc1, 0, 0, dx, dy
    
    Scope(0).ForeColor = 33023: Scope(1).ForeColor = 33023
    MoveToEx dc0, 0, dy2, 0: MoveToEx dc1, 0, dy2, 0
    MaxL = 0: MaxR = 0
    For X = 0 To UBound(Arr)
        If Abs(CLng(Arr(X).L)) > MaxL Then MaxL = Abs(CLng(Arr(X).L))
        If Abs(CLng(Arr(X).R)) > MaxR Then MaxR = Abs(CLng(Arr(X).R))
        If X Mod 15 = 0 Then
            LineTo dc0, X \ 4, Arr(X).L \ 4000 + dy2
            LineTo dc1, X \ 8, Arr(X).R \ 4000 + dy2
        End If
    Next
End Sub

Private Sub DrawVU(Arr() As Stereo)
Dim dx, dy, dy2, dx2, dc0 As Long
Dim dAmp As Long
Dim lAmpSum As Long
Dim lAmpCount As Long
Dim i As Integer

    dx = VUBox.ScaleWidth
    dx2 = VUBox.ScaleWidth \ 2
    dy = VUBox.ScaleHeight
    dy2 = VUBox.ScaleHeight
    
    dc0 = VUBox.hdc
    VUBox.Cls
    MoveToEx dc0, dx2, dy2, 0
    
    lAmpCount = UBound(Arr)
    VUBox.ForeColor = 33023
    For i = 0 To UBound(Arr)
        If i Mod 15 = 0 Then
            lAmpSum = lAmpSum + Arr(i).L \ 250
        End If
    Next i
    
    dAmp = lAmpSum / (1024 / 15)
    dAmp = Abs(dAmp)
    LineTo dc0, dAmp, 2

End Sub

Private Sub DrawFFT(Buf As WavBuf)
Dim i&, j&, X&, Y&, dx&, dy&, dc&, xlo&, xro&
Static Fl&, xScale#, xl#, xr#, BIH As BITMAPINFOHEADER
    xScale = 1 / 11111123456#
    
    FFTAudio Buf.Arr, Buf.FFT
    
    dx = FFTPanel.ScaleWidth: dy = FFTPanel.ScaleHeight
    dc = FFTPanel.hdc
      
    If Fl = 0 Then 'init
        ReDim pArr(1 To dx, 1 To dy)
        BIH.biSize = 40: BIH.biBitCount = 32: BIH.biPlanes = 1
        BIH.biWidth = dx: BIH.biHeight = dy
    Else
        For X = 1 To dx
            For Y = 1 To dy
                If pArr(X, Y).R <> 0 Then
                    If pArr(X, Y).R >= 20 Then pArr(X, Y).R = pArr(X, Y).R - 20
                    If pArr(X, Y).g >= 10 Then pArr(X, Y).g = pArr(X, Y).g - 10
                End If
            Next Y
        Next X
    End If
    
    xlo = dx / 2 - 6: xro = dx / 2 + 6
    X = 0: Y = 5
    For i = 1 To 23
      xl = 0: xr = 0
      For j = 1 To CLng(1.194 ^ i)
        With Buf.FFT(X)
          xl = xl + .L: xr = xr + .R
        End With
        X = X + 1
      Next j
      Y = Y + 5
    '    Rect pArr, xlo - Sqr(xl * xScale), y, xlo, y + 1
      RECT pArr, 5, dx - 6 - Y, 5 + Sqr(xr * xScale), dx - 4 - Y, dy - 3, "Vertical"
      RECT pArr, 5, Y, 5 + Sqr(xl * xScale), Y + 2, dy - 6, "Vertical"
    Next i
    '  RECT pArr, xlo + 3, 3, xlo + 5, MaxL / 32768 * (dy - 6), dy - 6
    '  Rect pArr, xro - 5, 3, xro - 3, MaxR / 32768 * (dy - 6), dy - 6
    
    StretchDIBits dc, 0, 0, dx, dy, 0, 0, dx, dy, pArr(1, 1), BIH, 0, vbSrcCopy
    Fl = 1
End Sub

Private Sub RECT(pArr() As BGRQuad, ByVal xs&, ByVal ys&, ByVal xe&, ByVal ye&, Optional Max&, Optional Spec$)
Dim X&, Y&, Mx&, MM&, D As Byte

    If xs < 1 Then xs = 1
    If ys < 1 Then ys = 1
    If xe > UBound(pArr, 1) Then xe = UBound(pArr, 1)
    If ye > UBound(pArr, 2) Then ye = UBound(pArr, 2)
    If Max And Spec = vbNullString Then
        Mx = Max * 0.85
        MM = Max * 0.5
        For X = xs To xe
            For Y = ys To ye
                If Y > Mx Then D = 0 Else If Y > MM Then D = 222 Else D = 128
                pArr(X, Y).R = 255
                pArr(X, Y).g = D
            Next Y
        Next X
    Else
        Mx = Max * 0.85
        MM = Max * ((0.4 - 0.35) * Rnd + 0.35)
        For Y = xs To xe
            For X = ys To ye
                If Y > Mx Then D = 0 Else If Y > MM Then D = 222 Else D = 128
                pArr(X, Y).R = 255
                pArr(X, Y).g = D
            Next X
        Next Y
    End If
End Sub



Private Sub List1_DblClick()
Dim i As Integer
Dim lRet As Long
    For i = 0 To List1.ListCount - 1 Step 1
        If List1.Selected(i) Then
            'change current track
            curTrack = i + 1
            If CheckIfPlaying = 1 Then
                mciSendString "play cd from " & CStr(curTrack), 0, 0, 0
            Else
                mciSendString "seek cd to " & CStr(curTrack), 0, 0, 0
            End If
       End If
    Next i
End Sub

Private Sub XT_Timer()
Dim i&
Dim sPos As String * 30
Dim sTrack As Integer
Dim sMin As Integer
Dim sSec As Integer
Dim lRet As Integer
Dim c1(2) As POINTAPI
Static lTicCount As Long

    lTicCount = lTicCount + 1
    If lTicCount > 10 Then
        If Num_Tracks <> GetNumTracks Then
            Num_Tracks = GetNumTracks
            SetTrackList Num_Tracks
        End If
        
        lRet = mciSendString("status cd position", sPos, Len(sPos), 0)
        If lRet = 0 Then
            sTrack = CInt(Mid$(sPos, 1, 2))
            sMin = CInt(Mid$(sPos, 4, 2))
            sSec = CInt(Mid$(sPos, 7, 2))
        End If
        
        lblstatus.Caption = Format(sTrack, "00") & " " & Format(sMin, "00") & ":" & Format(sSec, "00")
        
        lTicCount = 0
    
    End If

    Do While RBuf(RIdx).Hdr.dwFlags And WHDR_DONE
        ProcessBuf RBuf(RIdx)
        RIdx = CIdx(RIdx + 1)
    Loop
    
    DoEvents
    
    For i = RIdx To RIdx + UBound(RBuf)
        InitBuf RBuf(CIdx(i))
    Next i
    
End Sub

Private Sub InitBuf(Buf As WavBuf)
  If Buf.Hdr.dwUser Then Exit Sub
  With Buf.Hdr
    .dwUser = 1
    If .lpData = 0 Then .lpData = VarPtr(Buf.Arr(0))
    If .dwBufferLength = 0 Then .dwBufferLength = (UBound(Buf.Arr) + 1) * 4
    waveInPrepareHeader DevHandle, Buf.Hdr, Len(Buf.Hdr)
    waveInAddBuffer DevHandle, Buf.Hdr, Len(Buf.Hdr)
  End With
End Sub

Private Sub ProcessBuf(Buf As WavBuf)
Static lTicCount As Long

    waveInUnprepareHeader DevHandle, Buf.Hdr, Len(Buf.Hdr)
    Buf.Hdr.dwFlags = 0
    Buf.Hdr.dwUser = 0
    DrawOsz Buf.Arr
    DrawFFT Buf
    
'    lTicCount = lTicCount + 1
'    If lTicCount > 3 Then
        DrawVU Buf.Arr
        lTicCount = 0
'    End If
'    WriteFile Buf
End Sub

Private Function CIdx&(ByVal Idx&)
    CIdx = Idx Mod (UBound(RBuf) + 1)
End Function

Private Sub Form_Unload(Cancel As Integer)
Dim retvalue As Integer
Dim returnstring As Long
    xt.Enabled = False
    cmdStop_Click
    retvalue = mciSendString("close all", returnstring, 127, 0)
    mixerClose hMixer
    
    If DevHandle Then DoStop

    Set xt = Nothing
    Close
End Sub

Private Sub WriteFile(Buf As WavBuf)
Static L&(10)
    FSize = FSize + 4096
    If L(0) = 0 Then
        L(0) = 1179011410: L(2) = 1163280727: L(3) = 544501094: L(4) = 16: L(5) = 131073
        L(6) = 44100: L(7) = 176400: L(8) = 1048580: L(9) = 1635017060
    End If
    L(10) = FSize
    L(1) = FSize + 36
    Put FNr, FSize + 44 - 4095, Buf.Arr
    Put FNr, 1, L
End Sub

Private Function CheckIfPlaying() As Integer
Dim strTemp As String * 30

    CheckIfPlaying = 0
    mciSendString "status cd mode", strTemp, Len(strTemp), 0
    If Mid(strTemp, 1, 7) = "playing" Then CheckIfPlaying = 1
    
End Function
