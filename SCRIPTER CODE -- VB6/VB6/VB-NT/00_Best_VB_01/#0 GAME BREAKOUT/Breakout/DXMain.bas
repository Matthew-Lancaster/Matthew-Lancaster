Attribute VB_Name = "DXMain"
'---------------------------------------------------------------------------------------
' Module    : DXMain
' DateTime  : 4/2/2005 12:27
' Author    : Jason
' Purpose   : The main DX setup form
'---------------------------------------------------------------------------------------

Option Explicit

Public dx As DirectX7
Public dd As DirectDraw7
Public PixFormat As DDPIXELFORMAT

Private Primary As DirectDrawSurface7
Public Backbuffer As DirectDrawSurface7


Public bRECT As RECT             'Backbuffer RECT


Private ddsd1 As DDSURFACEDESC2  'Primary surface description
Public ddsd4 As DDSURFACEDESC2   'Backbuffer surface description

Public bRunning As Boolean       'For the main loop
Public bRestore As Boolean       'For restoring surfaces
Public bInit As Boolean          'Has DirectDraw been initialized?

Public bkColor As Long           'The fill color for the Backbuffer
Public lMagenta As Long          'The color key (magenta)

Public DX_WIDTH As Long          'DX screen width
Public DX_HEIGHT As Long         'DX screen height
Public DX_BPP As Long            'DX bits per pixel

Public Sub InitDX()

    On Error GoTo errOut
    
    Set dx = New DirectX7
    Set dd = dx.DirectDrawCreate("")
    
    'indicate that we dont need to change display depth
    Call dd.SetCooperativeLevel(frmDX.hwnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    
    'Set the display mode
    dd.SetDisplayMode DX_WIDTH, DX_HEIGHT, DX_BPP, 0, DDSDM_DEFAULT
    
    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd1.lBackBufferCount = 1
    
    'Create the primary surface
    Set Primary = dd.CreateSurface(ddsd1)
    
    'Attach a Backbuffer to the primary surface
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set Backbuffer = Primary.GetAttachedSurface(caps)
    
    'Get the backbuffer surface description
    Backbuffer.GetSurfaceDesc ddsd4
    
    'We create a DrawableSurface class from our backbuffer
    'that makes it easy to draw text
    Backbuffer.SetForeColor vbGreen
    Backbuffer.SetFontTransparency True
    
    'For getting the 16-bit long color value
    GetColorShiftValues Primary '<- For getting the 16-bit RGB long
    
    'Get the pixel format of the primary surface
    Primary.GetPixelFormat PixFormat
    
    'Set the lMagenta color key
    lMagenta = PixFormat.lRBitMask + PixFormat.lBBitMask
    
    'Set the bkColor to orange
    bkColor = 0 'DDRGB(200, 100, 0)
    
       'Set up a Backbuffer RECT
    bRECT.Bottom = DX_HEIGHT
    bRECT.Right = DX_WIDTH
    
    'Initialization is done
    bInit = True
     
    Exit Sub
errOut:
    Debug.Print "InitDX ; " & Err.Number & " ; " & Err.Description
    bInit = False
    
    'If there's an error...
    EndIt
    
End Sub

Public Sub EndIt()

    'Turn off the main loop
    bRunning = False
    
    'Restore the prior display mode
    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(frmDX.hwnd, DDSCL_NORMAL)
    
    'Destroy objects and surfaces
    Set dd = Nothing
    Set dx = Nothing
    Set Primary = Nothing
    Set Backbuffer = Nothing
    KillSurfaces
    StopBeatz
    
    'Unload forms
    Unload frmDX
    Unload Form1
    
    End
    
End Sub

Public Sub Render(ByVal X As Long, ByVal Y As Long, _
   ByVal Surface As DirectDrawSurface7, srcRECT As RECT, _
   ByVal Transparent As Boolean, Optional Stretch As Boolean = False)
   
   On Error GoTo errOut
   
   If Stretch Then
   
      Dim temp As RECT
      
      temp.Right = DX_WIDTH
      temp.Bottom = DX_HEIGHT
      
      If Transparent Then
         Backbuffer.Blt temp, Surface, srcRECT, DDBLT_KEYSRC Or DDBLT_WAIT
      Else
         Backbuffer.Blt temp, Surface, srcRECT, DDBLT_WAIT
      End If
      
      Exit Sub
      
   End If
      
   
   'If DirectDraw is not initialized then skip this procedure
   If bInit = False Then Exit Sub
   
   If Transparent = True Then
      Backbuffer.BltFast X, Y, Surface, srcRECT, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
   Else
      Backbuffer.BltFast X, Y, Surface, srcRECT, DDBLTFAST_WAIT
   End If
   
   Exit Sub
errOut:
   
   Debug.Print "Render ; " & Err.Number & " ; " & Err.Description

End Sub

Public Sub Flip()

    'flip the backbuffer to the screen
    Primary.Flip Nothing, DDFLIP_WAIT
    
End Sub

Public Sub FillBackGround()

   'Fill the Backbuffer with the orange color
   Call Backbuffer.BltColorFill(bRECT, bkColor)
   
End Sub

Public Sub CheckForLostSurfaces()

    'this will keep us from trying to blt in case we lose the surfaces (alt-tab)
    bRestore = False
    Do Until ExModeActive
        myDoEvents 0
        bRestore = True
    Loop

    ' if we lost and got back the surfaces, then restore them
    If bRestore Then
        bRestore = False
        dd.RestoreAllSurfaces
        InitSurfaces
    End If

End Sub

Private Function ExModeActive() As Boolean

    Dim TestCoopRes As Long
    
    TestCoopRes = dd.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If
    
End Function




