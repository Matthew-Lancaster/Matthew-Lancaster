VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "vbVidCap"
   ClientHeight    =   5445
   ClientLeft      =   2850
   ClientTop       =   3405
   ClientWidth     =   6750
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadPal 
         Caption         =   "&Load Palette..."
      End
      Begin VB.Menu mnuSetCapFile 
         Caption         =   "&Set Capture File..."
      End
      Begin VB.Menu mnuAllocFileSpace 
         Caption         =   "&Allocate File Space"
      End
      Begin VB.Menu mnuspacer0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveFileAs 
         Caption         =   "Save &Captured Video As..."
      End
      Begin VB.Menu mnuSavePalette 
         Caption         =   "Save &Palette..."
      End
      Begin VB.Menu mnuSaveFrame 
         Caption         =   "Save Single &Frame..."
      End
      Begin VB.Menu mnuspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste Palette"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Pre&ferences..."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAudioFmt 
         Caption         =   "&Audio Format..."
      End
      Begin VB.Menu mnuspacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "&Format..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSource 
         Caption         =   "S&ource..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Display..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompression 
         Caption         =   "&Compression..."
      End
      Begin VB.Menu mnuspacer6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOverlay 
         Caption         =   "&Overlay"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDriver 
         Caption         =   "<none>"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "&Capture"
      Begin VB.Menu mnuCapFrame 
         Caption         =   "&Single Frame"
      End
      Begin VB.Menu mnuCapFrames 
         Caption         =   "&Frames..."
      End
      Begin VB.Menu mnuCapVid 
         Caption         =   "&Video..."
      End
      Begin VB.Menu mnuCapPal 
         Caption         =   "&Palette..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSysInfo 
         Caption         =   "&System Info..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   +{F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hCapWnd As Long       ' Handle to the Capture Windows
Private nDriverIndex As Long  ' video driver index (default 0)
Private m_CapParams As CAPTUREPARMS
'Public property to prevent reentrancy in Form_Resize event
Public AutoSizing As Boolean
'read-only public property to allow other forms to retrieve hwnd of Cap Window
Public Property Get capwnd() As Long
    capwnd = hCapWnd
End Property
'read-only properties for sizing
Public Property Get MenuHeight() As Long
    MenuHeight = GetSystemMetrics(SM_CYMENU)
End Property
Public Property Get CaptionHeight() As Long
    CaptionHeight = GetSystemMetrics(SM_CYCAPTION)
End Property
Public Property Get XBorder() As Long
    If Me.Appearance = 0 Then   'flat
        XBorder = GetSystemMetrics(SM_CXBORDER)
    Else                        '3D
        XBorder = GetSystemMetrics(SM_CXEDGE)
    End If
End Property
Public Property Get YBorder() As Long
    If Me.Appearance = 0 Then   'flat
        YBorder = GetSystemMetrics(SM_CYBORDER)
    Else                        '3D
        YBorder = GetSystemMetrics(SM_CYEDGE)
    End If
End Property


Private Sub Form_Load()
    Dim retVal As Boolean
    Dim numDevs As Long
    Dim left As Long, top As Long
'    Me.Show
    
    'load trivial settings first
    'Me.BackColor = Val(GetSetting(App.Title, "preferences", "backcolor", "&H404040")) 'default to dk gray
    On Error Resume Next
    left = (Screen.Width - Me.Width) / 2 'center window by default
    top = (Screen.Height - Me.Height) / 2
    On Error GoTo 0
    'left = Val(GetSetting(App.Title, "preferences", "left", left))
    'top = Val(GetSetting(App.Title, "preferences", "top", top))
    If left < 0 Then left = 0 'just make sure app isn't off the screen
    If top < 0 Then top = 0
    'If left > Screen.Width - Me.Width Then left = Screen.Width - Me.Width
    'If top > Screen.Height - Me.Height Then top = Screen.Height - Me.Height
    Me.left = left
    Me.top = top
    
    numDevs = VBEnumCapDrivers(Me)
    If 0 = numDevs Then
        'MsgBox "No capture hardware detected", vbCritical, App.Title
       End
       Exit Sub
    End If
'    nDriverIndex = Val(GetSetting(App.Title, "driver", "index", "0"))
'    if invalid entry is in registry use default (0)
'    If mnuDriver.UBound < nDriverIndex Then
        'Set This to Whatever is Best
        'If you can Enumerate them and Find the Seach Text String of one you Want
        
        'WELL YOU MIGHT THING SETTING THIS TO 1 FOR FIRST DRIVER CORRECT BUT IF ONLY 1 THEN SET 0 NAUGHT
        If numDevs = 1 Then nDriverIndex = 0 Else Stop
'    End If
   mnuDriver(nDriverIndex).Checked = True
    '//Create Capture Window
    'only get this if you want a description for the driver name
'    Call capGetDriverDescription(nDriverIndex, lpszName, 100, lpszVer, 100)    '// Retrieves driver info
    
    hCapWnd = capCreateCaptureWindow("VB CAP WINDOW", WS_CHILD Or WS_VISIBLE, 0, 0, 160, 120, Me.hWnd, 0)
    If 0 = hCapWnd Then
        'MsgBox "could not create capture window", vbCritical, App.Title
        Exit Sub
    End If
    
    'THIS TAKES A COUPLE OF SECONDS
    retVal = ConnectCapDriver(hCapWnd, nDriverIndex)
    If False = retVal Then
       ' MsgBox "could not connect to capture driver", vbInformation, App.Title
    Else
        #If USECALLBACKS = 1 Then
            ' if we have a valid capwnd we can enable our status callback function
            Call capSetCallbackOnStatus(hCapWnd, AddressOf StatusProc)
            Debug.Print "---Callback set on capture status---"
        #End If
    End If
        '// Set the video stream callback function
'    capSetCallbackOnVideoStream lwndC, AddressOf MyVideoStreamCallback
'    capSetCallbackOnFrame lwndC, AddressOf MyFrameCallback
 

End Sub


Public Sub Form_Resize()
     

    Exit Sub
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    'kludgy way to restrict min form size - better way is to subclass MINMAXINFO messages
    If Me.WindowState = vbMinimized Then Exit Sub 'runtime error was happening when user minimized app...
    If Me.ScaleWidth < 320 Then Me.Width = (320 + (Me.XBorder * 2)) * Screen.TwipsPerPixelX
    If Me.ScaleHeight < 240 Then Me.Height = (240 + (Me.YBorder * 2) + Me.MenuHeight + Me.CaptionHeight) * Screen.TwipsPerPixelY
    'Get the capture window attributes
    retVal = capGetStatus(hCapWnd, capStat)
        
    If retVal Then
        'center the capture window on the form
        Call SetWindowPos(hCapWnd, _
                    0&, _
                    (Me.ScaleWidth - capStat.uiImageWidth) / 2, _
                    (Me.ScaleHeight - capStat.uiImageHeight) / 2, _
                    0&, _
                    0&, _
                    SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOSENDCHANGING) 'by telling Windows not to send
                                                                    'WM_WINDOWPOSCHANGING messages we
                                                                    'eliminate the need for a reentrancy flag
    End If
      
End Sub

Private Sub Form_Unload(Cancel As Integer)

'save trivial settings
If Me.WindowState = vbDefault Then
    'Call SaveSetting(App.Title, "preferences", "left", Me.left)
    'Call SaveSetting(App.Title, "preferences", "top", Me.top)
End If

'unsubclass if necessary
#If USECALLBACKS = 1 Then
    ' Disable status callback
    Call capSetCallbackOnStatus(hCapWnd, 0&)
    Debug.Print "---Capture status callback released---"
#End If

'disconnect VFW driver
Call mVFW.capDriverDisconnect(hCapWnd)
'destroy CapWnd
If hCapWnd <> 0 Then Call DestroyWindow(hCapWnd)
'Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long 'C BOOL

End Sub


Private Sub mnuAbout_Click()
    Dim AboutWnd As frmAbout
    Set AboutWnd = New frmAbout
    
    AboutWnd.Show vbModal, Me
    
    Set AboutWnd = Nothing
End Sub

Private Sub mnuAllocFileSpace_Click()
    Dim AllocWnd As frmAlloc
    Set AllocWnd = New frmAlloc
    
    AllocWnd.Show vbModal, Me
    
    Set AllocWnd = Nothing

End Sub

Private Sub mnuAudioFmt_Click()
    Call SetAudioFormatDlg(Me.hWnd)
End Sub

Private Sub mnuCapFrame_Click()

    Call capGrabFrame(hCapWnd)

End Sub

Private Sub mnuCapFrames_Click()
    Dim FrameCapWnd As frmCapFrame
    
    Set FrameCapWnd = New frmCapFrame
    FrameCapWnd.Show vbModal, Me
    
    Set FrameCapWnd = Nothing
    
End Sub

Private Sub mnuCapPal_Click()
    Dim PalCapWnd As frmCapPal
    
    Set PalCapWnd = New frmCapPal
    PalCapWnd.Show vbModal, Me
    
    Set PalCapWnd = Nothing
End Sub

Private Sub mnuCapVid_Click()
    Dim retVal As Boolean
    Dim VidCapWnd As frmCapVid
    
    Set VidCapWnd = New frmCapVid
    VidCapWnd.Show vbModal, Me
    If VidCapWnd.Tag <> "" Then 'use tag to indicate whether user has pressed OK or not
'            // Capture video sequence
        retVal = capCaptureSequence(hCapWnd)
        Unload VidCapWnd 'reclaim mem
    End If
    Set VidCapWnd = Nothing
End Sub

Private Sub mnuCompression_Click()

    Call capDlgVideoCompression(hCapWnd)

End Sub

Private Sub mnuCopy_Click()
    
    Call capEditCopy(hCapWnd)

End Sub

Private Sub mnuDisplay_Click()

    Call capDlgVideoDisplay(hCapWnd)
    
End Sub

Private Sub mnuDriver_Click(Index As Integer)
    Dim retVal As Boolean
    
    retVal = ConnectCapDriver(hCapWnd, Index)
    If False = retVal Then
        'MsgBox "could not connect to capture driver", vbInformation, App.Title
    Else
        Call SaveSetting(App.Title, "driver", "index", CStr(Index)) 'save selected device index
    End If
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub mnuFormat_Click()

    Call capDlgVideoFormat(hCapWnd)
    Call ResizeCaptureWindow(hCapWnd)

End Sub

Private Sub mnuLoadPal_Click()
Dim PalFile As String
Dim PalFileTitle As String
Dim retVal As Boolean

retVal = VBGetOpenFileName(PalFile, _
                            PalFileTitle, _
                            Filter:="Palette Files (*.pal)|*.pal", _
                            InitDir:=App.Path, _
                            DlgTitle:="Load Palette", _
                            DefaultExt:="Pal", _
                            HideReadOnly:=True, _
                            Owner:=Me.hWnd)
If True = retVal Then 'user did not cancel
    retVal = capPaletteOpen(hCapWnd, PalFile)
    If 0 = retVal Then
        'MsgBox "Could not load palette file: " & PalFileTitle, vbInformation, App.Title
    End If
End If
        

End Sub

Private Sub mnuOverlay_Click()
    
    mnuOverlay.Checked = Not (mnuOverlay.Checked)
    Call capOverlay(hCapWnd, mnuOverlay.Checked)
    
End Sub

Private Sub mnuPreferences_Click()
    Dim PrefsWnd As frmPrefs
    
    Set PrefsWnd = New frmPrefs
    PrefsWnd.Show vbModal, Me
    
    Set PrefsWnd = Nothing
End Sub

Private Sub mnuPreview_Click()

    mnuPreview.Checked = Not (mnuPreview.Checked)
    Call capPreview(hCapWnd, mnuPreview.Checked)

End Sub


Private Sub mnuSaveFileAs_Click()
Dim Filename As String
Dim retVal As Boolean

retVal = VBGetSaveFileNamePreview(Filename, _
                            FileMustExist:=False, _
                            HideReadOnly:=True, _
                            Filter:="AVI Files (*.avi)|*.avi", _
                            DefaultExt:="avi", _
                            Owner:=Me.hWnd)
If False <> retVal Then
    retVal = capFileSaveAs(hCapWnd, Filename)
    If True <> retVal Then
        'MsgBox "Problems saving capture file", vbInformation, App.Title
    End If
End If
End Sub

Private Sub mnuSaveFrame_Click()
Dim Filename As String
Dim retVal As Boolean

'retVal = VBGetSaveFileName(FileName, _
'                            filter:="DIB Bitmap Files (*.bmp)|*.bmp", _
'                            DlgTitle:="Save Single Frame", _
'                            DefaultExt:="bmp", _
'                            Owner:=Me.hWnd)
retVal = VBGetSaveFileName(Filename, _
                            Filter:="DIB Bitmap Files (*.jpg)|*.jpg", _
                            DlgTitle:="Save Single Frame", _
                            DefaultExt:="jpg", _
                            Owner:=Me.hWnd)
If False <> retVal Then
    retVal = capFileSaveDIB(hCapWnd, Filename)
    If True <> retVal Then
        'MsgBox "Problem saving frame", vbInformation, App.Title
    End If
End If
End Sub

Private Sub mnuSavePalette_Click()
Dim Filename As String
Dim retVal As Boolean

retVal = VBGetSaveFileName(Filename, _
                            Filter:="Palette Files (*.pal)|*.pal", _
                            DlgTitle:="Save Palette", _
                            DefaultExt:="pal", _
                            Owner:=Me.hWnd)
If False <> retVal Then
    retVal = capPaletteSave(hCapWnd, Filename)
    If True <> retVal Then
        'MsgBox "Problem saving palette", vbInformation, App.Title
    End If
End If
End Sub

Private Sub mnuSetCapFile_Click()
Dim CapFile As String
Dim CapFileTitle As String
Dim CapFileDir As String
Dim retVal As Boolean
Dim nfileLen As Long

CapFile = mVFW.capFileGetCaptureFile(hCapWnd)
CapFileTitle = VBGetFileTitle(CapFile)
CapFileDir = left$(CapFile, Len(CapFile) - Len(CapFileTitle))
retVal = VBGetOpenFileNamePreview(CapFile, _
                            FileTitle:=CapFileTitle, _
                            Filter:="AVI Files (*.avi)|*.avi", _
                            InitDir:=CapFileDir, _
                            DlgTitle:="Set Capture File", _
                            FileMustExist:=False, _
                            HideReadOnly:=True, _
                            DefaultExt:="avi", _
                            Owner:=Me.hWnd)
If True = retVal Then 'user did not cancel
    retVal = mVFW.capFileSetCaptureFile(hCapWnd, CapFile)
    If 0 = retVal Then
        'MsgBox "Could not set new capture file: " & CapFileTitle, vbInformation, App.Title
        Exit Sub
    Else
        'capture file was changed successfully let's allocate some disk space for it
        'but only if it doesn't already exist
        On Error Resume Next
        nfileLen = FileLen(CapFile)
        If Err.Number = 53 Then 'file does not exist
            Call mnuAllocFileSpace_Click
        End If
    End If
End If
End Sub

Private Sub mnuSource_Click()
'   /*
'    * Display the Video Source dialog when "Source" is selected from the
'    * menu bar.
'    */
    
    Call capDlgVideoSource(hCapWnd)

End Sub



Private Sub mnuSysInfo_Click()
Call ShellAbout(Me.hWnd, _
                App.Title & " System Info Window#OS Information:", _
                vbCrLf & _
                "vbVidCap.exe is Copyright(C) 1998-2000", _
                Me.Icon)
End Sub
