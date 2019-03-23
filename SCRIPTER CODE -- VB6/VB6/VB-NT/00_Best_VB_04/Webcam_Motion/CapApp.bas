Attribute VB_Name = "mCapApp"
'****************************************************************
'*  VB file:   CapApp.bas...
'*
'*  created:        1998 by Ray Mercer
'*  last modified:  12/2/98 by Ray Mercer (added comments)
'*
'*  Useful routines for creating a video capture application in
'*  Visual Basic.  Loosely based on routines found in the Microsoft
'*  VidCap32 application in the C-Language VFW Developer's kit
'*
'*
'*  Copyright (c) 1998 Ray Mercer.  All rights reserved.
'****************************************************************


Option Explicit

'application specific routines are here

Public Const ONE_MEGABYTE As Long = 1048576
Public Const MMSYSERR_NOERROR As Long = 0
Public Const INDEX_15_MINUTES As Long = 27000 '(30fps * 60sec * 15min)
Public Const INDEX_3_HOURS As Long = 324000 ' (30fps * 60sec * 60min * 3hr)


Public Function GetFreeSpace() As Long
'this function gets the amount of free disk space and adds the size
'of the current capture file
Dim capfilesize As Long
Dim Path As String

'get Cap File length
Path = capFileGetCaptureFile(frmMain.capwnd)
If Path <> "" Then
    On Error Resume Next 'if user has deleted file this is necessary
    capfilesize = FileLen(Path)
    capfilesize = capfilesize / ONE_MEGABYTE
End If

'now get free disk space from that drive
Path = left$(Path, 3)
GetFreeSpace = CLng(vbGetAvailableMBytesAsString(Path)) - capfilesize  'Use getfree.bas to handle large drives

End Function

Sub ResizeCaptureWindow(ByVal hCapWnd As Long)
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    
    
    
    'Get the capture window attributes
    'retVal = capGetStatus(hCapWnd, capStat)
        
    Exit Sub
    If retVal Then
        'Resize the main form to fit
        Call SetWindowPos(frmMain.hWnd, _
                    0&, _
                    0&, _
                    0&, _
                    capStat.uiImageWidth + (frmMain.XBorder * 2), _
                    capStat.uiImageHeight + (frmMain.YBorder * 4) _
                    + frmMain.CaptionHeight + frmMain.MenuHeight, _
                    SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
        'Resize the capture window to format size
        Call SetWindowPos(hCapWnd, _
                    0&, _
                    0&, _
                    0&, _
                    capStat.uiImageWidth, _
                    capStat.uiImageHeight, _
                    SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING)
    End If
    Call frmMain.Form_Resize
End Sub

Public Function VBEnumCapDrivers(ByRef frm As frmMain) As Long
'/*
' * Enumerate the potential capture drivers and add the list to the Options
' * menu.  This function is only called once at startup.
' * Returns 0 if no drivers are available.
' */
Const MAXVIDDRIVERS As Long = 9
Const CAP_STRING_MAX As Long = 128
Dim numDrivers As Long
Dim driverStrings(0 To MAXVIDDRIVERS - 1) As String
Dim Index As Long
Dim Device As String
Dim Version As String
Dim menu As VB.menu
 
'Debug.Print frm.Name
 
Device = String$(CAP_STRING_MAX, 0)
Version = String$(CAP_STRING_MAX, 0)
numDrivers = 0
For Index = 0 To (MAXVIDDRIVERS - 1) Step 1
    If 0 <> capGetDriverDescription(Index, _
                                    Device, _
                                    CAP_STRING_MAX, _
                                    Version, _
                                    CAP_STRING_MAX) _
                                                            Then
        'extend the menu
        If Index > 0 Then
            Load frm.mnuDriver(Index)
        End If
        Set menu = frm.mnuDriver(Index) 'get an object pointer to the new menu
        'Concatenate the device name and version strings to the new menu item
        menu.Caption = left$(Device, InStr(Device, vbNullChar) - 1)
        menu.Caption = menu.Caption & " "
        menu.Caption = menu.Caption & left$(Version, InStr(Version, vbNullChar) - 1)
        menu.Enabled = True
        numDrivers = numDrivers + 1
        Exit For
    End If

Next
VBEnumCapDrivers = numDrivers
End Function

Public Function ConnectCapDriver(ByVal hCapWnd As Long, ByVal nDriverIndex As Long) As Boolean
   Dim retVal As Boolean
   Dim Caps As CAPDRIVERCAPS
   Dim i As Long
   
   Debug.Assert (nDriverIndex < 10) And (nDriverIndex >= 0)
   '// Connect the capture window to the driver
    retVal = capDriverConnect(hCapWnd, nDriverIndex)
    If False = retVal Then
       'return False
       Exit Function
    End If
    '// Get the capabilities of the capture driver
    retVal = capDriverGetCaps(hCapWnd, Caps)
    
    If False <> retVal Then
        'reset menus (very app-specific)
        With frmMain
            For i = 0 To .mnuDriver.UBound
                .mnuDriver(i).Checked = False 'make sure all drivers are unchecked
            Next
            .mnuDriver(nDriverIndex).Checked = True 'then check the new driver
            'disable all hardware feature menu items
            .mnuSource.Enabled = False
            .mnuFormat.Enabled = False
            .mnuDisplay.Enabled = False
            .mnuOverlay.Enabled = False
            'Then enable the ones which are supported by the new driver
            If Caps.fHasDlgVideoSource <> 0 Then .mnuSource.Enabled = True
            If Caps.fHasDlgVideoFormat <> 0 Then .mnuFormat.Enabled = True
            If Caps.fHasDlgVideoDisplay <> 0 Then .mnuDisplay.Enabled = True
            If Caps.fHasOverlay <> 0 Then .mnuOverlay.Enabled = True
            
        End With
    End If
    '// Set the preview rate in milliseconds
    Call capPreviewRate(hCapWnd, 66) '15 FPS
    
    '// Start previewing the image from the camera
    Call capPreview(hCapWnd, True)
    'default to showing a preview each time
    frmMain.mnuPreview.Checked = True
        
    '// Resize the capture window to show the whole image
    Call ResizeCaptureWindow(hCapWnd)
    ConnectCapDriver = True
End Function
Public Function StatusProc(ByVal hCapWnd As Long, ByVal StatusCode As Long, ByVal lpStatusString As Long) As Long
        Select Case StatusCode
            Case 0 'this is recommended in docs
                'when zero is sent, clear old status messages
                'frmMain.Caption = App.Title
            Case IDS_CAP_END ' Video Capture has finished
                frmMain.Caption = App.Title
            Case IDS_CAP_STAT_VIDEOAUDIO, IDS_CAP_STAT_VIDEOONLY
                'MsgBox LPSTRtoVBString(lpStatusString), vbInformation, App.Title
            Case Else
                'use this function if you need a real VB string
                'frmMain.Caption = LPSTRtoVBString(lpStatusString)
                
                'or, just pass the LPCSTR to a WINAPI function
                Call SetWindowTextAsLong(frmMain.hWnd, lpStatusString)
        End Select
    Debug.Print "Driver returned code " & StatusCode & " to StatusProc"
    StatusProc = -(True) '- converts Boolean to C BOOL
End Function

'****************************************************************
'* FUNCTION LPSTRtoVBString()
'* ===============
'* by Ray Mercer
'* generic function to convert an LPCSTR to a VB String (BSTR)
'*
'* INPUTS:
'* LPSTR - a C language LPCSTR (returned from an API)
'* maxLen - optional parameter with a default value of 256
'*          defines the maximum possible length of the string
'*          pointed to by LPSTR
'*
'* RETURNS:
'* a VBString containing the string pointed to by LPSTR
'*  (works on DBCS systems too)
'****************************************************************
Private Function LPSTRtoVBString(ByVal LPSTR As Long, Optional ByVal maxlen As Long = 256) As String
    Dim sBuff  As String
    
    If LPSTR <> 0 Then 'quick and dirty input validation
        sBuff = String$(maxlen, 0) 'MCI_MAX
        If 0 <> lstrcpy(StrPtr(sBuff), LPSTR) Then 'copy mem directly
            LPSTRtoVBString = left$(sBuff, InStr(sBuff, vbNullChar) - 1) 'trim at NULL
            LPSTRtoVBString = StrConv(LPSTRtoVBString, vbUnicode) 'Convert to Unicode
        End If
    End If

End Function


