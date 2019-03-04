VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonDialogCopyFileClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ************************************************************************************************************
' * Custom CommonDialogCopyFileClipboard Object with use of API
' * Replacement for vb Common Dialog Object
' ************************************************************************************************************

' ************************************************************************************************************
' * Copyright ©2016, Frank Donckers
' * All source code in this file is licensed under a modified BSD license.
' * This means you may use the code in your own projects IF you provide attribution.
' ************************************************************************************************************

' ************************************************************************************************************
' * This Class can be used in order to avoid loading the vb CommonDialogCopyFileClipboard component
' ************************************************************************************************************

' API:
' ----
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal aint As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer

' Public constant variables:
' --------------------------
'Public Const WM_CLOSE = &H10
Private Const GW_HWNDNEXT = 2
Private Const WM_GETTEXTLENGTH = &HE
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDPREV = 3

' Public variabeles:
' ------------------
Private lngApphWnd       As Long

' API:
' ----
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As Long, _
ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' Common Dialog
Private Declare Function SHInvokePrinterCommand Lib "shell32.dll" Alias "SHInvokePrinterCommandA" (ByVal hWnd As Long, ByVal uAction As Long, ByVal lBuf1 As String, ByVal lBuf2 As String, ByVal fModal As Boolean) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

' ShowFont function
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Const LF_FACESIZE = 32
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

' Adapted Types:
' --------------
Private Type DevMode
        dmDeviceName        As String * 32
        dmSpecVersion       As Integer
        dmDriverVersion     As Integer
        dmSize              As Integer
        dmDriverExtra       As Integer
        dmFields            As Long
        dmOrientation       As Integer
        dmPaperSize         As Integer
        dmPaperLength       As Integer
        dmPaperWidth        As Integer
        dmScale             As Integer
        dmCopies            As Integer
        dmDefaultSource     As Integer
        dmPrintQuality      As Integer
        dmColor             As Integer
        dmDuplex            As Integer
        dmYResolution       As Integer
        dmTTOption          As Integer
        dmCollate           As Integer
        dmFormName          As String * 32
        dmUnusedPadding     As Integer
        dmBitsPerPel        As Integer
        dmPelsWidth         As Long
        dmPelsHeight        As Long
        dmDisplayFlags      As Long
        dmDisplayFrequency  As Long
End Type
'Dialog open file
Private Type OPENFILENAME
        lStructSize         As Long
        hWndOwner           As Long
        hInstance           As Long
        lpstrFilter         As String
        lpstrCustomFilter   As String
        nMaxCustFilter      As Long
        nFilterIndex        As Long
        lpstrFile           As String
        nMaxFile            As Long
        lpstrFileTitle      As String
        nMaxFileTitle       As Long
        lpstrInitialDir     As String
        lpstrTitle          As String
        Flags               As Long
        nFileOffset         As Integer
        nFileExtension      As Integer
        lpstrDefExt         As String
        lCustData           As Long
        lpfnHook            As Long
        lpTemplateName      As String
End Type

'Dialog font
Private Type CHOOSEFONT
        lStructSize         As Long
        hWndOwner           As Long
        HDC                 As Long
        lpLogFont           As Long
        iPointSize          As Long
        Flags               As Long
        rgbColors           As Long
        lCustData           As Long
        lpfnHook            As Long
        lpTemplateName      As String
        hInstance           As Long
        lpszStyle           As String
        nFontType           As Integer
        MISSING_ALIGNMENT   As Integer
        nSizeMin            As Long
        nSizeMax            As Long
End Type

'Dialog ColorPicker
Private Type ChooseColor
        lStructSize         As Long
        hWndOwner           As Long
        hInstance           As Long
        rgbResult           As Long
        lpCustColors        As Long
        Flags               As Long
        lCustData           As Long
        lpfnHook            As Long
        lpTemplateName      As String
End Type

'Dialog Printer select
Private Type PRINTDLG_TYPE
        lStructSize         As Long
        hWndOwner           As Long
        hDevMode            As Long
        hDevNames           As Long
        HDC                 As Long
        Flags               As Long
        nFromPage           As Integer
        nToPage             As Integer
        nMinPage            As Integer
        nMaxPage            As Integer
        nCopies             As Integer
        hInstance           As Long
        lCustData           As Long
        lpfnPrintHook       As Long
        lpfnSetupHook       As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate      As Long
        hSetupTemplate      As Long
End Type

Private Type DEVNAMES_TYPE
        wDriverOffset       As Integer
        wDeviceOffset       As Integer
        wOutputOffset       As Integer
        wDefault            As Integer
        extra               As String * 100
End Type


'Point
Private Type POINTAPI
        X   As Long
        Y   As Long
End Type

'Recangle
Private Type RECT
        Left    As Long
        Top     As Long
        Right   As Long
        Bottom  As Long
End Type

'Dialog page setup
Private Type PageSetupDlg
        lStructSize             As Long
        hWndOwner               As Long
        hDevMode                As Long
        hDevNames               As Long
        Flags                   As Long
        ptPaperSize             As POINTAPI
        rtMinMargin             As RECT
        rtMargin                As RECT
        hInstance               As Long
        lCustData               As Long
        lpfnPageSetupHook       As Long
        lpfnPagePaintHook       As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate      As Long
End Type

'Dialog font
Private Type LOGFONT
        lfHeight            As Long
        lfWidth             As Long
        lfEscapement        As Long
        lfOrientation       As Long
        lfWeight            As Long
        lfItalic            As Byte
        lfUnderline         As Byte
        lfStrikeOut         As Byte
        lfCharSet           As Byte
        lfOutPrecision      As Byte
        lfClipPrecision     As Byte
        lfQuality           As Byte
        lfPitchAndFamily    As Byte
        lfFaceName          As String * 31
End Type


' Constants for the common dialog
Private Const OFN_ALLOWMULTISELECT = &H200  'Allow multi select (Open Dialog)
Private Const OFN_EXPLORER = &H80000        'Set windows style explorer
Private Const OFN_FILEMUSTEXIST = &H1000    'File must exist
Private Const OFN_HIDEREADONLY = &H4        'Hide read-only check box (Open Dialog)
Private Const OFN_OVERWRITEPROMPT = &H2     'Promt beafore overwritning file (Save Dialog)
Private Const OFN_PATHMUSTEXIST = &H800     'Path must exist
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const DEFAULT_CHARSET = 1
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_QUALITY = 0
Private Const FW_BOLD = 700
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.
Private Const FW_NORMAL = 400
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const REGULAR_FONTTYPE = &H400
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&

' Constants for the GlobalAllocate
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_OUT_BUFFER = 2
Private Const DM_IN_BUFFER = 8

Private Const MAX_PATH = 260 'Constant for maximum path

Public cFileName As Collection   'Filename collection
Public cFileTitle As Collection  'Filetitle collection

' Property Variables:
Dim m_CancelError       As Boolean
Dim m_FileName          As String
Dim m_DialogTitle       As String
Dim m_InitialDir        As String
Dim m_Filter            As String
Dim m_FilterIndex       As Integer
Dim m_MultiSelect       As Boolean
Dim m_FontName          As String
Dim m_FontSize          As Integer
Dim m_FontColor         As Long
Dim m_FontBold          As Boolean
Dim m_FontItalic        As Boolean
Dim m_FontUnderline     As Boolean
Dim m_FontStrikeThru    As Boolean
Dim m_DialogOpened      As Boolean
Dim m_Canceled          As Boolean
'PageSetup
Dim m_PageLeftMargin    As Double
Dim m_PageRightMargin   As Double
Dim m_PageTopMargin     As Double
Dim m_PageBottomMargin  As Double
Dim m_PaperSizeWidth    As Double
Dim m_PapersizeHeight   As Double
Dim m_PaperSize         As String
Dim m_PaperOrientation  As String

'Public Methods
'--------------

'Open file dialog wrapper
Public Function ShowOpen(hWndOwner As Long, Optional NewFilter As String = "") As String
    Dim epOFN   As OPENFILENAME
    Dim lngRet  As Long

    m_Canceled = False
    ShowOpen = ""

    With epOFN

        If MultiSelect Then
            .Flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
            .lpstrFile = DefaultFileName & Space(9999 - Len(DefaultFileName)) & vbNullChar
            .lpstrFileTitle = Space(9999) & vbNullChar
        Else
            .Flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
            .lpstrFile = DefaultFileName & String(MAX_PATH - Len(DefaultFileName), 0) & vbNullChar
            .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
        End If

        .hWndOwner = hWndOwner 'Handle to window
        If NewFilter = "" Then
            .lpstrFilter = SetFilter(Filter) & vbNullChar 'File filter
        Else
            .lpstrFilter = SetFilter(NewFilter) & vbNullChar 'File filter
        End If
        .lpstrInitialDir = InitialDir & vbNullChar 'Initial directory
        .lpstrTitle = DialogTitle & vbNullChar 'Dialog title
        .lStructSize = Len(epOFN) 'Structure size in bytes
        .nFilterIndex = FilterIndex 'Filter index
        .nMaxFile = Len(.lpstrFile) 'Maximum file length
        .nMaxFileTitle = Len(.lpstrFileTitle) 'Maximum file title length
    End With

    lngRet = GetOpenFileName(epOFN) 'Call open dialog

    If lngRet <> 0 Then 'If there are no errors continue with opening file
        ParseFileName epOFN.lpstrFile
        ShowOpen = cFileName(1)
    Else
        m_Canceled = True
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
End Function


'Save File Dialog wrapper
Public Function ShowSave(hWndOwner As Long, Optional NewFilter As String = "")
    Dim epOFN   As OPENFILENAME
    Dim lngRet  As Long

    m_Canceled = False
    With epOFN
        .hWndOwner = hWndOwner 'Handle to window
        .Flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .lpstrFile = DefaultFileName & String(MAX_PATH - Len(DefaultFileName), 0) & vbNullChar
        .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
        If NewFilter = "" Then
            .lpstrFilter = SetFilter(Filter) & vbNullChar 'File filter
        Else
            .lpstrFilter = SetFilter(NewFilter) & vbNullChar 'File filter
        End If
        .lpstrInitialDir = InitialDir & vbNullChar 'Initial directory
        .lpstrTitle = DialogTitle & vbNullChar 'Dialog title
        .lStructSize = Len(epOFN) 'Structure size in bytes
        .nFilterIndex = FilterIndex 'Filter index
        .nMaxFile = Len(.lpstrFile) 'Maximum file length
        .nMaxFileTitle = Len(.lpstrFileTitle) 'Maximum file title length
    End With

    lngRet = GetSaveFileName(epOFN) 'Call save dialog

    If lngRet <> 0 Then 'If there are no errors continue with saving file
        ParseFileName epOFN.lpstrFile
    Else
        m_Canceled = True
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If

End Function

'Font dialog wrapper
Public Function ShowFont(hWndOwner As Long)
    Dim CF          As CHOOSEFONT
    Dim LF          As LOGFONT
    Dim lMemHandle  As Long
    Dim lLogFont    As Long
    Dim lngRet  As Long

    m_Canceled = False
    With LF
        .lfCharSet = DEFAULT_CHARSET 'Default character set
        .lfClipPrecision = CLIP_DEFAULT_PRECIS 'Clipping precision
        .lfFaceName = "Arial" & vbNullChar 'Font name
        .lfHeight = 13 'Height
        .lfOutPrecision = OUT_DEFAULT_PRECIS 'Precision mapping
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN 'Default pitch
        .lfQuality = DEFAULT_QUALITY 'Default quality
        .lfWeight = FW_NORMAL 'Regular font type
    End With

    ' Create the memory block
    lMemHandle = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(LF))
    lLogFont = GlobalLock(lMemHandle)
    CopyMemory ByVal lLogFont, LF, Len(LF)

    With CF
        .Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
        .HDC = Printer.HDC 'Device context of default printer
        .hWndOwner = hWndOwner 'Handle to window
        .iPointSize = 120 'Set font size to 12 size
        .lpLogFont = lLogFont 'Log font
        .lStructSize = Len(CF) 'Size of structure in bytes
        .nFontType = REGULAR_FONTTYPE 'Regular font type
        .nSizeMax = 72 'Maximum font size
        .nSizeMin = 10 'Minimum font size
        .rgbColors = RGB(0, 0, 0) 'Font color
    End With

    lngRet = CHOOSEFONT(CF) 'Call font dialog
    If lngRet <> 0 Then 'If there are no errors continue with font
        CopyMemory LF, ByVal lLogFont, Len(LF)

        FontName = Left(LF.lfFaceName, InStr(LF.lfFaceName, vbNullChar) - 1)
        FontSize = CF.iPointSize / 10
        FontColor = CF.rgbColors
        If LF.lfWeight = FW_NORMAL Then
            FontBold = False
            FontItalic = False
            FontUnderline = False
            FontStrikeThru = False
        Else
            If LF.lfWeight = FW_BOLD Then FontBold = True
            If LF.lfItalic <> 0 Then FontItalic = True
            If LF.lfUnderline <> 0 Then FontUnderline = True
            If LF.lfStrikeOut <> 0 Then FontStrikeThru = True
        End If
    Else
        m_Canceled = True
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If

    ' Unlock and free the memory block
    ' Note this must be done
    GlobalUnlock lMemHandle
    GlobalFree lMemHandle
End Function

'Colorpicker dialog wrapper
Public Function ShowColor(hWndOwner As Long)
    Dim epCC            As ChooseColor
    Dim lngRet          As Long
    Dim CusCol(0 To 16) As Long
    Dim i               As Integer

    m_Canceled = False
    ' Fills custom colors with white
    For i = 0 To 15
        CusCol(i) = vbWhite
    Next
    With epCC
        .hWndOwner = hWndOwner 'Handle to window
        .lStructSize = Len(epCC) 'Structure size in bytes
        .lpCustColors = VarPtr(CusCol(0)) 'Custom colors
        .rgbResult = 0 'RGB result
    End With
    lngRet = ChooseColor(epCC) 'Call color dialog
    If lngRet <> 0 Then 'If there are no errors continue with color
        ShowColor = epCC.rgbResult
    Else
        m_Canceled = True
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
End Function

'Page setup dialog wrapper
Public Function ShowPageSetup(hWndOwner As Long)
    Dim epPSD   As PageSetupDlg
    Dim lngRet  As Long

    m_Canceled = False
    epPSD.lStructSize = Len(epPSD) 'Structure size in bytes
    epPSD.hWndOwner = hWndOwner 'Handle to window
    lngRet = PageSetupDlg(epPSD) 'Call page setup dialog
    If lngRet <> 0 Then 'If there are no errors continue
        PageLeftMargin = CDbl(epPSD.rtMargin.Left / 100)
        PageRightMargin = CDbl(epPSD.rtMargin.Right / 100)
        PageTopMargin = CDbl(epPSD.rtMargin.Top / 100)
        PageBottomMargin = CDbl(epPSD.rtMargin.Bottom / 100)
        PaperSizeWidth = CDbl(epPSD.ptPaperSize.X / 100)
        PaperSizeHeight = CDbl(epPSD.ptPaperSize.Y / 100)
        If PaperSizeHeight > PaperSizeWidth Then
            PaperOrientation = "Portrait"
        Else
            PaperOrientation = "Landscape"
        End If
        PaperSize = GetPaperSizeName(PaperSizeWidth, PaperSizeHeight)
    Else
        m_Canceled = True
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
End Function

'Printer select dialog wrapper
Public Function ShowPrinter(hWndOwner As Long) As String
    ' Note:
    ' This is not my function it's from KPD-Team 1998 URL: http://www.allapi.net -> Code by Donald Grover
    ' and i have modified it a little
    Dim PrintDlg    As PRINTDLG_TYPE
    Dim DevMode     As DevMode
    Dim DevName     As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    m_Canceled = False
    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures
    ShowPrinter = ""

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hWndOwner = hWndOwner ' frmOwner.hWnd  'UserControl.ContainerHwnd 'Handle to window

    On Error Resume Next
    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    DevMode.dmCopies = Printer.Copies
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))

    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
            Next
        End If

        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex '1=uit  2=lange zijden  3=korte zijden
        Printer.Orientation = DevMode.dmOrientation '1=Portrait 2 =landscape
        Printer.PaperSize = DevMode.dmPaperSize '9=A4 8=A3
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        On Error GoTo 0
    Else
        m_Canceled = True
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
    ShowPrinter = Printer.DeviceName
End Function


'Functions for wrappers
'----------------------

' Remove null chars from filename and parse multi filename
Private Function ParseFileName(sFileName As String)
    Dim i           As Long
    Dim sPath       As String
    Dim sFiles()    As String
    Dim Pos         As Integer
    Dim sFile       As String
    Dim sFileTitle  As String

    ' Create new collections
    Set cFileName = New Collection
    Set cFileTitle = New Collection
    ' Found position of two last null chars
    Pos = InStr(sFileName, vbNullChar & vbNullChar)
    ' Remove from filename last two chars
    sFile = Left(sFileName, Pos - 1)
    ' Check to see if filename is single or multi
    If InStr(1, sFile, vbNullChar) <> 0 Then
    ' Multi file
        sFile = Left(sFileName, Pos) & vbNullChar 'Add null char at end of filename
        sPath = Left(sFileName, InStr(1, sFileName, Chr(0)) - 1) 'Get file path
        sFiles = Split(sFile, Chr(0)) 'Split file where is nullchar

        ' Add all filenames to collection
        For i = LBound(sFiles) To UBound(sFiles) - 2
            ' If path doesent contain separator then add it
            If Right(sPath, 1) = "\" Then
                cFileName.Add sPath & sFiles(i)
            Else
                cFileName.Add sPath & "\" & sFiles(i)
            End If
            ' Add file title
            cFileTitle.Add sFiles(i)
            ' Remove first item from collections
            If i = 1 Then cFileName.Remove 1: cFileTitle.Remove 1
        Next
    Else ' Single file
        'Add file name to collection
        cFileName.Add sFile
        ' Add file title
        cFileTitle.Add Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
    End If
End Function

' Replace "|" with Null Character
Private Function SetFilter(sFlt As String) As String
    Dim sLen    As Long
    Dim Pos     As Long

    sLen = Len(sFlt) 'Get filter length
    Pos = InStr(1, sFlt, "|") 'Find first position of "|"
    ' Loop while Pos > 0
    While Pos > 0
        ' Replace "|" with null char
        sFlt = Left(sFlt, Pos - 1) & vbNullChar & Mid(sFlt, Pos + 1, sLen - Pos)
        ' Find next position of "|"
        Pos = InStr(Pos + 1, sFlt, "|")
    Wend
    SetFilter = sFlt ' Set filter
End Function

' Get window name
Public Function GetCaption4Win(hWnd)
    Dim hwndLength%, hwndTitle$, a%
    hwndLength% = GetWindowTextLength(hWnd)
    hwndTitle$ = String$(hwndLength%, 0)
    a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))
    GetCaption4Win = hwndTitle$
End Function

'Properties:
'-----------

'Property FILENAME(S)
Public Property Get FileName() As String
    On Error Resume Next
    FileName = ""
    If m_MultiSelect Then
        Dim i As Integer
        For i = 1 To cFileName.Count
            FileName = FileName & cFileName(i)
            If i <> cFileName.Count Then FileName = FileName & ","
        Next i
    Else
        FileName = cFileName(1)
    End If
End Property

'Property CANCELED
Public Property Get Canceled() As Boolean
    Canceled = m_Canceled
End Property

'Property CANCEL ERROR
Public Property Get CancelError() As Boolean
    CancelError = m_CancelError
End Property
Public Property Let CancelError(ByVal New_CancelError As Boolean)
    m_CancelError = New_CancelError
End Property

'Property MULTI SELECT
Public Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property
Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    m_MultiSelect = New_MultiSelect
End Property

'Property DEFAULT FILENAME
Public Property Get DefaultFileName() As String
    DefaultFileName = m_FileName
End Property
Public Property Let DefaultFileName(ByVal New_Filename As String)
    m_FileName = New_Filename
End Property

'Property DIALOG TITLE
Public Property Get DialogTitle() As String
    DialogTitle = m_DialogTitle
End Property
Public Property Let DialogTitle(ByVal New_DialogTitle As String)
    m_DialogTitle = New_DialogTitle
End Property

'Property INITIAL DIRECTORY
Public Property Get InitialDir() As String
    InitialDir = m_InitialDir
End Property
Public Property Let InitialDir(ByVal New_InitialDir As String)
    m_InitialDir = New_InitialDir
End Property

'Property FILTER
Public Property Get Filter() As String
    Filter = m_Filter
End Property
Public Property Let Filter(ByVal New_Filter As String)
    m_Filter = New_Filter
End Property

'Property FILTER INDEX
Public Property Get FilterIndex() As Integer
    FilterIndex = m_FilterIndex
End Property
Public Property Let FilterIndex(ByVal New_FilterIndex As Integer)
    m_FilterIndex = New_FilterIndex
End Property

'Property FONT NAME
Public Property Get FontName() As String
    FontName = m_FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
End Property

'Property FONT SIZE
Public Property Get FontSize() As Integer
    FontSize = m_FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Integer)
    m_FontSize = New_FontSize
End Property

'Property FONT COLOR
Public Property Get FontColor() As Long
    FontColor = m_FontColor
End Property
Public Property Let FontColor(ByVal New_FontColor As Long)
    m_FontColor = New_FontColor
End Property

'Property FONT BOLD
Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_FontBold = New_FontBold
End Property

'Property FONT ITALIC
Public Property Get FontItalic() As Boolean
    FontItalic = m_FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_FontItalic = New_FontItalic
End Property

'Property FONT UNDERLINE
Public Property Get FontUnderline() As Boolean
    FontUnderline = m_FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    m_FontUnderline = New_FontUnderline
End Property

'Property FONT STRIKETHRU
Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = m_FontStrikeThru
End Property
Public Property Let FontStrikeThru(ByVal New_FontStrikeThru As Boolean)
    m_FontStrikeThru = New_FontStrikeThru
End Property


'Property PAGESETUP RIGHT MARGIN
Public Property Get PageRightMargin() As Double
    PageRightMargin = m_PageRightMargin
End Property
Public Property Let PageRightMargin(ByVal New_Value As Double)
    m_PageRightMargin = New_Value
End Property

'Property PAGESETUP LEFT MARGIN
Public Property Get PageLeftMargin() As Double
    PageLeftMargin = m_PageLeftMargin
End Property
Public Property Let PageLeftMargin(ByVal New_Value As Double)
    m_PageLeftMargin = New_Value
End Property

'Property PAGESETUP TOP MARGIN
Public Property Get PageTopMargin() As Double
    PageTopMargin = m_PageTopMargin
End Property
Public Property Let PageTopMargin(ByVal New_Value As Double)
    m_PageTopMargin = New_Value
End Property

'Property PAGESETUP BOTTOM MARGIN
Public Property Get PageBottomMargin() As Double
    PageBottomMargin = m_PageBottomMargin
End Property
Public Property Let PageBottomMargin(ByVal New_Value As Double)
    m_PageBottomMargin = New_Value
End Property

'Property PAGESETUP BOTTOM MARGIN
Public Property Get PaperSizeWidth() As Double
    PaperSizeWidth = m_PaperSizeWidth
End Property
Public Property Let PaperSizeWidth(ByVal New_Value As Double)
    m_PaperSizeWidth = New_Value
End Property

'Property PAGESETUP BOTTOM MARGIN
Public Property Get PaperSizeHeight() As Double
    PaperSizeHeight = m_PapersizeHeight
End Property
Public Property Let PaperSizeHeight(ByVal New_Value As Double)
    m_PapersizeHeight = New_Value
End Property

'Property PAGESETUP BOTTOM MARGIN
Public Property Get PaperSize() As String
    PaperSize = m_PaperSize
End Property
Public Property Let PaperSize(ByVal New_Value As String)
    m_PaperSize = New_Value
End Property

'Property PAGESETUP BOTTOM MARGIN
Public Property Get PaperOrientation() As String
    PaperOrientation = m_PaperOrientation
End Property
Public Property Let PaperOrientation(ByVal New_Value As String)
    m_PaperOrientation = New_Value
End Property

'Property PAGESETUP BOTTOM MARGIN
Public Property Get Directory() As String
    Dim iPos As Integer
    Dim sRet As String
    Directory = ""
    If FileName = "" Then Exit Property
    sRet = Trim$(FileName)
    iPos = InStrRev(sRet, "\", , vbTextCompare)
    If iPos > 0 Then sRet = Trim$(Left$(sRet, iPos))
    Directory = sRet
End Property

'Property PAGESETUP BOTTOM MARGIN
Public Property Get File() As String
    Dim iPos As Integer
    Dim sRet As String
    File = ""
    If FileName = "" Then Exit Function
    sRet = FileName
    iPos = InStrRev(sRet, "\", , vbTextCompare)
    If iPos > 0 Then sRet = Trim$(Right$(sRet, Len(sRet) - iPos))
    File = sRet
End Property

'Get the name for a papersize (A?)
Private Function GetPaperSizeName(ByVal lWidth As Long, ByVal lHeight As Long)
    'ToDo: add all papersizes other than Din Format
    'Always use portrait sizes to compare
    If lWidth > lHeight Then
        Dim lTemp As Long
        lTemp = lHeight
        lHeight = lWidth
        lWidth = lTemp
    End If
    If lHeight <= 37 And lWidth <= 26 Then
        GetPaperSizeName = "A10"
    ElseIf lHeight <= 52 And lWidth <= 37 Then
        GetPaperSizeName = "A9"
    ElseIf lHeight <= 74 And lWidth <= 52 Then
        GetPaperSizeName = "A8"
    ElseIf lHeight <= 105 And lWidth <= 74 Then
        GetPaperSizeName = "A7"
    ElseIf lHeight <= 148 And lWidth <= 105 Then
        GetPaperSizeName = "A6"
    ElseIf lHeight <= 210 And lWidth <= 148 Then
        GetPaperSizeName = "A5"
    ElseIf lHeight <= 297 And lWidth <= 210 Then
        GetPaperSizeName = "A4"
    ElseIf lHeight <= 420 And lWidth <= 297 Then
        GetPaperSizeName = "A3"
    ElseIf lHeight <= 594 And lWidth <= 420 Then
        GetPaperSizeName = "A2"
    ElseIf lHeight <= 841 And lWidth <= 594 Then
        GetPaperSizeName = "A1"
    ElseIf lHeight <= 1189 And lWidth <= 841 Then
        GetPaperSizeName = "A0"
    Else
        GetPaperSizeName = "A?"
    End If
End Function

'Initialize class
Private Sub Class_Initialize()
    On Error Resume Next
    FontBold = False
    FontSize = 10
    FontColor = vbBlack
    FontName = "Arial"
    Filter = "All Files (*.*)|*.*"
    InitialDir = App.Path
    Err.Clear
End Sub

'Terminate class
Private Sub Class_Terminate()
    On Error Resume Next
    Err.Clear
End Sub


