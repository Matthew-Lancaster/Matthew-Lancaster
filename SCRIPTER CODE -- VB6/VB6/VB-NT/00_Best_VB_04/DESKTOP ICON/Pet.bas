Attribute VB_Name = "Module1"
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' Locate Desktop Icons Source Code VB6
' Created 11 April 1999    Updated 24 June 1999,3 August 1999
' Now also obtains text associated with each icon.
' Retrieves Icon Images.
' Changes resolution while preserving icon pattern.
' Hides and UnHides Icons
' Detects screen resolution changes and rebuilds icon pattern

' By Paul Pavlic
' abuse and advice to    pepp@cyberdude.com
' Feel free to use anyway you can
' FreeWare
' Works on Win95
'
' Special thanks to Bruce McKinney
' Author of Hardcore Visual Basic
' A wonderful book with a great many insights
'
' Why?
' I looked all over the Internet, and never once saw any working
' code to do this, thought I would make a contribution

' Why call it Paulies Pet?
' Was the name of my Desktop Pet in VB3 rewriting it now to VB6
' Ported to VB6 Source available at
' http://www.vbcode.com/
'
' This code resolves Explorer page faults when trying to send
' LVM_GETITEMPOSITION to the Desktop Listview
' Also LVM_GETITEMTEXT

'Explores a number of issues involved with sending
'and retrieving messages between different processes.
'*************************************

'constants
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_ALWAYS = 4
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const SECTION_MAP_WRITE = &H2
Public Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Public Const LVFI_PARAM = 1
Public Const LVIF_TEXT = &H1
Public Const LVIF_IMAGE = &H2


'???? NOT documented in Win32api.txt ????
Public Const PAGE_READWRITE As Long = &H4
'???? NOT documented in Win32api.txt ????



Const LVM_GETTITEMCOUNT& = (&H1000 + 4)
Public Const LVM_SETITEMPOSITION& = (&H1000 + 15)
Const LVM_FIRST = &H1000
Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Const LVM_GETITEMTEXT = LVM_FIRST + 45
Public Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
Public Const LVSIL_NORMAL = 0
Public Const LVM_GETITEM = (LVM_FIRST + 5)
Public Const LVM_ARRANGE = (LVM_FIRST + 22)
'LVA_DEFAULT
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_BITSPERPEL = &H40000

Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17

':-)
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_GLOBAL = &H8
Public Const CDS_FULLSCREEN = &H4
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const WM_SETHOTKEY = &H32
Public Const WM_HOTKEY = &H312



'damn  hell of a lot of declares
'copymemory *3 avoid byval in code - bug? works this way
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Sub CopyMemoryOne Lib "kernel32" Alias "RtlMoveMemory" _
(ByVal hpvDest&, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Sub CopyMemoryTwo Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, ByVal hpvSource&, ByVal cbCopy As Long)




'display declares
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function EnumDisplaySettings Lib "user32" Alias _
   "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, _
   ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias _
   "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias _
"lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function FindWindowEx& Lib "user32" Alias _
"FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter _
As Long, ByVal lpClassName As String, ByVal lpWindowName As String)
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
lprect As Any, ByVal bErase As Long) As Long

'messaging declares
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam _
As Any) As Long
Declare Function SendMessageByLong& Lib "user32" Alias "SendMessageA" _
(ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Public Declare Function SendMessageBywParamAny Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam _
As Any) As Long



'declares for printing to the desktop or other window for debug purposes
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal _
hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As _
String, ByVal nCount As Long) As Long


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'declares for memory-mapped files
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As _
Long, ByVal hTemplateFile As Long) As Long
' changed lpFileMappigAttributes to Any, makes life much easier
Public Declare Function CreateFileMappingTwo Lib "kernel32" Alias _
"CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes _
As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, _
ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal _
hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal _
dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal _
dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress _
As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
As Long) As Long
Public Declare Function FlushViewOfFile Lib "kernel32" (ByVal lpBaseAddress _
As Long, ByVal dwNumberOfBytesToFlush As Long) As Long

Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

'declare to grab icons from imagelist
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal flags&) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lprect As RECT) As Long



'>>>>>  type declarations   <<<<<<
Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Global DevM As DEVMODE

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Public Type ScreenDisplayModes
    bits As Integer
    nwidth As Long
    nheight As Long
End Type




Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Type POINTAPI2
        x As Integer
        y As Integer
End Type

Public Type StringBuf
        x As String * 40
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Dim d As StringBuf
Dim c As POINTAPI


Public IconPosition() As POINTAPI
Public IconPosition2() As POINTAPI
Public IconName() As String
Public TempIconPosition2 As POINTAPI
Public IconNumber() As Long
'dimension some variables
Dim pNull As Long
Global StartWidth
Global StartHeight
Global CurrentWidth As Long
Global CurrentHeight As Long


Dim MyValue&, MyValue2&
Dim sFileName As String
Global CurrentDirectory As String
Global hdesk As Long
Dim i&
Global icount&
Global MovedIcon&
Global WroteToDeskTop&
'yes I know global is a dirty word
'one day I might even study programming
'except I do it for fun <g>
Dim objItem As LV_ITEM
Dim imageLVI As LV_ITEM
Dim lngLength As Long
Global hImageList
Public started As Long
Global Hidden() As String * 1
Global ScrMode() As ScreenDisplayModes
Global UnloadFlag As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Public Const WM_QUIT = &H12

Public Sub FindIcons()

'no error trapping done  quick and dirty code
'this is Proof of Concept Code
'a problem for the reader <g>

Form1.List1.Clear

Dim lprect As RECT
Dim DeskHwnd As Long
DeskHwnd = GetDesktopWindow()
ss& = GetWindowRect(DeskHwnd, lprect) 'get the size
StartWidth = lprect.Right             'of the Desktop
StartHeight = lprect.Bottom
CurrentWidth = StartWidth
CurrentHeight = StartHeight

hdesk = FindWindow("progman", vbNullString)
hdesk = FindWindowEx(hdesk, 0, "shelldll_defview", vbNullString)
hdesk = FindWindowEx(hdesk, 0, "syslistview32", vbNullString)
'hdesk is the handle of the Desktop's listview

'find out how many icons reside on the desktop
icount = SendMessageByLong(hdesk, LVM_GETTITEMCOUNT, 0, 0&)

'bail if we get a zero count
If icount = 0 Then MsgBox "Error occurred: No icons found", _
vbOKOnly, "PauliesPet": Unload Form1: End

'tell me how many icons we found
Form1.Text1.Text = Str(icount) + "  icons detected    "

'redimension arrays to the number of icons found
ReDim IconPosition(icount) As POINTAPI
ReDim IconPosition2(icount) As POINTAPI
ReDim IconName(icount) As String
ReDim IconNumber(icount) As Long
ReDim Hidden(icount) As String * 1

'window being called must be in large or small icon mode
'the desktop is  so we don't check
'LVSIL_NORMAL = 0

'returns handle to ImageList
hImageList = SendMessageByLong(hdesk, LVM_GETIMAGELIST, 0&, 0&)
'we'll use this later to draw the icons


'get our current directory and append \
CurrentDirectory = App.Path
If Right$(CurrentDirectory, 1) <> "\" Then _
CurrentDirectory = CurrentDirectory + "\"

'set to null
pNull = 0

'+++++  create memory mapped files * 2  +++++

'///// create a memory-mapped file (1) /////
sFileName = CurrentDirectory + "TEMPPPPP.PPP"
' Open file
hFile = CreateFile(sFileName, GENERIC_READ Or GENERIC_WRITE, 0, _
                   ByVal pNull, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, _
                   pNull)
' get handle
hFileMap = CreateFileMappingTwo(hFile, ByVal pNull, PAGE_READWRITE, _
0, 40, "MyMapping")
' Get pointer to memory representing file
pFileMap = MapViewOfFile(hFileMap, FILE_MAP_WRITE, 0, 0, 0)

'///// create a memory-mapped file (2) /////
sFileName2 = CurrentDirectory + "TEMPPPPP.PP2"
' Open file
hFile2 = CreateFile(sFileName2, GENERIC_READ Or GENERIC_WRITE, 0, _
                   ByVal pNull, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, _
                   pNull)
' get handle
hFileMap2 = CreateFileMappingTwo(hFile2, ByVal pNull, PAGE_READWRITE, _
0, 40, "MyMapping2")
' Get pointer to memory representing file
pFileMap2 = MapViewOfFile(hFileMap2, FILE_MAP_WRITE, 0, 0, 0)






'*********************************************** here we go
'loop through icon count
For i = 0 To icount - 1
    
    'pFileMap <lparam> is mem-map file Pointer
    'send message to explorer listview (desktop)
    Call SendMessageByLong(hdesk, LVM_GETITEMPOSITION, i, pFileMap)
    'copy returned to our POINTAPI (c.x,c.y)
    CopyMemoryTwo c, pFileMap, 8
       
    'put value in our arrays
    IconPosition(i) = c
    
    'back up array for swapping later
    IconPosition2(i) = c

    
    If IconPosition(i).x > StartWidth Or IconPosition(i).y > StartHeight Then
        'icon is hidden  I.E. NOT visible on desktop
        Hidden(i) = "Y"
    Else
        Hidden(i) = "N"
    End If
    
    
    
    
'==== Obtain the name of the specified list view item ===
    
    'first load what we want into our LV_ITEM type   objItem
    objItem.mask = LVIF_TEXT
    objItem.pszText = pFileMap2 'pointer to second mem-map file
    objItem.cchTextMax = 40     'max length we want to deal with
    ' copy LV_ITEM structure to memory mapped file
    CopyMemoryOne pFileMap, objItem, 40
    'send message to explorer listview (desktop)
    lngLength = SendMessageByLong(hdesk, LVM_GETITEMTEXT, i, _
                                pFileMap)
    'returns the number of characters in the text
    'copy returned to our StringBuf (d.x)
    'we use a type here to avoid problems with CopyMemory API
    CopyMemoryTwo d, pFileMap2, 40
              
    'make sure that the text length is not larger than 40 characters
    If lngLength > 40 Then lngLength = 40
    strName = Left$(d.x, lngLength)
    IconName(i) = strName
        
    '==== Obtain the icon number of the specified list view item ===
    '     for later retrieval from imagelist
    
    'first load what we want into our LV_ITEM type   objItem
    objItem.mask = LVIF_TEXT Or LVIF_IMAGE 'needs to be both
    objItem.iItem = i 'which item info do we want ?
    objItem.pszText = pFileMap2 'better load this
    objItem.cchTextMax = 40     'but we won't use it
    ' copy LV_ITEM structure to memory mapped file
    CopyMemoryOne pFileMap, objItem, 40
    lngLength = SendMessageByLong(hdesk, LVM_GETITEM, i, _
                                pFileMap)
    
    
    
    CopyMemoryTwo objItem, pFileMap, 40
    'ok put the icon imagelist number into memory
    IconNumber(i) = objItem.iImage
'HKEY_CURRENT_USER\Software\Microsoft\Windows\Current Version\Explorer\Shell Folders
    
    ' ok show all the info we have received
    Form1.List1.AddItem Str(i + 1) + Chr$(9) + Str(c.x) + Chr$(9) + _
    Str(c.y) + Chr$(9) + Str(objItem.iImage) + Chr$(9) + Hidden(i) + Chr$(9) + strName
    
    
Next i

'phew

'IMPORTANT:
'Release resources back to windows
FlushViewOfFile pFileMap, 40
UnmapViewOfFile pFileMap
CloseHandle hFileMap
CloseHandle hFile
FlushViewOfFile pFileMap2, 40
UnmapViewOfFile pFileMap2
CloseHandle hFileMap2
CloseHandle hFile2

End Sub
Sub RefreshDesktop()

' refresh the whole desktop , we'll just invalidate everything
xcc& = InvalidateRect(0&, ByVal 0, 0&)

End Sub
Sub RefreshDesktopFully()
' refresh the whole desktop , we'll just invalidate everything
'including background
xcc& = InvalidateRect(0&, ByVal 0, 1&)
End Sub
Sub MarkIcons()

' get the desktop device context for debug print to desktop
hwndSrc& = 0
hSrcDC& = GetDC(hwndSrc&) 'same as GetDc(Null)  Desktop dc
For i = 0 To icount - 1
    Test$ = Str(i + 1) 'print this to the the pos. on the desktop
    zs& = TextOut(hSrcDC&, IconPosition(i).x, _
    IconPosition(i).y, Test$, Len(Test$))
Next i
' **** IMPORTANT **** release resources back to Windows
Dmy& = ReleaseDC(hwndSrc&, hSrcDC&)

End Sub
Sub SwapIcon()

'seed random number generator
Randomize Timer
'get a random number from 0 to icount
MyValue = Int((icount * Rnd))
'get a 2nd random number from 0 to icount
'must be a different one
Randomize Timer
Do
    MyValue2 = Int((icount * Rnd))
Loop While MyValue = MyValue2

'swap icon positions
Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, MyValue, _
CLng(IconPosition2(MyValue2).x + IconPosition2(MyValue2).y * &H10000))
Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, MyValue2, _
CLng(IconPosition2(MyValue).x + IconPosition2(MyValue).y * &H10000))
  'update our second array
  TempIconPosition2 = IconPosition2(MyValue2)
  IconPosition2(MyValue2) = IconPosition2(MyValue)
  IconPosition2(MyValue) = TempIconPosition2
'tell me what we swapped
Form1.Text1.Text = "Swapping Icon Position " + Str(MyValue + 1) + _
" with Icon Position " + Str(MyValue2 + 1)

End Sub
Sub RefreshIconPositions()

For i = 0 To icount - 1
'reset our icon switching array
IconPosition2(i) = IconPosition(i)
'Set icon postions back to what we found originally
Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, i, _
CLng(IconPosition(i).x + IconPosition(i).y * &H10000))
Next i

End Sub
Sub CentreFormInScreen()
'straight out of the book
Dim TopCorner As Integer
Dim LeftCorner As Integer
  
  If Form1.WindowState <> 0 Then Exit Sub
  TopCorner = (Screen.Height - Form1.Height) \ 2
  LeftCorner = (Screen.Width - Form1.Width) \ 2
  Form1.Move LeftCorner, TopCorner

End Sub
Sub PutAllIconsOnDeskTop()
Call SendMessageByLong(hdesk, LVM_ARRANGE, 0, 0)
End Sub
