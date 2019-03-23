Attribute VB_Name = "CRCMod"

'IN THE CLASS
Public FLAG_DO_IN_CHUNK_OUT_OF_MEMORY

Public WORK_DONE
Public Filename_Script
Public Filename_Script_FILE

Public MaxFileSizeCRCCheck
Public MP3_HEADER_TAG_SIZE


Public OLDlblCount6

Public Option_Mnu_Any_Selected
Public CommandExit
Public CommandPause
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public mCancelScan As Boolean
Public PathToLoad

Public Option_Mnu_Dont_Del_From_First_Source_Folder1
Public Option_MNU_CRC_Check_After_First
Public Option_Mnu_Simulate
Public Option_MNU_Check_FOR_NULL_STRING
Public MENU_OPTION

Public m_CRC As clsCRC

Public RDClip
'Put in Sub Of Code
Public Scan_Individual
Public BIGGEST_FILESIZE
Rem


Public SCAN_REFINE1()
Public SCAN_REFINE2()
Public USE_SCAN_REFINE


'Option Explicit
Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type


Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type

'-----------------------------------------------
'AUTOSIZE LISTVIEW COLUMN
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000
'-----------------------------------------------
'Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
' As ColumnHeader = Nothing)
'
' Dim C As ColumnHeader
' If Column Is Nothing Then
'  For Each C In LV.ColumnHeaders
'   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
'  Next
' Else
'  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
' End If
' LV.Refresh
'End Sub
'-----------------------------------------------



Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
 As ColumnHeader = Nothing)

 Dim C As ColumnHeader
 If Column Is Nothing Then
  For Each C In LV.ColumnHeaders
   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
  Next
 Else
  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 LV.Refresh
End Sub


'------------
'MAIN RUSSIAN ROLLETTE
'------------
'MAIN ROUTINE
'------------

Sub CRCDUPE()



Dim WXHEX$


Dim ResultDel
Dim FR4
Dim FR1

ChDrive "C:"
ChDir "\"

Dim FileSize, FSize


'FileSizeDelTotal
'RESULT BE NOTHING AT WHILE BEGIN HERE
FileSizeDelTotal = 0
ScanPath.lblCount10 = Format(FileSizeDelTotal, "###,###,###,##0") + " Byte Deleted Total"



TVarNow = Now

Filename_Script_FILE = "Dupe's Deleted Script -- " + Format(TVarNow, "YYYY-MM-DD HH-MM-SS") + ".txt"

If Dir(ScanPath.txtPath + "z-Dupe's Deleted Script", vbDirectory) = "" Then
    MkDir ScanPath.txtPath + "z-Dupe's Deleted Script"
End If

Filename_Script = ScanPath.txtPath + "z-Dupe's Deleted Script\" + Filename_Script_FILE



ScanPath.Show
ScanPath.lblCount7.BackColor = RGB(255, 255, 0) '&HFFFF&

Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32


'FileName = App.Path + "\#_DATA\Visual Basic Code Loggs.txt"
'If Dir(FileName) <> "" Then
'    'Fs.getfile (FileName)
'    Set F = Fs.getfile(FileName)
'    xd1 = F.Size
'    xd2 = F.datelastmodified
'    '1024^2 ! Meg
'    If F.Size > 1024 ^ 2 Then
'        datefile = App.Path + "\#_DATA\" + Format(xd2, "YYYY-MM-DD HH-MM-SS") + " -- Visual Basic Code Loggs.txt"
'        Name FileName As datefile
'        MsgBox "Renamed and Archive the Logg Over a MEG Big" + vbCrLf + vbCrLf + FileName + vbCrLf + vbCrLf + datefile, vbOKOnly, ScanPath.Caption
'    End If
'End If





ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath = PathToLoad 'z_MENU_Form1.Label20.Caption 'Clipboard.GetText
ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked
ScanPath.lblCount7 = "Gather Files"
ScanPath.chkRefreshListView = vbUnchecked

'If IsIDE = True Then
'    ia = vbYes
'Else
'    If Fs.FolderExists(ScanPath.txtPath) = False Or Len(ScanPath.txtPath) < 4 Then
'        'ia = MsgBox("Not a Folder Set" + vbCrLf + Command$ + vbCrLf + "END", vbInformation Or vbMsgBoxSetForeground)
'        'End
'    ia = vbYes
'    End If
'    If Len(ScanPath.txtPath) < 4 And Fs.FolderExists(ScanPath.txtPath) = True Then
'        ia = MsgBox("Not Attempt Root" + vbCrLf + Command$ + vbCrLf + "END", vbInformation Or vbMsgBoxSetForeground)
'        End
'    End If
'End If
'
'
'If Fs.FolderExists(ScanPath.txtPath) = True And Len(ScanPath.txtPath) > 3 Then
'
'    If Mid(ScanPath.txtPath, Len(ScanPath.txtPath)) <> "\" Then ScanPath.txtPath = ScanPath.txtPath + "\"
'
'    ia = MsgBox("Run Hardcoded Jobs = Answer (Yes) or" + vbCrLf + "Run Command-line = Answer (No)" + vbCrLf + ScanPath.txtPath, vbYesNoCancel, vbMsgBoxSetForeground)
'
'    If ia = vbCancel Then End
'    xzag = 1
'End If
'If Fs.FolderExists(ScanPath.txtPath) = True And Len(ScanPath.txtPath) <= 3 Then MsgBox "NOT IN ROOT " + vbCrLf + ScanPath.txtPath: End
'
'If xzag = 0 Then
'    If ia = vbYes Then
'    IA1 = MsgBox("Run Hardcoded Jobs", vbYesNoCancel, vbMsgBoxSetForeground)
'    If IA1 = vbCancel Then End
'
'    End If
'    'vbyes
'    If ia = vbNo And IA1 = vbNo Then ia = MsgBox("Run Commandline Job" + vbcrlfScanPath.txtPath, vbYesNoCancel, vbMsgBoxSetForeground)
'    If ia = vbCancel Or ia = vbNo Then End
'End If
'
'If Fs.FolderExists(ScanPath.txtPath) = False And xzag = 0 And ia = vbNo Then MsgBox "NOT IN ROOT " + vbCrLf + ScanPath.txtPath: End
'

Dim XStartTime As Date
XStartTime = Now
'MENU_OPTION = 2 --- HARDCODED

If MENU_OPTION = 2 Then


'    ScanPath.txtPath = "M:\0 00 ART LOGGERS - WEBCAM\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
    
'    ScanPath.txtPath = "U:\0 00 ART LOGGERS\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    ScanPath.txtPath = "Y:\0 00 Art Loggers\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click

'    ScanPath.txtPath = "U:\0 00 ART LOGGERS\# TEMP INTERNET\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click


'w$ = "U:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
'w$ = "M:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
'w$ = "f:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
'w$ = "F:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\02 VB6 An MSDN 02\"
'w$ = "F:\03 MICROSOFT\# VB 6\"
'w$ = "M:\03 MICROSOFT\# OP SYS\Test\"
'w$ = "M:\DSC\"

'w$ = "I:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\02 VB6 An MSDN 02\"
'    'ScanPath.txtPath = "U:\00 0 VIDEO CAMERA'S\"
'    'ScanPath.txtPath = "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
''w$ = "F:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
    
    
    
'2015 jul 11
'    w$ = "I:\03 MICROSOFT\# VB 6\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
'2015 jul 11
'    w$ = "M:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
'2015 JUL 13 done
'    w$ = "U:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
 
''2015 jul 13 DONE
'    w$ = "I:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
''2015 jul 13 DONE
'    w$ = "I:\00 0 VIDEO CAMERA'S\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'
''2015 jul 13 DONE
'    w$ = "I:\0 00 MEDIA\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
    
''2015 JUL 14
'    w$ = "I:\0 00 ART LOGGERS\# FEED STATION\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "M:\0 00 ART LOGGERS\# FEED STATION\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "U:\0 00 ART LOGGERS\# FEED STATION\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    

'2015 JUL 14 start working on one drive - DONE JUN 16
'    w$ = "I:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "M:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "U:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
'2015 JUL 16 start working NEXT WAIT FOR SUPERCOPY
'subst i: w:
    
    
'    w$ = "D:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
    
'    w$ = "T:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "w:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "X:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click

'NEXT JOB
'
'    w$ = "T:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "W:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "X:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
'NEXT JOB
 
'    w$ = "T:\Videos\00 Vid XXX\AUTOPIX 01\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "W:\Videos\00 Vid XXX\AUTOPIX 01\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "X:\Videos\00 Vid XXX\AUTOPIX 01\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'

'NEXT
'    w$ = "T:\0 00 HTTrack\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "W:\0 00 HTTrack\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "X:\0 00 HTTrack\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'





'Next do in exe
'    w$ = "T:\0 00 ART WORK IN BOUND\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click

    
'DONE
'    w$ = "D:\VB6\VB6 Icons\HTTRACK - 00 ICONS\Free Icons Web\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
    
    
'----------------- done
'T:\0 00 Art\# 00 My Pictures\02 HTTRACK\NASA
'M: SUBST As U:
'U:\0 00 Art\# 00 My Pictures\02 HTTRACK\NASA
'X:\0 00 Art\00 HTTrack\00 NASA
    
'    w$ = "T:\0 00 Art\# 00 My Pictures\02 HTTRACK\NASA\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "U:\0 00 Art\# 00 My Pictures\02 HTTRACK\NASA\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "X:\0 00 Art\00 HTTrack\00 NASA\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
    
'
    
    
    'NEXT
'    w$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click



'
'
'    ScanPath.txtPath = "U:\0 00 MEDIA\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'

'
'    ScanPath.txtPath = "U:\KAT MP3 RECORDER\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    ScanPath.txtPath = "Y:\KAT MP3 RECORDER\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'

'    ScanPath.txtPath = "Y:\#\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click


'Done 2015-07-22
'    w$ = "T:\0 00 Music ---\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "X:\0 00 Music ---\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click

'w$ = "d:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click

'X:\0 00 ART LOGGERS - WEB CAM
'w$ = "T:\0 00 ART LOGGERS - WEB CAM\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$: End
'    Call ScanPath.cmdScan_Click
''w$ = "U:\0 00 ART LOGGERS - WEB CAM\"
''    ScanPath.txtPath = w$
''    If Fs.FolderExists(ScanPath.txtPath) = False Then End
''    Call ScanPath.cmdScan_Click
'w$ = "X:\0 00 ART LOGGERS - WEBCAM\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
'    Call ScanPath.cmdScan_Click
'w$ = "T:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
'    Call ScanPath.cmdScan_Click
'w$ = "X:\DSC\" ' D DRIVE
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
'    Call ScanPath.cmdScan_Click
'w$ = "U:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
'    Call ScanPath.cmdScan_Click
'w$ = "V:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
'    Call ScanPath.cmdScan_Click
'w$ = "Y:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
'    Call ScanPath.cmdScan_Click



'W$ = "D:\0 00 MUSIC ---\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "M:\0 00 MUSIC ---\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "T:\0 00 MUSIC ---\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "T:\0 00 Music --Z\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "V:\0 00 MUSIC ---\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "U:\0 00 MUSIC ---\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'


'W$ = "D:\# MY DOCS\# 01 My Documents\Downloads\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "D:\# MY DOCS\# 01 My Documents\Downloads ZZZ\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "D:\#0 1 INSTALLATIONS\Installation programs\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click

'
'W$ = "D:\0 00 MUSIC ---\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "D:\0 00 MUSIC -Z0\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "D:\0 00 MUSIC -Z1\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "D:\0 00 MUSIC -Z2\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "D:\0 00 MUSIC -Z3\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'W$ = "D:\0 00 MUSIC -Z4\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'

'w$ = "T:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
'    Call ScanPath.cmdScan_Click
'
'w$ = "X:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
'    Call ScanPath.cmdScan_Click


'GoTo JUMP_H3
GoTo JUMP_H2
GoTo JUMP_HIGHCORE

w$ = "D:\0 00 Art Loggers\# APP AND SCREEN -- SHOT"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click
w$ = "T:\0 00 ART LOGGERS\# APP AND SCREEN -- SHOT"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click
w$ = "M:\0 00 ART LOGGERS\# APP AND SCREEN -- SHOT"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click


JUMP_HIGHCORE:

w$ = "D:\0 00 Art Loggers\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click
w$ = "M:\0 00 ART LOGGERS\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click
w$ = "T:\0 00 ART LOGGERS\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click



'FILTER OUT AND DON'T EVEN LIST
ReDim SCAN_REFINE1(14)

SCAN_REFINE1(1) = "D:\0 00 Art\# 00 My Pictures - LARGE COUNT"
SCAN_REFINE1(2) = "D:\#\"
SCAN_REFINE1(3) = "D:\#0 "
SCAN_REFINE1(4) = "D:\0 00 ART WORK IN BOUND"
SCAN_REFINE1(5) = "D:\0 00 HTTrack"
SCAN_REFINE1(6) = "D:\0 00 MUSIC -"
SCAN_REFINE1(7) = "D:\0 00 VIDEO"
SCAN_REFINE1(8) = "D:\03 MICROSOFT"
SCAN_REFINE1(9) = "D:\DOCUM"
SCAN_REFINE1(10) = "D:\Email_OE"
SCAN_REFINE1(11) = "D:\FeedDemon"
SCAN_REFINE1(12) = "D:\RECYCLER"
SCAN_REFINE1(13) = "D:\System Volume Information"
SCAN_REFINE1(14) = "D:\VB6"

For we = 1 To UBound(SCAN_REFINE1)
    SCAN_REFINE1(we) = UCase(SCAN_REFINE1(we))
Next


JUMP_H3:


'DON'T DELETE IN MAIN ROUTINE
ReDim SCAN_REFINE2(3)

SCAN_REFINE2(1) = "D:\"
SCAN_REFINE2(2) = "T:\DSC\"
SCAN_REFINE2(3) = "T:\00 0 VIDEO CAMERA'S\"

ReDim SCAN_REFINE2(1)

SCAN_REFINE2(1) = "D:\# MY DOCS"


For we = 1 To UBound(SCAN_REFINE2)
    SCAN_REFINE2(we) = UCase(SCAN_REFINE2(we))
Next


ReDim SCAN_REFINE1(0)
'ReDim SCAN_REFINE2(0)

USE_SCAN_REFINE = False
'USE_SCAN_REFINE = True

GoTo JUMP10


w$ = "D:\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click

'OFFENDING JOB TO SCAN
w$ = "T:\DSC Z COMPARE1\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click

w$ = "T:\DSC\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click

w$ = "T:\00 0 VIDEO CAMERA'S\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click


JUMP10:



'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------

JUMP_H2:


'W$ = "D:\DSC\DCIM\IMAGES\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'
'W$ = "D:\# MY DOCS\# 01 My Documents\My Pictures\MP Navigator EX\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'
'W$ = "\\Asuseee\asus eee d\# MY DOCS\# 01 My Documents\My Pictures\MP Navigator EX\"
'W$ = "\\Linda-pc\d acer\# MY DOCS\# 01 My Documents\My Pictures\MP Navigator EX\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'
'    ReDim SCAN_REFINE1(0)
'    ReDim SCAN_REFINE2(0)

'W$ = "D:\# MY DOCS\"



'W$ = "D:\# MY DOCS\# 01 My Documents\Downloads\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'
'W$ = "D:\#MY DOCS OF ASUS BIGGER\"
'W$ = "D:\# MY DOCS\# 01 My Documents\Downloads Archive Previous Year\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
'
'W$ = "D:\# MY DOCS\# 01 My Documents\Downloads ZZZ\"
'    ScanPath.txtPath = W$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + W$:  End
'    Call ScanPath.cmdScan_Click
    
    
w$ = "D:\#\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then MsgBox "NOT A FOLDER END" + vbCrLf + w$:  End
    Call ScanPath.cmdScan_Click
    
    
    ReDim SCAN_REFINE1(0)
    ReDim SCAN_REFINE2(0)


End If

Dim D1$
Dim Source_Path_Folder_Filter


'-------------------------------------
'PROBLEM WITH ACCESS LOCKED FILES
Reset
'-------------------------------------


If MENU_OPTION = 1 Then
    
    ReDim SCAN_REFINE1(0)
    ReDim SCAN_REFINE2(0)
    
    If Mid(ScanPath.txtPath, Len(ScanPath.txtPath)) <> "\" Then ScanPath.txtPath = ScanPath.txtPath + "\"
    If Fs.FolderExists(ScanPath.txtPath) = False Then End
    
    
    '--------------------
    If ScanPath.mnu_Dont_Del_From_First_Source_Folder.Checked = True Then
        Option_Mnu_Dont_Del_From_First_Source_Folder1 = True
    End If
    If ScanPath.Mnu_Simulate.Checked = True Then
        Option_Mnu_Simulate = True
    End If
    If ScanPath.MNU_CRC_Check_After_First.Checked = True Then
        Option_MNU_CRC_Check_After_First = True
    End If
    If ScanPath.MNU_Check_FOR_NULL_STRING.Checked = True Then
        Option_MNU_Check_FOR_NULL_STRING = True
    End If
    '--------------------
'    If Option_MNU_CRC_Check_After_First = True Then
'        Option_Mnu_Dont_Del_From_First_Source_Folder1 = True
'    End If
'    If Option_Mnu_Simulate = True Then
'        Option_Mnu_Dont_Del_From_First_Source_Folder1 = True
'    End If
    '--------------------
    
    
    '-----------------------------------------------
    If ScanPath.Mnu_Repeat_Questions_Toggle.Checked Then
        If Option_Mnu_Dont_Del_From_First_Source_Folder1 = True Then
            Option_Mnu_Any_Selected = True
        End If
        If Option_Mnu_Any_Selected = False Then MsgBox "You Havn't Selected any Options" + vbCrLf + "So You Will Be Asked Them Anyway After The Scan Path", , ScanPath.Caption
    End If
    '-----------------------------------------------
    
    Call ScanPath.cmdScan_Click
    
    If mCancelScan = True Then GoTo ENDE
    
    ScanPath.ListView1.Refresh
    
    'PROBLEM WITH CMDSCAN LEAVE FILES OPEN MAYBE
    Reset
        
    
    '------------------------------
    'OLD STYLE OF FIND FIRST FOLDER
    '------------------------------
    If 1 = 2 Then
    
'    If ScanPath.mnu_Dont_Del_From_First_Source_Folder.Checked = True Then
'        Source_Path_Folder_Filter = D1$
        
        For we = 1 To ScanPath.ListView1.ListItems.Count
            DoEvents
            ScanPath.lblCount2 = Trim(Str(we))
            
            A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
            B1$ = ScanPath.ListView1.ListItems.Item(we)
            'GET FOLDER PATH FROM PATH FILE
            
            
            'IF THE 1ST FOLDER IS NOT REAL FIRST FOLDER
            'BECUASE NOT FILES PRESENT IN FIRST THEN
            'CHECK AGAIN
            If A1$ <> ScanPath.txtPath Then TRIG1 = True
            'MIGHT NEED THIS ANOTHER TIME
            
            If TRIG1 = True Then
                
            
                'd1$ = A1$ 'Mid(A1$, 1, InStr(Len(ScanPath.txtPath) + 1, A1$, "\"))
                
                ScanPath.Dir1.Path = ScanPath.txtPath
                
                'REM 1ST FOLDER
                D1$ = ScanPath.Dir1.List(0)
                
            
            End If
            
            'If Len(A1$) <> Len(D1$) Then Exit For
            If TRIG1 = True Then Exit For
        
        Next
    End If
    '------------------------------
    'OLD STYLE OF FIND FIRST FOLDER
    '------------------------------
        
        
        
     
    '------------------------------
    '1ST FOLDER PROTECTED
    '------------------------------
    ScanPath.Dir1.Path = ScanPath.txtPath
    D1$ = ScanPath.Dir1.List(0)
    
    Source_Path_Folder_Filter = D1$
    
    ScanPath.WindowState = vbNormal
    
    
    '----------------------------------------------------
    'QUESTIONS TO ASK
    '----------------------------------------------------
    'DONT DEL FROM 1ST FOLDER CHECK
    '----------------------------------------------------
    
    If ScanPath.Mnu_Repeat_Questions_Toggle.Checked = True Then
    
        'Has Menu Option 1 Been Looked At
        '------------------------------
        If Option_Mnu_Dont_Del_From_First_Source_Folder1 = False Then
        
            i = MsgBox("Option 1 of 3 Question" + vbCrLf + "Remember Menu Option 1 " + vbCrLf + vbCrLf + "The Next Sorted Level Path From Source is " + vbCrLf + vbCrLf + Source_Path_Folder_Filter + vbCrLf + vbCrLf + "EveryThing in This Folder -- and The Source Path Base Root are Read Only " + vbCrLf + "Double Check is Right Because Sort Order Not Same as Explorer" + vbCrLf + ScanPath.txtPath + vbCrLf + vbCrLf + "Without SubFolders - Root Path - Both Will Be Read Only Not Deleted" + vbCrLf + "HardCoded Filters Will Be Off" + vbCrLf + vbCrLf + "Do You Want This Option -----", vbYesNoCancel, ScanPath.Caption)
            If i <> vbYes And i <> vbNo Then MsgBox "Option 1 of 3 Note 1" + vbCrLf + "Not Any Action Taken", , ScanPath.Caption: Exit Sub
            
            If i = vbYes Then
                ScanPath.mnu_Dont_Del_From_First_Source_Folder.Checked = True
            Else
                ScanPath.mnu_Dont_Del_From_First_Source_Folder.Checked = False
            
            End If
        
        Else
            
            'OPTION HAS BEEN LOOKED AT HERE IS RESULT
            i = MsgBox("Option 1 of 3 Note 1" + vbCrLf + "The Next Sorted Level Path From Source is " + vbCrLf + vbCrLf + D1$ + vbCrLf + vbCrLf + "EveryThing in This Folder -- and The Source Path Base Root are Read Only " + vbCrLf + "Double Check is Right Because Sort Order Not Same as Explorer" + vbCrLf + vbCrLf + ScanPath.txtPath + vbCrLf + vbCrLf + "Without SubFolders - Root Path - Both Will Be Read Only Not Deleted" + vbCrLf + "HardCoded Filters Will Be Off", vbOKCancel, ScanPath.Caption)
            If i <> 1 Then MsgBox "Not Any Action Taken - Terminate Work", , ScanPath.Caption: Exit Sub
            
        
        End If
                
        
        If ScanPath.mnu_Dont_Del_From_First_Source_Folder.Checked = True Then
            i = MsgBox("Option 1 of 3 Note 2" + vbCrLf + "With This Option --> Don't Delete From Source Folder Selected" + vbCrLf + "Both Other HardCoded Filters Will Be Off" + vbCrLf + "#1 BlackList Remove From Checking and" + vbCrLf + "#2 Filter Check for CRC32 Compares But Don't Delete Some Filtering - Read Only - Indicated in Delete Status", vbOKCancel, ScanPath.Caption)
            If i <> 1 Then MsgBox "Option 1 of 3 " + vbCrLf + "Not Any Action Taken - Terminate Work": Exit Sub
        End If



        '--------------------------------
        'Has Menu Option 2 Been Looked At
        '--------------------------------
        'PARTIAL CRC CHECK
        '--------------------------------
        textv$ = "I Also Added It Dont Skip Large Files on Partital It Truncicates Them at 400 Meg" + vbCrLf
        textv$ = textv$ + "Lot Of Conditions Added When Paital Checks Both Whole File And Partial Conflict Asks Question or Find In Logg The Bigg CRC Is Not Acted On - You Get a Warning Message Box Info" + vbCrLf
        textv$ = textv$ + "More Logging of Deletions and Simulated" + vbCrLf
        textv$ = textv$ + "Add WhiteSpace Nulls and Space Files Protected Safe Read Omly" + vbCrLf
        textv$ = textv$ + "Added 3 Program Improvements 2015-Nov Took All-Night and 1 To Do Another to My Drive Mapping" + vbCrLf
        
        If Option_MNU_CRC_Check_After_First = False Then
        
            
            i = MsgBox("Option 2 of 3 Question" + vbCrLf + vbCrLf + "With This Option You Can CRC32 Check Partly From After the First 800 Bytes" + vbCrLf + "Good For Checking IDTAG Headers On MP3's and Media" + vcbrlf + "Also I Added Only to Check If They Are MP3 and WAV Files With This" + vcbrlf + "Also If they Are Over > 20K and Under < 400Meg Or the Fall Back On Full CRC32" + vbCrLf + "Also All Files Will Be Checked But Unless Media File Will Be Read Only - This is the Prioity Against Other Options Aside From Simulate" + vbCrLf + textv$ + vbCrLf + "Do You Want This Option -----", vbYesNoCancel, ScanPath.Caption)
            If i = 2 Then MsgBox "Option 2 of 3 Note 1" + vbCrLf + "Terminate the Dupe Process", , ScanPath.Caption: Exit Sub
            
            If i <> vbYes And i <> vbNo Then MsgBox "Option 2 of 3 Note 1" + vbCrLf + "Not Any Action Taken With This Option Will Be Left Not Checked and Check the Whole File - Which is Easier Because You Don't Have to Load The File into Memory in a String Variable", , ScanPath.Caption ': Exit Sub
            
            If i = vbYes Then
                ScanPath.MNU_CRC_Check_After_First.Checked = True
            Else
                ScanPath.MNU_CRC_Check_After_First.Checked = False
            
            End If
        
        Else
            
            'OPTION HAS BEEN LOOKED AT HERE IS RESULT
            i = MsgBox("Option 2 of 3 Note 2" + vbCrLf + "You Have Selected CRC32 Dupe Checking After the First 800 Bytes - For MP3 IDtag Skipping" + vbCrLf + textv$, vbOKCancel, ScanPath.Caption)
            If i <> 1 Then MsgBox "Option 2 of 3 " + vbCrLf + "Terminate the Dupe Process", , ScanPath.Caption: Exit Sub
            
        
        End If
                

        '--------------------------------
        'Has Menu Option 3 Been Looked At
        '--------------------------------
        'ALL READ ONLY
        '--------------------------------
        
        'TEST
        Check3.Value = vbChecked
        Option_Mnu_Simulate = True
        
        If Option_Mnu_Simulate = False Then

            i = MsgBox("Option 3 of 3 Question" + vbCrLf + "Do You Want Simulation While Safe All Ready Only", vbYesNoCancel, ScanPath.Caption)
            
            If i = 2 Then MsgBox "Option 3 of 3 Note 1" + vbCrLf + "Terminate the Dupe Process", , ScanPath.Caption: Exit Sub
            
            If i <> vbYes And i <> vbNo Then MsgBox "Option 3 of 3 Note 1" + vbCrLf + "Not Any Action Taken With This Option - Will Be Left Not Checked and Real Mode Delete's Will Happen if They Are a Match the Same and Taking Into Account Other Options", , ScanPath.Caption
            
            If i = vbYes Then
                ScanPath.Mnu_Simulate.Checked = True
            Else
                ScanPath.Mnu_Simulate.Checked = False
            
            End If
        
        Else
            
            'OPTION HAS BEEN LOOKED AT HERE IS RESULT
            i = MsgBox("Option 3 of 3 Note 2" + vbCrLf + "Yes In Safe Read Only Simulate Mode See the Results", vbOKCancel, ScanPath.Caption)
            If i <> 1 Then MsgBox "Terminate the Dupe Process", , ScanPath.Caption: Exit Sub
        End If
    End If
    '----------------------------------------------------
    'QUESTIONS OVER
    '----------------------------------------------------

End If


'If ScanPath.MNU_CRC_Check_After_First.Checked = True Then
'    MP3_HEADER_TAG_SIZE = &H12948
'End If



'ScanPath.chkRefreshListView = vbChecked

Call ScanPath.LOCK_WINDOW_UPDATE_SUB(False)

DoEvents
If mCancelScan = True Then GoTo ENDE


'ScanPath.Show

Reset
yy = "--"
    
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " -- Total"

'-------------------------------------
'RENAME A BATCH
'-------------------------------------
If 1 = 2 Then

    ScanPath.lblCount7 = "Rename A Batch"
    
    For we = 1 To ScanPath.ListView1.ListItems.Count
        DoEvents
        ScanPath.lblCount2 = Trim(Str(we))
        
        A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(we)
        
        D1$ = B1$
        e1$ = A1$
        
        If InStr(B1$, "HotKeyApp") > 0 Then
            D1$ = Replace(B1$, "HotKeyApp", "HotKey-")
            D1$ = Replace(D1$, ".jpg", "-App.jpg")
            D1$ = Replace(D1$, ".txt", "-App.txt")
        End If
        If InStr(B1$, "HotKeyScreen") > 0 Then
            D1$ = Replace(B1$, "HotKeyScreen", "HotKey-")
            D1$ = Replace(D1$, ".jpg", "-Screen.jpg")
            D1$ = Replace(D1$, ".txt", "-Screen.txt")
        End If
        
        D1$ = Replace(B1$, "  ", " ")
        D1$ = Replace(LCase(D1$), "eddie lancaster - ", "Eddie Lancaster_")
        
        
        
        If D1$ <> B1$ Then
            On Error Resume Next
            Name A1$ + B1$ As A1$ + D1$
            work = True
            On Error GoTo 0
        End If
    
    Next
    
    If work = True Then MsgBox "Done With Work Files" Else MsgBox "Done Files", , ScanPath.Caption
    work = -False
    
    For we = 1 To ScanPath.ListView1.ListItems.Count
        DoEvents
        ScanPath.lblCount2 = Trim(Str(we))
        
        A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(we)
        
        D1$ = B1$
        e1$ = A1$
    
    
        If A1$ <> olda1$ Then
            olda1$ = A1$
            If InStr(e1$, "#Bad Pix") > 0 Then
                e1$ = Replace(e1$, "#Bad Pix", "")
            End If
            If InStr(e1$, "HotKeyScreenCapture") > 0 Then
                e1$ = Replace(e1$, "HotKeyScreen", "HotKey-")
                e1$ = e1$ + "-Screen"
            End If
            If InStr(e1$, "HotKeyAppCapture") > 0 Then
                e1$ = Replace(e1$, "HotKeyApp", "HotKey-")
                e1$ = e1$ + "-App"
            End If
            If InStr(e1$, "\-Screen") > 0 Then
                e1$ = Replace(e1$, "\-Screen", "-Screen")
            End If
            If InStr(e1$, "\-App") > 0 Then
                e1$ = Replace(e1$, "\-App", "-App")
            End If
        
        End If
    
        'HotKeyScreenCapture_2009-07-13-Mon
        If InStr(A1$, "#Bad Pix") > 0 Then e1$ = A1$
        If e1$ <> A1$ Then
            Debug.Print A1$
            Debug.Print e1$
            xc = xc + 1
            ScanPath.lblCount3 = Trim(Str(xc))
            
            'Name Mid(A1$, 1, Len(A1$) - 1) As e1$
            Set F = Fs.GetFolder(A1$)
            On Error Resume Next
                F.Name = Mid(e1$, InStrRev(e1$, "\") + 1)
            'If Err.Number = 0 Then MsgBox "Done Err=0"
            On Error GoTo 0
            
            work = True
        End If
    
    Next
    If work = True Then MsgBox "Done With Work Folders", , ScanPath.Caption Else MsgBox "Done Folders", , ScanPath.Caption
    
    Exit Sub

End If
'-------------------------------------
'RENAME A BATCH END OF SUB ROUTINE
'-------------------------------------


DoEvents
If mCancelScan = True Then GoTo ENDE




'-------------------------------------------------------
'"Remove Our BlackScript 01 of 02"
'-------------------------------------------------------
'01 OF 02
'02 of 02 WHEN NOT PROTECT 1ST SOURCE FOLDER NOT

ScanPath.lblCount7 = "Remove BlackScript 01 of 02"
'-------------------------------------
'REMOVE TO BLACKLIST
'-------------------------------------

Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    DoEvents
    If we Mod 100 = 0 Then ScanPath.lblCount2 = Trim(Str(we))
    If mCancelScan = True Then Exit For
    
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)

    XGO = 1
    
    'DONT DELETE FROM MASTER AND NEVER CRC DUPE WITH OTHER FOLDERS
    'DONT DELETE ANY FROM THESE FOLDERS AND LEAVE BE ANY THE DUPE ELSEWHERE WITHIT
    
    'IN MY DOCUMENTS 2006
'    If InStr(A1$, "# COPY FAMILY 1991-2006") > 0 Then xgo = 0
'    If InStr(A1$, "\0 00 ART WORK IN BOUND\AUTOGRAPHS & MOVIE POSTERS\# TEST") > 0 Then xgo = 0
'    If InStr(A1$, "\VB6\VB6 Icons") > 0 And InStr(A1$, " - Folders") > 0 Then xgo = 0
'    If InStr(A1$, "T:\DSC\# ME") > 0 Then xgo = 0
'    If InStr(A1$, "T:\DSC\# Docus Proofs Texts") > 0 Then xgo = 0
'    If InStr(A1$, "D:\#\#E") > 0 Then xgo = 0
    'If InStr(LCase(A1$), LCase("\I Explorer Web Pages")) > 0 Then xgo = 0
    'If InStr(LCase(A1$), LCase("\My Webs")) > 0 Then xgo = 0
    
    
    
    '----------------------------------------------------------------------
    'THIS QUITE SPECIAL NOW AS REALLY DEL ALL WHEN PROTECT SOURCE FOLDER ON
    
    '--------------------------------------------
    'NOT ANY OF THESE AS WE WANT THE DUPE TO SHOW
    'WHERE ORGINAL CAME FROM
    'If InStr(LCase(A1$), LCase("\AUTOPIX 00")) > 0 Then XGO = 0
    '--------------------------------------------
    
    
    If InStr(A1$, "\z-Dupe's Deleted Script") > 0 Then XGO = 0
    If InStr(A1$, "\_gsdata_") > 0 Then XGO = 0
    
    
    '---------------------------------------------
    If XGO = 0 Then
        ScanPath.ListView1.ListItems.Remove (we)
    End If
    '---------------------------------------------

Next

Call ScanPath.LOCK_WINDOW_UPDATE_SUB(0)
DoEvents
Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)


If mCancelScan = True Then GoTo ENDE




'-------------------------------------------------------
'"Remove BlackScript"  02 of 02 WHEN NOT PROTECT 1ST SOURCE FOLDER NOT
'-------------------------------------------------------

If ScanPath.mnu_Dont_Del_From_First_Source_Folder.Checked = False Then

    ScanPath.lblCount7 = "Remove BlackScript 02 of 02"
    '-------------------------------------
    'REMOVE TO BLACKLIST
    '-------------------------------------
    
    Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)
    
    For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
        DoEvents
        If we Mod 100 = 0 Then ScanPath.lblCount2 = Trim(Str(we))
        A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(we)
    
        XGO = 1
        
        'DONT DELETE FROM MASTER AND NEVER CRC DUPE WITH OTHER FOLDERS
        'DONT DELETE ANY FROM THESE FOLDERS AND LEAVE BE ANY THE DUPE ELSEWHERE WITHIT
        
        If InStr(A1$, "# COPY FAMILY 1991-2006") > 0 Then XGO = 0
        If InStr(A1$, "\0 00 ART WORK IN BOUND\AUTOGRAPHS & MOVIE POSTERS\# TEST") > 0 Then XGO = 0
        If InStr(A1$, "\VB6\VB6 Icons") > 0 And InStr(A1$, " - Folders") > 0 Then XGO = 0
        If InStr(A1$, "T:\DSC\# ME") > 0 Then XGO = 0
        If InStr(A1$, "T:\DSC\# Docus Proofs Texts") > 0 Then XGO = 0
        If InStr(A1$, "D:\#\#E") > 0 Then XGO = 0
        
        'IN MY DOCUMENTS 2006
        If InStr(LCase(A1$), LCase("\I Explorer Web Pages")) > 0 Then XGO = 0
        If InStr(LCase(A1$), LCase("\My Webs")) > 0 Then XGO = 0
        
        '--------------------------------------------
        'NOT ANY OF THESE AS WE WANT THE DUPE TO SHOW
        'WHERE ORGINAL CAME FROM
        If InStr(LCase(A1$), LCase("\AUTOPIX 00")) > 0 Then XGO = 0
        '--------------------------------------------
        
        If InStr(A1$, "\_gsdata_") > 0 Then XGO = 0
        If XGO = 0 Then
            ScanPath.ListView1.ListItems.Remove (we)
        End If
    
    Next
    
    Call ScanPath.LOCK_WINDOW_UPDATE_SUB(0)
    DoEvents
    Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)
    
End If

If mCancelScan = True Then GoTo ENDE

'-------------------------------------------------------------
'TEST REMOVE SMALL FILE SO CAN WORK LARGE MEMORY OVERFLOW
'-------------------------------------------------------------
'TAKE TO LONG THIS WAY
'SEEM READ IN QUICKER THAN REMOVE
'-------------------------------------------------------------
'ANOTHER METHOD AS READ IN
'-------------------------------------------------------------

If Test = Test + 10 Then

    ScanPath.lblCount7 = "Test Remove Small"
    '-------------------------------------
    'REMOVE SMALL
    '-------------------------------------
    
    Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)
    
    For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
        DoEvents
        If we Mod 100 = 0 Then ScanPath.lblCount2 = Trim(Str(we))
        'A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        'A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        X1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
    
        XGO = 1
        If Val(X1$) < 100000 Then XGO = 0
        
        If XGO = 0 Then
            ScanPath.ListView1.ListItems.Remove (we)
        End If
    
    Next
    
    Call ScanPath.LOCK_WINDOW_UPDATE_SUB(0)
    DoEvents
    Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)
    
End If

If mCancelScan = True Then GoTo ENDE






'-------------------------------------------------------------
'"Format The Sizes TO BE ABLE FOR Sorting"
'-------------------------------------------------------------
ScanPath.lblCount7 = "Format The Sizes For Sorting"

'PUT THE SIZES SORTABLE
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " -- Total"
NAUGHT_STRING = String$(Len(Trim(Replace(ScanPath.lblCount6, " Biggest Scanning FSize", ""))), "0")
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    DoEvents
    If we Mod 100 = 0 Then ScanPath.lblCount2 = Trim(Str(we))
    If mCancelScan = True Then Exit For
    
    D1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
    
    RS = Format(Val(D1$), NAUGHT_STRING)
    ScanPath.ListView1.ListItems.Item(we).SubItems(3) = RS

Next
'---------------------------------------------------------
'---------------------------------------------------------
'DONE
'---------------------------------------------------------

Call ScanPath.LOCK_WINDOW_UPDATE_SUB(0)

'ScanPath.ListView1.Refresh
'DoEvents

'LAST 3 COLUMNS -- SIZE AND CRC AND DELETE STATUS
Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(3))
Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(4))
'Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(5))
'DoEvents


'Shell "explorer.exe /select," + A1$ + B1$, vbNormalFocus


Call ScanPath.LOCK_WINDOW_UPDATE_SUB(0)
DoEvents
Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)


DoEvents
If mCancelScan = True Then GoTo ENDE



'----------------------------------------------------
'REMOVE NAUGHT SIZE AND RAR'S
'REMOVE RAR'S
'----------------------------------------------------

ScanPath.lblCount7 = "If Naught && Check Rename DOT Ext"

'ScanPath.lblCount7 = "Remove WildCard"

Rdclip2 = "Delete Naught Size - User Intervention Applied About Double Dot - Nov 2015 --------" + vbCrLf
Rdclip2 = Rdclip2 + "Remove From Script If Naught An Then Check and Rename DOT Ext" + vbCrLf
Rdclip2 = Rdclip2 + "Remove From Script If Header Tag Mode and if Less than Header Size" + vbCrLf
Rdclip2 = Rdclip2 + "-------------------------------------------------" + vbCrLf

XS2 = Len(Rdclip2)
'          ----------
Dim XDone
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    'ScanPath.lblCount2 = Trim(Str(WE))
    If we Mod 100 = 0 Then ScanPath.lblCount2 = Trim(Str(we)): DoEvents
    
    
    If mCancelScan = True Then Exit For
    X1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(3)

    XDone = 0

    If Val(X1$) = 0 Then
        A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(we)
        G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
        
        'CHECK IT EXISTS AGAIN TO BE SURE
        If Fs.FileExists(A1$ + G1$) = True Then
            Set F = Fs.getfile(A1$ + G1$)
            dfname = GetLongName(A1$ + G1$)
            shortname = F.ShortPath
            'ShowProperties
            
            extn = Fs.GetExtensionName(A1$ + G1$)

            'Bug in microsoft wont delete with filename ends with dot
            'But not a extension
            'Rename to sort problem workaround
            If Mid(B1$, Len(B1$), 1) = "." And extn = "" Then
                'Name A1$ + G1$ As A1$ + G1$ + ".Renamed"
                'BETTER
                If Mid(B1$, Len(B1$), 1) = "." Then Mid(B1$, Len(B1$), 1) = " ": B1$ = Trim(B1$)
                
                i = MsgBox("DOUBLE DOT ISSUE " + vbCrLf + A1$ + B1$ + vbCrLf + A1$ + G1$ + vbCrLf + vbCrLf + A1$ + B1$ + ".Renamed" + vbCrLf + vbCrLf + "CAN'T DELETE THESE FILES UNLESS RENAMED" + vbCrLf + "Do You Want to Rename It", vbYesNo, ScanPath.Caption)
                
                If i = vbYes Then
                    Name A1$ + G1$ As A1$ + B1$ + ".Renamed"
                    Rdclip2 = Rdclip2 + A1$ + G1$ + vbCrLf
                    Rdclip2 = Rdclip2 + A1$ + B1$ + ".Renamed" + vbCrLf
                
                Else
                    Rdclip2 = Rdclip2 + A1$ + G1$ + vbCrLf
                    Rdclip2 = Rdclip2 + A1$ + B1$ + ".Renamed" + vbCrLf
                    Rdclip2 = Rdclip2 + "User Objected to Rename" + vbCrLf
                
                End If
                'Msgbox_alert = True
                
                
                'ScanPath.ListView1.ListItems.Item(WE).SubItems(2) = G1$ + ".Renamed"
                'Set F = Fs.getfile(A1$ + G1$)
                'shortname = F.ShortPath
                'extn = Fs.GetExtensionName(A1$ + G1$)
                'ScanPath.ListView1.ListItems.Item(WE) = F.Name
                
                'WITCH DUNK SHALL I DELETE FIRST TO TEST
                'Fs.DeleteFile A1$ + G1$ + ".txt", True
                'Rdclip2 = Rdclip2 + A1$ + G1$ + vbCrLf
                'Rdclip2 = Rdclip2 + A1$ + B1$ + ".Renamed" + vbCrLf
            Else
                On Error Resume Next
                '2015-08-12 Don't Delete Them Without Checking
                'Fs.DeleteFile A1$ + G1$, True
                If Err.Number > 0 Then
                    On Error GoTo 0
                    'Fs.DeleteFile shortname, True
                End If
                On Error GoTo 0

                'Rdclip2 = Rdclip2 + A1$ + G1$ + vbCrLf
            End If


            On Error Resume Next
            'RmDir A1$
            On Error GoTo 0
        End If
        ScanPath.ListView1.ListItems.Remove (we)
        'XDone = 1
    End If
Next
'----------------------------------------------------
'REMOVE NAUGHT SIZE AND RAR'S
'REMOVE RAR'S
'----------------------------------------------------
'DONE
'----------------------------------------------------


'----------------------------------------------------
'LOGG WRITING
'----------------------------------------------------
ScanPath.Text1.Text = Filename_Script_FILE
FR4 = FreeFile
Open Filename_Script For Append As #FR4
    
    STRING_LEN = 55
    Print #FR4, String(STRING_LEN, "-")
    Print #FR4, "LOGG STARTER ---------- " + Format(TVarNow, "DD MMM YYYY HH:MM:SS")
    Print #FR4, "Folder Given ---------- " + ScanPath.txtPath
    Print #FR4, "1ST FOLDER FILTERED  -- " + Source_Path_Folder_Filter
    Print #FR4, String(STRING_LEN, "-")
    
    If ScanPath.Mnu_Simulate.Checked = True Then
        Print #FR4, "Option Set --- SIMULATE ONLY ALL READ ONLY"
    End If
    If ScanPath.mnu_Dont_Del_From_First_Source_Folder.Checked = True Then
        Print #FR4, "Option Set --- DONT DEL FROM 1ST FOLDER CHECK"
    End If
    If ScanPath.MNU_CRC_Check_After_First.Checked = True Then
        Print #FR4, "Option Set --- To CRC Check after 1st " + Hex$(MP3_HEADER_TAG_SIZE) + " Hex Bytes for Media Tag Heading Like MP3"
    End If
    If ScanPath.MNU_Check_FOR_NULL_STRING.Checked = True Then
        Print #FR4, "Option Set --- Check FOR NULLSTRING FILES ALL SAME BYTE CHARACTOR"
    End If
    
    Print #FR4, String(STRING_LEN, "-")

    Print #FR4, "STATUS -- CRC Dupe & Source Folder Protected"
    'Print #FR4, "NOT BEING REPORTED ANYMORE DUE TO OVER LOGGING           --"
    Print #FR4, String(STRING_LEN, "-")
Close #FR4


'----------------------------------------------------
'LOGG WRITING STEP 2 MAYBE
'----------------------------------------------------

If mCancelScan = True Then GoTo ENDE

'GOT SOME RESULTS
If Len(Rdclip2) > XS2 Then
    
    FR1 = FreeFile
    Open Filename_Script For Append As #FR1
        IVAR = 26
        IVAR = 55
        Print #FR1, String(IVAR, "-")
        Print #FR1, Format(Now, "YYYY-MM-DD HH:MM:SS") + " ------"
        Print #FR1, String(IVAR, "-")
        Print #FR1, "RESULTS GIVEN FROM RENAME OVER DOUBLE DOT WITHOUT A EXTENSION ISSUE"
        Print #FR1, String(IVAR, "-")
        Print #FR1, Rdclip2;
        Print #FR1, String(IVAR, "-")
    Close #FR1
    LOGG_INFO_DONE = True
End If

'----------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------



Call ScanPath.LOCK_WINDOW_UPDATE_SUB(0)
DoEvents
Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)
DoEvents
If mCancelScan = True Then GoTo ENDE



ScanPath.lblCount7 = "Remove WildCards None to Do"

If 1 = 2 Then
    For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
        'DoEvents
        If we Mod 100 = 0 Then ScanPath.lblCount2 = Trim(Str(we)): DoEvents
        
        X1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
    
        'XDone = 1
    
        If XDone = 0 Or True = False Then
        '    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
        '    B1$ = ScanPath.ListView1.ListItems.Item(WE)
        '
        '    If InStr(LCase(B1$), ".rar") > 0 And XDONE = 0 Then
        ''        ScanPath.ListView1.ListItems.Remove (WE)
        '    End If
        '    If InStr(UCase(A1$), "\ME") > 0 And XDONE = 0 Then
        '        ScanPath.ListView1.ListItems.Remove (WE)
        '    End If
        End If
    
    Next
End If

If ScanPath.MNU_CRC_Check_After_First.Checked = True Then

    ScanPath.lblCount7 = "Remove Tag Header MP3 Low Size"

    For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
        
        'DoEvents
        If we Mod 100 = 0 Then ScanPath.lblCount2 = Trim(Str(we)): DoEvents
        If mCancelScan = True Then Exit For
        
        'G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
        'A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(we)
        'FILESIZE
        X1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
        FileSize = Val(X1$)
        EXTNAME = ""
        '----------------------------------------------------------------
        'DO WHILE MP3 AND MEDIA FILES IN CHECK AFTER TAG HEADER 1ST BLOCK
        If ScanPath.MNU_CRC_Check_After_First.Checked = True Then
            EXTNAME = ""
            If InStr(B1$, ".") > 0 Then
                EXTNAME = UCase(Mid(B1$, InStrRev(B1$, ".")))
            End If
            If EXTNAME <> "" Then
                If InStr(" .MP3 .WAV .AVI .MPG .MP4", EXTNAME) = 0 Then
                        EXTNAME = ""
                    Else
                        'A = A
                End If
            End If
        End If
        '----------------------------------------------------------------
        EXCePT = True
        If FileSize <= MP3_HEADER_TAG_SIZE Then EXCePT = False
        
        If EXTNAME = "" Or EXCePT = False Then
            ScanPath.ListView1.ListItems.Remove (we)
        End If
        '----------------------------------------------------------------
    
    Next
End If

'ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " - After 0 and *.Rar"
ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " - After Removing"


Call ScanPath.LOCK_WINDOW_UPDATE_SUB(0)
DoEvents
Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)
DoEvents
If mCancelScan = True Then GoTo ENDE





'----------------------------------------------------
'SORT ON SIZE
'----------------------------------------------------

ScanPath.lblCount7 = "Sort on Size"
DoEvents

'SORT ON SIZE
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

'----------------------------------------------------
'SORT ON SIZE
'----------------------------------------------------
'DONE
'----------------------------------------------------

'LAST 3 COLUMNS -- SIZE AND CRC AND DELETE STATUS
Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(3))
Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(4))
'Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(5))


DoEvents
Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)
DoEvents
If mCancelScan = True Then GoTo ENDE



'------------------------------------------
'SORT SIZE DONE NOW GOT BIGGEST WORKER FILE
'------------------------------------------
BIGGEST_FILESIZE_WORKER = Val(ScanPath.ListView1.ListItems.Item(ScanPath.ListView1.ListItems.Count).SubItems(3))
ScanPath.lblCount_BIGGEST_WORKING_SIZE = Format(BIGGEST_FILESIZE_WORKER, "###,###,###,##0") + " Biggest Worker FSize"
'------------------------------------------
'------------------------------------------



'------------------------------------------------------------
'REMOVE NON DUPE SIZE'S ONE ABOVE AND ONE BELOW pass 01 of 02
'------------------------------------------------------------

ScanPath.lblCount7 = "Remove Non Dupes"

TOPWE = ScanPath.ListView1.ListItems.Count
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    
    If mCancelScan = True Then Exit For
    
    If we - 1 >= 1 Then
        X1$ = ScanPath.ListView1.ListItems.Item(we - 1).SubItems(3)
    Else
        X1$ = "X"
    End If
    X2$ = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
    If we + 1 > TOPWE Then
        X3$ = "X"
    Else
        X3$ = ScanPath.ListView1.ListItems.Item(we + 1).SubItems(3)
    End If


    If X2$ = X1$ Or X2$ = X3$ Then
        RESULT = True
    Else
        ScanPath.ListView1.ListItems.Remove (we)
        NonDupeRemove = NonDupeRemove + 1
        TOPWE = ScanPath.ListView1.ListItems.Count
        'ScanPath.lblCount4 = Trim(Str(NonDupeRemove)) + " Non Dupes Removed"

        DoEvents
    End If
    
    If we Mod 100 = 0 Then ScanPath.lblCount3 = Trim(Str(we)) + " + " + Trim(Str(NonDupeRemove)) + " NonDupes Removed"
    
    'If WE Mod 10 = 0 Then ScanPath.lblCount3 = Trim(Str(WE)): DoEvents
    'DoEvents

Next

ScanPath.lblCount3 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " + " + Trim(Str(NonDupeRemove)) + " NonDupes Removed"




Call ScanPath.LOCK_WINDOW_UPDATE_SUB(0)

DoEvents
If mCancelScan = True Then GoTo ENDE


On Error Resume Next
'-------------------------------------
'PROBLEM WITH ACCESS LOCKED FILES
Reset
Close
'-------------------------------------
If Err.Number > 0 Then Stop


'-------------------------------------------------
'ONE MORE OF PROCESS THESE CODE TO TAKE ASJUSTMENT
'-------------------------------------------------
'SORT SIZE DONE NOW GOT BIGGEST WORKER FILE
'-------------------------------------------------
BIGGEST_FILESIZE_WORKER = Val(ScanPath.ListView1.ListItems.Item(ScanPath.ListView1.ListItems.Count).SubItems(3))
ScanPath.lblCount_BIGGEST_WORKING_SIZE = Format(BIGGEST_FILESIZE_WORKER, "###,###,###,##0") + " Biggest Worker FSize"
'-------------------------------------------------
'-------------------------------------------------



'------------------------------------------------
' Main Routine - Main Form - Main Code
'------------------------------------------------




ScanPath.lblCount7 = "Main Routine"

we3 = ScanPath.ListView1.ListItems.Count

'LAST 3 COLUMNS -- SIZE AND CRC AND DELETE STATUS
Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(3))
Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(4))
'Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(5))

Dim DeleteError, ReadError

Dim XXCoulmn, TT5$

ScanPath.lblCount9.Caption = Format(BIGGEST_FILESIZE, "####,###,###,###") + " Biggest Scanning FSize"

TOTAL_COUNT = ScanPath.ListView1.ListItems.Count
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    DoEvents
    ScanPath.lblCount4 = Str(we) + " of" + Str(TOTAL_COUNT)
    If mCancelScan = True Then Exit For
    
    If ScanPath.WindowState = vbMaximized Or 1 = 1 Then
        ScanPath.ListView1.ListItems(we).Selected = True
        ScanPath.ListView1.SelectedItem.EnsureVisible
    End If

    '--------------------------
    DoEvents
    If CommandPause = True Then
        Do
            DoEvents
            Sleep 50
            DoEvents
        Loop Until CommandPause = False Or CommandExit = True
    End If
    If CommandExit = True Then Exit For
    '--------------------------

    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    FileSize = Val(ScanPath.ListView1.ListItems.Item(we).SubItems(3))
    ScanPath.lblCount6 = Format(FileSize, "####,###,###,###") + " File Size @"
    
    On Error Resume Next

    Err.Clear
    
    If OldFSize <> FileSize Then
        yy = ""
        'DUPE_COMPARE_AGAINST = "": DUPE_COMPARE_AGAINST_COUNT = 0
        ScanPath.List1.Clear
    End If
    
    OldFSize = FileSize
    
    WXHEXXX$ = ""
    WXHEX$ = ""
    
    'LENGHT OF DATE STRING
    'ERRORLOGGS2 = String$(19, "-") + vbCrLf
    ERRORLOGGS2 = ""
    ERRORLOGGS2 = ERRORLOGGS2 + Format(Now, "YYYY-MM-DD HH-MM-SS") + vbCrLf
    ERRORLOGGS2 = ERRORLOGGS2 + A1$ + B1$ + vbCrLf
    ERRORLOGGS2 = ERRORLOGGS2 + A1$ + G1$ + vbCrLf
    ERRORLOGGS2 = ERRORLOGGS2 + "FileSize -- " + Trim(Str(FileSize)) + vbCrLf
    
    ERRORLOGGS1 = ""
    ERRORLOGGS3 = ""
    ResultDel1 = ""
    
    If FileSize = 0 Then Stop
'    If FileSize > 0 Then
    
    'C1 = B1$ = "2008-11-23 - Pearsall - Squat Rocking 3.mp3"
    'If Fs.FileExists(A1$ + G1$) = True And C1 = True Then
    
    If Fs.FileExists(A1$ + G1$) = True Then 'And C1 = True Then
        
        EXTNAME = ""
        '----------------------------------------------------------------
        'DO WHILE MP3 AND MEDIA FILES IN CHECK AFTER TAG HEADER 1ST BLOCK
        If ScanPath.MNU_CRC_Check_After_First.Checked = True Then
            EXTNAME = ""
            If InStr(B1$, ".") > 0 Then
                EXTNAME = UCase(Mid(B1$, InStrRev(B1$, ".")))
            End If
            If EXTNAME <> "" Then
                If InStr(" .MP3 .WAV .AVI .MPG .MP4", EXTNAME) = 0 Then
                        EXTNAME = ""
                    Else
                        A = A
                End If
            End If
        End If
        '----------------------------------------------------------------
        ERRORLOGGS3 = ""
        
        EXCePT = False
        If ScanPath.MNU_CRC_Check_After_First.Checked = True Then EXCePT = True
        If FileSize <= MP3_HEADER_TAG_SIZE Then EXCePT = False
        If EXTNAME = "" Then EXCePT = False
        
        If EXCePT = True Then
                    
            
            '1 GIG
            '1024 ^ 3
            'MEG
            '1024 ^ 2

'            BUFFERSIZE = FileSize
            'MaxFileSizeCRCCheck = (1024 ^ 2) * 100
            
            '1 GIG
'            If FileSize >= MaxFileSizeCRCCheck Then
'                BUFFERSIZE = MaxFileSizeCRCCheck
'            End If
            
            
            WXHEX$ = m_CRC.CalculateFile_Part(A1$ + G1$, MP3_HEADER_TAG_SIZE)
            'RETURNS HEX OR DECIMAL BUT AFTER HEX COMMAND 8 DIGITS
            'OF NAUGHT ARE REDUCED TO 1
            
            'TT5$ = Space$(BUFFERSIZE)
            
'            FR5 = FreeFile
'            Open A1$ + G1$ For Binary As #FR5
'                On Error Resume Next
'                Err.Clear
'                Get FR5, MP3_HEADER_TAG_SIZE, TT5$
'                If Err.Number > 0 Then
'                    MsgBox "Error Read String From File " + vbCrLf + Err.Description
'                End If
'
'            Close #FR5
            
            'Clipboard.Clear
            'Clipboard.SetText TT5$
            'mid(tt5$,7000,50)
            
            
'            DoAbort = False
'            If StripSpaces(TT5$) = "" Then
'                DoAbort = True
'                ERRORLOGGS3 = ERRORLOGGS3 + "Partial File Is a Void of Spaces - Safe Won't Delete" + vbCrLf
'
'            End If
'
'            If StripNulls(TT5$) = "" Then
'                DoAbort = True
'                ERRORLOGGS3 = ERRORLOGGS3 + "Partial File Is a Void of NullString - Safe Won't Delete" + vbCrLf
'            End If
            
            
            'WXHEX$ = Hex(m_CRC.CalculateString(TT5$))
            'Debug.Print WXHEX$
            
            'NOTHING THE STRING AFTER TO SAVE MEMORY
            'TT5$ = "*"
            
            'ERRORLOGGS3 = ERRORLOGGS3 + "Partial CRC Performed" + vbCrLf
            
'            If FileSize >= MaxFileSizeCRCCheck Then
'                ERRORLOGGS3 = ERRORLOGGS + "Partial CRC Compare Is Truncated Because Bigger than" + Str(Int(MaxFileSizeCRCCheck / 1024 ^ 2)) + " Meg" + vbCrLf
'            End If
            
            If Trim(WXHEX$) <> "" And WXHEX$ <> "00000000" Then
                If Len(yy) > 3 Then
                    If InStr(yy, WXHEX$) > 0 Then
                        ERRORLOGGS3 = ERRORLOGGS3 + "Partial CRC IS a Duplicate" + vbCrLf
                        
                        
                        If COMMAND_WINSTATE_RESIZE.BackColor <> Label_COMMAND_WINSTATE_RESIZE_COLOR.BackColor Then
                            ScanPath.WindowState = vbNormal
                        End If

                    Else
                        'ERRORLOGGS3 = ERRORLOGGS3 + "Partial CRC Not Duplicate" + vbCrLf
                    End If
                    'FIND A CHECK AGAINST PREVIOUS FILE SAME SIZE FIRST BEFORE
                    'RESULT OF COMPARE CAN BE MADE
                
                End If
            Else
                If WXHEX$ = "00000000" Then
                    ERRORLOGGS3 = ERRORLOGGS3 + "CRC RETURNED RESULT NAUGHT 00000000 ERROR" + vbCrLf
                End If
                
                'NOT A RESULT FROM CRC CALC
                'Stop
            End If
        End If
        
        
        
        
        EXCePT = False
        If ScanPath.MNU_CRC_Check_After_First.Checked = True Then EXCePT = True
        
        
        If WXHEX$ = "" And EXCePT = False Then
            
            'CRC NOT DONE PARTIAL SO DO WHOLE FILE HERE
            'BECAUSE MAYBE FILE TOO SMALL
            'OVERRIDE - WANT - USE - NOT CHECK SMALL FILE ON PARTIAL
            'QUICKER
            
'            WXHEX$ = Hex(m_CRC.CalculateFile(A1$ + G1$))
'        Else

            
            
            
            WXHEX$ = m_CRC.CalculateFile(A1$ + G1$)
            
            CalculateFile_ERROR = ""
            If InStr(WXHEX$, "Err.Description") > 0 Then
                MsgBox WXHEX$
                If IsIDE = True Then Stop
                CalculateFile_ERROR = WXHEX$
            End If
            
            Source_Var = False
            
            If InStr(yy, WXHEX$) > 0 And WXHEX$ <> "00000000" And CalculateFile_ERROR = "" Then
                'GOT A MATCH
                ERRORLOGGS3 = ERRORLOGGS3 + "WHOLE FILE IS A CRC32 Duplicate" + vbCrLf
                
                If ScanPath.COMMAND_WINSTATE_RESIZE.BackColor <> ScanPath.Label_COMMAND_WINSTATE_RESIZE_COLOR.BackColor Then
                    ScanPath.WindowState = vbNormal
                End If
            
            Else
                If WXHEX$ = "00000000" Then
                    ERRORLOGGS3 = ERRORLOGGS3 + "CRC RETURNED RESULT NAUGHT 00000000 ERROR" + vbCrLf
                End If
                If WXHEX$ = "00000000" Then
                    ERRORLOGGS3 = ERRORLOGGS3 + "CRC RETURNED RESULT NAUGHT 00000000 ERROR" + vbCrLf
                End If
                
                If InStr(WXHEX$, "Err.Description") > 0 Then
                    ERRORLOGGS3 = ERRORLOGGS3 + "CRC RETURNED ERROR CODE DESCRIPTION" + vbCrLf + Mid(WXHEX$, 20)
                End If
                
                
                'NOT GOT A MATCH
                'ERRORLOGGS3 = ERRORLOGGS3 + "WHOLE FILE NOT A CRC32 Duplicate" + vbCrLf
            
                'ScanPath.WindowState = vbNormal
                'MsgBox ERRORLOGGS3, vbOKOnly, ScanPath.Caption
            
            End If
        
        
            we2 = we2 + 1
                
            'SCRIPT OPT OUT
            'DONT DELETE FROM MASTER FOLDER BUT DO FOR DUPES ELSEWHERE WITHIT
            XGO = 1
'                If InStr(A1$, "\03 MICROSOFT\# VB 6\01 VB6 An MSDN 01") > 0 Then xgo = 0
'                If InStr(A1$, "\03 MICROSOFT\# VB 6\01 VB6 An MSDN 02") > 0 Then xgo = 0
'                If InStr(A1$, "\03 MICROSOFT\# OP SYS\Test\00 Windows Xp Pro") > 0 Then xgo = 0
'                If InStr(A1$, "D:\0 00 MUSIC ---") > 0 Then xgo = 0
''                If InStr(A1$, "D:\# MY DOCS\# 01 My Documents\Downloads\") > 0 Then xgo = 0 ': MsgBox "CHECK"
'                If InStr(A1$, "D:\#0 1 INSTALLATIONS\Installation programs\") > 0 Then xgo = 0
'                If InStr(A1$, "D:\0 00 MUSIC ---\") > 0 Then xgo = 0

'                If InStr(A1$, "D:\# MY DOCS\") > 0 Then xgo = 0
            
            For WE5 = 1 To UBound(SCAN_REFINE2)
                If InStr(UCase(A1$), SCAN_REFINE2(WE5)) > 0 Then XGO = 0: Exit For
            Next
                    
                
                    
                    
            
            Source_Var = False
            If ScanPath.mnu_Dont_Del_From_First_Source_Folder.Checked = True Then

                XGO = 1
'                If InStr("*" + A1$, Source_Path_Folder_Filter) > 0 Then Stop
                If InStr("*" + A1$, Source_Path_Folder_Filter + "\") > 0 Then
                    XGO = 0
                    Source_Var = True
                                    
                End If
            Else
                'GOT ONE TO DELETE -- CHECK IS NOT HERE
'                A = A
            End If
            
            
            
            DoEvents
            If CommandPause = True Then
                Do
                    DoEvents
                    Sleep 500
                Loop Until CommandPause = False Or CommandExit = True
            End If

            If CommandExit = True Then Exit For
            
            
                            
            'OPPOSITE ---- XXXXX BEWARE THIS ONE ----------------
'                If InStr(A1$, "T:\DSC Z COMPARE1\") = 0 Then xgo = 0
            
            
            'D:\# MY DOCS\# 01 My Documents\Downloads\
'                If InStr(A1$, "\DSC\# Docus Proofs Texts\") > 0 Then xgo = 0
            
            'If InStr(A1$, "T:\DSC\# ME") > 0 Then xgo = 0
            'If InStr(A1$, "T:\DSC\# Docus Proofs Texts") > 0 Then xgo = 0
            
            DeleteError = False
            ReadError = False
            
            Debug.Print Err.Description
            
            If Err.Number > 0 Then ReadError = True
            
            
            '--------------------------------------------------------
            'BIG MOMENT -- DELETE LOOKING IMINANT -------------------
            '--------------------------------------------------------
            UpForDel = False
            If XGO = 1 And Err.Number = 0 And CalculateFile_ERROR = "" Then
                'Err.Clear
                wex4 = wex4 + 1
                'DoAbort = Not if Nullstring or Space-String
                UpForDel = True
                If ScanPath.Mnu_Simulate.Checked = False And DoAbort = False Then
                    
                    Err.Clear
                    
                    i = ShellFileDelete(A1$ + G1$, True)
                    
                    '-----------------------------------
                    'DO MORE DELETE IF FAIL
                    'NOT ANYMORE
                    '-----------------------------------
                    If 1 = 2 Then
                        Set F = Fs.getfile(A1$ + G1$)
                        If Err.Number = 0 Then
                            tt = F.Name
                            
                            F.Delete True
                            Set F = Nothing
                        
                        Else
                            'need to say error fault
                            'MsgBox "Error Get File " + A1$ + B1$ + vbCrLf + Err.Description
                        End If
                        
                        Set F = Nothing
                        
                        If Err.Number > 0 Then
                            Err.Clear
                            Fs.DeleteFile (A1$ + G1$), True
                        
                        End If
                        
                        
                        If Err.Number > 0 Then
                            'Close ': Reset
                            Close
                            Err.Clear
                            Kill A1$ + G1$
                            Reset
                        End If
                    End If
                    '-----------------------------------
                    'DO MORE DELETE IF FAIL
                    'NOT ANYMORE
                    '-----------------------------------
                    
                    
                    
                    
                    
                    '-------------------------------------------
                    ResultDel2 = ""
                    If Err.Number = 0 And i = True Then
                        FileSizeDelTotal = FileSize + FileSizeDelTotal
                        If FileSizeDelTotal < 1024 ^ 2 Then ''MaxFileSizeCRCCheck Then
                            ScanPath.lblCount10 = Format(FileSizeDelTotal, "###,###,###,##0") + " Byte Deleted Total"
                        Else
                            ScanPath.lblCount10 = Format(FileSizeDelTotal / 1024 ^ 2, "###,###,###,###.0") + " Meg Deleted Total"
                        End If
                        KillDel = KillDel + 1
                        ResultDel1 = "Del Success --- ---"
                        ResultDel2 = "+"
                    Else
                        ResultDel1 = "Del Fail ATTEMPT --"
                        KillDelFAIL = KillDelFAIL + 1
                        ResultDel2 = "-"
                    End If
                    
                    '----------------------------------------------------
                    'SAVE THE HEADER INFO OF LABLE'S USED TOTALS
                    'USUAL ONLY WHEN A WRITE DONE WHEN CHANGE AFTER A DEL
                    '----------------------------------------------------
                    Call SAVE_RD_CLIP
                    
                                    
                    'ScanPath.ListView2.ListItems.Add.SubItems(1) = A1$
                    'ScanPath.ListView2.ListItems.Add = B1$
                    If FLAG_DO_IN_CHUNK_OUT_OF_MEMORY = True Then
                        ResultDEL_STATUS3 = "--Chunk Mode"
                    End If
                    
                    If ResultDel2 = "+" Then
                        ResultDEL_STATUS2 = "+Del"
                    Else
                        ResultDEL_STATUS2 = "-Fail"
                    End If
                    
                    Set LV = ScanPath.ListView2.ListItems.Add(, , ResultDEL_STATUS2 + ResultDEL_STATUS3)

                    LV.SubItems(1) = B1$
                    LV.SubItems(2) = A1$
                        
'                        UpForDel = True
                Else
                    'GOT ONE FOR DELETE BUT SIMULATE PROTECT READ ONLY SAFE MODE
                    A = A
                
                End If
                
                If Err.Number > 0 Then DeleteError = True
                
'-----------------------------------------------
                If CalculateFile_ERROR = "" Then
                    If ScanPath.Mnu_Simulate.Checked = False Then
                        ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "DELETED"
                    Else
                        'KillDelBlock = KillDelBlock + 1
    
                        ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "Simulate - Deleted - Read Only"
                    End If
                End If
'------------------------------------------------
                If CalculateFile_ERROR <> "" Then
                    'KillDelBlock = KillDelBlock + 1
                    ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "CRC32 RESULT -- " + CalculateFile_ERROR
                End If
                
                
                'ERRORLOGGS = ERRORLOGGS + ScanPath.ListView1.ListItems.Item(we).SubItems(4) + vbCrLf
            
                If ScanPath.WindowState = vbMaximized Or ScanPath.WindowState = vbNormal Then
                    ScanPath.ListView2.ListItems(ScanPath.ListView2.ListItems.Count).Selected = True
                    ScanPath.ListView2.SelectedItem.EnsureVisible
                End If
            
            End If
            
            'If UpForDel = True Or Err.Number > 0 Or source_var = True Then
                
                
                
            KillDelBlock_FLAG = False
            edge21 = False
            edge22 = False
                        
                        
            If WXHEX$ = "00000000" Then
            'VarCRCMSGBOXFLAG
                ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "Error Deleting File Not Sure, Answer RETURN vALUE ""00000000""" + Err.Description '-Reading File Safe Not To Deleted"
                
                ERRORLOGGS1 = ERRORLOGGS1 + ScanPath.ListView1.ListItems.Item(we).SubItems(4) + vbCrLf
                edge21 = True
                edge22 = True
                KillDelBlock_FLAG = True
            End If
            
            If DeleteError = True Then
                ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "Error Deleting File Not Sure, Answer " + Err.Description '-Reading File Safe Not Deleted"
                
                ERRORLOGGS1 = ERRORLOGGS1 + ScanPath.ListView1.ListItems.Item(we).SubItems(4) + vbCrLf
                edge21 = True
                edge22 = True
                KillDelBlock_FLAG = False
            End If
            
            If ReadError = True Then
                ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "Error Read File Before Delete - Best Not To Delete"
                
                ERRORLOGGS1 = ERRORLOGGS1 + ScanPath.ListView1.ListItems.Item(we).SubItems(4)
                edge21 = True
                KillDelBlock_FLAG = False
            End If
            
            
            If Source_Var = True Then
                ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "CRC Dupe & Source Folder Protected"
                'KillDelBlock = KillDelBlock + 1
                
                'ERRORLOGGS1 = ERRORLOGGS1 + ScanPath.ListView1.ListItems.Item(we).SubItems(4) + vbCrLf
                
                edge21 = True
                edge22 = True
                KillDelBlock_FLAG = False
            End If
            
            '--------------------------------------
            If ScanPath.Mnu_Simulate.Checked = True And UpForDel = True And edge22 = False Then
                ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "CRC Dupe & Simulate - Filter - Read Only Mode"
                
                ERRORLOGGS1 = ERRORLOGGS1 + ScanPath.ListView1.ListItems.Item(we).SubItems(4)
                
                If DoAbort = True Then
                    ERRORLOGGS1 = ERRORLOGGS1 + "If Simulate was Not Enable The File Would be Safe Read Only By NullString and Or Space-String Condition Decision Making Act"
                End If
                
                ERRORLOGGS1 = ERRORLOGGS1 + vbCrLf
                KillDelBlock_FLAG = False
            End If
            
            
            FLAGNOTSHOWTWICE = False
            If Len(A1$ + G1$) >= 256 Then
                If ResultDel2 <> "+" Then
                    ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "DOS_8.3>256 Path - " + ScanPath.ListView1.ListItems.Item(we).SubItems(4)
                    ERRORLOGGS1 = ERRORLOGGS1 + ScanPath.ListView1.ListItems.Item(we).SubItems(4)
                    FLAGNOTSHOWTWICE = True
                    KillDelBlock_FLAG = False
                End If
            End If
            
            If Len(A1$ + B1$) >= 256 And FLAGNOTSHOWTWICE = False Then
                If ResultDel2 <> "+" Then
                    KillDelBlock_FLAG = False
                    ScanPath.ListView1.ListItems.Item(we).SubItems(4) = ">256 Path - " + ScanPath.ListView1.ListItems.Item(we).SubItems(4)
                    ERRORLOGGS1 = ERRORLOGGS1 + ScanPath.ListView1.ListItems.Item(we).SubItems(4)
                End If
            End If
            
            'CalculateFile_ERROR
            If CalculateFile_ERROR <> "" Then
                'KillDelBlock = KillDelBlock + 1
                ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "CRC32 RESULT -- " + CalculateFile_ERROR
                ERRORLOGGS1 = ERRORLOGGS1 + ScanPath.ListView1.ListItems.Item(we).SubItems(4)
                KillDelBlock_FLAG = False
            End If
            
            If KillDelBlock_FLAG = True Then
                KillDelBlock = KillDelBlock + 1
            End If
            
            '----------------------------------------
            'WRITE APPEND FILE
            'WHAT FILE WAS COMPARED AGAINST AND OTHER
            '----------------------------------------
            If ResultDel1 <> "" Or ERRORLOGGS1 <> "" Then
                'BEFORE -- ERRORLOGGS3 = ""
                'WHOLE FILE IS A CRC DUPLICATE
                
                
                
                'ERRORLOGGS3 = String(55, "-") + vbCrLf + ERRORLOGGS3
                'MsgBox ERRORLOGGS3
                '----------------------------------------------
                
                ERRORLOGGS3 = ERRORLOGGS3 + String(55, "-") + vbCrLf
                
                If FLAG_DO_IN_CHUNK_OUT_OF_MEMORY = False Then
                    ERRORLOGGS3 = ERRORLOGGS3 + "CRC32 -- READ WHOLE FILE" + vbCrLf
                    ERRORLOGGS3 = ERRORLOGGS3 + "CRC32 -- NOT IN MEMORY OVERFLOW PROBLEM - UNDERFLOW - GLASS HALF FULL - WATER" + vbCrLf
                Else
                    ERRORLOGGS3 = ERRORLOGGS3 + "CRC32 -- READ CHUNK FILE -- MEMORY OVERFLOW MODE" + vbCrLf
                    ERRORLOGGS3 = ERRORLOGGS3 + "CRC32 -- READ CHUNK FILE -- WATER" + vbCrLf
                End If
                
                ERRORLOGGS3 = ERRORLOGGS3 + "CRC32 -- " + WXHEX + vbCrLf
                ERRORLOGGS3 = String(55, "-") + vbCrLf + ERRORLOGGS3
                
                'TRIM_STRING = "# " + Format(DUPE_COMPARE_AGAINST_COUNT, "000") + " -- " + TRIM_STRING
                'ScanPath.List1.AddItem WXHEX$ + "# -- " + A1$ + G1$ + vbCrLf
                
                'TRIM_STRING = ERRORLOGGS2
                'WXHEX$ "# -- " + A1$ + G1$ + vbCrLf
'                MsgBox ERRORLOGGS2
                
                
                ERRORLOGGS3 = ERRORLOGGS3 + ERRORLOGGS2
 '               MsgBox ERRORLOGGS3
                
'                MsgBox ERRORLOGGS3
                '----------------------------------------------
                'STATUS OF DELETE RESULT SUCCESS
                ERRORLOGGS3 = ERRORLOGGS3 + ResultDel1 + vbCrLf
                '----------------------------------------------
                ERRORLOGGS3 = ERRORLOGGS3 + String(55, "-") + vbCrLf
'                MsgBox ERRORLOGGS3
                
                If ERRORLOGGS1 <> "" Then
                    ERRORLOGGS3 = ERRORLOGGS3 + ERRORLOGGS1 + vbCrLf
                    ERRORLOGGS3 = ERRORLOGGS3 + String(55, "-") + vbCrLf
                Else
                    'ERRORLOGGS3 = ERRORLOGGS3 + "Long Path Problem at Delete -- Nothing to Report" + vbCrLf
                    'ERRORLOGGS3 = ERRORLOGGS3 + String(55, "-") + vbCrLf
                End If
                DUPE_COMPARE_AGAINST_COUNT = 0
                
                '-------------------------------------------
                'FIND COMPARE TO WHAT WAS NEXT TO EACH OTHER
                '-------------------------------------------
                TEST_MSG_BOX = False
                TRIM_STRING = ""
                
                'WANT THE MOST PREVIOUS FIRST LIKE DO IN REVERSE STEP ORDER
                
                For RX22 = ScanPath.List1.ListCount - 1 To 0 Step -1
                    If InStr(ScanPath.List1.List(RX22), WXHEX$) Then
                        DUPE_COMPARE_AGAINST_COUNT = DUPE_COMPARE_AGAINST_COUNT + 1
                        
                        If DUPE_COMPARE_AGAINST_COUNT > 1 Then TEST_MSG_BOX = True  ': Stop
                        'Stop
                        
                        TRIM_STRING1 = Mid(ScanPath.List1.List(RX22), InStr(ScanPath.List1.List(RX22), "# -- ") + 5)
                        TRIM_STRING2 = Mid(ScanPath.List1.List(RX22), 1, InStr(ScanPath.List1.List(RX22), "# -- ") - 1)
                        '-------------------------------------------------
                        'STILL KEPP VBCRLF IN LIST BOX WHEN NOT MULTI-LINE
                        'If InStr(TRIM_STRING1, vbCrLf) > 0 Then Stop
                        TRIM_STRING1 = Replace(TRIM_STRING1, vbCrLf, "")
                        '-------------------------------------------------
                        TRIM_STRING3 = Format(DUPE_COMPARE_AGAINST_COUNT, "000") + " # " + TRIM_STRING2
                        If DUPE_COMPARE_AGAINST_COUNT > 1 Then
                            TRIM_STRING3 = vbCrLf + TRIM_STRING3
                        End If
                        TRIM_STRING4 = TRIM_STRING3 + vbCrLf + TRIM_STRING1
                        'WXHEX$ "# -- " + A1$ + G1$ + vbCrLf
                        
                        TRIM_STRING = TRIM_STRING + TRIM_STRING4 + vbCrLf + String(55, "-")
                        'MsgBox TRIM_STRING4 '+ vbCrLf + String(55, "-")
                        
                        
                        If TEST_MSG_BOX = True Then
                            'MsgBox TRIM_STRING
                            'Stop
                        End If
                    End If
                Next
                
                If TEST_MSG_BOX = False Then
                    ERRORLOGGS3 = ERRORLOGGS3 + "Duplicate Compared With A SINGLE File Previous BELOW" + vbCrLf
                    If TRIM_STRING = "" Then Stop
                Else
                    ERRORLOGGS3 = ERRORLOGGS3 + "Duplicate Compared With " + Format(DUPE_COMPARE_AGAINST_COUNT, "000") + " Multiple PREVIOUS in Script BELOW" + vbCrLf
                End If
                
                ERRORLOGGS3 = ERRORLOGGS3 + String(55, "-") + vbCrLf
                ERRORLOGGS3 = ERRORLOGGS3 + TRIM_STRING + vbCrLf
                
                '--------------------------------------------------
                'INSERT THIS COUNT AND AT TOP OF THESE HEADER BLOCK
                'TO END WITH
                '--------------------------------------------------
                'COUNT AMOUNT OF CRC DEL
                '-----------------------
                DELETE_COUNT_NUMBER = DELETE_COUNT_NUMBER + 1
                
                'FUCK WORK AROUND
                If DELETE_COUNT_NUMBER = 1 Then
                    BEGIN_CRLF = vbCrLf
                Else
                    BEGIN_CRLF = ""
                End If
                
                ERRORLOGGS3 = BEGIN_CRLF + String(55, "-") + vbCrLf + "## -- " + Format(DELETE_COUNT_NUMBER) + vbCrLf + ERRORLOGGS3 + vbCrLf
                '-----------------------



                'MULTIPLE
'                If TEST_MSG_BOX = True Then
                    'MsgBox ERRORLOGGS3
'                End If
'
'                'SINGLE
'                If TEST_MSG_BOX = False Then
'                    MsgBox ERRORLOGGS3
'                End If

                
                'ERRORLOGGS3 = ERRORLOGGS3 + DUPE_COMPARE_AGAINST
                'ERRORLOGGS3 = ERRORLOGGS3 + String(55, "-") + vbCrLf
                '----------------------------------------------
                
                'FINISH LATER WHY TWO LOTS OF DUPE RESULT DATA
                'ResultDel = ERRORLOGGS2 + ResultDel
                
                FR4 = FreeFile
                Open Filename_Script For Append As #FR4
                    'MsgBox ERRORLOGGS3
                    Print #FR4, ERRORLOGGS3;
                    'Print #FR4, "Duplicate Was Compare With These File"
                    'Print #FR4, DUPE_COMPARE_AGAINST
                Close #FR4
                LOGG_INFO_DONE = True
            End If
            
            'NEW VAR USE HERE
            'ERRORLOGGS4 = ERRORLOGGS4 + ERRORLOGGS3
            
            
'            If ScanPath.Mnu_Simulate.Checked = True And UpForDel = True And edge22 = False Then
'                ERRORLOGGS = ERRORLOGGS + ScanPath.ListView1.ListItems.Item(WE).SubItems(4)
'
'                If DoAbort = True Then
'                    ScanPath.ListView1.ListItems.Item(WE).SubItems(4) = "Simulate - Space or Null Void File Exception Safe"
'
'                    ERRORLOGGS = ERRORLOGGS + "If Simulate was Not Emabaled The File Would be Safe Read Only By NullString and Or Space-String Condition Decision Making Act"
'
'                End If
'                ERRORLOGGS = ERRORLOGGS + vbCrLf
'            End If
        
        
        End If
        
        If ScanPath.MNU_CRC_Check_After_First.Checked = True Then
            If TT5$ = "*" Then
                'ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "Part CRC OPT -- " + ScanPath.ListView1.ListItems.Item(we).SubItems(4)
            End If
        End If
        'End If
                        
        If KillDelFAIL > 0 Then VarStringDel = ":" + Trim(Str(KillDelFAIL)) + "-"
            
        ScanPath.lblCount5 = Trim(Str(we2)) + " /" + Str(wex4) + "+  " + Trim(Str(KillDelFAIL)) + "-  " + Trim(Str(KillDelBlock)) + "*Blok  Del Dupes " + Format((we2 / we3) * 100, "0.00") + "%"

        'Delete Status
        If Len(ScanPath.ListView1.ListItems.Item(we).SubItems(4)) > XXCoulmn Then
            
            XXCoulmn = Len(ScanPath.ListView1.ListItems.Item(we).SubItems(4))
            
            'LAST COLUMNS -- DELETE STATUS
            Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(5))
        
        End If

        Set F = Nothing
        
        
        
        
        
        If WXHEX$ <> "" Then
            yy = yy + " - " + WXHEX$
            
            'DUPE_COMPARE_AGAINST_COUNT = DUPE_COMPARE_AGAINST_COUNT + 1
            'DUPE_COMPARE_AGAINST = DUPE_COMPARE_AGAINST + "# " + Format(DUPE_COMPARE_AGAINST_COUNT, "000") + " -- " + A1$ + G1$ + vbCrLf
            
            FNAMESHORT = G1$
            FNAMESHORT = FNAMESHORT + Space(12 - Len(G1$))
            'LONG AND SHORT OF IT
            'SHORT AND SWEET
            ScanPath.List1.AddItem WXHEX$ + "# -- " + FNAMESHORT + " -- " + A1$ + B1$ + vbCrLf
        End If
        
        If WXHEXXX$ <> "" Then yy = yy + " - " + WXHEXXX$
    
    End If

    'oldfilename = A1$ + B1$

endb:
    
Next


If FileSizeDelTotal < 1024 ^ 2 Then ''MaxFileSizeCRCCheck Then
ScanPath.lblCount10 = Format(FileSizeDelTotal, "###,###,###,##0") + " Byte Deleted Total"
Else
ScanPath.lblCount10 = Format(FileSizeDelTotal / 1024 ^ 2, "###,###,###,###.0") + " Meg Deleted Total"
End If


'LAST TWO COLUMNS I THINK CRC AND DELETE STATUS
'Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(3))
'Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(4))

ScanPath.WindowState = vbNormal
ScanPath.lblCount7.BackColor = RGB(0, 255, 0)

'If mCancelScan = True Then GoTo ENDE

ENDE:

If mCancelScan = True Then
'    ScanPath.lblCount7 = "Main Routine - Aborted"
Else
    ScanPath.lblCount7 = "Main Routine - Finished"
End If


'----------------------------------------------------
'SAVE THE HEADER INFO OF LABLE'S USED TOTALS
'USUAL ONLY WHEN A WRITE DONE WHEN CHANGE AFTER A DEL
'----------------------------------------------------
Call SAVE_RD_CLIP

'----------------------------------------------------------------------
'BEEN TRYING TO BE QUICK BEFORE NOW THESE ARE GOING TO BE BETTER METHOD
'----------------------------------------------------------------------

On Error Resume Next
Err.Clear
If Dir(Filename_Script) <> "" Then
    
    'RENAME TO TEMP
    '--------------
    If Dir(Filename_Script + ".TMP") Then
        Kill Filename_Script + ".TMP"
    End If
    '--------------
    'RENAME TO TEMP
    '--------------
    Name Filename_Script As Filename_Script + ".TMP"
    
    '-------------------------------------
    'WRITE A INITIAL HEADER INDEX AT TOP
    '-------------------------------------
    'WHY DOES THIS NOT APPEAR TO BE WRITEN
    '-------------------------------------

    FR4 = FreeFile
    Open Filename_Script For Output As #FR4
        STRING_LEN = 55
        Print #FR4, String(STRING_LEN, "-") '------------------------
        
        Print #FR4, "THIS LOGG FILE IS IN SOURCE AND ALSO APP.PATH FOLDER"
        Print #FR4, "SCAN SOURCE LOGG IS APPENED DURING PROGRESS AND "
        Print #FR4, "COPY MADE AT END TO APP.PATH"
        Print #FR4, "NOTE - UNLESS DON'T REACH END BY WAY OF CRASH "
        Print #FR4, "NOT BE A PROCESS SUMMERY END HEADER MERGED TOGETHER"
        Print #FR4, "THE HEADER BLOCK STATUS IS WRITTEN AT "
        Print #FR4, "EVERY DELETE AT TEMP FILE IN APP.PATH"
        Print #FR4, String(STRING_LEN, "-") '------------------------
        Print #FR4, "SOURCE PATH " + Filename_Script
        Print #FR4, String(STRING_LEN, "-") '------------------------
        FILENAMESCRIPT_PATHSTRIP = Mid(Filename_Script, InStrRev(Filename_Script, "\") + 1)
        Print #FR4, "APP PROGRAM PATH " + App.Path + "\" + FILENAMESCRIPT_PATHSTRIP
        Print #FR4, String(STRING_LEN, "-") '------------------------
        Print #FR4, Format(XStartTime, "YYYY-MM-DD HH:MM:SS") + " Start Time ----"
        Print #FR4, Format(Now, "YYYY-MM-DD HH:MM:SS") + " Finish Time ---"
        Print #FR4, String(STRING_LEN, "-") '------------------------
        If CommandExit = True Then FLAG5 = 1: Print #FR4, "LOGG END --- BY User Button CommandExit = True"
        If mCancelScan = True And CommandExit = False Then FLAG5 = 1: Print #FR4, "LOGG END --- BY mCancelScan = True"
        If FLAG5 = 0 Then Print #FR4, "LOGG END --- BY NORMAL FINISH ----"
        Print #FR4, String(STRING_LEN, "-") '------------------------
        Print #FR4, "Delete Count Total  =" + Str(KillDel)
        Print #FR4, "Delete Count Failed =" + Str(KillDelFAIL)
        Print #FR4, "Delete Count Block  =" + Str(KillDelBlock)
        Print #FR4, String(STRING_LEN, "-") '------------------------
    Close #FR4
    
    '-------------------------------------------------
    'OPEN ORGINAL FILENAME LOCATION OF LOGG FROM CLEAR
    '-------------------------------------------------
    'PUT THE HEADER BLOCK LINE BY LINE
    '---------------------------------
    'AND A SMALL PRE HEADER ALSO WRITTEN BEFORE
    '------------------------------------------
    
    STRING_LEN = 55
    FF1 = FreeFile
    Open Filename_Script For Append As #FF1
    
        Print #FR4, String(STRING_LEN, "-") '------------------------
        Print #FF1, "LABEL INFO ON FORM AT END"
        Print #FR4, String(STRING_LEN, "-") '------------------------
        
        If Dir(App.Path + "\TEMP_INFO_RD_CLIP.TXT") <> "" Then
            FRRD1 = FreeFile
            Open App.Path + "\TEMP_INFO_RD_CLIP.TXT" For Input As #FRRD1
                Do
                    Line Input #FRRD1, LINE_TEXT_INFO
                    Print #FF1, LINE_TEXT_INFO
                
                Loop Until EOF(FRRD1)
            Close FRRD1
            Close FF1
        Else
            MsgBox "HEADER BLOCK MISSING" + vbCrLf + App.Path + "\TEMP_INFO_RD_CLIP.TXT"
            Close FRRD1
            Close FF1
        End If
'---------

    Err.Clear
    FF1 = FreeFile
    Open Filename_Script For Append As #FF1
    
        If Err.Number > 0 Then
            MsgBox "ERROR OF MESSAGE = " + Err.Description + vbCrLf + "HAPPEN WHEN WRITE LAST PART OF LOGG FOR APPEND INFO"
        End If
    
        If Dir(Filename_Script + ".TMP") <> "" Then
            '-----------------------------------------------------------------
            'ADD APPEND THE TEMP LOGG SCRIPT TO THE ORGINAL LOCATION FROM TEMP
            '-----------------------------------------------------------------
            FR1 = FreeFile
            Open Filename_Script + ".TMP" For Input As #FR1
                Do
                    Line Input #FR1, LINE_TEXT_INFO
                    
                    If Mid(LINE_TEXT_INFO, 1, 5 + 1) = "## -- " Then
                        XCOUNT1 = Val(Mid(LINE_TEXT_INFO, 7))
                        XCOUNT2 = Format(XCOUNT1, String(Len(Str(DELETE_COUNT_NUMBER) - 1), "0"))
                        LINE_TEXT_INFO = Mid(LINE_TEXT_INFO, 1, InStr(LINE_TEXT_INFO, " -- ")) + " -- " + XCOUNT2 + " -- OF --" + Str(DELETE_COUNT_NUMBER)
                    End If
                    
                    
                    Print #FF1, LINE_TEXT_INFO
                Loop Until EOF(FRRD1)
            Close FR1
        Else
            MsgBox "TEMP LOGG SCRIPT MISSING" + vbCrLf + Filename_Script + ".TMP"
            Close FR1
        End If
    Close FF1

End If

If Dir(Filename_Script) <> "" Then
'    FR4 = FreeFile
'    Open Filename_Script For Input As #FR4
'        ERRORLOGGS4 = Input(LOF(FR4), FR4)
'    Close #FR4
'    If Trim(ERRORLOGGS4) <> "" Then
'        RDClip = RDClip + ERRORLOGGS4 + vbCrLf
'    End If
End If

'Debug.Print RDClip

'NEW VAR USE HERE
'ERRORLOGGS4 = ERRORLOGGS4 + ERRORLOGGS3



'Clipboard.Clear
'Clipboard.SetText (RDClip)

'---------------------------------------
'---------------------------------------
'If KillDel > 0 Then






If LOGG_INFO_DONE = False Then
    If Dir(Filename_Script) <> "" Then
        'Kill Filename_Script
    End If
End If

If KillDel = 0 And Dir(Filename_Script) <> "" And LOGG_INFO_DONE = True Then
    i = MsgBox("NOTHING DUPE DELETED DO YOU WANT TO DELETE LOGG FILE", vbYesNo)
    If i = vbYes Then Kill Filename_Script
'Else
'    FR4 = FreeFile
'    Open Filename_Script For Input As #FR4
'        Data_String = Input(LOF(FR4), FR4)
'    Close #FR4
End If
'---------------------------------------
'---------------------------------------


If Dir(Filename_Script) <> "" Then
    Set F = Fs.getfile(Filename_Script)
    F.Copy App.Path + "\" + Filename_Script_FILE
Else
    LOGG_INFO_DONE = False
End If

'FR1 = FreeFile
'Open App.Path + "\#_DATA\Visual Basic Code Loggs.txt" For Append As #FR1
'    Print #FR1, "-----------------------------------"
'    Print #FR1, "THIS LOGG FILE IS ALSO IN THE FOLDER SCAN REQUESTED AND IS APPENED DURING PROGRESS FOR RESULTS INCASE PROGRAM CRASH"
'    Print #FR1, Filename_Script
'    Print #FR1, "-----------------------------------"
'    Print #FR1, Format(XStartTime, "YYYY-MM-DD HH:MM:SS") + " Start Time ---"
'    Print #FR1, Format(Now, "YYYY-MM-DD HH:MM:SS") + " Finish Time ---"
'    Print #FR1, "-----------------------------------"
'
'    Print #FR1, RDClip;
'    Print #FR1, "-----------------------------------"
'
'    If Data_String <> "" Then
'        Print #FR1, "---- and Another Data -------------"
'        Print #FR1, "-----------------------------------"
'        Print #FR1, Data_String
'        Print #FR1, "-----------------------------------"
'    End If
'Close #FR1

If LOGG_INFO_DONE = True Then
    'i = MsgBox("Do You Want To Open The Error Log OR BOTH - Scroll to Bottom", vbYesNo, ScanPath.Caption)
'    i = MsgBox("Do You Want To Open The Error Logg" + vbCrLf + "Scroll to Bottom for Appended Index of Result Info", vbYesNo, ScanPath.Caption)
    i = MsgBox("Do You Want To Open The Error Logg" + vbCrLf + "Scroll to Bottom", vbYesNo, ScanPath.Caption)
    
    If i = vbYes Then Call ScanPath.Mnu_NotePad_Logg_Click

End If

'If mCancelScan = True Then Unload Me.hWnd 'z_MENU_Form1: Unload ScanPath

WORK_DONE = True
If mCancelScan <> True Then
    ScanPath.WindowState = vbNormal
End If

End Sub


Sub SAVE_RD_CLIP()

'1ST USE HERE
RDClip5 = RDClip5 + Trim(ScanPath.lblCount1) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount2) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount3) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount4) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount5) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount9) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount6) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount_BIGGEST_WORKING_SIZE) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount11_MEMORY_OVERFLOW1) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount12_MEMORY_OVERFLOW2) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount12_MEMORY_OVERFLOW3) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount6) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount10) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount7) + vbCrLf
RDClip5 = RDClip5 + Trim(ScanPath.lblCount8) + vbCrLf
'RDClip5 = RDClip5 + ERRORLOGGS4 + vbCrLf
'ERROR LOGGS PROPERBLY NOT WRITEN IN MEMORY ANYMORE AS SIZE


'Debug.Print RDClip5

On Error Resume Next
FRRD1 = FreeFile
Open App.Path + "\TEMP_INFO_RD_CLIP.TXT" For Output As #FRRD1
    Print #FRRD1, RDClip5
Close FRRD1


End Sub


Public Function StripSpaces(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStrRev(OriginalStr, Chr(32)) > 0) Then
        OriginalStr = Left(OriginalStr, InStrRev(OriginalStr, Chr(0)) - 1)
        End If
        StripSpaces = OriginalStr
End Function


Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStrRev(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStrRev(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
End Function



Public Function GetShortName(ByVal sLongFileName As String) As String
' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/kb/q175512/
' ---> The original comments were by them :

       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
End Function

Public Function GetLongName(ByVal sShortName As String) As String
' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/default.aspx?scid=kb;EN-US;154822
' ---> The original comments were by them :
     Dim sLongName As String
     Dim sTemp As String
     Dim iSlashPos As Integer

     'Add \ to short name to prevent Instr from failing
     sShortName = sShortName & "\"

     'Start from 4 to ignore the "[Drive Letter]:\" characters
     iSlashPos = InStr(4, sShortName, "\")
        
        If Mid(sShortName, 1, 2) = "\\" Then
            iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
            iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
        
        
        End If

     'Pull out each string between \ character for conversion
     While iSlashPos
       sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
       If sTemp = "" Then
         'Error 52 - Bad File Name or Number
         GetLongName = ""
         Exit Function
       End If
       sLongName = sLongName & "\" & sTemp
       iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
     Wend

     'Prefix with the drive letter
     GetLongName = Left$(sShortName, 2) & sLongName

   End Function




'***********************************************
'# Check, whether we are in the IDE
Public Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Public Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************

