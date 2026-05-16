VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   14265
   LinkTopic       =   "Form2"
   ScaleHeight     =   5550
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   4350
      Left            =   4635
      TabIndex        =   4
      Top             =   330
      Width           =   9510
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   2745
      TabIndex        =   3
      Top             =   345
      Width           =   1710
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   780
      Left            =   360
      TabIndex        =   2
      Top             =   2310
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   765
      Left            =   375
      TabIndex        =   1
      Top             =   1365
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   840
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   345
      Width           =   2130
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   1425
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Type FILE_NOTIFY_INFORMATION
  NextEntryOffset As Long
  Action As Long
  FileNameLength As Long
  FileName(1 To 255) As Byte
End Type

Private Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1&
Private Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2&
Private Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4&
Private Const FILE_NOTIFY_CHANGE_SIZE = &H8&
Private Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10&
Private Const FILE_NOTIFY_CHANGE_LAST_ACCESS = &H20&
Private Const FILE_NOTIFY_CHANGE_CREATION = &H40&
Private Const FILE_NOTIFY_CHANGE_SECURITY = &H100&

Private Const FILE_ACTION_ADDED = &H1&
Private Const FILE_ACTION_REMOVED = &H2&
Private Const FILE_ACTION_MODIFIED = &H3&
Private Const FILE_ACTION_RENAMED_OLD_NAME = &H4&
Private Const FILE_ACTION_RENAMED_NEW_NAME = &H5&

Private Const MAX_DIRS = 10 '//Only handle maximum upto 10 dir watch
Private Const MAX_BUFFER = 4096
Private Const MAX_PATH = 256 '//Your code will not work if Less than 256 (Yeppppp)
Private Type OVERLAPPED
  Internal As Long
  InternalHigh As Long
  offset As Long
  OffsetHigh As Long
  hEvent As Long
End Type

'// this is the all purpose structure that contains
'// the interesting directory information and provides
'// the input buffer that is filled with file change data

Private Type DIRECTORY_INFO
  hDir As Long
  lpszDirName As String * MAX_PATH
  lpBuffer(MAX_BUFFER) As Byte
  dwBufLength As Long
  oOverLapped As OVERLAPPED
  lComplKey As Long
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function ReadDirectoryChangesW Lib "kernel32" ( _
    ByVal hDirectory As Long, _
    lpBuffer As Any, _
    ByVal nBufferLength As Long, _
    ByVal bWatchSubtree As Long, _
    ByVal dwNotifyFilter As Long, _
    lpBytesReturned As Long, _
    lpOverlapped As Any, _
    lpCompletionRoutine As Any) As Long

Private Declare Function CreateIoCompletionPort Lib "kernel32" ( _
    ByVal FileHandle As Long, _
    ByVal ExistingCompletionPort As Long, _
    ByVal CompletionKey As Long, _
    ByVal NumberOfConcurrentThreads As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Declare Function PostQueuedCompletionStatus Lib "kernel32" ( _
    ByVal CompletionPort As Long, _
    lpNumberOfBytesTransferred As Long, _
    lpCompletionKey As Long, _
    lpOverlapped As Long) As Long

Private Declare Function GetQueuedCompletionStatus Lib "kernel32" ( _
    ByVal CompletionPort As Long, _
    lpNumberOfBytesTransferred As Long, _
    lpCompletionKey As Long, _
    lpOverlapped As OVERLAPPED, _
    ByVal dwMilliseconds As Long) As Long

Private Const FILE_LIST_DIRECTORY = &H1
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_DELETE = &H4
Private Const FILE_SHARE_WRITE = &H2

Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const OPEN_EXISTING = 3

Private Const INVALID_HANDLE_VALUE = -1
Private Const INFINITE = &HFFFF

Private hDir As Long
Private lpSecurityAttributes As SECURITY_ATTRIBUTES

'Private hEvent As Long
Private hIOCompPort As Long
Private oOverLapped As OVERLAPPED
Private InfoBuffer(MAX_BUFFER) As Byte
Private lpBytesReturned As Long

Private Type FileChange
  NextPos As Long
  Name As String
  Type As Long
End Type

Private hThread As Long, lngTid As Long
Private DirInfo(MAX_DIRS) As DIRECTORY_INFO  '// Buffer for all of the directories
Private numDirs As Integer
Private ret As Long, lngFilter As Long, bWatchSubDir As Boolean

Private Function WatchDirectoryOrFilename(colDirPaths As Collection) As Boolean
  Dim i As Integer

  If colDirPaths.Count <= 0 Then Exit Function
  numDirs = colDirPaths.Count

  lpSecurityAttributes.nLength = Len(lpSecurityAttributes)

  For i = 1 To colDirPaths.Count
    '//[Step 1] Obtain each dir handle which you want to monitor
    DirInfo(i - 1).hDir = CreateFile(colDirPaths(i), _
        FILE_LIST_DIRECTORY, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
        lpSecurityAttributes, _
        OPEN_EXISTING, _
        FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OVERLAPPED, _
        0)

    DirInfo(i - 1).lpszDirName = colDirPaths(i)

    '//[Step 2] Set up a key(directory info) for directory

    '//Note: Use hIOCompPort as an exisiting port so CreateIoCompletionPort dont create new port for each dir

    DirInfo(i - 1).lComplKey = VarPtr(DirInfo(i - 1))  '//Get a completion Key for each directory
    '//Here we have used mem loc for each dirinfo record
    hIOCompPort = CreateIoCompletionPort(DirInfo(i - 1).hDir, hIOCompPort, DirInfo(i - 1).lComplKey, 0)

    '// Start watching each of the directories of interest

    ret = ReadDirectoryChangesW(DirInfo(i - 1).hDir, _
        DirInfo(i - 1).lpBuffer(0), _
        MAX_BUFFER, _
        bWatchSubDir, _
        lngFilter, _
        lpBytesReturned, _
        DirInfo(i - 1).oOverLapped, _
        ByVal 0&)
  Next

  If ret <> 0 Then WatchDirectoryOrFilename = True
End Function

Private Sub HandleDirectoryChanges()
  On Error Resume Next

  Dim fContinue As Boolean
  Dim lPointer As Long, iPos As Integer
  Dim lngCompKey As Long
  Dim lpOverlapped As OVERLAPPED
  Dim fni As FILE_NOTIFY_INFORMATION
  Dim DI As DIRECTORY_INFO
  Dim FileInfo As FileChange

  '// Retrieve the directory info for this directory
  '// through the completion key
  ret = GetQueuedCompletionStatus(hIOCompPort, _
      lpBytesReturned, _
      lngCompKey, _
      oOverLapped, _
      100)       '//Put INFINITE if synchronous

  If (ret <> 0) Then
    lPointer = 0

    '//Findout which directory is this based on completionkey
    For iPos = 0 To numDirs - 1
      If DirInfo(iPos).lComplKey = lngCompKey Then Exit For
    Next

    Do
      FileInfo = GetFileName(DirInfo(iPos).lpBuffer, lPointer)

      strPath = Trim(DirInfo(iPos).lpszDirName) & "\" & FileInfo.Name

      Select Case FileInfo.Type
        Case FILE_ACTION_ADDED
          strmsg = "[" & strPath & "] was added to the directory."
        Case FILE_ACTION_REMOVED
          strmsg = "[" & strPath & "] was removed from the directory."
        Case FILE_ACTION_MODIFIED
          strmsg = "[" & strPath & "] was modified. (i.e. timestamp/attributes chage)."
        Case FILE_ACTION_RENAMED_OLD_NAME
          strmsg = "[" & strPath & "] was renamed and this is the old name."
        Case FILE_ACTION_RENAMED_NEW_NAME
          strmsg = "[" & strPath & "] was renamed and this is the new name."
      End Select
      
      If strmsg <> "" Then List2.AddItem strmsg
      
      lPointer = FileInfo.NextPos
    Loop While (lPointer > 0)
        
    ret = ReadDirectoryChangesW(DirInfo(iPos).hDir, _
        DirInfo(iPos).lpBuffer(0), _
        MAX_BUFFER, _
        bWatchSubDir, _
        lngFilter, _
        lpBytesReturned, _
        DirInfo(iPos).oOverLapped, _
        ByVal 0&)
  End If

End Sub

Private Function GetLongFromBuf(Buf() As Byte, lPos As Long) As Long
  On Error Resume Next

  Dim lTemp As Long
  Dim lResult As Long
  lResult = Buf(lPos)
  lTemp = Buf(lPos + 1): lResult = lResult + lTemp * 256
  lTemp = Buf(lPos + 2): lResult = lResult + lTemp * 65536
  lTemp = Buf(lPos + 3): lResult = lResult + lTemp * 65536 * 256
  GetLongFromBuf = lResult
End Function

Private Function GetFileName(Buf() As Byte, lPointer As Long) As FileChange
  On Error Resume Next

  Dim lNextOffset As Long
  Dim lFileSize As Long
  Dim sTemp As String
  Dim bTemp() As Byte
  Dim i As Long

  lNextOffset = GetLongFromBuf(Buf, lPointer)
  GetFileName.Type = GetLongFromBuf(Buf, lPointer + 4)
  lFileSize = GetLongFromBuf(Buf, lPointer + 8)

  If lFileSize = 0 Then
    GetFileName.NextPos = 0
    Exit Function
  End If
  ReDim bTemp(1 To lFileSize)

  For i = 1 To lFileSize
    bTemp(i) = Buf(lPointer + 11 + i)
  Next i

  GetFileName.Name = bTemp

  If lNextOffset = 0 Then
    GetFileName.NextPos = 0
  Else
    GetFileName.NextPos = lNextOffset + lPointer
  End If
End Function

Sub CleanUp()
  On Error Resume Next
  
  Call PostQueuedCompletionStatus(hIOCompPort, 0, 0, ByVal 0&)

  For i = 0 To numDirs - 1
    CloseHandle (DirInfo(i).hDir)
    DirInfo(i).hDir = 0
  Next

  CloseHandle (hIOCompPort)
  hIOCompPort = 0
End Sub

Private Sub Check1_Click()
  bWatchSubDir = Check1.Value
End Sub

Private Sub Command1_Click()
  On Error Resume Next
  If List1.ListCount >= MAX_DIRS Then
    MsgBox "Maximum " & MAX_DIRS & " directories can be monitored."
    Exit Sub
  End If
  
  '//Directories to watch
  Dim colDirPaths As New Collection
    
  List1.AddItem Text1
    
  For i = 0 To List1.ListCount - 1
    colDirPaths.Add List1.List(i)
  Next
  
  Call CleanUp
  Timer1.Enabled = WatchDirectoryOrFilename(colDirPaths)
End Sub

Private Sub Form_Load()
  On Error Resume Next

  '//Directories to watch
  Dim colDirPaths As New Collection
  
  Command1.Caption = "<< Add"
  Check1.Caption = "Monitor All Sub Directories"
  Check1.Value = 1
  
  'MkDir "D:\TEMP1"
  'MkDir "D:\TEMP2"

  '//Lets watch for 2 different directories
  'colDirPaths.Add "D:\TEMP1"
  'colDirPaths.Add "D:\TEMP2"
  colDirPaths.Add "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures"


  'List1.AddItem "D:\TEMP1"
  'List1.AddItem "D:\TEMP2"
  List1.AddItem "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures"

  '//Create Notification filter
  lngFilter = FILE_NOTIFY_CHANGE_FILE_NAME _
      Or FILE_NOTIFY_CHANGE_SIZE _
      Or FILE_NOTIFY_CHANGE_ATTRIBUTES _
      Or FILE_NOTIFY_CHANGE_DIR_NAME _
      Or FILE_NOTIFY_CHANGE_FILE_NAME _
      Or FILE_NOTIFY_CHANGE_LAST_ACCESS _
      Or FILE_NOTIFY_CHANGE_LAST_WRITE _
      Or FILE_NOTIFY_CHANGE_SECURITY



  '//Watch for sub directory changes
  bWatchSubDir = Check1.Value

  Timer1.Enabled = False
  Timer1.Interval = 100  '//Check for Queued Changes every 100 ms
  Timer1.Enabled = WatchDirectoryOrFilename(colDirPaths)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call CleanUp
End Sub

Private Sub Timer1_Timer()
  Me.Caption = "Monitoring ..."
  Call HandleDirectoryChanges
End Sub
