VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CHECK HERE
'HAD TO REM OUT
'Form1.MNU_SCANPATH_COUNTER.Caption = Str(mFileCount) + " Files"
'AND ANOTHER


'Option Explicit

'Dim FS


'##############################################################################################
'Purpose:       Scan a specified path for files/folders matching the Filter Criteria
'Author:        Richard Mewett ©2003
'Version:       2.00 (14.10.2005)
'Dependencies:  None!

'This allows filtered scan operations since you can define Min/Max file sizes, Date Ranges
'multiple File Extensions & File Attributes

'Copyright ©2005 Richard Mewett. All rights reserved.
'
'This software is provided "as-is," without any express or implied warranty.
'In no event shall the author be held liable for any damages arising from the
'use of this software.
'If you do not agree with these terms, do not install "ScanPath". Use of
'the program implicitly means you have agreed to these terms.
'
'Permission is granted to anyone to use this software for any purpose,
'including commercial use, and to alter and redistribute it, provided that
'the following conditions are met:
'
'1. All redistributions of source code files must retain all copyright
'   notices that are currently in place, and this list of conditions without
'   any modification.
'
'2. All redistributions in binary form must retain all occurrences of the
'   above copyright notice and web site addresses that are currently in
'   place (for example, in the About boxes).
'
'3. Modified versions in source or binary form must be plainly marked as
'   such, and must not be misrepresented as being the original software.

'##############################################################################################

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const NULL_CHAR = 0
Private Const EXT_SEPERATOR_CHAR = "."
Private Const MASK_SEPERATOR_CHAR = ";"

Private Const ALL_FILES = "*.*"
Private Const PARENT_DIR = "."
Private Const ROOT_DIR = ".."

Private Const DIR_KEY = "1"
Private Const FILE_KEY = "0"

Private Type SYSTEMTIME
    wYear             As Integer
    wMonth            As Integer
    wDayOfWeek        As Integer
    wDay              As Integer
    wHour             As Integer
    wMinute           As Integer
    wSecond           As Integer
    wMilliseconds     As Integer
End Type

Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type

Public Enum DateTypeEnum
    Modified = 0
    Created = 1
    LastAccessed = 2
End Enum

Private mWFD As WIN32_FIND_DATA

Private mFromDate As Date
Private mToDate As Date
Private mMaximumSize As Long
Private mMinimumSize As Long

Private mArchive As Boolean
Private mCompressed As Boolean
Private mFolder As Boolean
Private mHidden As Boolean
Private mNormal As Boolean
Private mReadOnly As Boolean
Private mSystem As Boolean
Private mFilter As String


Private mDirCount As Long
Private mFileCount As Long
Private mDateType As DateTypeEnum
Private mDateCheck As Integer
Private mSizeCheck As Integer
Private mCancelScan As Boolean
Private mPatternMatching As Boolean
Private mFilters() As String

Public Event DirMatchxx(Directory As String, DDirectory As String, Path As String)
Public Event DirMatchFAST(Directory As String, DDirectory As String, Path As String, DDate As Date)
Public Event DirMatch(Directory As String, DDirectory As String, Path As String)
Public Event FileMatch(fileName As String, DFilename As String, Path As String)
Public Event FileMatchFast(fileName As String, DFilename As String, Path As String, DDate As Date)


Public Property Get DirCount() As Long
    DirCount = mDirCount
End Property


Public Property Get FileCount() As Long
    FileCount = mFileCount
End Property



Private Function FileMatches(udtWFD As WIN32_FIND_DATA) As Boolean
    '##############################################################################################
    'Purpose: Used to verify a file matches the Filter Criteria
    '##############################################################################################
    
    Dim dtFileDate As Date
    Dim nFilter As Integer
    Dim nPos As Integer
    Dim bFolder As Boolean
    Dim bOK As Boolean
    Dim sExt As String
    
    bFolder = (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
    
    If bFolder Or (mFilter = ALL_FILES) Then
        bOK = True
    Else
        'Remove the padding from filename by truncating at first null char
        sExt = UCase$(Mid$(udtWFD.cFileName, 1, InStr(udtWFD.cFileName, Chr$(NULL_CHAR)) - 1))
        
        If mPatternMatching Then
            'Complex searches such as *test??.text
            For nFilter = LBound(mFilters) To UBound(mFilters)
                If sExt Like mFilters(nFilter) Then
                    bOK = True
                    Exit For
                End If
            Next nFilter
        Else
            'Simple extension only search
            'Get the Extension - Scan backwards since Windows uses the last "." to determine extension
            '>> i.e. datafile.txt.bmp is a bmp file not a txt file
            nPos = InStrRev(sExt, EXT_SEPERATOR_CHAR)
            If nPos > 0 Then
                'Append seperator char to Extension prevent false matches
                '>> i.e. *.rtf should not match a file with a .r extension
                'Looping through an array of Filters would negate the need for this but Instr is faster
                sExt = UCase$(Mid$(sExt, nPos + 1)) & MASK_SEPERATOR_CHAR
                bOK = (InStr(mFilter, sExt) > 0)
            End If
        End If
    End If
    
    If bOK Then
        If (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_ARCHIVE) Then
            bOK = mArchive
        End If
        If (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_COMPRESSED) Then
            bOK = bOK And mCompressed
        End If
        If (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            bOK = bOK And mFolder
        End If
        If (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN) Then
            bOK = bOK And mHidden
        End If
        If (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) Then
            bOK = bOK And mNormal
        End If
        If (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_READONLY) Then
            bOK = bOK And mReadOnly
        End If
        If (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM) Then
            bOK = bOK And mSystem
        End If
    End If
        
    If bOK And Not bFolder Then
        Select Case mSizeCheck
        Case 1
            bOK = (udtWFD.nFileSizeLow >= mMinimumSize)
        Case 2
            bOK = (udtWFD.nFileSizeLow <= mMaximumSize)
        Case 3
            bOK = (udtWFD.nFileSizeLow >= mMinimumSize) And (udtWFD.nFileSizeLow <= mMaximumSize)
        End Select
                
        If bOK And (mDateCheck > 0) Then
            Select Case mDateType
            Case 0
                dtFileDate = FormatFileDate(udtWFD.ftLastWriteTime)
            Case 1
                dtFileDate = FormatFileDate(udtWFD.ftCreationTime)
            Case 2
                dtFileDate = FormatFileDate(udtWFD.ftLastAccessTime)
            End Select
            
            Select Case mDateCheck
            Case 1
                bOK = (dtFileDate >= mFromDate)
            Case 2
                bOK = (dtFileDate <= mToDate)
            Case 3
                bOK = (dtFileDate >= mFromDate) And (dtFileDate <= mToDate)
            End Select
        End If
    End If
                    
    FileMatches = bOK
End Function

Private Function FormatFileDate(CT As FILETIME) As String
    Const SHORT_DATE = "Short Date"
    Const SHORT_TIME = "Short Time"
    
    Dim ST As SYSTEMTIME
    Dim ds(1) As Single
       
    'CT.dwHighDateTime
       
    If FileTimeToSystemTime(CT, ST) Then
          ds(0) = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
          ds(1) = TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
          
          FormatFileDate = Format$(ds(0), SHORT_DATE) & " " & Format$(ds(1), SHORT_TIME)
    End If
End Function




Public Property Get Filter() As String
    If Right$(mFilter, 1) = MASK_SEPERATOR_CHAR Then
        Filter = Left$(mFilter, Len(mFilter) - 1)
    Else
        Filter = mFilter
    End If
End Property

Public Property Let Filter(ByVal sNewValue As String)
    mFilter = Trim$(sNewValue)
End Property

Public Property Get FOLDER() As Boolean
    FOLDER = mFolder
End Property

Public Property Let FOLDER(ByVal bNewValue As Boolean)
    mFolder = bNewValue
End Property

Public Property Get Archive() As Boolean
    Archive = mArchive
End Property


Public Property Let Archive(ByVal bNewValue As Boolean)
    mArchive = bNewValue
End Property


Public Property Get System() As Boolean
    System = mSystem
End Property

Public Property Let System(ByVal bNewValue As Boolean)
    mSystem = bNewValue
End Property

Public Property Get Hidden() As Boolean
    Hidden = mHidden
End Property

Public Property Let Hidden(ByVal bNewValue As Boolean)
    mHidden = bNewValue
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property

Public Property Let ReadOnly(ByVal bNewValue As Boolean)
    mReadOnly = bNewValue
End Property

Public Property Get Normal() As Boolean
    Normal = mNormal
End Property

Public Property Let Normal(ByVal bNewValue As Boolean)
    mNormal = bNewValue
End Property


Private Function DirectScan_USE_NO_LIST(Path As String, SubFolders As Boolean) As Long
    Dim colDirs As Collection
    Dim lCount As Long
    Dim lResult As Long
    Dim bOK As Boolean
    Dim sDir As String
    Dim sFile As String
    Dim dFile As String
    
    Set colDirs = New Collection
    
    colDirs.Add Path
    Do While colDirs.Count And Not mCancelScan
        sDir = colDirs.Item(1)
        colDirs.Remove 1
        
        lResult = FindFirstFile(sDir & ALL_FILES, mWFD)
        If lResult = INVALID_HANDLE_VALUE Then
            Exit Do
        End If
        
        Do
            sFile = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
            dFile = Left$(mWFD.cAlternate, InStr(mWFD.cAlternate, Chr$(NULL_CHAR)) - 1)
            If dFile = "" Then dFile = sFile
            Select Case sFile
            Case ROOT_DIR, PARENT_DIR
            
            Case Else
                bOK = FileMatches(mWFD)
                
                If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    If bOK Then
                        If SubFolders Then
                            colDirs.Add sDir & sFile & "\"
                        End If
                    
                        mDirCount = mDirCount + 1
                        ScanPath.lblCount2 = mDirCount
                        
                        ScanPath.lblCount7 = sDir + sFile
                        RaiseEvent DirMatch(sFile, dFile, sDir)
                    End If
                Else
                    If bOK Then
                        mFileCount = mFileCount + 1
                        If (mFileCount Mod 50) = 0 Then
                            ScanPath.lblCount1 = mFileCount
                        End If
                        
                        'LAST FILE =
                        ScanPath.lblCount7 = sDir + sFile
                        'xDATEPATH = sDir + sFile
                        
                        RaiseEvent FileMatch(sFile, dFile, sDir)
                        'RaiseEvent FileMatch(sFile, dFile, sDir)
                    End If
                End If
                
            End Select
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
            End If
        Loop While FindNextFile(lResult, mWFD)
        FindClose lResult
    Loop
    
    ScanPath.lblCount1 = mFileCount
    
    
    Set colDirs = Nothing
End Function

Private Function DirectScan_USE_NO_LIST_FAST(Path As String, SubFolders As Boolean) As Long
    Dim colDirs As Collection
    Dim lCount As Long
    Dim lResult As Long
    Dim bOK As Boolean
    Dim sDir As String
    Dim sFile As String
    Dim dFile As String
    Dim sDate As Date
    
    Set colDirs = New Collection
    mFileCount = 0
    
    colDirs.Add Path
    Do While colDirs.Count 'And Not mCancelScan
        sDir = colDirs.Item(1)
        colDirs.Remove 1
        
        lResult = FindFirstFile(sDir & ALL_FILES, mWFD)
        If lResult = INVALID_HANDLE_VALUE Then
            Exit Do
        End If
        
        Do
            sFile = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
            dFile = Left$(mWFD.cAlternate, InStr(mWFD.cAlternate, Chr$(NULL_CHAR)) - 1)
            If dFile = "" Then dFile = sFile
            Select Case sFile
            Case ROOT_DIR, PARENT_DIR
            
            Case Else
                bOK = FileMatches(mWFD)
                
                If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    If bOK Then
                        If SubFolders Then
                            colDirs.Add sDir & sFile & "\"
                        End If
                    
                        mDirCount = mDirCount + 1
                        'ScanPath.lblCount2 = mDirCount
                        
                        
                        'ScanPath.lblCount7 = sDir + sFile
                        'RaiseEvent DirMatch(sFile, dFile, sDir)
                    End If
                Else
                    If bOK Then
                        
                        mFileCount = mFileCount + 1
                        tFileCount = mFileCount
                        'If (mFileCount Mod 50) = 0 Then
                        '    ScanPath.lblCount1 = mFileCount
                        'End If
                        
                        
                        'NEWEST FILE =
                        'sDate = mWFD.ftLastWriteTime
                        sDate = FormatFileDate(mWFD.ftLastWriteTime)
                        'sDate = sDate + TimeValue(mWFD.ftLastWriteTime, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
                        
                        'If sDate >= LAST_FILE_DATE_TIME Then
                        '    LAST_FILE_DATE_TIME = sDate
                        '    LAST_FILE_DATE_PATH = sDir + sFile
                        'End If
                        RaiseEvent FileMatchFast(sFile, dFile, sDir, sDate)
                        
                        'RaiseEvent FileMatchFast(sFile, dFile, sDate, sDate)
                    
                    End If
                End If
                
            End Select
            
            lCount = lCount + 1
            If (lCount Mod 100) = 0 Then
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
                
                'ScanPath.lblCount1 = mFileCount
                'Form1.MNU_SCANPATH_COUNTER.Caption = Str(mFileCount) + " Files"

            
            End If
        Loop While FindNextFile(lResult, mWFD)
        FindClose lResult
    Loop
    
    ScanPath.lblCount1 = mFileCount
    'Form1.MNU_SCANPATH_COUNTER.Caption = mFileCount
    'Form1.MNU_SCANPATH_COUNTER.Caption = Str(lCount) + "+" + Str(mFileCount) + " Files"
'    Form1.MNU_SCANPATH_COUNTER.Caption = Str(mFileCount) + " Files"
    
    
    Set colDirs = Nothing
End Function

Private Function DirectScan_USE_NO_LIST_FAST_FORM_CAPTURE(Path As String, SubFolders As Boolean) As Long
    
    Set FS = CreateObject("Scripting.FileSystemObject")

    Dim colDirs As Collection
    Dim lCount As Long
    Dim lResult As Long
    Dim bOK As Boolean
    Dim sDir As String
    Dim sFile As String
    Dim dFile As String
    Dim sDate As Date
    Dim F
    
    Set colDirs = New Collection
    mFileCount = 0
    
    colDirs.Add Path
    Do While colDirs.Count 'And Not mCancelScan
        sDir = colDirs.Item(1)
        colDirs.Remove 1
        
        lResult = FindFirstFile(sDir & ALL_FILES, mWFD)
        If lResult = INVALID_HANDLE_VALUE Then
            Exit Do
        End If
        
        Do
            sFile = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
            dFile = Left$(mWFD.cAlternate, InStr(mWFD.cAlternate, Chr$(NULL_CHAR)) - 1)
            If dFile = "" Then dFile = sFile
            Select Case sFile
            Case ROOT_DIR, PARENT_DIR
            
            Case Else
                bOK = FileMatches(mWFD)
                
                If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    If bOK Then
                        If SubFolders Then
                            colDirs.Add sDir & sFile & "\"
                        End If
                    
                        mDirCount = mDirCount + 1
                        'ScanPath.lblCount2 = mDirCount
                        
                        
                        'ScanPath.lblCount7 = sDir + sFile
                        'RaiseEvent DirMatch(sFile, dFile, sDir)
                    End If
                Else
                    If bOK Then
                        
                        mFileCount = mFileCount + 1
                        tFileCount = mFileCount
                        'If (mFileCount Mod 50) = 0 Then
                        '    ScanPath.lblCount1 = mFileCount
                        'End If
                        
                        
                        'NEWEST FILE =
                        'sDate = mWFD.ftLastWriteTime
                        
                        'FORMAT FDATE NOT WORKING PROPER
                        'sDate = FormatFileDate(mWFD.ftLastWriteTime)
                        
                        'sDate = FormatFileDate(mWFD.ftCreationTime)
                        'TimeValue(mWFD.ftLastWriteTime)
                        'sDate = sDate + TimeValue(mWFD.ftLastWriteTime, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
                        
                        Set F = FS.GetFile(sDir + dFile)
                        sDate = F.DateLastModified

                        
                        'IS DATE FILE BIGGER NEWER THAN TIME SET
                        If sDate >= LAST_FILE_DATE_TIME Then
                            'Debug.Print "YES"
                            Exit Do
                            'DON'T NEED RUN ANYMORE IF DATE FILE NEWER THAN LIMIT
                        Else
                            'FILE IF IN AFTER DATE RANGE
                            'WILL WANT DELETE
                                
                                'ERROR DON'T CARE
                                On Error Resume Next
                                Kill sDir + sFile
                                If Err.Number > 0 And IsIDE = True Then Stop
                                On Error GoTo 0
                            
                            'Stop
                            'Kill
                        '    LAST_FILE_DATE_TIME = sDate
                        '    LAST_FILE_DATE_PATH = sDir + sFile
                        End If
                        
                        'RaiseEvent FileMatchFast(sFile, dFile, sDir, sDate)
                        
                        'RaiseEvent FileMatchFast(sFile, dFile, sDate, sDate)
                    End If
                End If
                
            End Select
            
            lCount = lCount + 1
            If (lCount Mod 100) = 0 Then
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
                'ScanPath.lblCount1 = mFileCount
                'Form1.MNU_SCANPATH_COUNTER.Caption = Str(mFileCount) + " Files"
            End If
        Loop While FindNextFile(lResult, mWFD)
        FindClose lResult
    Loop
    
    'ScanPath.lblCount1 = mFileCount
    'Form1.MNU_SCANPATH_COUNTER.Caption = mFileCount
    'Form1.MNU_SCANPATH_COUNTER.Caption = Str(lCount) + "+" + Str(mFileCount) + " Files"
    'Form1.MNU_SCANPATH_COUNTER.Caption = Str(mFileCount) + " Files"
    
    Set colDirs = Nothing
End Function




Private Function DirectScan(Path As String, SubFolders As Boolean) As Long
    Dim colDirs As Collection
    Dim lCount As Long
    Dim lResult As Long
    Dim bOK As Boolean
    Dim sDir As String
    Dim sFile As String
    Dim dFile As String
    
    Set colDirs = New Collection
    
    colDirs.Add Path
    Do While colDirs.Count And Not mCancelScan
        sDir = colDirs.Item(1)
        colDirs.Remove 1
        
        lResult = FindFirstFile(sDir & ALL_FILES, mWFD)
        If lResult = INVALID_HANDLE_VALUE Then
            Exit Do
        End If
        
        Do
            sFile = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
            dFile = Left$(mWFD.cAlternate, InStr(mWFD.cAlternate, Chr$(NULL_CHAR)) - 1)
            If dFile = "" Then dFile = sFile
            Select Case sFile
            Case ROOT_DIR, PARENT_DIR
            
            Case Else
                bOK = FileMatches(mWFD)
                
                If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    If bOK Then
                        If SubFolders Then
                            colDirs.Add sDir & sFile & "\"
                        End If
                    
                        mDirCount = mDirCount + 1
'                        ScanPath.lblCount7 = sDir + sFile
                        
                        'NOT DIR MATCH OPTION
                        'RaiseEvent DirMatch(sFile, dFile, sDir)
                    End If
                Else
                    If bOK Then
                        mFileCount = mFileCount + 1
'                        ScanPath.lblCount7 = sDir + sFile
                        RaiseEvent FileMatch(sFile, dFile, sDir)
                        'RaiseEvent FileMatch(sFile, dFile, sDir)
                    End If
                End If
                
            End Select
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
            End If
        Loop While FindNextFile(lResult, mWFD)
        FindClose lResult
    Loop
    
    
    Set colDirs = Nothing
End Function

Private Function DirectScanDIR(Path As String, SubFolders As Boolean) As Long
   
    Dim colDirs As Collection
    Dim lCount As Long
    Dim lResult As Long
    Dim bOK As Boolean
    Dim sDir As String
    Dim Bdir As String
    Dim sFile As String
    Dim dFile As String
    Dim DirCounter
    Set colDirs = New Collection
    
    DirCounter = 1
    
    colDirs.Add Path
    Do While colDirs.Count And Not mCancelScan
        sDir = colDirs.Item(1)
            
        colDirs.Remove 1
        
        lResult = FindFirstFile(sDir & ALL_FILES, mWFD)
        If lResult = INVALID_HANDLE_VALUE Then
            Exit Do
        End If
        
        Do
            DoEvents
            sFile = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
            dFile = Left$(mWFD.cAlternate, InStr(mWFD.cAlternate, Chr$(NULL_CHAR)) - 1)
            If dFile = "" Then dFile = sFile


            'bDir = mWFD.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY
            
            Select Case sFile
            
            Case ROOT_DIR, PARENT_DIR
            
            Case Else
'                bOK = FileMatches(mWFD)
                
                If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    'If bOK Then
                  '      If bDir <> sDir Then
                            colDirs.Add sDir + sFile + "\"
                            Bdir = sDir
                    
                        mDirCount = mDirCount + 1
                        RaiseEvent DirMatchxx(sFile, dFile, sDir)
                End If
                
            End Select
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
            End If
        Loop While FindNextFile(lResult, mWFD)
        FindClose lResult
    Loop
    
    Set colDirs = Nothing
End Function




Private Function DirectScanDIR_FAST(Path As String, SubFolders As Boolean) As Long
   
    Dim colDirs As Collection
    Dim lCount As Long
    Dim lResult As Long
    Dim bOK As Boolean
    Dim sDir As String
    Dim Bdir As String
    Dim sFile As String
    Dim dFile As String
    Dim DirCounter
    Set colDirs = New Collection
    Dim sDate As Date

    DirCounter = 1
    mDirCount = 0
    colDirs.Add Path
    Do While colDirs.Count And Not mCancelScan
        sDir = colDirs.Item(1)
            
        colDirs.Remove 1
        
        lResult = FindFirstFile(sDir & ALL_FILES, mWFD)
        If lResult = INVALID_HANDLE_VALUE Then
            Exit Do
        End If
        
        Do
            DoEvents
            sFile = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
            dFile = Left$(mWFD.cAlternate, InStr(mWFD.cAlternate, Chr$(NULL_CHAR)) - 1)
            If dFile = "" Then dFile = sFile


            'bDir = mWFD.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY
            
            Select Case sFile
            
            Case ROOT_DIR, PARENT_DIR
            
            Case Else
'                bOK = FileMatches(mWFD)
                
                If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    'If bOK Then
                  '      If bDir <> sDir Then
                        
                        colDirs.Add sDir + sFile + "\"
                        
                        sDate = FormatFileDate(mWFD.ftLastWriteTime)

                        
                        'Bdir = sDir
                    
                        mDirCount = mDirCount + 1
                        RaiseEvent DirMatchFAST(sFile, dFile, sDir, sDate)
                End If
                
            End Select
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
'                Form1.MNU_SCANPATH_COUNTER.Caption = Str(lCount) + "+" + Str(mDirCount) + " Folders"

            End If
        Loop While FindNextFile(lResult, mWFD)
        FindClose lResult
    Loop
'    Form1.MNU_SCANPATH_COUNTER.Caption = Str(lCount) + "+" + Str(mDirCount) + " Folders"
    
    Set colDirs = Nothing
End Function




Public Sub StartScan_USE_NO_LIST(Path As String, Optional SubFolders As Boolean, Optional Sorted As Boolean, Optional PatternMatching As Boolean)
    
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    mDirCount = 0
    mFileCount = 0
    mPatternMatching = PatternMatching
    
    'Verify a filter (file mask) has been specified
    If Len(mFilter) = 0 Then
        ReDim mFilters(0)
        mFilter = ALL_FILES
    ElseIf mFilter <> ALL_FILES Then
        'Append trailing seperator char to Filter (see FileMatches for details)
        mFilter = UCase$(mFilter) & MASK_SEPERATOR_CHAR
        mFilters() = Split(Left$(mFilter, Len(mFilter) - 1), MASK_SEPERATOR_CHAR)
    End If
    
    If (mMinimumSize > 0) And (mMaximumSize > 0) Then
        mSizeCheck = 3
    ElseIf (mMaximumSize > 0) Then
        mSizeCheck = 2
    ElseIf (mMinimumSize > 0) Then
        mSizeCheck = 1
    Else
        mSizeCheck = 0
    End If
    
    If (CDbl(mFromDate) > 0) And (CDbl(mToDate) > 0) Then
        mDateCheck = 3
    ElseIf (CDbl(mToDate) > 0) Then
        mDateCheck = 2
    ElseIf (CDbl(mFromDate) > 0) Then
        mDateCheck = 1
    Else
        mDateCheck = 0
    End If
    
    mCancelScan = False
    
    'USE DIRECT SCAN
    If Sorted Then
        SortedScan Path, SubFolders
    Else
        DirectScan_USE_NO_LIST Path, SubFolders
    End If
End Sub

Public Sub StartScan_USE_NO_LIST_FAST(Path As String, Optional SubFolders As Boolean, Optional Sorted As Boolean, Optional PatternMatching As Boolean)
    
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    mDirCount = 0
    mFileCount = 0
    tFileCount = 0
    mPatternMatching = PatternMatching
    
    'Verify a filter (file mask) has been specified
    If Len(mFilter) = 0 Then
        ReDim mFilters(0)
        mFilter = ALL_FILES
    ElseIf mFilter <> ALL_FILES Then
        'Append trailing seperator char to Filter (see FileMatches for details)
        mFilter = UCase$(mFilter) & MASK_SEPERATOR_CHAR
        mFilters() = Split(Left$(mFilter, Len(mFilter) - 1), MASK_SEPERATOR_CHAR)
    End If
    
    If (mMinimumSize > 0) And (mMaximumSize > 0) Then
        mSizeCheck = 3
    ElseIf (mMaximumSize > 0) Then
        mSizeCheck = 2
    ElseIf (mMinimumSize > 0) Then
        mSizeCheck = 1
    Else
        mSizeCheck = 0
    End If
    
    If (CDbl(mFromDate) > 0) And (CDbl(mToDate) > 0) Then
        mDateCheck = 3
    ElseIf (CDbl(mToDate) > 0) Then
        mDateCheck = 2
    ElseIf (CDbl(mFromDate) > 0) Then
        mDateCheck = 1
    Else
        mDateCheck = 0
    End If
    
    mCancelScan = False
    
    DirectScan_USE_NO_LIST_FAST Path, SubFolders
    
    Exit Sub
    
    'USE DIRECT SCAN
    If Sorted Then
        SortedScan Path, SubFolders
    Else
        DirectScan_USE_NO_LIST_FAST Path, SubFolders
    End If
End Sub

Public Sub StartScan_USE_NO_LIST_FAST_FORM_CAPTURE(Path As String, Optional SubFolders As Boolean, Optional Sorted As Boolean, Optional PatternMatching As Boolean)
    
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    mDirCount = 0
    mFileCount = 0
    tFileCount = 0
    mPatternMatching = PatternMatching
    
    'Verify a filter (file mask) has been specified
    If Len(mFilter) = 0 Then
        ReDim mFilters(0)
        mFilter = ALL_FILES
    ElseIf mFilter <> ALL_FILES Then
        'Append trailing seperator char to Filter (see FileMatches for details)
        mFilter = UCase$(mFilter) & MASK_SEPERATOR_CHAR
        mFilters() = Split(Left$(mFilter, Len(mFilter) - 1), MASK_SEPERATOR_CHAR)
    End If
    
    If (mMinimumSize > 0) And (mMaximumSize > 0) Then
        mSizeCheck = 3
    ElseIf (mMaximumSize > 0) Then
        mSizeCheck = 2
    ElseIf (mMinimumSize > 0) Then
        mSizeCheck = 1
    Else
        mSizeCheck = 0
    End If
    
    If (CDbl(mFromDate) > 0) And (CDbl(mToDate) > 0) Then
        mDateCheck = 3
    ElseIf (CDbl(mToDate) > 0) Then
        mDateCheck = 2
    ElseIf (CDbl(mFromDate) > 0) Then
        mDateCheck = 1
    Else
        mDateCheck = 0
    End If
    
    mCancelScan = False
    
    '-------------------------------------------
    DirectScan_USE_NO_LIST_FAST_FORM_CAPTURE Path, SubFolders
    '-------------------------------------------
    
    Exit Sub
    
    'USE DIRECT SCAN
    If Sorted Then
        SortedScan Path, SubFolders
    Else
        DirectScan_USE_NO_LIST_FAST Path, SubFolders
    End If
End Sub

Public Sub StartScan(Path As String, Optional SubFolders As Boolean, Optional Sorted As Boolean, Optional PatternMatching As Boolean)
    
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    mDirCount = 0
    mFileCount = 0
    
    mPatternMatching = PatternMatching
    
    'Verify a filter (file mask) has been specified
    If Len(mFilter) = 0 Then
        ReDim mFilters(0)
        mFilter = ALL_FILES
    ElseIf mFilter <> ALL_FILES Then
        'Append trailing seperator char to Filter (see FileMatches for details)
        mFilter = UCase$(mFilter) & MASK_SEPERATOR_CHAR
        mFilters() = Split(Left$(mFilter, Len(mFilter) - 1), MASK_SEPERATOR_CHAR)
    End If
    
    If (mMinimumSize > 0) And (mMaximumSize > 0) Then
        mSizeCheck = 3
    ElseIf (mMaximumSize > 0) Then
        mSizeCheck = 2
    ElseIf (mMinimumSize > 0) Then
        mSizeCheck = 1
    Else
        mSizeCheck = 0
    End If
    
    If (CDbl(mFromDate) > 0) And (CDbl(mToDate) > 0) Then
        mDateCheck = 3
    ElseIf (CDbl(mToDate) > 0) Then
        mDateCheck = 2
    ElseIf (CDbl(mFromDate) > 0) Then
        mDateCheck = 1
    Else
        mDateCheck = 0
    End If
    
    mCancelScan = False
    
    If Sorted Then
        SortedScan Path, SubFolders
    Else
        DirectScan Path, SubFolders
    End If
End Sub

Public Sub StartScanDir(Path As String, Optional SubFolders As Boolean, Optional Sorted As Boolean, Optional PatternMatching As Boolean)
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    mDirCount = 0
    mPatternMatching = PatternMatching
    
    'Verify a filter (file mask) has been specified
    If Len(mFilter) = 0 Then
        ReDim mFilters(0)
        mFilter = ALL_FILES
    ElseIf mFilter <> ALL_FILES Then
        'Append trailing seperator char to Filter (see FileMatches for details)
        mFilter = UCase$(mFilter) & MASK_SEPERATOR_CHAR
        mFilters() = Split(Left$(mFilter, Len(mFilter) - 1), MASK_SEPERATOR_CHAR)
    End If
    
    If (mMinimumSize > 0) And (mMaximumSize > 0) Then
        mSizeCheck = 3
    ElseIf (mMaximumSize > 0) Then
        mSizeCheck = 2
    ElseIf (mMinimumSize > 0) Then
        mSizeCheck = 1
    Else
        mSizeCheck = 0
    End If
    
    If (CDbl(mFromDate) > 0) And (CDbl(mToDate) > 0) Then
        mDateCheck = 3
    ElseIf (CDbl(mToDate) > 0) Then
        mDateCheck = 2
    ElseIf (CDbl(mFromDate) > 0) Then
        mDateCheck = 1
    Else
        mDateCheck = 0
    End If
    
    mCancelScan = False
    
    If Sorted Then
        SortedScanDIR Path, SubFolders
    Else
        DirectScanDIR Path, SubFolders
    End If
End Sub


Public Sub StartScanDir_FAST(Path As String, Optional SubFolders As Boolean, Optional Sorted As Boolean, Optional PatternMatching As Boolean)
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    mDirCount = 0
    mPatternMatching = PatternMatching
    
    'Verify a filter (file mask) has been specified
    If Len(mFilter) = 0 Then
        ReDim mFilters(0)
        mFilter = ALL_FILES
    ElseIf mFilter <> ALL_FILES Then
        'Append trailing seperator char to Filter (see FileMatches for details)
        mFilter = UCase$(mFilter) & MASK_SEPERATOR_CHAR
        mFilters() = Split(Left$(mFilter, Len(mFilter) - 1), MASK_SEPERATOR_CHAR)
    End If
    
    If (mMinimumSize > 0) And (mMaximumSize > 0) Then
        mSizeCheck = 3
    ElseIf (mMaximumSize > 0) Then
        mSizeCheck = 2
    ElseIf (mMinimumSize > 0) Then
        mSizeCheck = 1
    Else
        mSizeCheck = 0
    End If
    
    If (CDbl(mFromDate) > 0) And (CDbl(mToDate) > 0) Then
        mDateCheck = 3
    ElseIf (CDbl(mToDate) > 0) Then
        mDateCheck = 2
    ElseIf (CDbl(mFromDate) > 0) Then
        mDateCheck = 1
    Else
        mDateCheck = 0
    End If
    
    mCancelScan = False
    
'    If Sorted Then
 '       SortedScanDIR Path, SubFolders
  '  Else
        DirectScanDIR_FAST Path, SubFolders
   ' End If
End Sub



Private Sub Class_Initialize()
    mArchive = True
    mCompressed = True
    mFolder = True
    mNormal = True
    mReadOnly = True
End Sub


Public Property Get Compressed() As Boolean
    Compressed = mCompressed
End Property

Public Property Let Compressed(ByVal bNewValue As Boolean)
    mCompressed = bNewValue
End Property

Private Function SortedScan(Path As String, SubFolders As Boolean)
    Dim lCount As Long
    Dim lResult As Long
    Dim nCount As Integer
    Dim nFile As Integer
    Dim bOK As Boolean
    Dim sFile As String
    Dim sFiles() As String
    Dim dFiles() As String
    Dim dFile
      
    If mCancelScan Then
        Exit Function
    End If
    
    ReDim sFiles(1 To 10)
    ReDim dFiles(1 To 10)
    
    lResult = FindFirstFile(Path & ALL_FILES, mWFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        Do
            sFile = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
            dFile = Left$(mWFD.cAlternate, InStr(mWFD.cAlternate, Chr$(NULL_CHAR)) - 1)
            If dFile = "" Then dFile = sFile
            
        
            Select Case sFile
            Case ROOT_DIR, PARENT_DIR
            
            Case Else
                bOK = FileMatches(mWFD)
                
                If bOK Then
                    nFile = nFile + 1
                    If nFile > UBound(sFiles) Then
                        ReDim Preserve sFiles(1 To nFile + 10)
                    End If
                    If nFile > UBound(dFiles) Then
                        ReDim Preserve dFiles(1 To nFile + 10)
                    End If
                    
                    If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                        sFiles(nFile) = DIR_KEY & sFile
                        dFiles(nFile) = DIR_KEY & dFile
                  Else
                        sFiles(nFile) = FILE_KEY & sFile
                        dFiles(nFile) = FILE_KEY & dFile
                    End If
                End If
            
            End Select
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                ScanPath.lblCount1 = Trim(Str(lCount))
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
            End If
        Loop While FindNextFile(lResult, mWFD)
        
        FindClose (lResult)
    End If
    
    If nFile > 0 Then
        SortStrings sFiles(), 1, nFile
        
        For nCount = 1 To nFile
            If Mid$(sFiles(nCount), 1, 1) = DIR_KEY Then
                RaiseEvent DirMatch(Mid$(sFiles(nCount), 2), Mid$(dFiles(nCount), 2), Path)
            
                If SubFolders Then
                    mDirCount = mDirCount + 1
                    SortedScan Path & Mid$(sFiles(nCount), 2) & "\", SubFolders
                End If
'            Else
'                mFileCount = mFileCount + 1
'                RaiseEvent FileMatch(Mid$(sFiles(nCount), 2), Mid$(dFiles(nCount), 2), Path)
'                RaiseEvent FileMatch(Mid$(sFiles(nCount), 2), Path)
            End If
            
            If mCancelScan Then
                Exit For
            End If
        Next nCount
    End If
End Function

Private Function SortedScanDIR(Path As String, SubFolders As Boolean)
    Dim lCount As Long
    Dim lResult As Long
    Dim nCount As Integer
    Dim nFile As Integer
    Dim bOK As Boolean
    Dim sFile As String
    Dim sFiles() As String
    Dim dFiles() As String
    Dim Bdir$
    Dim dFile
    
    If mCancelScan Then
        Exit Function
    End If
    
    ReDim sFiles(1 To 10)
    
    lResult = FindFirstFile(Path & ALL_FILES, mWFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        Do
            sFile = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(NULL_CHAR)) - 1)
            dFile = Left$(mWFD.cAlternate, InStr(mWFD.cAlternate, Chr$(NULL_CHAR)) - 1)
            If dFile = "" Then dFile = sFile
            
            Select Case sFile
            Case ROOT_DIR, PARENT_DIR
            
            Case Else
                
                
                bOK = FileMatches(mWFD)
                
                
               If bOK Then
                    nFile = nFile + 1
                    If nFile > UBound(sFiles) Then
                        ReDim Preserve sFiles(1 To nFile + 10)
                    End If
                    If nFile > UBound(sFiles) Then
                        ReDim Preserve dFiles(1 To nFile + 10)
                    End If
                    If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                        sFiles(nFile) = DIR_KEY & sFile
                        dFiles(nFile) = DIR_KEY & dFile
                    Else
                        sFiles(nFile) = FILE_KEY & sFile
                        dFiles(nFile) = FILE_KEY & dFile
                    End If
                End If
                
                
                
                
                
                
            
            End Select
            
            lCount = lCount + 1
            If (lCount Mod 10) = 0 Then
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
            End If
        Loop While FindNextFile(lResult, mWFD)
        
        FindClose (lResult)
    End If
    
    If nFile > 0 Then
        SortStrings sFiles(), 1, nFile
        
        For nCount = 1 To nFile
            If Mid$(sFiles(nCount), 1, 1) = DIR_KEY Then
                RaiseEvent DirMatchxx(Mid$(sFiles(nCount), 2), Mid$(dFiles(nCount), 2), Path)
            
                If SubFolders Then
                    mDirCount = mDirCount + 1
                    SortedScanDIR Path & Mid$(sFiles(nCount), 2) & "\", SubFolders
                End If
'            Else
'                mFileCount = mFileCount + 1
            End If
            
            DoEvents
            
            If mCancelScan Then
                Exit For
            End If
        Next nCount
    End If
End Function


Private Sub SortStrings(sArray() As String, ByVal nFirst As Integer, ByVal nLast As Integer)
    Dim nBoundary As Integer
    Dim nIndex As Integer

    If nLast <= nFirst Then Exit Sub

    SwapS sArray(nFirst), sArray((nFirst + nLast) / 2)
    
    nBoundary = nFirst

    For nIndex = nFirst + 1 To nLast
        If StrComp(sArray(nFirst), sArray(nIndex), vbTextCompare) = 1 Then
            nBoundary = nBoundary + 1
            SwapS sArray(nBoundary), sArray(nIndex)
        End If
    Next nIndex

    SwapS sArray(nFirst), sArray(nBoundary)

    SortStrings sArray(), nFirst, nBoundary - 1
    SortStrings sArray(), nBoundary + 1, nLast
End Sub

Private Sub SwapS(Str1 As String, Str2 As String)
    Dim sTemp As String
    
    sTemp = Str1
    Str1 = Str2
    Str2 = sTemp
End Sub

Public Property Get MinimumSize() As Long
    MinimumSize = mMinimumSize
End Property

Public Property Let MinimumSize(ByVal vNewValue As Long)
    mMinimumSize = vNewValue
End Property

Public Property Get MaximumSize() As Long
    MaximumSize = mMaximumSize
End Property

Public Property Let MaximumSize(ByVal vNewValue As Long)
    mMaximumSize = vNewValue
End Property

Public Property Get FromDate() As Date
    FromDate = mFromDate
End Property

Public Property Let FromDate(ByVal vNewValue As Date)
    mFromDate = vNewValue
End Property

Public Property Get ToDate() As Date
    ToDate = mToDate
End Property

Public Property Let ToDate(ByVal vNewValue As Date)
    mToDate = vNewValue
End Property

Public Property Get DateType() As DateTypeEnum
    DateType = mDateType
End Property

Public Property Let DateType(ByVal vNewValue As DateTypeEnum)
    mDateType = vNewValue
End Property

Public Sub StopScan()
    mCancelScan = True
End Sub

Public Function WFDPointer() As Long
    'This returns a pointer to the current dir/file found and
    'may be used to extract extra details (size, dates etc)
    'without the performancec hit of processing the file again
    
    'NOTE: This only works for DirectScan. SortedScan raises events
    'after sorting the results so the mWFD variable is no longer valid
    
    WFDPointer = VarPtr(mWFD)
End Function
