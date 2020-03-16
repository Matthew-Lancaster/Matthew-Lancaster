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
'Option Explicit
Dim lCount
Dim FoundPath As String

'MSDN HELP FILE NOW WORKING
'--------------------------
'----
'VBP GET WINDOWS VERSION NUMBER - Google Search
'https://www.google.co.uk/search?q=VBP+GET+WINDOWS+VERSION+NUMBER&rlz=1C1CHBD_en-GBGB744GB744&oq=VBP+GET+WINDOWS+VERSION+NUMBER&aqs=chrome..69i57.16343j0j7&sourceid=chrome&ie=UTF-8
'--------
'visual studio - I want to install VB6 on windows 10 OS. I need to work on VBP project so i have to install VB6 seamlessly - Stack Overflow
'https://stackoverflow.com/questions/38169307/i-want-to-install-vb6-on-windows-10-os-i-need-to-work-on-vbp-project-so-i-have
'--------
'Visual Studio Help Engine for MSDN: Enables MSDN menu functionality in Visual FoxPro - Free download and software reviews - CNET Download.com
'http://download.cnet.com/Visual-Studio-Help-Engine-for-MSDN-Enables-MSDN-menu-functionality-in-Visual-FoxPro/3000-2213_4-10727794.html
'----

Dim Old_FilePath

Dim FileSystem As Object
Dim HostFolder As String
Dim FOUND_SET_ONCE
Dim FLAG_NEVER_MESSENGER_BEFORE_01
Dim FLAG_NEVER_MESSENGER_BEFORE_02
Dim Last_Folder


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
'Public mCancelScan2

Private mPatternMatching As Boolean
Private mFilters() As String

Public Event DirMatchxx(Directory As String, DDirectory As String, Path As String)
Public Event DirMatch(Directory As String, DDirectory As String, Path As String)
Public Event FileMatch(Filename As String, DFilename As String, Path As String)




Private Function SortedScan_3(Path As String, SubFolders As Boolean)

    HostFolder = Path
    
    ScanPath.lblCount4 = ""
    LCOL = 50
    LOWER_COLOR = QBColor(15) - RGB(LCOL, LCOL, LCOL)
    
    ScanPath.lblCount1.BackColor = LOWER_COLOR
    ScanPath.lblCount2.BackColor = LOWER_COLOR
    ScanPath.lblCount3.BackColor = LOWER_COLOR
    ScanPath.lblCount4.BackColor = LOWER_COLOR
    ScanPath.lblCount7.BackColor = LOWER_COLOR
    ScanPath.lblCount8.BackColor = LOWER_COLOR
    
'    ScanPath.lblCount4 = "WORK COMPLETE OF ONE DRIVE"
    ScanPath.lblCount1.BackColor = QBColor(10)
    ScanPath.lblCount2.BackColor = QBColor(10)
    ScanPath.lblCount3.BackColor = QBColor(10)
    ScanPath.lblCount4.BackColor = QBColor(10)
    
    Call TIME_SLICER_DISPLAY
    
    COUNTPROCES2 = 0
    
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    '----------------------------------------------------------
    'DoFolder FileSystem.GetFolder(HostFolder)
    '----------------------------------------------------------
    
    Recurse HostFolder, "*.*"
    
    
    ' Call the first directory (root)
    'GetDirectory (HostFolder)
    Close #FF02
    

End Function


'----
'recursive api SUBFOLDERS - Google Search
'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=recursive+api+SUBFOLDERS&oq=recursive+api+SUBFOLDERS&gs_l=psy-ab.3...31961.35448.0.35637.11.11.0.0.0.0.68.641.11.11.0....0...1.1.64.psy-ab..0.10.586...0j0i22i30k1j33i160k1j33i21k1.0.hLwnAIt8vYU
'--------
'performance - Optimize Speed of Recursive File Search in Subdirectories? - Stack Overflow
'https://stackoverflow.com/questions/30511217/optimize-speed-of-recursive-file-search-in-subdirectories
'----
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
Function Recurse(folderPath As String, Filename As String)
    Dim FileHandle    As Long
    Dim searchPattern As String
    Dim FoundItem     As String
    Dim filedata      As WIN32_FIND_DATA

    searchPattern = folderPath & "\*"

    FoundPath = vbNullString
    FileHandle = FindFirstFile(searchPattern, filedata)
    If FileHandle <> INVALID_HANDLE_VALUE Then
        Do
            FoundItem = Left$(filedata.cFileName, InStr(filedata.cFileName, vbNullChar) - 1)

            If FoundItem = "." Or FoundItem = ".." Then 'Skip metadirectories
                'Found Directory
            ElseIf filedata.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                FoundPath = Recurse(folderPath & "\" & FoundItem, Filename)
                mDIRCount2 = mDIRCount2 + 1
                'Found File
                'ElseIf StrComp(foundItem, fileName, vbTextCompare) = 0 Then 'these seem about equal
            Else
            
                'InStr(1, foundItem, fileName, vbTextCompare) > 0 Then 'for performance
                'FoundPath = folderPath & "\" & FoundItem
                FoundPath = folderPath & "\" & FoundItem

                Call Routine_On_File(FoundPath, FoundItem)
                
                Call TIME_SLICER_DISPLAY
            
            End If
'
'            If FoundPath <> vbNullString Then
'                Recurse = FoundPath
'                Exit Function
'            End If

        Loop While FindNextFile(FileHandle, filedata)
    End If

    'No Match Found
    Recurse = vbNullString
End Function

Sub Routine_On_File(FoundPath, FoundItem)

            Err.Clear
            
            FoundPath = Mid(FoundPath, 1, 1) + Replace(Mid(FoundPath, 2), "\\", "\")
            
            'EDIT IN FIX 2017-10-31 04:14 AM
            '-------------------------------
            
            MDIR2 = FoundItem
            mFileCount = mFileCount + 1
            'mFILECount2 = mFILECount2 + 1
            
            bOK = FileMatches_FSO_(FoundItem)
            
            Call TIME_SLICER_DISPLAY
                        
            If bOK Then
                'COUNTPROCES2 = COUNTPROCES2 + 1

                If COUNTPROCES2 = 0 Then TIME_SLICER = True
                'RaiseEvent FileMatch(sFile.Name, "", SubFolder.Path + "\" + SubFolder.Name)
                
                'Call Form_SEND_TO.XSCRIPT(File.Path + "\" + File.Name)
                
                'Err.Clear
                
                'BLACKLIST BLOCKER SCRIPT 02 OF 02
                'BLACKLIST BLOCKER SCRIPT 02 OF 02
                ESCAPE_CODE_1 = True
                If Mid(UCase(HostFolder), 1, 8) = "D:\VIDEO" Then
                    ESCAPE_CODE_1 = False
                End If
                If Mid(UCase(HostFolder), 1, 11) = "D:\V_DSC_ME" Then
                    ESCAPE_CODE_1 = False
                End If
                If Mid(UCase(HostFolder), 1, 12) = "$RECYCLE.BIN" Then
                    ESCAPE_CODE_1 = False
                End If

                ESCAPE_CODE_2 = True
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "D:\VIDEO") Then
                    ESCAPE_CODE_2 = False
                End If
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "D:\V_DSC_ME") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = ESCAPE_CODE_1 And InStr(UCase(FoundPath), "_GSDATA_") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "$RECYCLE.BIN") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "\ISO\") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "\ISO_IMAGE_DISK\") Then
                    ESCAPE_CODE_2 = False
                End If
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "\0 00 MOBILE\") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "\IMAGES\") Then
                    If InStr(UCase(FoundPath), ".ISO") > 0 Then
                        ESCAPE_CODE_2 = False
                    End If
                End If
                
                'ENDS IN THIS LINE
                '-----------------
                If InStr(UCase(FoundPath), "NOT_ONLINE DOC\") Then
                    ESCAPE_CODE_2 = False
                End If
                
'                NOT_ONLINE DOC
                
                
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "MICROSOFT") Then
                    If InStr(UCase(FoundPath), ".ISO") > 0 Then
                        ESCAPE_CODE_2 = False
                    End If
                End If
                If ESCAPE_CODE_1 = True And InStr(UCase(FoundPath), "INSTALLATION") Then
                    If InStr(UCase(FoundPath), ".ISO") > 0 Then
                        ESCAPE_CODE_2 = False
                    End If
                End If
                
                
                'ERR.DESCRIPTION
                'GETTING ERROR SUBSCRIPT OUT OF RANGE
                'WHY NOT ON ERROR SET YET
                '----------------------------------------------------------------
                If Err.Number > 0 Then ESCAPE_CODE_2 = False
                
                If ESCAPE_CODE_2 = True Then
                
                    On Error Resume Next
                            Err.Clear
                            Print #FF02, FoundPath
                            Call ScanPath.ROUTINE_ADD_TO_LISTVEW(FoundPath)

                            COUNTPROCES2 = COUNTPROCES2 + 1
                        
                            If Err.Number > 0 Then
                                '-----------------------------------------------------------------
                                'MAYBE THE FILE WAS UNLOCKED SET UP AGAIN ONCE MORE AS NEED ARRIVE
                                'DONT STAY IN LOOP OR MESSENEGER CRASH ON ERROR
                                '-----------------------------------------------------------------
                                Err.Clear
                                FF02 = FreeFile ' LESS MODIFY
                                Open FILE12 For Append As #FF02
                                Print #FF02, FoundPath
                                Call ScanPath.ROUTINE_ADD_TO_LISTVEW(FoundPath)
                                If Err.Number > 0 Then
                                    MsgBox "ERROR MULTIPLE RETRY _ KEEP FILE OPEN " + vbCrLf + vbCrLf + File.Path + vbCrLf + vbCrLf + "CRASH IS HAPPEN NEXT AFTER" + vbCrLf + vbCrLf + "OR CONTINUE DROP A FILE FROM PLAYLIST" + vbCrLf + Time$, vbMsgBoxSetForeground, App.EXEName
                                End If
                            End If
                            Call TIME_SLICER_DISPLAY
                        'End If
                
                    
                    On Error GoTo 0

                    
                End If
        
            End If
            
            If mCancelScan2 Then
                mCancelScan = mCancelScan2
                Exit Sub
            End If


End Sub

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

Sub TIME_SLICER_DISPLAY()
            
            lCount = lCount + 1
            If (lCount Mod 2000) = 0 Then
                DoEvents
                
                '-------------------
                TIME_SLICER = True
                'PLUS SET BY A TIMER
                '-------------------
            End If
            
            '-----------------------------------------------------------------------
            '-----------------------------------------------------------------------
            If TIME_SLICER = True Then
            
                TIME_SLICER = False
                mFILECount2 = mFILECount2 + 1
                'ScanPath.lblCount1 = Trim(Str(mFILECount2) + "   FILE")
                ScanPath.lblCount1 = Format(mFileCount, "###,###,###,###,###,###,0") + "  FILE"
                ScanPath.lblCount2 = Format(mDIRCount2, "###,###,###,###,###,###,0") + "  DIR"
                ScanPath.lblCount5 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + "  FOUNDING PROCESSOR"
                DoEvents
            End If
            
'            If COUNTPROCES2 = 0 And FOUND_SET_ONCE = FLASE Then
'                ScanPath.lblCount5 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + " FOUNDING PROCESS"
'                FOUND_SET_ONCE = True
'            End If
            

            
            i4 = DateDiff("s", PROCESSBEGIN, Now) / 60
            ScanPath.lblCount4 = Format(i4, "0.000") + " MIN"
            
'            ScanPath.lblCount7 = Replace(MDIR2, "\", "\ ")
'            ScanPath.lblCount8 = Replace(MDIR3, "\", "\ ")
'            If ScanPath.lblCount7 = ScanPath.lblCount8 Then
'                ScanPath.lblCount7.BackColor = ScanPath.Label_BACK_COLOR2.BackColor
'                ScanPath.lblCount8.BackColor = ScanPath.Label_BACK_COLOR2.BackColor
'            Else
'                ScanPath.lblCount7.BackColor = ScanPath.Label_BACK_COLOR1.BackColor
'                ScanPath.lblCount8.BackColor = ScanPath.Label_BACK_COLOR1.BackColor
'            End If
            '-----------------------------------------------------------------------
            '-----------------------------------------------------------------------
            'DoEvents
            
            TAP = TAP + 1
            If TAP = 1 Then YCH = "-"
            If TAP = 2 Then YCH = "/"
            If TAP = 3 Then YCH = "|"
            If TAP = 4 Then YCH = "\": TAP = 0
            
            ScanPath.lblCount3 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + " " + YCH

End Sub




Function GetDirectory(strDirectory)

    ' Switch off error checking
    On Error Resume Next

    ' Create FS object
    Dim oFSO
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    ' Get current folder
    Dim StrFolder
    Set StrFolder = oFSO.GetFolder(strDirectory)

    ' Display current folder
    WScript.Echo StrFolder.Path

    ' Cycle through all subfolders for current Folder
    Dim currentsubfolder
    For Each currentsubfolder In StrFolder.SubFolders

        ' Call recursive func to get subfolders for this folder
        GetDirectory currentsubfolder

    Next

    ' Set the pointer to nothing
    Set oFSO = Nothing

End Function



Public Property Get DirCount() As Long
    DirCount = mDirCount
End Property


Public Property Get FileCount() As Long
    FileCount = mFileCount
End Property



Private Function FileMatches_FSO_(FILENAME_PASSER) As Boolean
    '##############################################################################################
    'Purpose: Used to verify a file matches the Filter Criteria
    '##############################################################################################
    
    Dim dtFileDate As Date
    Dim nFilter As Integer
    Dim nPos As Integer
    Dim bFolder As Boolean
    Dim bOK As Boolean
    Dim sExt As String
    
    'bFolder = (udtWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
    
    'If bFolder Or (mFilter = ALL_FILES) Then
    '    bOK = True
    'Else
        'Remove the padding from filename by truncating at first null char
        
        
        sExt = UCase$(FILENAME_PASSER)
        
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
    
    
    'End If
    
    FileMatches_FSO_ = bOK

End Function



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
    
    
    'FileMatches = bOK
    
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
       
'    ct.dwHighDateTime
       
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

Public Property Get Folder() As Boolean
    Folder = mFolder
End Property

Public Property Let Folder(ByVal bNewValue As Boolean)
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
    
    ScanPath.lblCount4 = ""
    LCOL = 50
    LOWER_COLOR = QBColor(15) - RGB(LCOL, LCOL, LCOL)
    ScanPath.lblCount1.BackColor = LOWER_COLOR
    ScanPath.lblCount2.BackColor = LOWER_COLOR
    ScanPath.lblCount3.BackColor = LOWER_COLOR
    ScanPath.lblCount4.BackColor = LOWER_COLOR
    ScanPath.lblCount7.BackColor = LOWER_COLOR
    ScanPath.lblCount8.BackColor = LOWER_COLOR
    
    Do While colDirs.Count And Not mCancelScan
        If mCancelScan2 Then
            mCancelScan = mCancelScan2: Exit Function
        End If
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
            
'            If sFile = ROOT_DIR Then Stop
'            If sFile = PARENT_DIR Then Stop

            Select Case sFile
            Case ROOT_DIR, PARENT_DIR
            Case Else

                bOK = FileMatches(mWFD)
                If SubFolders = False Then
                    If bOK = True Then Stop
                End If
                COUNTPROCES2 = COUNTPROCES2 + 1
                MDIR3 = sDir + sFile ' IRFAN
                
                If MDIR3 = "D:\DSC" Then Stop
                'If InStr(4, MDIR3, "\") > 0 Then Stop
                'If InStr(MDIR3, "D:\DSC") > 0 And InStr(4, MDIR3, "\") > 0 Then Stop
                
                If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    If bOK Then
                        If SubFolders Then
                            colDirs.Add sDir & sFile & "\"
                        End If
                    
                        mDirCount = mDirCount + 1
                        mDIRCount2 = mDirCount
                        
                        RaiseEvent DirMatch(sFile, dFile, sDir)
                    End If
                Else
                    If bOK Then
                        mFileCount = mFileCount + 1
                        mFILECount2 = mFileCount
                        
                        MDIR2 = sDir + sFile ' IRFAN
                        RaiseEvent FileMatch(sFile, dFile, sDir)
                    
                    End If
                End If
                
            End Select
                
            If mCancelScan2 Then
                Exit Do
            End If
            
            
            lCount = lCount + 1
            If (lCount Mod 100) = 0 Then
                DoEvents
                If mCancelScan Then
                    Exit Do
                End If
            End If
        Loop While FindNextFile(lResult, mWFD)
        FindClose lResult
    Loop
    
'    ScanPath.lblCount1 = mFileCount
    
    DoEvents
    
    ScanPath.lblCount2 = Trim(Str(mDirCount) + "   DIR")
    ScanPath.lblCount1 = Trim(Str(mFileCount) + "   FILE")
    
    If COUNTPROCES2 > 0 Then
        ScanPath.lblCount3 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + "  PROCESS"
    End If
    
'    ScanPath.lblCount4 = "WORK COMPLETE OF ONE DRIVE"
    ScanPath.lblCount1.BackColor = QBColor(10)
    ScanPath.lblCount2.BackColor = QBColor(10)
    ScanPath.lblCount3.BackColor = QBColor(10)
    ScanPath.lblCount4.BackColor = QBColor(10)
'    ScanPath.lblCount7.BackColor = QBColor(10)
'    ScanPath.lblCount8.BackColor = QBColor(10)
    DoEvents
    
    Set colDirs = Nothing
End Function




Private Function SortedScan_2(Path As String, SubFolders As Boolean)
    Dim lCount As Long
    Dim lResult As Long
    Dim nCount As Integer
    Dim nFile As Integer
    Dim bOK As Boolean
    Dim sFile As String
    Dim sDir As String
    Dim dFile As String
    Dim sFiles() As String
    Dim dFiles() As String
    
      
    ScanPath.lblCount4 = ""
    LCOL = 50
    LOWER_COLOR = QBColor(15) - RGB(LCOL, LCOL, LCOL)
    ScanPath.lblCount1.BackColor = LOWER_COLOR
    ScanPath.lblCount2.BackColor = LOWER_COLOR
    ScanPath.lblCount3.BackColor = LOWER_COLOR
    ScanPath.lblCount4.BackColor = LOWER_COLOR
    ScanPath.lblCount7.BackColor = LOWER_COLOR
    ScanPath.lblCount8.BackColor = LOWER_COLOR
      
      
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
                    
                    COUNTPROCES2 = COUNTPROCES2 + 1
                    MDIR3 = sDir + sFile ' IRFAN
                    
                    If MDIR3 = "D:\DSC" Then Stop
                    'If InStr(4, MDIR3, "\") > 0 Then Stop
                    'If InStr(MDIR3, "D:\DSC") > 0 And InStr(4, MDIR3, "\") > 0 Then Stop

                    
                    If nFile > UBound(sFiles) Then
                        ReDim Preserve sFiles(1 To nFile + 10)
                    End If
                    If nFile > UBound(dFiles) Then
                        ReDim Preserve dFiles(1 To nFile + 10)
                    End If
                    
                    If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                        sFiles(nFile) = DIR_KEY & sFile
                        dFiles(nFile) = DIR_KEY & dFile
                        
'                        mDirCount = mDirCount + 1
'                        mDIRCount2 = mDirCount
                        
                        'RaiseEvent DirMatch(sFile, dFile, sDir)
                        
                  Else
                        sFiles(nFile) = FILE_KEY & sFile
                        dFiles(nFile) = FILE_KEY & dFile
                        
                        mFileCount = mFileCount + 1
                        mFILECount2 = mFileCount
                        
                        MDIR2 = sDir + sFile ' IRFAN
                        'RaiseEvent FileMatch(sFile, dFile, sDir)
                        
                        
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
    
    
    DoEvents
'
'    ScanPath.lblCount2 = Trim(Str(mDirCount) + "   DIR")
'    ScanPath.lblCount1 = Trim(Str(mFileCount) + "   FILE")
'
    If COUNTPROCES2 > 0 Then
        ScanPath.lblCount3 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + "  PROCESS"
    End If
    
    If nFile > 0 Then
        SortStrings sFiles(), 1, nFile
        
        For nCount = 1 To nFile
            If Mid$(sFiles(nCount), 1, 1) = DIR_KEY Then
                RaiseEvent DirMatch(Mid$(sFiles(nCount), 2), Mid$(dFiles(nCount), 2), Path)
            
                If SubFolders Then
                    mDirCount = mDirCount + 1
                    mDIRCount2 = mDirCount
                    SortedScan_2 Path & Mid$(sFiles(nCount), 2) & "\", SubFolders
                End If
            Else
                mFileCount = mFileCount + 1
                'MDIR2 = sDir + sFile ' IRFAN
                MDIR2 = Path + Mid$(sFiles(nCount), 2)
                RaiseEvent FileMatch(Mid$(sFiles(nCount), 2), Mid$(dFiles(nCount), 2), Path)
            End If
            
            If mCancelScan Then
                Exit For
            End If
        Next nCount
    End If
    
    
    ScanPath.lblCount2 = Trim(Str(mDirCount) + "   DIR")
    ScanPath.lblCount1 = Trim(Str(mFileCount) + "   FILE")
    
    If COUNTPROCES2 > 0 Then
        ScanPath.lblCount3 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + "  PROCESS"
    End If
    
'    ScanPath.lblCount4 = "WORK COMPLETE OF ONE DRIVE"
    ScanPath.lblCount1.BackColor = QBColor(10)
    ScanPath.lblCount2.BackColor = QBColor(10)
    ScanPath.lblCount3.BackColor = QBColor(10)
    ScanPath.lblCount4.BackColor = QBColor(10)
    
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
    
    If Sorted Then
        SortedScan Path, SubFolders
    Else
        DirectScan_USE_NO_LIST Path, SubFolders
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
        SortedScan_3 Path, SubFolders
    Else
        'DirectScan Path, SubFolders
        DirectScan_USE_NO_LIST Path, SubFolders
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
    Dim sDir As String
    Dim dFile As String
    Dim sFiles() As String
    Dim dFiles() As String
    
      
    ScanPath.lblCount4 = ""
    LCOL = 50
    LOWER_COLOR = QBColor(15) - RGB(LCOL, LCOL, LCOL)
    ScanPath.lblCount1.BackColor = LOWER_COLOR
    ScanPath.lblCount2.BackColor = LOWER_COLOR
    ScanPath.lblCount3.BackColor = LOWER_COLOR
    ScanPath.lblCount4.BackColor = LOWER_COLOR
    ScanPath.lblCount7.BackColor = LOWER_COLOR
    ScanPath.lblCount8.BackColor = LOWER_COLOR
      
      
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
                    
                    COUNTPROCES2 = COUNTPROCES2 + 1
                    MDIR3 = sDir + sFile ' IRFAN
                    
                    If MDIR3 = "D:\DSC" Then Stop
                    'If InStr(4, MDIR3, "\") > 0 Then Stop
                    'If InStr(MDIR3, "D:\DSC") > 0 And InStr(4, MDIR3, "\") > 0 Then Stop

                    
                    If nFile > UBound(sFiles) Then
                        ReDim Preserve sFiles(1 To nFile + 10)
                    End If
                    If nFile > UBound(dFiles) Then
                        ReDim Preserve dFiles(1 To nFile + 10)
                    End If
                    
                    If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                        sFiles(nFile) = DIR_KEY & sFile
                        dFiles(nFile) = DIR_KEY & dFile
                        
'                        mDirCount = mDirCount + 1
'                        mDIRCount2 = mDirCount
                        
                        'RaiseEvent DirMatch(sFile, dFile, sDir)
                        
                  Else
                        sFiles(nFile) = FILE_KEY & sFile
                        dFiles(nFile) = FILE_KEY & dFile
                        
                        mFileCount = mFileCount + 1
                        mFILECount2 = mFileCount
                        
                        MDIR2 = sDir + sFile ' IRFAN
                        'RaiseEvent FileMatch(sFile, dFile, sDir)
                        
                        
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
    
    
    DoEvents
'
'    ScanPath.lblCount2 = Trim(Str(mDirCount) + "   DIR")
'    ScanPath.lblCount1 = Trim(Str(mFileCount) + "   FILE")
'
    If COUNTPROCES2 > 0 Then
        ScanPath.lblCount3 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + "  PROCESS"
    End If
    
    If nFile > 0 Then
        SortStrings sFiles(), 1, nFile
        
        For nCount = 1 To nFile
            If Mid$(sFiles(nCount), 1, 1) = DIR_KEY Then
                RaiseEvent DirMatch(Mid$(sFiles(nCount), 2), Mid$(dFiles(nCount), 2), Path)
            
                If SubFolders Then
                    mDirCount = mDirCount + 1
                    mDIRCount2 = mDirCount
                    SortedScan Path & Mid$(sFiles(nCount), 2) & "\", SubFolders
                End If
            Else
                mFileCount = mFileCount + 1
                'MDIR2 = sDir + sFile ' IRFAN
                MDIR2 = Path + Mid$(sFiles(nCount), 2)
                RaiseEvent FileMatch(Mid$(sFiles(nCount), 2), Mid$(dFiles(nCount), 2), Path)
            End If
            
            If mCancelScan Then
                Exit For
            End If
        Next nCount
    End If
    
    
        ScanPath.lblCount2 = Trim(Str(mDirCount) + "   DIR")
    ScanPath.lblCount1 = Trim(Str(mFileCount) + "   FILE")
    
    If COUNTPROCES2 > 0 Then
        ScanPath.lblCount3 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + "  PROCESS"
    End If
    
'    ScanPath.lblCount4 = "WORK COMPLETE OF ONE DRIVE"
    ScanPath.lblCount1.BackColor = QBColor(10)
    ScanPath.lblCount2.BackColor = QBColor(10)
    ScanPath.lblCount3.BackColor = QBColor(10)
    ScanPath.lblCount4.BackColor = QBColor(10)
    
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




Sub DoFolder(strDirectory)
    
        Dim SubFolder
        Dim lCount As Long
        'ERR.NUMBER
        On Error Resume Next
        Err.Clear
        
        Dim oFSO
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Dim StrFolder
        Set StrFolder = oFSO.GetFolder(strDirectory)
        
        For Each SubFolder In StrFolder.SubFolders
                
                
            'If Err.Number = 70 And FLAG_NEVER_MESSENGER_BEFORE_01 = False Then
            If Err.Number = 70 Then
                FLAG_NEVER_MESSENGER_BEFORE_01 = True
                FLAG_NEVER_MESSENGER_BEFORE_02 = True
                Debug.Print Time$
                Debug.Print Time$ + " " + "FILE COUNT WITH ERROR @ " + Str(mFileCount)
                Debug.Print Time$ + " ERROR PERMISSION 01 OF 02 __ " + Last_Folder
'                MsgBox "PERMISSION ERROR AT LAST FOLDER " + vbCrLf + vbCrLf + Last_Folder + vbCrLf + vbCrLf + "MORE INFO IN DEBUGGING _ AND " + vbCrLf + vbCrLf + "FILE COUNT WITH ERROR @ " + Str(mFileCount)
                ' RESUME NEXT FOLLOWS THROUGH THE CODE LANDING A LOT OF ERROR SO EXIT SUB GOTO
'                GoTo JUMP_UP_TO_NEXT
            End If
            If Err.Number = 0 Then
                If SubFolder.Path <> "" Then Last_Folder = SubFolder.Path
            End If
            If Err.Number > 0 And Err.Number <> 70 Then
                'Stop
                Debug.Print Time$ + "_" + Str(Err.Number)
                Debug.Print Time$ + "  " + Err.Description
                Debug.Print Time$ + "  " + Last_Folder
            End If
            'err.description 424 -- Object Require
            DoEvents
                
            If Err.Number = 0 Then
                MDIR3 = SubFolder.Path + "\" + SubFolder.Name
                mDirCount = mDirCount + 1
                mDIRCount2 = mDirCount
                'ScanPath.lblCount2 = Trim(Str(mDirCount) + "   DIR")
                'Debug.Print ScanPath.lblCount2.Caption & "  " & SubFolder.Name
                TIME_SLICER = True
            End If
            
            Call TIME_SLICER_DISPLAY

                
            DoFolder SubFolder
            'ERROR 92 ERR.DESCRIPTION __ "For loop not initialized"
            If mCancelScan2 Then Exit For
'        Next
        
        If mCancelScan2 Then Exit Sub
        
        Dim File
        For Each File In StrFolder.Files
            ' Operate on each file
            MDIR2 = File.Name
            mFileCount = mFileCount + 1
            'mFILECount2 = mFILECount2 + 1
            bOK = FileMatches_FSO_(File.Name)
            
            Call TIME_SLICER_DISPLAY
                        
            TAP = TAP + 1
            If TAP = 1 Then YCH = "-"
            If TAP = 2 Then YCH = "/"
            If TAP = 3 Then YCH = "|"
            If TAP = 4 Then YCH = "\": TAP = 0
            
            ScanPath.lblCount3 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + " " + YCH

                        
            If bOK Then
                'COUNTPROCES2 = COUNTPROCES2 + 1

                If COUNTPROCES2 = 0 Then TIME_SLICER = True
                'RaiseEvent FileMatch(sFile.Name, "", SubFolder.Path + "\" + SubFolder.Name)
                
                'Call Form_SEND_TO.XSCRIPT(File.Path + "\" + File.Name)
                
                
                'BLACKLIST BLOCKER SCRIPT 02 OF 02
                'BLACKLIST BLOCKER SCRIPT 02 OF 02
                ESCAPE_CODE_1 = True
                If Mid(UCase(HostFolder), 1, 8) = "D:\VIDEO" Then
                    ESCAPE_CODE_1 = False
                End If
                If Mid(UCase(HostFolder), 1, 11) = "D:\V_DSC_ME" Then
                    ESCAPE_CODE_1 = False
                End If
                If Mid(UCase(HostFolder), 1, 12) = "$RECYCLE.BIN" Then
                    ESCAPE_CODE_1 = False
                End If

                ESCAPE_CODE_2 = True
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "D:\VIDEO") Then
                    ESCAPE_CODE_2 = False
                End If
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "D:\V_DSC_ME") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = ESCAPE_CODE_1 And InStr(UCase(File.Path), "_GSDATA_") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "$RECYCLE.BIN") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "\ISO\") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "\ISO_IMAGE_DISK\") Then
                    ESCAPE_CODE_2 = False
                End If
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "\0 00 MOBILE\") Then
                    ESCAPE_CODE_2 = False
                End If
                
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "\IMAGES\") Then
                    If InStr(UCase(File.Name), ".ISO") > 0 Then
                        ESCAPE_CODE_2 = False
                    End If
                End If
                
                'ENDS IN THIS LINE
                '-----------------
                If InStr(UCase(File.Path), "NOT_ONLINE DOC\") Then
                    ESCAPE_CODE_2 = False
                End If
                
'                NOT_ONLINE DOC
                
                
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "MICROSOFT") Then
                    If InStr(UCase(File.Name), ".ISO") > 0 Then
                        ESCAPE_CODE_2 = False
                    End If
                End If
                If ESCAPE_CODE_1 = True And InStr(UCase(File.Path), "INSTALLATION") Then
                    If InStr(UCase(File.Name), ".ISO") > 0 Then
                        ESCAPE_CODE_2 = False
                    End If
                End If
                
                If Err.Number > 0 Then ESCAPE_CODE_2 = False
                
                If ESCAPE_CODE_2 = True Then
                
                
                    On Error Resume Next
                        If File.Path <> Old_FilePath Then
                            Err.Clear
                            Print #FF02, File.Path
                            COUNTPROCES2 = COUNTPROCES2 + 1
                            Old_FilePath = File.Path
                            Debug.Print Old_FilePath
                        
                            If Err.Number > 0 Then
                                '-----------------------------------------------------------------
                                'MAYBE THE FILE WAS UNLOCKED SET UP AGAIN ONCE MORE AS NEED ARRIVE
                                'DONT STAY IN LOOP OR MESSENEGER CRASH ON ERROR
                                '-----------------------------------------------------------------
                                Err.Clear
                                FF02 = FreeFile ' LESS MODIFY
                                Open FILE12 For Append As #FF02
                                Print #FF02, File.Path
                                If Err.Number > 0 Then
                                    MsgBox "ERROR MULTIPLE RETRY _ KEEP FILE OPEN " + vbCrLf + vbCrLf + File.Path + vbCrLf + vbCrLf + "CRASH IS HAPPEN NEXT AFTER" + vbCrLf + vbCrLf + "OR CONTINUE DROP A FILE FROM PLAYLIST" + vbCrLf + Time$, vbMsgBoxSetForeground, App.EXEName
                                End If
                            End If
                            Call TIME_SLICER_DISPLAY
                        End If
                
                    
                    On Error GoTo 0

                    
                End If
        
            End If
            
            If mCancelScan2 Then
                mCancelScan = mCancelScan2
                Exit Sub
            End If
            
            
            
            
        Next
       Next
JUMP_UP_TO_NEXT:

    'ScanPath.lblCount2 = Trim(Str(mDirCount) + "   DIR")
    
    
    
'ScanPath.lblCount1 = Format(mFileCount, "###,###,###,###,###,###,0") + "  FILE"
'ScanPath.lblCount2 = Format(mDIRCount2, "###,###,###,###,###,###,0") + "  DIR"
'ScanPath.lblCount5 = Format(COUNTPROCES2, "###,###,###,###,###,###,0") + "  FOUNDING PROCESSOR"


'----
'ASP FileSystem Object
'https://www.w3schools.com/asp/asp_ref_filesystem.asp
'----
'----
'VBP filesystemobject recursive folders - Google Search
'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=VBP+filesystemobject+recursive+folders&oq=VBP+filesystemobject+recursive+folders&gs_l=psy-ab.3..0i8i13i30k1l2.11342.11797.0.12604.4.4.0.0.0.0.72.269.4.4.0....0...1.1.64.psy-ab..0.4.268...0i8i7i30k1j0i13i30k1.0.Z-NUZwasRaw
'----
'----
'filesystemobject - Loop Through All Subfolders Using VBA - Stack Overflow
'https://stackoverflow.com/questions/22645347/loop-through-all-subfolders-using-vba
'----
'Wed 20 September 2017 19:00:50----------
'JUMP_UP_TO_NEXT:

EXIT_SUB_ROUTINER:


End Sub


