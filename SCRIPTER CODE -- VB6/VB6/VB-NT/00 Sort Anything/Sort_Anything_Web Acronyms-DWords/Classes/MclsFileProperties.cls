VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mclsFileProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//******************************************
'// Author:   Chong Long Choo               *
'//******************************************
'// Use this module for your apps, I've made only 10% of this code

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800

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

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const MAX_PATH = 260

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Const SHGFI_ICON = &H100                         '  get icon
Private Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Private Const SHGFI_TYPENAME = &H400                     '  get type name
Private Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Private Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Private Const SHGFI_EXETYPE = &H2000                     '  return exe type
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Private Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Private Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const SHGFI_OPENICON = &H2                       '  get open icon
Private Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Private Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Private Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private mvarCompanyName As String
Private mvarFileDescription As String
Private mvarFileVersion As String
Private mvarInternalName As String
Private mvarLegalCopyright As String
Private mvarOriginalFileName As String
Private mvarProductName As String
Private mvarProductVersion As String
Private mvarFileName As String
Private mvarByte As String
Private mvarCreationTime As String
Private mvarLastAccessTime As String
Private mvarLastWriteTime As String
Private mvarReadOnly As Boolean
Private mvarHidden As Boolean
Private mvarSystem As Boolean
Private mvarDirectory As Boolean
Private mvarArchive As Boolean
Private mvarNormal As Boolean
Private mvarTemporary As Boolean
Private mvarCompress As Boolean
Private mvarFileType As String
Private mvarIconIndex As Long

Public Property Get IconIndex() As Long
    IconIndex = mvarIconIndex
End Property

Public Property Get FileType() As String
    FileType = mvarFileType
End Property

Public Property Get Compress() As Boolean
    Compress = mvarCompress
End Property

Public Property Get Temporary() As Boolean
    Temporary = mvarTemporary
End Property

Public Property Get Normal() As Boolean
    Normal = mvarNormal
End Property

Public Property Get Archive() As Boolean
    Archive = mvarArchive
End Property

Public Property Get Directory() As Boolean
    Directory = mvarDirectory
End Property

Public Property Get System() As Boolean
    System = mvarSystem
End Property

Public Property Get Hidden() As Boolean
    Hidden = mvarHidden
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mvarReadOnly
End Property

Public Property Get LastWriteTime() As String
    LastWriteTime = mvarLastWriteTime
End Property

Public Property Get LastAccessTime() As String
    LastAccessTime = mvarLastAccessTime
End Property

Public Property Get CreationTime() As String
    CreationTime = mvarCreationTime
End Property

Public Property Get ByteSize() As String
    ByteSize = mvarByte
End Property

Public Property Get ProductVersion() As String
    ProductVersion = mvarProductVersion
End Property

Public Property Get ProductName() As String
    ProductName = mvarProductName
End Property

Public Property Get OriginalFileName() As String
    OriginalFileName = mvarOriginalFileName
End Property

Public Property Get LegalCopyright() As String
    LegalCopyright = mvarLegalCopyright
End Property

Public Property Get InternalName() As String
    InternalName = mvarInternalName
End Property

Public Property Get FileVersion() As String
    FileVersion = mvarFileVersion
End Property

Public Property Get FileDescription() As String
    FileDescription = mvarFileDescription
End Property

Public Property Get CompanyName() As String
    CompanyName = mvarCompanyName
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Function FindFileInfo(strFileName As String, bLargeIcon As Boolean) As Long
    'On Error GoTo GetFileVersionData_Error

    Dim sInfo As String, lSizeof As Long
    Dim lResult As Long, intDel As Integer
    Dim lHandle As Long
    Dim ftime As SYSTEMTIME
    Dim filedata As WIN32_FIND_DATA
    Dim intStrip As Integer
    Dim SHFI As SHFILEINFO
    Dim lSizeSHFI As Long
    Dim lFlags As Long

    If strFileName <> "" Then
        
        ' Get CreationTime, LastWriteTime and LastAccessTime
        filedata = Findfile(strFileName)
        Call FileTimeToSystemTime(filedata.ftCreationTime, ftime)
        mvarCreationTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
        Call FileTimeToSystemTime(filedata.ftLastWriteTime, ftime)  ' Determine Last Modified date and time
        mvarLastWriteTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
        Call FileTimeToSystemTime(filedata.ftLastAccessTime, ftime) ' Determine Last accessed date (note no time is recorded)
        mvarLastAccessTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear
        
        ' Get file's attributes
        mvarHidden = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN)
        mvarSystem = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM)
        mvarReadOnly = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY)
        mvarArchive = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE)
        mvarTemporary = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY)
        mvarNormal = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL)
        mvarCompress = ((filedata.dwFileAttributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED)
        mvarFileName = UCase$(strFileName)
        
        ' Get size of the file
        If filedata.nFileSizeHigh = 0 Then
            mvarByte = Format$(filedata.nFileSizeLow, "###,###,###")
        Else
            mvarByte = Format$(filedata.nFileSizeHigh, "###,###,###")
        End If
        
        ' Get CompanyName, FileDescription, FileVersion, InternalName
        ' LegalCopyright, OriginalFilename, ProductName, ProductVersion
        lHandle = 0
        lSizeof = GetFileVersionInfoSize(strFileName, lHandle)
        If lSizeof > 0 Then
            sInfo = String$(lSizeof, 0)
            lResult = GetFileVersionInfo(ByVal strFileName, 0&, ByVal lSizeof, ByVal sInfo)
            If lResult Then
                intDel = InStr(sInfo, "CompanyName")
                If intDel > 0 Then
                    intDel = intDel + 12
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    mvarCompanyName = Mid$(sInfo, intDel, intStrip - intDel)
                End If
                
                intDel = InStr(sInfo, "FileDescription")
                If intDel > 0 Then
                    intDel = intDel + 16
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    mvarFileDescription = Mid$(sInfo, intDel, intStrip - intDel)
                End If
                
                intDel = InStr(sInfo, "FileVersion")
                If intDel > 0 Then
                    intDel = intDel + 12
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    mvarFileVersion = Mid$(sInfo, intDel, intStrip - intDel)
                End If
                
                intDel = InStr(sInfo, "InternalName")
                If intDel > 0 Then
                    intDel = intDel + 16
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    mvarInternalName = Mid$(sInfo, intDel, intStrip - intDel)
                End If
                
                intDel = InStr(sInfo, "LegalCopyright")
                If intDel > 0 Then
                    intDel = intDel + 16
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    mvarLegalCopyright = Mid$(sInfo, intDel, intStrip - intDel)
                End If
                
                intDel = InStr(sInfo, "OriginalFilename")
                If intDel > 0 Then
                    intDel = intDel + 20
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    mvarOriginalFileName = Mid$(sInfo, intDel, intStrip - intDel)
                End If

                intDel = InStr(sInfo, "ProductName")
                If intDel > 0 Then
                    intDel = intDel + 12
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    mvarProductName = Mid$(sInfo, intDel, intStrip - intDel)
                End If

                intDel = InStr(sInfo, "ProductVersion")
                If intDel > 0 Then
                    intDel = intDel + 16
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    mvarProductVersion = Mid$(sInfo, intDel, intStrip - intDel)
                End If
            End If
        End If
        
        ' Get file's type and the index of icon of the file
        lSizeSHFI = Len(SHFI)
        lFlags = SHGFI_SYSICONINDEX Or SHGFI_TYPENAME
        If bLargeIcon Then
            lFlags = lFlags Or SHGFI_LARGEICON
        Else
            lFlags = lFlags Or SHGFI_SMALLICON
        End If
        SHGetFileInfo strFileName, 0&, SHFI, lSizeSHFI, lFlags
        mvarFileType = Left$(SHFI.szTypeName, InStr(1, SHFI.szTypeName, vbNullChar) - 1)
        mvarIconIndex = SHFI.iIcon
        
        FindFileInfo = 1
    Else
        FindFileInfo = 0
    End If
    
GetFileVersionData_Error:
    FindFileInfo = 0
    Exit Function
    
invalid_file_info_error:
    FindFileInfo = 1
    Exit Function
End Function
Public Function FindFileInfoView(strFileName As String, bLargeIcon As Boolean) As Long
    'On Error GoTo GetFileVersionData_Error

    Dim sInfo As String, lSizeof As Long
    Dim lResult As Long, intDel As Integer
    Dim lHandle As Long
    Dim ftime As SYSTEMTIME
    Dim filedata As WIN32_FIND_DATA
    Dim intStrip As Integer
    Dim SHFI As SHFILEINFO
    Dim lSizeSHFI As Long
    Dim lFlags As Long

    If strFileName <> "" Then
        
        ' Get CreationTime, LastWriteTime and LastAccessTime
        filedata = Findfile(strFileName)
        Call FileTimeToSystemTime(filedata.ftCreationTime, ftime)
        mvarCreationTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
        Call FileTimeToSystemTime(filedata.ftLastWriteTime, ftime)  ' Determine Last Modified date and time
        mvarLastWriteTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
        Call FileTimeToSystemTime(filedata.ftLastAccessTime, ftime) ' Determine Last accessed date (note no time is recorded)
        mvarLastAccessTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear
        
        
        filedata = Findfile(strFileName)
        Call FileTimeToSystemTime(filedata.ftCreationTime, ftime)
        mvarCreationTime = ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
        
        ' Get file's attributes
        '/
        
        ' Get size of the file
        If filedata.nFileSizeHigh = 0 Then
            mvarByte = Format$(filedata.nFileSizeLow, "###,###,###")
        Else
            mvarByte = Format$(filedata.nFileSizeHigh, "###,###,###")
        End If
        
        ' Get CompanyName, FileDescription, FileVersion, InternalName
        ' LegalCopyright, OriginalFilename, ProductName, ProductVersion
        '/
        
        ' Get file's type and the index of icon of the file
        lSizeSHFI = Len(SHFI)
        lFlags = SHGFI_SYSICONINDEX Or SHGFI_TYPENAME
        If bLargeIcon Then
            lFlags = lFlags Or SHGFI_LARGEICON
        Else
            lFlags = lFlags Or SHGFI_SMALLICON
        End If
        SHGetFileInfo strFileName, 0&, SHFI, lSizeSHFI, lFlags
        mvarFileType = Left$(SHFI.szTypeName, InStr(1, SHFI.szTypeName, vbNullChar) - 1)
        mvarIconIndex = SHFI.iIcon
        
        FindFileInfoView = 1
    Else
        FindFileInfoView = 0
    End If
    
GetFileVersionData_Error:
    FindFileInfoView = 0
    Exit Function
    
invalid_file_info_error:
    FindFileInfoView = 1
    Exit Function
End Function

Private Function Findfile(xstrfilename) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    plngFirstFileHwnd = FindFirstFile(xstrfilename, Win32Data)
    ' Get information of file using API call
    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error"   ' If file was not found, return error as name
    Else
        Findfile = Win32Data    ' Else return results
    End If
    plngRtn = FindClose(plngFirstFileHwnd) ' It is important that you close the handle
                'for FindFirstFile
End Function


Public Function FindFileSize(FileName As String) As Long
        On Error Resume Next
        Dim filedata As WIN32_FIND_DATA

        filedata = Findfile(FileName)
        
        If filedata.nFileSizeHigh = 0 Then
            FindFileSize = Format$(filedata.nFileSizeLow, "###,###,###")
        Else
            FindFileSize = Format$(filedata.nFileSizeHigh, "###,###,###")
        End If
End Function

