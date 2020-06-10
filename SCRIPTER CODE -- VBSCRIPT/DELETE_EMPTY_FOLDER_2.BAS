Attribute VB_Name = "Module2"
'----
'FreeVBCode code snippet: Customized Recursive File Search
'http://www.freevbcode.com/ShowCode.asp?ID=8299
'----

'Option Explicit
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Sub FILE_FOUND(VAR_PATH, VAR_FileName)
    Debug.Print VAR_PATH & VAR_FileName
End Sub

'Public Sub FindFiles(strPath As String, strWildcardPassed As String, bSubFolders As Boolean, ByVal lvwListViewControl As ListView)
Public Sub FindFiles(strPath As String, strWildcardPassed As String, bSubFolders As Boolean)
'    Screen.MousePointer = 11
    
'    Dim litmListItem As ListItem
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, strLocalPathBiffer As String, strWildCard As String
    Dim strFileNameNullsStriped As String
    
    strLocalPathBiffer = AddBackslash(strPath)
    strWildCard = strLocalPathBiffer & strWildcardPassed
    
    '---> Searching the current directory
    '---> Find the first file, if any
    hFile = FindFirstFile(strWildCard, WFD)
    '---> If a FIRST file was found and it isn't a directory
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
'        Set litmListItem = lvwListViewControl.ListItems.Add(, , strLocalPathBiffer & StripNulls(WFD.cFileName))
        Call FILE_FOUND(strLocalPathBiffer, StripNulls(WFD.cFileName))
    End If
    '<--- If a FIRST file was found and it isn't a directory
    '<--- Find the first file, if any

    '---> Cycle finding NEXT-files, if any
    While FindNextFile(hFile, WFD)
        '---> If a NEXT file was found and it isn't a directory
        If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            'Set litmListItem = lvwListViewControl.ListItems.Add(, , strLocalPathBiffer & StripNulls(WFD.cFileName))
            Call FILE_FOUND(strLocalPathBiffer, StripNulls(WFD.cFileName))
        End If
        '<--- If a NEXT file was found and it isn't a directory
    Wend
    '<--- Cycle finding NEXT-files, if any
    '<--- Searching the current directory
    
    '---> Current directory search was finished,
    '---> keep looking for subfolders and their contents
    '<--- if subfolders check box is checked
    If bSubFolders Then
        
        '---> Find FIRST file (actually - folder)
        hFile = FindFirstFile(strLocalPathBiffer & "*.*", WFD)
        '---> If a folder was found as a FIRST file
        If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
        StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
           '---> Recursive self-invocation to search the subfolder
            strFileNameNullsStriped = StripNulls(WFD.cFileName)
            strLocalPathBiffer = strLocalPathBiffer & strFileNameNullsStriped
            FindFiles strLocalPathBiffer, strWildcardPassed, True ', lvwListViewControl
        End If '<--- If (hFile > 0) And ....
        '<--- If a folder was found as a FIRST file
        
        '---> Cycle looking for-NEXT files (actually - folders)
        While FindNextFile(hFile, WFD)
            If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
            StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
                '---> Recursive self-invocation to search the subfolder
            strFileNameNullsStriped = StripNulls(WFD.cFileName)
            strLocalPathBiffer = AddBackslash(strPath)
            strLocalPathBiffer = strLocalPathBiffer & strFileNameNullsStriped
                FindFiles strLocalPathBiffer, strWildcardPassed, True ', lvwListViewControl
            End If '<--- If ((WFD.dwFileAttributes And ...
        Wend '<--- While FindNextFile(hFile, WFD)
        '<--- Cycle looking for-NEXT files (actually - folders)
        
    End If '<--- If bSubFolders
    
    FindClose hFile
    
    Set litmListItem = Nothing
    
    Screen.MousePointer = 0
End Sub

Private Function StripNulls(strString As String) As String
    StripNulls = Left$(strString, InStr(1, strString, Chr$(0)) - 1)
End Function

Private Function AddBackslash(strString As String) As String
    If Len(strString) Then
       If Right$(strString, 1) <> "\" Then
          AddBackslash = strString & "\"
       Else
          AddBackslash = strString
       End If
    Else
          AddBackslash = "\"
    End If
End Function

