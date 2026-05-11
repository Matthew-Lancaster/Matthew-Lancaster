Attribute VB_Name = "GetFolder"
Public lovely



Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
'Public Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long




Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


'txtPath.Text = GetFolder(Me.hWnd, "Scan Path:", txtPath.Text)

Public Function GetFolderFromPath(sPath As String) As String
    Dim nPos As Long
    Dim sFolder As String

    nPos = InStrRev(sPath, "\")
    
    If nPos > 0 Then
        sFolder = Left$(sPath, nPos - 1)
        If Right$(sFolder, 1) = ":" Then
            sFolder = sFolder & "\"
        End If
        GetFolderFromPath = sFolder
    End If
End Function


Public Function GetFolder(hWnd As Long, Optional sPrompt As String, Optional sStartFolder As String) As String
    '##############################################################################################
    'Displays a Folder Browser to select a Folder
    '##############################################################################################
    
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim sFolder As String
    Dim pos As Long
    
    'Fill the BROWSEINFO structure with the needed data
    With bi
        'hwnd of the window that receives messages from the call. Can be your application or the handle from GetDesktopWindow().
        .hOwner = hWnd
        
        'Pointer to the item identifier list specifying the location of the "root" folder to browse from.
        'If NULL, the desktop folder is used.
        .pidlRoot = 0&
    
        'message to be displayed in the Browse dialog
        If Len(sPrompt) = 0 Then
            .lpszTitle = "Select the folder:"
        Else
            .lpszTitle = sPrompt
        End If
    
        'the type of folder to return. - the constants perform differently for non networked pc's
        .ulFlags = BIF_RETURNONLYFSDIRS
        
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)
        .lParam = SHSimpleIDListFromPath(StrConv(sStartFolder, vbUnicode))
    End With
        
    'show the browse for folders dialog
     pidl = SHBrowseForFolder(bi)
    
    'the dialog has closed, so parse & display the user's returned folder selection contained in pidl
    sFolder = Space$(MAX_PATH)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal sFolder) Then
        pos = InStr(sFolder, Chr$(0))
        GetFolder = Left$(sFolder, pos - 1)
    End If
    
    Call CoTaskMemFree(pidl)
End Function


Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    '############################################################################
    'Purpose: Required by GetGolder() Function
    '############################################################################
 
    Select Case uMsg
    Case BFFM_INITIALIZED
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, 0&, ByVal lpData)
                    
    Case Else
    
    End Select
End Function


Private Function FARPROC(pfn As Long) As Long
    '############################################################################
    'Purpose: Required by GetGolder() Function
    
    'A dummy procedure that receives and returns
    'the value of the AddressOf operator.
    
    'This workaround is needed as you can't
    'assign AddressOf directly to a member of a
    'user-defined type, but you can assign it
    'to another long and use that instead!
    '############################################################################
 
    FARPROC = pfn
End Function

