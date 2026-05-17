Attribute VB_Name = "SPECIAL_FOLDER_Mod"
Public Enum SpecialFolderConstants
    spfldDesktop = &H0
    spfldInternet = &H1
    spfldPrograms = &H2
    spfldControls = &H3
    spfldPrinters = &H4
    spfldPersonal = &H5
    spfldFavorites = &H6
    spfldStartUp = &H7
    spfldRecent = &H8
    spfldSendTo = &H9
    spfldBitBucket = &HA
    spfldStartMenu = &HB
    spfldDesktopDirectory = &H10
    spfldDrives = &H11
    spfldNetwork = &H12
    spfldNethood = &H13
    spfldFonts = &H14
    spfldTemplates = &H15
    spfldCommonStartMenu = &H16
    spfldCommonPrograms = &H17
    spfldCommonStartup = &H18
    spfldCommonDesktopDirectory = &H19
    spfldAppData = &H1A
    spfldPrintHood = &H1B
    spfldAltStartup = &H1D                          '' DBCS
    spfldCommonAltStartup = &H1E                   '' DBCS
    spfldCommonFavorites = &H1F
    spfldInternetCache = &H20
    spfldCookies = &H21
    spfldHistory = &H22
    spfldWindows = &H24
    spfldWindowSystem = &H25
    
    '&H27& My Pictures
    '40 User Profile
    '43 Common Files
    '46 All Users Templates
    '47 Administrative Tools
    '49 Network Connections
    
End Enum


Public FS
Public sSystemFolder As String
Public sTempFolder As String
Public sWindowsFolder As String


Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type



'Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
'Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
'Public Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
'Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long



Public Function GetSpecialfolder(CSIDL As Long) As String
    '##############################################################################################
    'Returns the Path to a "Special" Folder (i.e. Internet History)
    '##############################################################################################
    
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function

