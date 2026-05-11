VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Special_Folder 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxFolders 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "Special_Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type CdislInfoType
    Value As Long
    Name As String
    Description As String
    Path As String
End Type
Private m_CdislInfo() As CdislInfoType
Private m_NumCdislInfo As Integer

Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Const MAX_PATH = 260

Private Const CSIDL_ADMINTOOLS = &H30
Private Const CSIDL_ALTSTARTUP = &H1D
Private Const CSIDL_APPDATA = &H1A
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_COMMON_ADMINTOOLS = &H2F
Private Const CSIDL_COMMON_ALTSTARTUP = &H1E
Private Const CSIDL_COMMON_APPDATA = &H23
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Private Const CSIDL_COMMON_DOCUMENTS = &H2E
Private Const CSIDL_COMMON_FAVORITES = &H1F
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_TEMPLATES = &H2D
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_COOKIES = &H21
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_HISTORY = &H22
Private Const CSIDL_INTERNET = &H1
Private Const CSIDL_INTERNET_CACHE = &H20
Private Const CSIDL_LOCAL_APPDATA = &H1C
Private Const CSIDL_MYMUSIC = &HD
Private Const CSIDL_MYPICTURES = &H27
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PRINTHOOD = &H1B
Private Const CSIDL_PROFILE = &H28
Private Const CSIDL_PROGRAM_FILES = &H26
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_SYSTEM = &H25
Private Const CSIDL_TEMPLATES = &H15
Private Const CSIDL_WINDOWS = &H24
' Make the FlexGrid's columns big enough to hold all values.
Private Sub SizeColumns(ByVal flx As MSFlexGrid)
Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer

    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
        max_wid = 0
        For r = 0 To max_row
            wid = TextWidth(flx.TextMatrix(r, c))
            If max_wid < wid Then max_wid = wid
        Next r
        flx.ColWidth(c) = max_wid + 240
    Next c
End Sub

' Load information about the values.
Private Sub LoadData()
    SaveCdislInfo CSIDL_ADMINTOOLS, "CSIDL_ADMINTOOLS", "Stores administrative tools for an individual user"
    SaveCdislInfo CSIDL_ALTSTARTUP, "CSIDL_ALTSTARTUP", "The user's nonlocalized Startup program group"
    SaveCdislInfo CSIDL_APPDATA, "CSIDL_APPDATA", "Common repository for application-specific data"
    SaveCdislInfo CSIDL_BITBUCKET, "CSIDL_BITBUCKET", "Recycle Bin"
    SaveCdislInfo CSIDL_COMMON_ADMINTOOLS, "CSIDL_COMMON_ADMINTOOLS", "Administrative tools for all users"
    SaveCdislInfo CSIDL_COMMON_ALTSTARTUP, "CSIDL_COMMON_ALTSTARTUP", "Nonlocalized Startup program group for all users (NT)"
    SaveCdislInfo CSIDL_COMMON_APPDATA, "CSIDL_COMMON_APPDATA", "Application data for all users"
    SaveCdislInfo CSIDL_COMMON_DESKTOPDIRECTORY, "CSIDL_COMMON_DESKTOPDIRECTORY", "Folder containing files and folders on the desktop (NT)"
    SaveCdislInfo CSIDL_COMMON_DOCUMENTS, "CSIDL_COMMON_DOCUMENTS", "Documents common to all users (95, 98, NT)"
    SaveCdislInfo CSIDL_COMMON_FAVORITES, "CSIDL_COMMON_FAVORITES", "Common repository for all users' favorite items (NT)"
    SaveCdislInfo CSIDL_COMMON_PROGRAMS, "CSIDL_COMMON_PROGRAMS", "Program groups that appear on the Start menu for all users (NT)"
    SaveCdislInfo CSIDL_COMMON_STARTMENU, "CSIDL_COMMON_STARTMENU", "Programs and folders that appear on the Start menu for all users (NT)"
    SaveCdislInfo CSIDL_COMMON_STARTUP, "CSIDL_COMMON_STARTUP", "Programs that appear in the Startup folder for all users (NT)"
    SaveCdislInfo CSIDL_COMMON_TEMPLATES, "CSIDL_COMMON_TEMPLATES", "Templates available for all users (NT)"
    SaveCdislInfo CSIDL_CONTROLS, "CSIDL_CONTROLS", "Icons for the Control Panel applets"
    SaveCdislInfo CSIDL_COOKIES, "CSIDL_COOKIES", "Repository for Internet cookies"
    SaveCdislInfo CSIDL_DESKTOP, "CSIDL_DESKTOP", "Desktop root namespace"
    SaveCdislInfo CSIDL_DESKTOPDIRECTORY, "CSIDL_DESKTOPDIRECTORY", "Directory that holds physical desktop items"
    SaveCdislInfo CSIDL_DRIVES, "CSIDL_DRIVES", "My Computer"
    SaveCdislInfo CSIDL_FAVORITES, "CSIDL_FAVORITES", "Repository for the user's favorite items"
    SaveCdislInfo CSIDL_FONTS, "CSIDL_FONTS", "Fonts"
    SaveCdislInfo CSIDL_HISTORY, "CSIDL_HISTORY", "Internet history items"
    SaveCdislInfo CSIDL_INTERNET, "CSIDL_INTERNET", "Virtual folder representing the Internet"
    SaveCdislInfo CSIDL_INTERNET_CACHE, "CSIDL_INTERNET_CACHE", "Temporary Internet files"
    SaveCdislInfo CSIDL_LOCAL_APPDATA, "CSIDL_LOCAL_APPDATA", "Data repository for local (non-roaming) applications"
    SaveCdislInfo CSIDL_MYMUSIC, "CSIDL_MYMUSIC", "Music files"
    SaveCdislInfo CSIDL_MYPICTURES, "CSIDL_MYPICTURES", "Picture files"
    SaveCdislInfo CSIDL_NETHOOD, "CSIDL_NETHOOD", "My Network Places"
    SaveCdislInfo CSIDL_NETWORK, "CSIDL_NETWORK", "Network Neighborhood"
    SaveCdislInfo CSIDL_PERSONAL, "CSIDL_PERSONAL", "Common repository for documents"
    SaveCdislInfo CSIDL_PRINTERS, "CSIDL_PRINTERS", "Installed printers"
    SaveCdislInfo CSIDL_PRINTHOOD, "CSIDL_PRINTHOOD", "Printers virtual folder"
    SaveCdislInfo CSIDL_PROFILE, "CSIDL_PROFILE", "The user's profile folder"
    SaveCdislInfo CSIDL_PROGRAM_FILES, "CSIDL_PROGRAM_FILES", "Program Files"
    SaveCdislInfo CSIDL_PROGRAM_FILES_COMMON, "CSIDL_PROGRAM_FILES_COMMON", "Components shared across applications (NT, 2000)"
    SaveCdislInfo CSIDL_PROGRAMS, "CSIDL_PROGRAMS", "The user's program groups"
    SaveCdislInfo CSIDL_RECENT, "CSIDL_RECENT", "Recently used documents"
    SaveCdislInfo CSIDL_SENDTO, "CSIDL_SENDTO", "SendTo menu items"
    SaveCdislInfo CSIDL_STARTMENU, "CSIDL_STARTMENU", "Start menu items"
    SaveCdislInfo CSIDL_STARTUP, "CSIDL_STARTUP", "The user's Startup group"
    SaveCdislInfo CSIDL_SYSTEM, "CSIDL_SYSTEM", "System folder"
    SaveCdislInfo CSIDL_TEMPLATES, "CSIDL_TEMPLATES", "Common repository for document templates"
    SaveCdislInfo CSIDL_WINDOWS, "CSIDL_WINDOWS", "Windows directory (SYSROOT)"
End Sub
' Save information about a CDISL value.
Private Sub SaveCdislInfo(ByVal new_value As Long, ByVal new_name As String, ByVal new_description As String)
    m_NumCdislInfo = m_NumCdislInfo + 1
    ReDim Preserve m_CdislInfo(1 To m_NumCdislInfo)
    With m_CdislInfo(m_NumCdislInfo)
        .Value = new_value
        .Name = new_name
        .Description = new_description
    End With
End Sub

' Get a special folder's path.
Public Function GetSpecialFolderPath(ByVal folder_number As Long) As String
Dim Path As String

    Path = Space$(MAX_PATH)
    If SHGetSpecialFolderPath(hwnd, Path, _
        folder_number, False) _
    Then
        GetSpecialFolderPath = Left$(Path, InStr(Path, Chr$(0)))
        GetSpecialFolderPath = Mid$(GetSpecialFolderPath, 1, Len(GetSpecialFolderPath) - 1)
    Else
        GetSpecialFolderPath = "???"
    End If
End Function


Private Sub Form_Load()
Dim i As Integer

    ' Load the data.
    LoadData

    ' Fill the grid.
    flxFolders.Rows = m_NumCdislInfo + 1
    flxFolders.Cols = 4
    flxFolders.FixedCols = 0

    flxFolders.TextMatrix(0, 0) = "Folder"
    flxFolders.TextMatrix(0, 1) = "Value"
    flxFolders.TextMatrix(0, 2) = "Path"
    flxFolders.TextMatrix(0, 3) = "Description"

    ' Display the list of values.
    For i = 1 To m_NumCdislInfo
        With m_CdislInfo(i)
            flxFolders.TextMatrix(i, 0) = .Name
            flxFolders.TextMatrix(i, 1) = "&H" & Hex$(.Value)
            flxFolders.TextMatrix(i, 2) = _
                GetSpecialFolderPath(.Value)
            flxFolders.TextMatrix(i, 3) = .Description
        End With
    Next i

    ' Resize the FlexGrid columns.
    SizeColumns flxFolders
End Sub
Private Sub Form_Resize()
    flxFolders.Move 0, 0, ScaleWidth, ScaleHeight

    ' Resize the FlexGrid columns.
    SizeColumns flxFolders
End Sub


