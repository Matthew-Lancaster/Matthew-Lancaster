VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5352
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   14208
   Icon            =   "SYNCRONIZER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5352
   ScaleWidth      =   14208
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   4140
      Left            =   168
      TabIndex        =   2
      Top             =   948
      Width           =   13128
      _ExtentX        =   23156
      _ExtentY        =   7303
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.FileListBox File1 
      Height          =   456
      Left            =   11316
      TabIndex        =   1
      Top             =   132
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Timer Timer_INIT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12552
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   444
      Left            =   120
      TabIndex        =   0
      Top             =   132
      Width           =   10788
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB_FOLDER"
   End
   Begin VB.Menu MNU_PROJECT_FOLDER 
      Caption         =   "PROJECT_FOLDER"
   End
   Begin VB.Menu MNU_DONE 
      Caption         =   "DONE..."
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FLAG_QUIETLY_COMMAND_STRING

Dim R_R
Dim T_1
Dim T_2

Dim objStartFolder_2
Dim FSO

Dim RecursiveSearch_Most_Recent_VAR

Public EXIT_TRUE

Dim STORE_DATE As Date
Dim SEARCH_VAR As String
Private wmi As Object, Locator As Object
Dim SampleRate As Long

Private PrevCpuTime As Single

Private Const MAX_PATH As Long = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelpSnapshot Lib "Kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
    Private Const TH32CS_SNAPPROCESS As Long = 2&
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "Kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "Kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20
     
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

'----------------------------------------------------------------------------------
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type

Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'----------------------------------------------------------------------------------


Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long

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

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100


Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function


Private Sub Form_Load()
    
    If GetComputerName = "1-ASUS-X5DIJ" Then End
    If GetComputerName = "2-ASUS-EEE" Then End
    If GetComputerName = "3-LINDA-PC" Then End
    If GetComputerName = "5-ASUS-P2520LA" Then End
    
    If IsIDE = True Then
        If Mid(App.Path, 1, 2) = "\\" Then
            Me.WindowState = vbNormal
            'Me.Show
            DoEvents
            MsgBox "Don't Run on a Network Path as Less Access Right Given in ISIDE=True" + vbCrLf + vbCrLf + App.Path + "\" + vbCrLf + App.EXEName + ".VBP __ Project", vbMsgBoxSetForeground
            End
        End If
    End If
    
    If Command$ = "QUIETLY" Then
        FLAG_QUIETLY_COMMAND_STRING = True
    End If
    
    If App.PrevInstance = True Then End
    Call Project_Check_Date.VB_PROJECT_CHECKDATE("FORM LOAD")
    
    If Dir(App.Path + "\PROJECT SET", vbDirectory) = "" Then
        CreateFolderTree (App.Path + "\PROJECT SET")
    End If
    If Dir(App.Path + "\PROJECT ARCHIVER", vbDirectory) = "" Then
        CreateFolderTree (App.Path + "\PROJECT ARCHIVER FOLDER")
    End If
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Label1.Top = 0
    Label1.Height = 800
    
    ListView1.Top = Label1.Top + Label1.Height
    'ListView_INDEX.Top = Label1.Top + Label1.Height
    'ListView_INDEX.Left = 0
    

'    List1.Top = Label1.Top + Label1.Height
'    List2.Top = Label1.Top + Label1.Height
'    List4.Top = Label1.Top + Label1.Height
'    List2.Left = 0
'    List2.Width = 750
'    List4.Left = List2.Width
'    List4.Width = 2700
'    List1.Left = List4.Width + List4.Left
'    List1.Height = 6000
'    List2.Height = List1.Height
'    List4.Height = List1.Height
'    List1.Width = 7800
    ListView1.Width = 14000
    ListView1.Height = 5900
    'ListView_INDEX.Width = 500
    'ListView_INDEX.Height = ListView1.Height
    ListView1.Left = 0 'ListView_INDEX.Width
    Me.Width = ListView1.Left + ListView1.Width + 250
    Me.Height = ListView1.Top + ListView1.Height + Menu_Height * Screen.TwipsPerPixelY + 580
    Label1.Width = ListView1.Left + ListView1.Width
    
    ListView1.Font.Name = "ARIAL"
    ListView1.Font.Size = 11
    Label1.Font.Name = "ARIAL"
    Label1.Font.Size = 11
    
    With ListView1
        .ColumnHeaders.Add , "#", "#", 300, lvwColumnLeft
        .ColumnHeaders.Add , ">", ">", 400, lvwColumnLeft
        .ColumnHeaders.Add , "DATE", "DATE", 2400, lvwColumnLeft
        .ColumnHeaders.Add , "NAME", "NAME", 3400, lvwColumnLeft
        .ColumnHeaders.Add , "PATH", "PATH", 5000, lvwColumnLeft
        .View = lvwReport
    End With
    
    SetFullRowSelection ListView1.hWnd, True
    Me.Caption = "SYNCRONIZER"
    Timer_INIT.Enabled = True
    
  '  If Me.WindowState = vbMinimized Then End


'    FLAG_QUIETLY_COMMAND_STRING = True
    
    If FLAG_QUIETLY_COMMAND_STRING = True Then Me.Hide

    
End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.EXIT_TRUE = True



'If Me.WindowState = vbMinimized Then End



On Error Resume Next
For Each Form In Forms
Err.Clear
    If Form.EXIT_TRUE = True Then
        'Form.Name
        If Err.Number = 0 Then
            Me.EXIT_TRUE = True
        End If
    End If
Next Form
On Error GoTo 0

If IsIDE = False Then
    If Me.WindowState <> vbMinimized And Me.EXIT_TRUE = False Then
        Me.WindowState = vbMinimized
        Cancel = True
        Exit Sub
    End If
End If

For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

End Sub

Private Sub Timer_INIT_Timer()
    Timer_INIT.Enabled = False
    
    T_1 = "PROJECT ARCHIVER FOLDER"
    T_2 = "PROJECT ARCHIVER PATH_FILE"
    File1.Path = App.Path + "\PROJECT SET\"
    File1.Pattern = "*.*"
    For R_R = 0 To File1.ListCount - 1
    
        ListView1.ListItems.Clear
    
        Call SYNC("D:\VB6\VB-NT\00_Best_VB_01", File1.List(R_R))
        If EXIT_TRUE = True Then Unload Me: Exit Sub
        
        Call SYNCRO_ARCHIVER
        If EXIT_TRUE = True Then Unload Me: Exit Sub
        
        Call SYNCRO_COPY
        If EXIT_TRUE = True Then Unload Me: Exit Sub
        
        Load Project_Check_Date
        If EXIT_TRUE = True Then Unload Me: Exit Sub

    Next

Label1.Caption = Label1.Caption + " __Done"
Label1.BackColor = RGB(127, 255, 127)
MNU_DONE.Visible = True


If FLAG_QUIETLY_COMMAND_STRING Then End

End Sub


Private Sub SYNC(objStartFolder, SEARCH_VAR As String)
    
    'Beep
    'Call MNU_LAST_ART_PIC_IVIEW_Click
    'Me.WindowState = vbMinimized
    
    Dim FS, F, F1, fc, S
    Dim objFiles As Files, objFile As File
    Dim objFolders As Folders, objFolder As Folder
    Dim fld As Folder
    
    STORE_DATE = 0
    objStartFolder_2 = objStartFolder
    Set fld = FSO.GetFolder(objStartFolder)
    Call RecursiveSearch_Most_Recent(fld, SEARCH_VAR)
    LAST_CLIPBOARD_FILE = RecursiveSearch_Most_Recent_VAR
    If EXIT_TRUE = True Then Unload Me

End Sub

Sub SYNCRO_ARCHIVER()

For R_L = 1 To ListView1.ListItems.Count
    SET_GO = True
    If InStr(ListView1.ListItems.Item(R_L).SubItems(4), "\PROJECT SET") > 0 Then SET_GO = False
    If InStr(ListView1.ListItems.Item(R_L).SubItems(4), "\" + T_1) > 0 Then SET_GO = False
    If SET_GO = True Then
        PATH_FILE_NAME_SET = ""
        PATH_FILE_NAME_SET = PATH_FILE_NAME_SET + Trim(ListView1.ListItems.Item(R_L).SubItems(4)) + "\"
        PATH_FILE_NAME_SET = PATH_FILE_NAME_SET + Trim(ListView1.ListItems.Item(R_L).SubItems(3))
    
        'PATH_FILE_NAME1 = RecursiveSearch_Most_Recent_VAR
        PATH_FILE_NAME1 = PATH_FILE_NAME_SET
        PATH_FILE_NAME2 = App.Path + "\" + T_1 + "\" + Mid(PATH_FILE_NAME1, InStrRev(PATH_FILE_NAME1, "\") + 1)
        PATH_FILE_NAME4_ST = App.Path + "\" + T_2 + "\" + Mid(PATH_FILE_NAME1, InStrRev(PATH_FILE_NAME1, "\") + 1)
        PATH_FILE_NAME_FILE = Mid(PATH_FILE_NAME1, InStrRev(PATH_FILE_NAME1, "\") + 1)
        'PATH_FILE_NAME4 =
        'PATH_FILE_NAME_FILE_ST = Mid(PATH_FILE_NAME_FILE, 1, InStrRev(PATH_FILE_NAME_FILE, ".") - 1)
        PATH_FILE_NAME_FILE_ST = PATH_FILE_NAME_FILE
        PATH_FILE_NAME_EXT = Mid(PATH_FILE_NAME2, InStrRev(PATH_FILE_NAME2, ".") + 1)
        PATH_FILE_NAME_STRIP_EXT = Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, ".") - 1)
        PATH_FILE_NAME_STRIP_EXT4 = Mid(PATH_FILE_NAME4_ST, InStrRev(PATH_FILE_NAME4_ST, "\") + 1)
        PATH_FILE_NAME2_APPEND = App.Path + "\" + T_1 + "\" + Mid(PATH_FILE_NAME1, Len(objStartFolder_2) + 2)
        PATH_FILE_NAME2_APPEND_ST = Mid(PATH_FILE_NAME2_APPEND, 1, InStrRev(PATH_FILE_NAME2_APPEND, "\") - 1)
        
        Set FN = FSO.GetFile(PATH_FILE_NAME1)
        PATH_FILE_NAME2_APPEND = PATH_FILE_NAME2_APPEND_ST + "\" + PATH_FILE_NAME_FILE_ST + "__" + Format(FN.DateLastModified, "YYYY-MM-DD -- HH-MM-SS") + "." + PATH_FILE_NAME_EXT
        PATH_FILE_NAME2 = PATH_FILE_NAME_STRIP_EXT + "__" + Format(FN.DateLastModified, "YYYY-MM-DD -- HH-MM-SS") + "." + PATH_FILE_NAME_EXT
        PATH_FILE_NAME4 = PATH_FILE_NAME4_ST + "\" + PATH_FILE_NAME_STRIP_EXT4 + "__" + Format(FN.DateLastModified, "YYYY-MM-DD -- HH-MM-SS") + "." + PATH_FILE_NAME_EXT
        
        If FN.Size > 0 Then
            '__01 OF 02 __ DIR STRUCTURE
            
            If Dir(PATH_FILE_NAME2_APPEND_ST, vbDirectory) = "" Then
                CreateFolderTree (PATH_FILE_NAME2_APPEND_ST)
            End If
            If Dir(PATH_FILE_NAME2_APPEND) = "" Then
                FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2_APPEND
            End If
            
            
            '__02 OF 02 __ COLUMN
            If Dir(PATH_FILE_NAME4_ST, vbDirectory) = "" Then
                CreateFolderTree (PATH_FILE_NAME4_ST)
            End If
            If Dir(PATH_FILE_NAME4) = "" Then
                FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME4
            End If
        End If
    End If
Next
End Sub

Sub SYNCRO_COPY()

For R_L = 1 To ListView1.ListItems.Count
    SET_GO = True
    If InStr(ListView1.ListItems.Item(R_L).SubItems(4), "\PROJECT SET") > 0 Then SET_GO = False
    If InStr(ListView1.ListItems.Item(R_L).SubItems(4), "\" + T_1) > 0 Then SET_GO = False
    If SET_GO = True Then
        PATH_FILE_NAME_SET = ""
        PATH_FILE_NAME_SET = PATH_FILE_NAME_SET + Trim(ListView1.ListItems.Item(R_L).SubItems(4)) + "\"
        PATH_FILE_NAME_SET = PATH_FILE_NAME_SET + Trim(ListView1.ListItems.Item(R_L).SubItems(3))
    
        PATH_FILE_NAME1 = RecursiveSearch_Most_Recent_VAR
        PATH_FILE_NAME2 = PATH_FILE_NAME_SET
        Set FN_1 = FSO.GetFile(PATH_FILE_NAME1)
        Set FN_2 = FSO.GetFile(PATH_FILE_NAME2)
        If FN_1.Size > 0 And FN_1.DateLastModified >= FN_2.DateLastModified Then
            If PATH_FILE_NAME1 <> PATH_FILE_NAME2 Then
                FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
            End If
        End If
    End If
Next

PATH_FILE_NAME1 = PATH_FILE_NAME1
PATH_FILE_NAME_FILE = Mid(PATH_FILE_NAME1, InStrRev(PATH_FILE_NAME1, "\") + 1)
PATH_FILE_NAME2 = App.Path + "\PROJECT SET\" + PATH_FILE_NAME_FILE

'-----------------------------------------------------------------------------
'CAN'T DO ITSELF
'-----------------------------------------------------------------------------
'"D:\VB6\VB-NT\00_Best_VB_01\10 SYNCRONIZE\PROJECT SET\Form_VB_Check_Date.frm"
'-----------------------------------------------------------------------------
'If IsIDE = True Then
'IF
If PATH_FILE_NAME1 <> "" Then
' If IsIDE = True And PATH_FILE_NAME1 = "" Then MsgBox "CHECK IF RUN HERE A WHY IF LINE WANTER": Stop
' MEAN TO CHECK HERE FOR DEBUG IF LINE REQUIRE AS SOMETIME ENTER WITH PATH EMPTY
FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
End If

'End If
End Sub

Sub RecursiveSearch_Most_Recent(fld As Folder, search As String)
' -------------------------------------
' Compile Error:
' User-defined type not defined
' -------------------------------------
' DO HERE
' PROJECT -- REFERENCE
' FOLDER -- WINDOWS SCRIPT OBJECT MODEL
' -------------------------------------

Dim fold As Folder
Dim fil As File

    For Each fold In fld.SubFolders
        RecursiveSearch_Most_Recent fold, search
        If EXIT_TRUE = True Then Exit For
        COUNTING_UP = COUNTING_UP + 1
        If COUNTING_UP Mod 10 = 0 Then DoEvents
    Next fold
    
    For Each fil In fld.Files
        COUNTING_UP = COUNTING_UP + 1
        If COUNTING_UP Mod 10 = 0 Then DoEvents
        If EXIT_TRUE = True Then Exit For
        If fil.Name = search Then
            If fil.DateLastModified > STORE_DATE Then
                STORE_DATE = fil.DateLastModified
                RecursiveSearch_Most_Recent_VAR = fil.Path
                Label1.Caption = RecursiveSearch_Most_Recent_VAR + vbCrLf + Format(fil.DateLastModified, "YYYY-MM-DD -- HH-MM-SS") + vbCrLf + "PROJECT SET #" + Str(R_R + 1) + " OF #" + Str(File1.ListCount)
                HITT = "#"
            End If
            'LockWindowUpdate ListView1.hWnd
            LISTVIEW_ITEM_1 = Trim(Str(ListView1.ListItems.Count + 1)) + "  "
            LISTVIEW_ITEM_2 = HITT + "   "
            LISTVIEW_ITEM_3 = Format(fil.DateLastModified, "YYYY-MM-DD  HH-MM-SS") + "  "
            LISTVIEW_ITEM_4 = fil.Name + "  "
            LISTVIEW_ITEM_5 = fld.Path + "  "
            SET_GO = True
            If InStr(fld.Path, "\PROJECT SET") > 0 Then SET_GO = False
            If InStr(fld.Path, "\PROJECT ARCHIVER") > 0 Then
                SET_GO = False
            End If
            If SET_GO = True Then
                With ListView1
                    Set LV2 = .ListItems.Add(, , LISTVIEW_ITEM_1)
                    LV2.SubItems(1) = LISTVIEW_ITEM_2
                    LV2.SubItems(2) = LISTVIEW_ITEM_3
                    LV2.SubItems(3) = LISTVIEW_ITEM_4
                    LV2.SubItems(4) = LISTVIEW_ITEM_5
                End With
                
                If HITT = "#" Then
                    For R_L = ListView1.ListItems.Count - 1 To 1 Step -1
                      ListView1.ListItems.Item(R_L).SubItems(1) = "" + "  "
                    Next
                End If
                
                DoEvents
                'LockWindowUpdate 0
                WIDTH_ADJUST = Len(LISTVIEW_ITEM_1) + Len(LISTVIEW_ITEM_2) + Len(LISTVIEW_ITEM_3) + Len(LISTVIEW_ITEM_4)
                If OLD_WIDTH_ADJUST <> WIDTH_ADJUST Then
                    For R_L = 1 To ListView1.ColumnHeaders.Count
                        Call LV_AutoSizeColumn(ListView1, ListView1.ColumnHeaders.Item(R_L))
                    Next
                End If
                OLD_WIDTH_ADJUST = WIDTH_ADJUST
            End If
        End If
    Next fil
'REFENCER
'----
'vb6 fso get sub folder and files - Google Search
'https://www.google.co.uk/search?q=vb6+fso+get+sub+folder+and+files&oq=vb6+fso+get+sub+folder+and+files&aqs=chrome..69i57.11696j0j7&sourceid=chrome&ie=UTF-8
'--------
'Loop through folders/subfolders-VBForums
'http://www.vbforums.com/showthread.php?613400-Loop-through-folders-subfolders
'----
End Sub

Private Sub MNU_PROJECT_FOLDER_Click()

XY = Dir(App.Path + "\PROJECT SET\*.*")

Shell "EXPLORER /SELECT, " + App.Path + "\PROJECT SET\" + XY, vbMaximizedFocus

End Sub

Private Sub MNU_VB_FOLDER_Click()
Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus
End Sub

Private Sub MNU_VB_ME_Click()

VBPATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VBPATH) = "" Then
    VBPATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
End If

Shell VBPATH + " """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End
End Sub

Private Sub MNU_EXIT_Click()
Me.WindowState = vbMinimized

EXIT_TRUE = True
Unload Me
'End
End Sub

Private Sub SetFullRowSelection(ByVal hWndListView, ByVal bFullRow As Boolean)
   SendMessage hWndListView, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal CLng(bFullRow)
End Sub

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


Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)

        End With
       
    End If
   
End Function



Public Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If Not FolderExists(Left$(sPath, nPos - 1)) Then
            MkDir Left$(sPath, nPos - 1)
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Not FolderExists(sPath) Then MkDir sPath
    
    CreateFolderTree = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function


Public Function FolderExists(sFolder As String) As Boolean
    '##############################################################################################
    'Returns True if the specified folder exists
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function



Public Function GetOperatingSystem()
    Dim localHost       As String
    Dim objWMIService   As Variant
    Dim colOperatingSystems As Variant
    Dim objOperatingSystem As Variant
    On Error GoTo Error_Handler
    localHost = "." 'Technically could be run against remote computers, if allowed
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & localHost & "\root\cimv2")
    Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each objOperatingSystem In colOperatingSystems
        GetOperatingSystem = objOperatingSystem.Caption & " " & objOperatingSystem.VERSION
        GetOperatingSystem = Format(Val(Mid(objOperatingSystem.VERSION, 1, InStr(objOperatingSystem.VERSION, "."))), "00")
        Exit Function
    Next
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: getOperatingSystem" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
'  IsIDE = False
'  Exit Function
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************




