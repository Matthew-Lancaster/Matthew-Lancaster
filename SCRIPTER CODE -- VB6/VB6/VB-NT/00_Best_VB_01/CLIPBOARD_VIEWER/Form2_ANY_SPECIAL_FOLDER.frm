VERSION 5.00
Begin VB.Form FRM_SPECIAL_FOLDER 
   Caption         =   "Jump Any Special Folder"
   ClientHeight    =   5640
   ClientLeft      =   228
   ClientTop       =   576
   ClientWidth     =   10128
   LinkTopic       =   "Form2"
   ScaleHeight     =   5640
   ScaleWidth      =   10128
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4188
      Top             =   2880
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4080
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "Press Jump Any Special Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "FRM_SPECIAL_FOLDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SP_FOLDER_COUNT

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


Private Sub Form_Load()
'Load Form2_ANY_SPECIAL_FOLDER


Label1.Left = 0
Label1.Top = 0
List1.Left = 0
List1.Top = Label1.Top + Label1.Height

Call SETUP_DIMENSIONS_NORM

'Me.Width = Screen.Width - 1100

Label1.Width = Me.Width '- 400
List1.Width = Me.Width - 200

'List1.Height = (Screen.Height / Screen.TwipsPerPixelY) - 100
'List1.Height = (Screen.Height) - 3000
'Me.Height = (List1.Height + List1.Top)

'Me.Top = Screen.Height / 2 - Me.Height / 2
'Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Refresh
DoEvents
List1.Height = Me.Height - 900
List1.Refresh
DoEvents
'Me.Height = (List1.Height + List1.Top) + 200

GetSpecialfolder_Script

Me.Show

End Sub

Function GetSpecialfolder_Script()
    
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Dim SPECIAL_FOLDER_PATH As String
    SP_FOLDER_COUNT = 0
    
    'DEBUG
    For R2 = 0 To 255
        r = SHGetSpecialFolderLocation(100, R2, IDL)
        If r = NOERROR Then
            Path$ = Space$(512)
            r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
            SPECIAL_FOLDER_PATH = Left$(Path, InStr(Path, Chr$(0)) - 1)
            If SPECIAL_FOLDER_PATH <> "" Then
                'Debug.Print Str(R2) + " -- " + SPECIAL_FOLDER_PATH
                HEX_NUM = Hex(R2)
                If Len(HEX_NUM) < 2 Then HEX_NUM = "0" + HEX_NUM
                HEX_NUM = "&H_" + HEX_NUM
                List1.AddItem Format(R2, "000") + " -- " + HEX_NUM + " -- " + SPECIAL_FOLDER_PATH
                SP_FOLDER_COUNT = SP_FOLDER_COUNT + 1
                
                If InStr(SPECIAL_FOLDER_PATH, "\Startup") > 0 Then
                    MODIFY_VAR = Replace(SPECIAL_FOLDER_PATH, "\Startup", "\Startup-01-Delayed")
                    If FS.FolderExists(MODIFY_VAR) = True Then
                        TOTAL_FOUND = TOTAL_FOUND + 1
                        List1.AddItem Format(R2, "000") + " -- " + HEX_NUM + " -- ----" + MODIFY_VAR
                    End If
                End If
                
                If InStr(DUPE_STRING, SPECIAL_FOLDER_PATH + " - ") > 0 Then
                    DUPE_COUNT = DUPE_COUNT + 1
                End If
                DUPE_STRING = DUPE_STRING + SPECIAL_FOLDER_PATH + " - "
            End If
        End If
    Next
    Label1.Caption = "Press -- Jump Any -- Special Folder *" + Str(SP_FOLDER_COUNT) + " -- Dupe Count *" + Str(DUPE_COUNT) + " -- Custom Additonal ---- *" + Str(TOTAL_FOUND)
    DUPE_STRING = ""
End Function

Sub SETUP_DIMENSIONS_NORM()

    Dim RECTT1 As RECT
    Dim RECTT2 As RECT
    Dim RECTT4 As RECT
    
    On Error Resume Next
    
   ' If Me.WindowState <> vbMinimized Then
        
        'DONT_RESIZE_WHILE_SETUP = True
        Me.WindowState = vbNormal
                
        Test1 = FindWindow("Shell_TrayWnd", vbNullString)
        T1 = GetWindowRect(Test1, RECTT1) ' BOTTOM BAR
        Test2 = FindWindow("MOM Class", vbNullString)
        T1 = GetWindowRect(Test2, RECTT2) ' TOP BAR
        'IRON BAR - DUMB BELLS
        TEST4 = FindWindowPart_BASEBAR("BaseBar")
        T1 = GetWindowRect(TEST4, RECTT4) ' LEFT BAR
        A = RECTT4.Top
        y1 = (RECTT1.Bottom - RECTT1.Top) * Screen.TwipsPerPixelY
        
        If RECTT2.Bottom - RECTT2.Top < 33 Then
        
            Y2 = (RECTT2.Bottom - RECTT2.Top) * Screen.TwipsPerPixelY
        
        Else
        
            Y2 = 0
        
        End If
        
        If TEST4 > 0 Then
            If RECTT4.Right = RECTT4.Left Then
                X4 = (RECTT4.Right) * Screen.TwipsPerPixelX
            Else
                X4 = (RECTT4.Right - RECTT4.Left) * Screen.TwipsPerPixelX
            End If
        End If
        
        SCREEN_WIDTH_SPACE = Screen.Width - X4
        SCREEN_HEIGHT_SPACE = Screen.Height - y1 - Y2
        
        Me.Width = SCREEN_WIDTH_SPACE - 1200
        'Me.Height = SCREEN_HEIGHT_SPACE - 800
        Me.Height = SCREEN_HEIGHT_SPACE - 1500
        'THIS THE FORM HEIGHT
        'THE BOX INSIDE IS ADJUST AFTER
        
        
        
        
        'Form1.Left = (Screen.Width - Me.Width) / 2
        Me.Left = X4 + (SCREEN_WIDTH_SPACE - Me.Width) / 2
        Me.Top = Y2 + (SCREEN_HEIGHT_SPACE - Me.Height) / 2

End Sub

Private Sub List1_DblClick()
    
    PATH_NAME = List1.List(List1.ListIndex)
    PATH_NAME = Mid(PATH_NAME, InStr(PATH_NAME, ":\") - 1)
    
    Shell "Explorer.exe /e," + PATH_NAME, vbNormalFocus
    Beep
    
'    If InStr(Label1.Caption, " -- FOLDER LOADING") = 0 Then
'        Label1.Caption = Label1.Caption + " -- FOLDER LOADING"
'    End If
    
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    
    'Unload Me
    Form1.WindowState = vbMinimized

End Sub
