VERSION 5.00
Begin VB.Form AWINAMP_CONTROL 
   BackColor       =   &H80000007&
   Caption         =   "WINAMP CONTROL"
   ClientHeight    =   1884
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   7416
   LinkTopic       =   "Form1"
   ScaleHeight     =   1884
   ScaleWidth      =   7416
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer WINAMP_Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1950
      Top             =   735
   End
   Begin VB.ListBox List1 
      Height          =   1008
      Left            =   2865
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   270
      Visible         =   0   'False
      Width           =   4110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   15
      Left            =   1065
      TabIndex        =   15
      Top             =   360
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   14
      Left            =   1155
      TabIndex        =   14
      Top             =   330
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   13
      Left            =   1050
      TabIndex        =   13
      Top             =   420
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   12
      Left            =   1140
      TabIndex        =   12
      Top             =   390
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   11
      Left            =   870
      TabIndex        =   11
      Top             =   480
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   10
      Left            =   1050
      TabIndex        =   10
      Top             =   360
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   9
      Left            =   945
      TabIndex        =   9
      Top             =   450
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   8
      Left            =   1035
      TabIndex        =   8
      Top             =   420
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   7
      Left            =   885
      TabIndex        =   7
      Top             =   435
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   6
      Left            =   975
      TabIndex        =   6
      Top             =   405
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   5
      Left            =   870
      TabIndex        =   5
      Top             =   495
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   465
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   3
      Left            =   780
      TabIndex        =   3
      Top             =   465
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   2
      Left            =   870
      TabIndex        =   2
      Top             =   435
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   1
      Left            =   765
      TabIndex        =   1
      Top             =   525
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   0
      Left            =   855
      TabIndex        =   0
      Top             =   495
      Width           =   210
   End
End
Attribute VB_Name = "AWINAMP_CONTROL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim G, NEWWINAMP
Dim cText As String
Dim FPATH As String
Private cProcesses As New clsCnProc

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long


Private Sub Form_Load()


Call WINAMP_Timer_Timer

If NEWWINAMP = True Then
    Call ADDLISTITEMS
End If

End Sub


Sub ADDLISTITEMS()

x = 0
For Each Control In Label1
    Control.Caption = ""
    Control.Visible = False
    Control.AutoSize = True
Next

R = 0
For Each Control In Label1
    
    If R >= List1.ListCount Then
        Control.Visible = False
    Else
        Control.Caption = Mid(List1.List(R), InStr(List1.List(R), " -- ") + 3)
        Control.AutoSize = False
        Control.Top = x
        Control.Left = 0
        x = x + Control.Height + 15
        Control.Visible = True
        
        If Control.Width > W Then W = Control.Width
    End If
    R = R + 1
Next

W = 0
x = 0
For Each Control In Label1
    If Control.Visible = True Then
        If Control.Height + Control.Top > x Then x = Control.Height + Control.Top
        If Control.Width > W Then W = Control.Width
    End If
Next

Me.Height = x + 470
Me.Width = W + 120


End Sub


Private Sub WINAMP_Timer_Timer()
i = FindWindow("Winamp v1.x", vbNullString)
NEWWINAMP = False
If i <> OLDi Then
    OLDi = i
    'ADD NEW ITEM
    Call FINDWINAMPPART
    Call FINDWINAMPPARTVIDEO
    
    
    
    'MAKE LIST
    For R = List1.ListCount - 1 To 0 Step -1
        HNDL = List1.List(R)

        TTx1 = GetWindowTitle(Val(HNDL))
        If TTx1 = "" Then
            List1.RemoveItem (R)
        End If

    Next
    
    'DUPE COPY REMOVED
    For R = List1.ListCount - 1 To 0 Step -1
        If Mid(List1.List(R), 1, 10) = Mid(List1.List(R + 1), 1, 10) And InStr(List1.List(R), "*ONE") = 0 Then
            List1.RemoveItem (R)
        End If
    
    Next
    
End If

Call ADDLISTITEMS

End Sub


Sub FINDWINAMPPART()

test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    
Do While test_hwnd <> 0
    
    cText = Space$(255)
    G = GetClassName(test_hwnd, cText, 255)
    cText = StripNulls(cText)
    If cText = "Winamp v1.x" Then
        DUPE = 0
        For R = 0 To List1.ListCount
            If InStr(List1.List(R), str(test_hwnd)) Then
                DUPE = 1
            End If
        Next
        If DUPE = 0 Then
            NEWWINAMP = True
            
            FPATH = GetFileFromHwnd(test_hwnd)
            FPATH = Mid(FPATH, InStrRev(FPATH, "\", Len(FPATH) - 15) + 1)
            FPATH = Mid(FPATH, 1, InStrRev(FPATH, "\", Len(FPATH) - 1) - 1)
            List1.AddItem (str(test_hwnd)) + " -- " + FPATH
        End If
    
    End If
    
    
    
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop






End Sub


Sub FINDWINAMPPARTVIDEO()

test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    
Do While test_hwnd <> 0
    
    cText = Space$(255)
    G = GetClassName(test_hwnd, cText, 255)
    cText = StripNulls(cText)
    
    If cText = "Winamp Video" Then
                fIsDBCVisible = IsWindowVisible(test_hwnd)
                If fIsDBCVisible = 1 Then
        
        XM = 0
        For R = 0 To List1.ListCount - 1
                    
        j2 = 0
          Do While i <> 0
             j2 = i
             i = GetParent(i)
        Loop
        i = j2
                
                If InStr(List1.List(R), str(test_hwnd)) > 0 Then XM = 1
                If InStr(List1.List(R), str(i)) > 0 And InStr(List1.List(R), "ONE **") > 0 Then XM = 1
                
        Next
    

                If XM = 0 Then
                    FPATH = GetFileFromHwnd(test_hwnd)
                    FPATH = Mid(FPATH, InStrRev(FPATH, "\", Len(FPATH) - 15) + 1)
                    FPATH = Mid(FPATH, 1, InStrRev(FPATH, "\", Len(FPATH) - 1) - 1)
                    List1.AddItem (str(test_hwnd)) + " -- " + FPATH
                    List1.AddItem (str(j2)) + " -- " + FPATH + " *ONE **" + str(test_hwnd)
            End If
            End If
        
    End If
    
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop






End Sub


Public Function GetFileFromHwnd(lngHwnd) As String

'MsgBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function


