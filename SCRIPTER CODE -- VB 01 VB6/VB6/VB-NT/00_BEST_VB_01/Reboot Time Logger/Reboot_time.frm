VERSION 5.00
Begin VB.Form Reboot_Time 
   Caption         =   "Boot Time Logger & Delayed Program Launcher"
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5640
   Icon            =   "Reboot_time.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   870
      Top             =   45
   End
End
Attribute VB_Name = "Reboot_Time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

'---In Module---'
Private Type TimeConv
    Days        As Long
    Hours       As Long
    Minutes     As Long
    Seconds     As Long
    MSeconds    As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

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



Private Sub Form_Load()

'CSIDL_STARTUP

'startup$ = GetSpecialFolderPath(CSIDL_STARTUP)

Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\2CidRunMe.txt" For Input As #1
'Open App.Path + "Reboot Loggon Logg.txt" For Input As #1
Seek 1, LOF(1) - 10000
wer$ = Input(10000, #1)
'wer$ = Input(LOF(1), #1)
Close #1

w1 = InStrRev(wer$, "ReBoot Time")
If w1 > 0 Then
    w2 = InStrRev(wer$, vbCr, w1) + 2
End If
If w1 > 0 Then
    pts$ = Mid$(wer$, w2, 19)
    Dim wtr As Long
    wtr = DateValue(pts$) + TimeValue(pts$)
End If


wtr = DateDiff("n", wtr, Now)

If wtr > 18 And _
    SystemUptime.Days = 0 And _
     SystemUptime.Hours = 0 And _
      SystemUptime.Minutes < 20 Then
    ballistic$ = "REBOOT"
Else
    ballistic$ = "LOGON"
End If

'Open App.Path + "Reboot Loggon Logg.txt" For Append As #1
Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\2CidRunMe.txt" For Append As #1
If ballistic$ = "REBOOT" Then
    Print #1, Now; " ***"; Format$(Now, "DDD dd-MMM-YYYY"); " Last System ReBoot Time.+++++++++"
End If
If ballistic$ = "LOGON" Then
    Print #1, Now; " ***"; Format$(Now, "DDD dd-MMM-YYYY"); " Last Logon Time.-----------------"
End If
Close #1

End

End Sub


' Get a special folder's path.
' CSIDL_STARTUP
Private Function GetSpecialFolderPath(ByVal folder_number As Long) As String
Dim Path As String

    Path = Space$(MAX_PATH)
    If SHGetSpecialFolderPath(hwnd, Path, _
        folder_number, False) _
    Then
        GetSpecialFolderPath = Left$(Path, InStr(Path, Chr$(0)))
    Else
        GetSpecialFolderPath = "???"
    End If
End Function


''---In Module---'
'Private Type TimeConv
'    Days        As Long
'    Hours       As Long
'    Minutes     As Long
'    Seconds     As Long
'    MSeconds    As Long
'End Type
 
'Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
 

Private Function ConvertTime(ByVal Tick As Long) As TimeConv
 
ConvertTime.MSeconds = Tick Mod 1000
 
Tick = Tick \ 1000
ConvertTime.Days = ((Tick) \ (24 * (60 ^ 2)))
 
If ConvertTime.Days > 0 Then Tick = (Tick - 24 * (60 ^ 2)) * ConvertTime.Days
ConvertTime.Hours = Tick \ (60 ^ 2)
 
If ConvertTime.Hours > 0 Then Tick = Tick - ((60 ^ 2) * ConvertTime.Hours)
ConvertTime.Minutes = Tick \ 60
 
ConvertTime.Seconds = Tick Mod 60
 
End Function
 
 
Private Property Get SystemUptime() As TimeConv
SystemUptime = ConvertTime(GetTickCount)
End Property
 
 
'---Form Usage Example:---'
Private Sub Timer2_Timer()

Label4 = "System Uptime: " + vbCr & SystemUptime.Days & " Days, " & _
        SystemUptime.Hours & ":" & _
        SystemUptime.Minutes & ":" & _
        SystemUptime.Seconds & "." & _
        SystemUptime.MSeconds & "."

End Sub


