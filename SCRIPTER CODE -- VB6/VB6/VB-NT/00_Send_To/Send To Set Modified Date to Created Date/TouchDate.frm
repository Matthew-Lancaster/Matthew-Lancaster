VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000016&
   Caption         =   "FormMODIFY DATE AS CREATED DATE1"
   ClientHeight    =   4488
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10716
   LinkTopic       =   "Form1"
   ScaleHeight     =   4488
   ScaleWidth      =   10716
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label8 
      Caption         =   "SET ONE DATE HARDCODED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   100
      TabIndex        =   7
      Top             =   3012
      Width           =   6348
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label7"
      Height          =   180
      Left            =   8304
      TabIndex        =   6
      Top             =   2328
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label6"
      Height          =   240
      Left            =   7512
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   576
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   100
      TabIndex        =   4
      Top             =   3744
      Width           =   10500
   End
   Begin VB.Label Label4 
      Caption         =   "FOLDER LABEL4"
      Height          =   528
      Left            =   100
      TabIndex        =   3
      Top             =   960
      Width           =   10500
   End
   Begin VB.Label Label3 
      Caption         =   "PERFORM ON ALL FILES IN FOLDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   100
      TabIndex        =   2
      Top             =   276
      Width           =   7176
   End
   Begin VB.Label Label2 
      Caption         =   "MODIFY DATE AS CREATED DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   100
      TabIndex        =   1
      Top             =   2340
      Width           =   6348
   End
   Begin VB.Label Label1 
      Caption         =   "NOW DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   100
      TabIndex        =   0
      Top             =   1632
      Width           =   4116
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Dim W$
Dim WORK
Dim FSO

'--------------------------------------

'Excellent Code Set the File Time To ANything you want touch date Pro Keep you file time after rotate JPg's




'Option Explicit

'THIS TOGGLES A SMALL PASSWORD FEATURE THAT I EMBEDDED
Const MPassword = False

'THIS IS THE CODE THAT MUST BE ENTERED (FINISHED BY PRESSING HASH)
'WHEN THE PASSWORD TOGGLE IS TRUE
Const MSequence = 22515

Dim MEntered As Variant

 Private Type FILETIME
     dwLowDate  As Long
     dwHighDate As Long
 End Type
 
 Private Type SYSTEMTIME
     wYear      As Integer
     wMonth     As Integer
     wDayOfWeek As Integer
     wDay       As Integer
     wHour      As Integer
     wMinute    As Integer
     wSecond    As Integer
     wMillisecs As Integer
 End Type
 
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
  
Private Declare Function CreateFile Lib "kernel32" Alias _
   "CreateFileA" (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) _
   As Long

Private Declare Function LocalFileTimeToFileTime Lib _
     "kernel32" (lpLocalFileTime As FILETIME, _
      lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
   (ByVal hFile As Long, ByVal MullP As Long, _
    ByVal NullP2 As Long, lpLastWriteTime _
    As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
    "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
    As FILETIME) As Long
    
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Public Function SetFileDateTime(ByVal Filename As String, _
  ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
If Dir(Filename) = "" Then Exit Function
If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

lFileHnd = CreateFile(Filename, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)
    
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
SetFileDateTime = lRet > 0

End Function














Private Sub Form_Load()
Set FSO = CreateObject("Scripting.FileSystemObject")


'If IsIDE = True Then GoTo Start2

W$ = "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe"

W$ = "E:\01 FAVS\Ministers\Images"

W$ = "M:\0 00 Art\Google Downloader\00 MICHELLE"
W$ = "U:\# MY DOCS\# 01 My Documents\02_GOOGLE_EARTH\#Google Earth\My Places ASUS BIG\files"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2009 FaceBook Photo's\#\Part 1\DSC02338.JPG"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-00 - SONY DSC-HX5V - DSC00321.JPG"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-14 - SONY DSC-HX5V - DSC00322.JPG"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-26 - SONY DSC-HX5V - DSC00323.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 21 13.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 20 38.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 19 43.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 17 52.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 14 11.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 10 51.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 03 54 49.JPG"

W$ = "M:\DSC\2015\10350614"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\Facebook Holiday Inn Hotel"
W$ = "M:\DSC\2005-2007\2005 NEXT"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-19 Kings Hotel"
W$ = "M:\DSC\2010\# TEST"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2008-11-29 Best Western Brighton Stay in On the Run"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2008-12-08 Best Western"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-09 Best Western\test"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-16 Holiday Inn"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-19 Kings Hotel\test"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-10-27 Meadowfield Hospital Maple Ward Asylum\test"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2009 FaceBook London"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Rutland Gardens"
W$ = "D:\Videos\00 Vid XXX\ME\2011 GALAXY SAMSUNG GT-P1000 MP4"
W$ = "D:\Videos\00 Vid XXX\ME"
W$ = "D:\Videos\1999 Accidents\"
W$ = "D:\0 00 ART LOGGERS - WEBCAM\VIDEO\"
W$ = "D:\DSC\# Docus Proofs Texts\# BHT\NOTICE BOARD"


If Command$ <> "" Or W$ <> "" Then
    If W$ = "" Then W$ = Command$
Else
    If W$ <> "" Then
        a = a
    Else
        End
    End If
End If


Label4.Caption = W$
End Sub


Sub BATCH_MODIFIED_TO_CREATED_TIME()

If Mid(W$, 1, 1) = """" Then
    W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
End If

If Dir$(W$) <> "" Then
    DIRW$ = Mid$(W$, 1, InStrRev(W$, "\") - 1)
End If
    
If Mid$(W$, Len(W$), 1) <> "\" Then
    W$ = W$ + "\"
End If
    
    
On Error Resume Next
XX = FSO.FolderExists(W$)

Dim DateSet As Date

If XX = False Then
'time2$ = "2011-11-01 10:00:00"
'DateSet = DateValue(time2$) + TimeValue(time2$)
'DateSet = Now
'
'tt = SetFileDateTime(W$, DateSet)

'tt = LastModifiedToCurrent(W$)
End
End If

If DIRW$ <> "" Then W$ = DIRW$

'MsgBox str(XX) + "--" + W$
'End

Start2:

Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.*"

'ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\#000 New\2009-08 Aug\Unet\Mine\Word Example Usage"
'ScanPath.txtPath.Text = "E:\01 FAVS\Ministers\Images"
ScanPath.txtPath.Text = W$

If Len(W$) < 5 Then End

'ScanPath.Show
Dim DT1 As Date
'Dim Ds1 As Date
Dim Ds2 As Date
'Dim DateSet As Date
For we = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFile(A1$ + B1$)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    
    DateSet = DT3
    
    Set F = Nothing
    
    tt = SetFileDateTime(A1$ + B1$, DateSet)
    XC = XC + 1
    
Next

MsgBox "Done = " + vbCrLf + str(XC) '+dd$
End


End Sub

Sub CODE_RUN()



If Mid(W$, 1, 1) = """" Then
    W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
End If

If Dir$(W$) <> "" Then
    DIRW$ = Mid$(W$, 1, InStrRev(W$, "\") - 1)
End If
    
If Mid$(W$, Len(W$), 1) <> "\" Then
    W$ = W$ + "\"
End If
    
    
On Error Resume Next
XX = FSO.FolderExists(W$)

Dim DateSet As Date

If XX = False Then
'time2$ = "2011-11-01 10:00:00"
'DateSet = DateValue(time2$) + TimeValue(time2$)
'DateSet = Now
'
'tt = SetFileDateTime(W$, DateSet)

'tt = LastModifiedToCurrent(W$)
End
End If

If DIRW$ <> "" Then W$ = DIRW$

'MsgBox str(XX) + "--" + W$
'End

Start2:

Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.avi"

'ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\#000 New\2009-08 Aug\Unet\Mine\Word Example Usage"
'ScanPath.txtPath.Text = "E:\01 FAVS\Ministers\Images"
ScanPath.txtPath.Text = W$

If Len(W$) < 5 Then End

'ScanPath.Show
Dim DT1 As Date
'Dim Ds1 As Date
Dim Ds2 As Date
'Dim DateSet As Date
For we = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFile(A1$ + B1$)
    DT1 = F.datelastmodified
    
    DateSet = DT1 - TimeSerial(1, 0, 0)
    
    Set F = Nothing
    
    'Ds1 = Mid(B1$, 1, 20)
'    Ds1 = Replace(Ds1, "+", ":") + ":00"
'    If InStr(B1$, "2008 011(Nov) 29 Sat 08-04-10") > 0 Then
    'If Year(DT1) = 2009 Then
        '2008 012(Dec) 04 Thu 18-22-52 - W880I - DSC00331.JPG
      '  Ds1 = DateValue("04-Dec-2008") + TimeValue("18:22:52")
        
'        xb1$ = B1$
'        xb1$ = Replace(xb1$, "MILL VIEW HOSPITAL-", "")
'
        If A1$ + B1$ = "D:\0 00 ART LOGGERS - WEBCAM\VIDEO\Web_Video_Capture- 2015-10-03 14-28-59 Sat -- Microsoft WDM Image Capture (Win32)2.avi" Then
'            Ds1 = DateValue(Mid(xb1$, 9, 2) + "-" + Mid(xb1$, 6, 2) + "-" + Mid(xb1$, 1, 4))
'            Ds2 = TimeValue(Replace(Mid(xb1$, 12, 8), "-", ":"))
        '    DateSet = DateValue(Ds1) + TimeValue(Ds1)
'            DateSet = Ds1 + Ds2
    '        DateSet = GetExif(a1$ + B1$)
            DateSet = DateValue("2015-10-03") + TimeValue("14:28:59")
                  
            '
            tt = SetFileDateTime(A1$ + B1$, DateSet)
            XC = XC + 1
            'tt = LastModifiedToCurrent(a1$ + B1$)
        End If
'        Stop
        'End
    'End If
    
Next

MsgBox "Done = " + vbCrLf + str(XC) '+dd$
End


End Sub


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************


Function GetExif(File$)
    Dim objExif As New ExifReader
    Dim txtExifInfo As String
    
    objExif.Load File$
    txtExifInfo = objExif.Tag(DateTimeOriginal)
    GetExif = txtExifInfo
End Function


'Function GetProperty(strFile, n)
'Dim objShell
'Dim objFolder
'Dim objFolderItem
'Dim i
'Dim strPath
'Dim strName
'Dim intPos
'
'On Error GoTo ErrHandler
'
'intPos = InStrRev(strFile, "")
'strPath = Left(strFile, intPos)
'strName = Mid(strFile, intPos + 1)
'Set objShell = CreateObject("Shell.Application")
'Set objFolder = objShell.Namespace(strPath)
'Set objFolderItem = objFolder.ParseName(strName)
'If Not objFolderItem Is Nothing Then
'GetProperty = objFolder.GetDetailsOf(objFolderItem, n)
'End If
'
'ExitHandler:
'Set objFolderItem = Nothing
'Set objFolder = Nothing
'Set objShell = Nothing
'Exit Function
'
'ErrHandler:
'MsgBox Err.Description, vbExclamation
'Resume ExitHandler
'End Function

Private Sub Label1_Click()

If Label1.BackColor = Label6.BackColor Then Exit Sub
If Label2.BackColor = Label6.BackColor Then Exit Sub
If Label8.BackColor = Label6.BackColor Then Exit Sub

Label1.BackColor = Label6.BackColor
Label5.BackColor = Label7.BackColor

WORK = "MOD_TO_NOW_DATE"

End Sub

Private Sub Label2_Click()

If Label1.BackColor = Label6.BackColor Then Exit Sub
If Label2.BackColor = Label6.BackColor Then Exit Sub
If Label8.BackColor = Label6.BackColor Then Exit Sub

Label2.BackColor = Label6.BackColor
Label5.BackColor = Label7.BackColor

WORK = "MOD_TO_CREATED_DATE"

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

If WORK = "MOD_TO_NOW_DATE" = True Then
    Call CODE_RUN
    Exit Sub
End If

If WORK = "MOD_TO_CREATED_DATE" = True Then
    Call BATCH_MODIFIED_TO_CREATED_TIME
    Exit Sub
End If

If WORK = "SET_ONE_DATE" = True Then
        
    a = "D:\DSC\# Docus Proofs Texts\# BHT\NOTICE BOARD\2009-09-22 21-23-03 - Sony Ericsson K800i - DSC03044.JPG"
    a = "D:\DSC\# Docus Proofs Texts\# BHT\SMS PROOFS\2009-09-22 23-09-55 - Sony Ericsson K800i - DSC03047.JPG"
    a = "D:\DSC\# Docus Proofs Texts\# BHT\SMS PROOFS\2009-09-22 23-10-12 - Sony Ericsson K800i - DSC03048.JPG"
    'A = ""
    'A = ""
    'A = ""
    'A = ""
    'A = ""
    'A = ""
    'A = ""
    'A = ""
    
    Set F = FSO.GetFile(a)
    DT1 = F.datelastmodified
    Set F = Nothing
    'DateSet = DT1 - TimeSerial(1, 0, 0)
    DATEVAR = "2009-09-22 21-23-03"
    DATEVAR = "2009-09-22 23-09-55"
    DATEVAR = "2009-09-22 23-10-12"
    'DATEVAR = ""
    'DATEVAR = ""
    'DATEVAR = ""
    'DATEVAR = ""
    'DATEVAR = ""
    'DATEVAR = ""
    'DATEVAR = ""
    'DATEVAR = ""
    For r = 1 To 10
    If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = "/"
    Next
    For r = 10 To Len(DATEVAR)
    If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = ":"
    Next
    DateSet = DateValue(DATEVAR) + TimeValue(DATEVAR)
    
'    Ds1 = DateValue(Mid(xb1$, 9, 2) + "-" + Mid(xb1$, 6, 2) + "-" + Mid(xb1$, 1, 4))
'    Ds2 = TimeValue(Replace(Mid(xb1$, 12, 8), "-", ":"))
'    DateSet = DateValue(Ds1) + TimeValue(Ds1)
'    DateSet = Ds1 + Ds2
'    DateSet = GetExif(a1$ + B1$)
            
    tt = SetFileDateTime(a, DateSet)
    XC = XC + 1
    
    MsgBox "Done = " + str(XC)
    End
    
End If




End Sub


Private Sub Label8_Click()
If Label1.BackColor = Label6.BackColor Then Exit Sub
If Label2.BackColor = Label6.BackColor Then Exit Sub
If Label8.BackColor = Label6.BackColor Then Exit Sub

Label8.BackColor = Label6.BackColor
Label5.BackColor = Label7.BackColor

WORK = "SET_ONE_DATE"

End Sub
