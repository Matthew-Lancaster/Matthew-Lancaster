VERSION 5.00
Begin VB.Form FormStart 
   Caption         =   "Form2"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TXmsg
Dim ExeName
Dim ExeFullPath
Dim ProjectName
Dim ProjectFullPath

Public Sub FormStartLoader()

'DaysToScan = 110

FontSizez = 10

Call OpenProfileScan

'Exit Sub

ScanPath.DTPicker1(0) = vbNullString
ScanPath.cboMask.Text = "*.*"
'ScanPath.cboMask.Text = "*.lnk;*.pif"
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath.Text = "E:\01 VB Shell Folders\00 Shell " + OIP2$ + " Loader"
Call ScanPath.cmdScan_Click
AStart = ScanPath.ListView1.ListItems.Count

'ScanPath.ListView1.SortOrder = lvwAscending
'ScanPath.ListView1.SortKey = 1
'ScanPath.ListView1.Sorted = True
'ScanPath.ListView1.Sorted = False
'ScanPath.ListView1.SortKey = 0
'ScanPath.ListView1.Sorted = True
'ScanPath.ListView1.Sorted = False

For we = 1 To ScanPath.ListView2.ListItems.Count
    A1$ = ScanPath.ListView2.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView2.ListItems.Item(we)
    If FS.FileExists(A1$ + B1$) = True Then
        
        '-- I think of everything dont allow is dont exist IDEA yerse
        'add's EXE's here
        With ScanPath.ListView1
            Set LV = .ListItems.Add(, , B1$)
            LV.SubItems(1) = A1$
            'LV.SubItems(2) = DDirectory
        End With
    End If
Next
ScanPath.ListView2.ListItems.Clear

'ScanPath.ListView1.SortOrder = lvwDescending
'ScanPath.ListView1.SortKey = 3
'ScanPath.ListView1.Sorted = True
'ScanPath.ListView1.Sorted = False
DoEvents


End Sub


Sub OpenProfileScan()

Dim LV As ListItem

'ScanPath.Show 'Test

ScanPath.chkCopyMemory.Value = vbChecked

TXmsg = OIP$ + " Auto Loader"
Form1.Caption = TXmsg

ScanPath.cboMask = "*.vbp"
ScanPath.cboDate.ListIndex = 0
DaysToScanYears = 18
DaysToScan = 365 * DaysToScanYears
ScanPath.DTPicker1(0) = Now - (DaysToScan)  'All in Last Ten Days

ScanPath.txtPath.Text = "D:\VB6\VB-NT"
Call ScanPath.cmdScan_Click

'ScanPath.txtPath.Text = "X:\00 Lists-Common-Words\Acronyms Code"
'Call ScanPath.cmdScan_Click
'ScanPath.txtPath.Text = "X:\00 Lists-Common-Words\Double Word Code"
'Call ScanPath.cmdScan_Click


On Local Error GoTo 0
On Error GoTo 0
'SORT ON DATE CHECK - YES - DOUBLE CHECKED YES KEY # 3
'ScanPath.Show
ScanPath.ListView1.SortOrder = lvwDescending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
DoEvents
'On Error

ProProjects = ScanPath.ListView1.ListItems.Count

'TRANSFER TO LISTVIEW3 MAKE IT EASIER TO HAVE THE MENU DISPLAY
For we = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    LINK_DATA = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    E1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
    
    With ScanPath.ListView3
        Set LV = .ListItems.Add(, , Format(we, "##0"))
        LV.SubItems(1) = B1$
        LV.SubItems(2) = A1$
        LV.SubItems(3) = LINK_DATA
        LV.SubItems(4) = E1$
        'LV.SubItems(2) = DDirectory
    End With
Next

ScanPath.ListView3.SortOrder = lvwDescending
ScanPath.ListView3.SortKey = 4
ScanPath.ListView3.Sorted = True
ScanPath.ListView3.Sorted = False
DoEvents


'Exit Sub

SCRIPT_TOTAL_PROJECTS = 50
''TRIM DOWN WE ONLY WANT MOST RECENT
'If ScanPath.ListView1.ListItems.Count > SCRIPT_TOTAL_PROJECTS Then
'    Do
'       ' DoEvents
'        ScanPath.ListView1.ListItems.Remove (ScanPath.ListView1.ListItems.Count)
'    Loop Until ScanPath.ListView1.ListItems.Count <= SCRIPT_TOTAL_PROJECTS
'End If

'Exit Sub

DoEvents
we = 0
If ScanPath.ListView1.ListItems.Count > 0 Then
Do
    we = we + 1
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    vFile = A1$ + B1$
    If InStr(LCase(Right(B1$, 4)), ".vbp") > 0 Then
        Call AddProject(vFile)
        
        If vFile <> "" Then
            LINK_DATA = Mid$(vFile, 1, InStrRev(vFile, "\"))
            D2$ = Mid$(vFile, InStrRev(vFile, "\") + 1)
                
            'ADD THE PROJECTS EXECUTABLE
            With ScanPath.ListView1
            Set LV = ScanPath.ListView1.ListItems.Add(we, , D2$)
            LV.SubItems(1) = LINK_DATA
            'LV.SubItems(2) = DDirectory
            End With
            we = we + 1
        End If
    End If
Loop Until we >= ScanPath.ListView1.ListItems.Count Or ProProjects2 >= SCRIPT_TOTAL_PROJECTS
'ProProjects = ScanPath.ListView1.ListItems.Count
'ProProjects2 = we
End If

'ScanPath.ListView1.SortOrder = lvwAscending
'ScanPath.ListView1.SortKey = 0
'ScanPath.ListView1.Sorted = True
'ScanPath.ListView1.SortKey = 1
'ScanPath.ListView1.Sorted = True
'ScanPath.ListView1.Sorted = False


'Exit Sub
'ProProjects = 0
'For we = 1 To ScanPath.ListView1.ListItems.Count
'    B1$ = ScanPath.ListView1.ListItems.Item(we)
'    If InStr(LCase(Right(B1$, 4)), ".vbp") > 0 Then
'        ProProjects = ProProjects + 1
'    End If
'Next

ProProjects2 = 0

For we = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    If InStr(LCase(Right(B1$, 4)), ".vbp") > 0 Then
        ProProjects2 = ProProjects2 + 1
        If ProProjects2 >= SCRIPT_TOTAL_PROJECTS Then Exit For
    End If
    
    With ScanPath.ListView2
        Set LV = .ListItems.Add(, , B1$)
        LV.SubItems(1) = A1$
        'LV.SubItems(2) = DDirectory
    End With

Next



ScanPath.ListView1.ListItems.Clear



End Sub

Sub AddProject(vFile)
  
  Dim pFile As String
  pFile = vFile
  
  
  Dim Readln As String
  Dim Num As Integer
  Dim Path32 As String
  Dim ExeName32 As String
  Dim ExeName As String
  Num = FreeFile()
  
  
sex = 1
  'Retreive Compile information from VBP file
  Open vFile For Input As #Num
    Do While Not EOF(Num)
      Line Input #Num, Readln
      
      If Left(Readln, 9) = "ExeName32" Then
        sex = 0
        ExeName32 = Replace(Readln, "ExeName32=" & Chr(34), "")
        ExeName32 = Left$(ExeName32, Len(ExeName32) - 1)
      ElseIf Left(Readln, 6) = "Path32" Then
        sex = 0
        Path32 = Replace(Readln, "Path32=" & Chr(34), "")
        Path32 = Left$(Path32, Len(Path32) - 1)
      End If
    Loop
  Close #Num
  
  If sex = 1 Then vFile = "": Exit Sub
  'Get Realtive Path to the project file
  Path32 = GetRelativePath(Path32, Left$(pFile, Len(pFile) - (Len(GetFileName(pFile))) - 1))
    
  If InStr(ExeName32, ":\") = 0 Then
      ExeName = Path32 & "\" & ExeName32
      ExeName = Replace(ExeName, "\\", "\")
    Else
    ExeName = ExeName32
    ExeName = Replace(ExeName, "\\", "\")
  End If
  
  
vFile = ExeName

  
End Sub


Function GetRelativePath(findPath As String, startPath As String) As String
  
  Dim l As Integer
  Dim I As Integer
  Dim Backs As Integer
  Dim vstartPath As String
  
    GetRelativePath = startPath
    If findPath = startPath Then Exit Function
  
  vstartPath = startPath
  
  'Find out how many BackDirs (..\) there are
  l = Len(findPath)
  findPath = Replace(findPath, "..\", "")
  Backs = (l - Len(findPath)) / 3
  
  'Back up BACKS BackDirs
  For I = 1 To Backs
    If I = 1 Then
      l = InStrRev(vstartPath, "\")
    Else
      l = InStrRev(vstartPath, "\", l - 1)
    End If
    vstartPath = Left(vstartPath, l - 1)
  Next
  
  GetRelativePath = vstartPath & "\" & Trim(Replace(findPath, Chr(34), ""))
  
End Function

Function GetFileName(vPath As String) As String
  Dim spot As Integer
  
  spot = InStrRev(vPath, "\")
  
  If spot <= 0 Then
    If Len(Trim(vPath)) = 0 Then
      GetFileName = ""
    Else
      GetFileName = vPath
    End If
  Else
    GetFileName = Mid(vPath, spot + 1)
  End If
End Function


Sub LabelClick(Index)
    
'If Check1.Value = vbUnchecked Then
'Open App.Path + "\Winloads.txt" For Append As #1
'Print #1, Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Label1(Index)
'Close #1
'Shell Dir1.List(Index) + "\Winamp.exe", vbNormalFocus

A1$ = ScanPath.ListView1.ListItems.Item(Index).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(Index)

ChDrive Mid$(A1$, 1, 2)
ChDir A1$

If LoadFolder = True Then Shell "Explorer.exe /select, " + A1$ + B1$, vbNormalFocus: Unload Form1: Exit Sub ': End

'Load File From Link
LINK_DATA = A1$ + B1$
'If InStr(LCase(B1$), ".lnk") > 0 And InStr(LCase(B1$), ".txt") Then
If InStr(LCase(B1$), ".lnk") > 0 Then
    
    Open A1$ + B1$ For Binary As #1
        rr$ = Space$(LOF(1))
        Get #1, 1, rr$
    Close #1
    'If Mid$(rr$, &H568, 3) = ":" + Chr$(0) + "\" Then
    'rr$ = Mid$(rr$, &H566)
    'End If
    If InStr(rr$, ":\") = 0 Then
    
        I = MsgBox("INVAILD LINK NAUGHT SIZE" + vbCrLf + LINK_DATA, vbMsgBoxSetForeground)
        
        I = MsgBox("INVAILD LINK" + vbCrLf + LINK_DATA + vbCrLf + "DO YOU WANT ME TO TAKE YOU THERE", vbYesNo + vbMsgBoxSetForeground)
        If I = vbYes Then
            Shell "Explorer.exe /select, " + LINK_DATA, vbNormalFocus
            Unload Form1
            Exit Sub
        End If
        Exit Sub
    End If
    
    tt$ = Mid$(rr$, InStr(105, rr$, ":\") - 1)
    LINK_DATA = Mid$(tt$, 1, InStr(tt$, Chr$(0)) - 1)
    
    If InStr(LINK_DATA, "Blue Tooth Sync Folder") > 0 Then MsgBox ("Wrong Link Points to Blue Tooth folder"): End

    Form1.LINK_DATA = LINK_DATA

    Call Form1.TEST_LINK_MODIFY
    
    LINK_DATA = Form1.LINK_DATA

End If

'If LoadFolderFile = True And InStr(LCase(B1$), ".txt") Then
''    LINK_DATA = Mid$(TT$, 1, InStr(TT$, Chr$(0)) - 1)
'    'LINK_DATA = Mid$(LINK_DATA, 1, InStrRev(LINK_DATA, "\"))
'    Shell "explorer /e, /select, " + LINK_DATA, vbNormalFocus
'    Unload Form1
'    Exit Sub
'End If

If LoadFolderFile = True Then
    Shell "explorer /e, /select, " + LINK_DATA, vbNormalFocus
    Unload Form1
    Exit Sub
End If



If InStr(LCase(Right(B1$, 4)), ".vbp") > 0 Then
    DumVar = LastModifiedToCurrent(A1$ + B1$)
    On Error GoTo 0

    ChDrive Mid$(A1$, 1, 2)
    ChDir A1$
    
    VBPATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VBPATH) = "" Then
        VBPATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    End If
    If Dir(VBPATH) = "" Then
        MsgBox "%PROGRAM FILES%\\Microsoft Visual Studio\VB98\VB6.EXE" + vbCrLf + "NOT FOUND _ END", vbMsgBoxSetForeground
        End
    End If
    
    
    
    ' --------------------------------------------------------------------
    ' FIND IF THE IS NOT CORRECT VERSION AND SIMPLE PUT CORRECT
    ' ONE MY COMPUTER SEEM PROBLEM WITH BASIC
    ' TALK HAS WRONG VERISON ONCE EVERY FEW DAY
    ' AND EDIT IT TO WHAT NORM VERISON SUPPOSED TO BE AND FINE
    ' DISCOVERY -- IT SEEM TO HAPPEN JUST AFTER A COMPILE SOMETIME_
    ' [ Monday 15:37:30 Pm_17 June 2019 ]
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    Dim R, FR
    Dim VAR_STRING As String

    FR = FreeFile
    Open A1$ + B1$ For Binary As FR
        VAR_STRING = Space(LOF(FR))
        Get #FR, , VAR_STRING
    Close FR

    XR = InStr(UCase(VAR_STRING), UCase("A1}#2.2#0; mscomctl.OCX"))
    If XR > 0 Then
        Mid(VAR_STRING, XR, Len("A1}#2.1#0; mscomctl.OCX")) = "A1}#2.1#0; MSCOMCTL.OCX"
        MsgBox UCase("#2.2# -- mscomctl.OCX") + vbCrLf + "WRONG VERSION -- AUTO CHANGED TO" + vbCrLf + UCase("#2.1# -- mscomctl.OCX"), vbMsgBoxSetForeground
        GO_NEXT_IN = True
    End If
    
    If GO_NEXT_IN = True Then
        If Dir(A1$ + B1$) <> "" Then
            Kill A1$ + B1$
        End If
        FR = FreeFile
        Open A1$ + B1$ For Binary As FR
            Put #FR, , VAR_STRING
        Close FR
    End If
    
    
    Shell VBPATH + " """ + A1$ + B1$ + """", vbNormalFocus
    End
End If

'This run for most VBP Progs


vLaunch A1$ + B1$, vbNullString
Unload Form1

Exit Sub
End


que = 0
If A1$ = "--Drive" Then
que = 1
Shell "Explorer.exe /e," + Mid$(B1$, 1, 3), vbNormalFocus
Else
Open A1$ + B1$ For Binary As #1
rr$ = Space$(LOF(1))
Get #1, 1, rr$
If Mid$(rr$, &H568, 3) = ":" + Chr$(0) + "\" Then
rr$ = Mid$(rr$, &H566)
End If

'que = 0
If B1$ = "My Computer.lnk" Then
vLaunch A1$ + B1$, vbNullString
End
Exit Sub
End If

If que = 0 Then
tq = InStrRev(rr$, ":\")
If tq = 0 Then
    tq = InStrRev(rr$, ":" + Chr(0) + "\")
    If tq > 0 Then tq = tq - 1
End If
If tq = 0 Then
'MsgBox "Not Right Path Not Found in shortcut Prog Launch Explorer"
vLaunch A1$ + B1$, vbNullString
End
End If
rr$ = Mid$(rr$, tq - 1)
End If

tq = InStr(rr$, Chr$(0) + Chr$(0))
rr$ = Mid$(rr$, 1, tq - 1)

Do
tq = InStr(rr$, Chr$(0))
If tq > 0 Then rr$ = Mid$(rr$, 1, tq - 1) + Mid$(rr$, tq + 1)
Loop Until tq = 0

Shell "Explorer.exe /e," + rr$, vbNormalFocus
'vLaunch a1$ + b1$

End If

End


End Sub

