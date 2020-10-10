Attribute VB_Name = "MULTI_MODULE"
'Option Explicit


Public SCAN_PARTMASK
Public XX, MSGBOX2 ', mFileCount2
Dim TG

Dim BLACKL(), TTH As String, StartX, TX5
Public m_CRC As clsCRC, OFSI1, A22 As String
Dim TdD As Long
Dim RXRIME



Private Type SHFILEOPSTRUCT
hWnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FOF_ALLOWUNDO = &H40&
Private Const FOF_NOCONFIRMATION = &H10&
Private Const FO_COPY = &H2&
Private Const FO_MOVE = &H1&
Private Const FO_DELETE = &H3&
Private Const FO_RENAME = &H4&
Private Const FOF_NOCONFIRMMKDIR = &H200&
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4          ' X - Cannot be set
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10      ' X - CreateDirectory to set
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800    ' X -


Dim WinType_SFO As SHFILEOPSTRUCT
Dim lRet As Long
Dim lFlags As Long


'----------------
'## MODIFY BATCH SHORT CUTS VARS
'Option Explicit

'---------------------------
'Skrol 29
'skrol29@freesurf.fr
'http://www.rezo.net/dir/skrol29/
'---------------------------
'Version 1.00, on 02/13/1999
'Version 1.01, on 04/19/1999
'---------------------------
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_FAVORITES = &H1F

Private Declare Function api_SHAddToRecentDocs Lib _
   "shell32.dll" Alias "SHAddToRecentDocs" (ByVal dwFlags As _
   Long, ByVal dwData As String) As Long

Private Declare Function api_SHGetSpecialFolderLocation Lib _
  "shell32.dll" Alias "SHGetSpecialFolderLocation" (ByVal _
  hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long

Private Declare Function api_SHGetPathFromIDList Lib _
   "shell32.dll" Alias "SHGetPathFromIDList" _
   (ByVal pidl As Long, ByValsPath As String) _
   As Long



'Option Explicit

'Private Sub Command1_Click()
'This will Create a ShortCut of Notepad in our desktop, its name is "Notepad", minimize windows when run, use the 2nd icon as the Shortcut icon.

'Create_ShortCut "C:\WINDOWS\NOTEPAD.EXE", "Desktop", "Notepad", , 7, 1
'End Sub



Sub Create_ShortCut(ByVal TargetPath As String, ByVal ShortCutPath As String, ByVal ShortCutname As String, Optional ByVal WorkPath As String, Optional ByVal Window_Style As Integer, Optional ByVal IconNum As Integer)

Dim VbsObj As Object
Set VbsObj = CreateObject("WScript.Shell")

Dim MyShortcut As Object
'ShortCutPath = VbsObj.SpecialFolders(ShortCutPath)
Set MyShortcut = VbsObj.CreateShortcut(ShortCutPath)

'yy = VbsObj.ExpandEnvironmentStrings(TargetPath)

'MyShortcut.fullname = Nothing
MyShortcut.TargetPath = TargetPath
MyShortcut.WorkingDirectory = WorkPath
MyShortcut.WindowStyle = Window_Style
MyShortcut.IconLocation = TargetPath & "," & IconNum
MyShortcut.TargetPath = "'" + TargetPath + "'"
'Debug.Print TargetPath


MyShortcut.Save

End Sub


'----------------
'## MODIFY BATCH SHORT CUTS VARS
'## END


Public Sub m_CreateShortcut(Filename As Variant, _
    TargetPath As String, Optional ScParam As String, _
   Optional StartFolder As String, Optional IcoNum As Integer, _
   Optional IcoPath As String, Optional WindowMode As Integer)

'If you want to use one of the windows folders for the shortcut
'location, you can pass one of the constants defined in the declarations, e.g.,
' CSIDL_PROGRAMS  = Programs
' CSIDL_STARTUP = Startup
' CSIDL_RECENT = RecentDocs
' CSIDL_DESKTOP = Desktop

'NOTE: AS WRITTEN THIS CODE MUST BE PLACED
'WITHIN A FORM MODULE

'Example:  Puts a shortcut to Notepad on the desktop with
'          a .txt document to be opened

' m_CreateShortcut CSIDL_DESKTOP, "MyFile", _
'   "C:\windows\Notepad.exe", "C:\MyFile.txt"
    
Dim Shortcut0 As String 'Full path for the temporary shortcut
                         'created in the RecentDocs folder.
Dim n0 As Integer       'Cusror position in Shortcut0.
Dim x0 As String * 1    'Variable while reading Shortcut0.
Dim l0 As Long          'Lenth of the Shortcut0 file.
Dim Shortcut1 As String 'Full path for the final shortcut.
Dim n1 As Integer       'Cusror position in Shortcut1
Dim x1 As String * 1    'Variable while reading Shortcut1.
Dim l1 As Long          'Lenth of the Shortcut1 file

Dim t As Double
Dim P As Long
Dim i As Integer
Dim x As String
Dim y0 As String * 2

''Check for the target folder
'If IsNumeric(ScFolder) Then
'    ScFolder = p_GetSpecialFolder(CInt(ScFolder))
'ElseIf Dir$(ScFolder, vbDirectory) = "" Then
'    MsgBox "Le r‚ø¾rtoire '" & ScFolder & "' est introuvable.", _
'       vbCritical, "Cr‚­ïion d'un raccrourci"
'    Exit Sub
'End If

'Create a temporary shortcut with only the
'target in the the RecentDocs.
If api_SHAddToRecentDocs(2, TargetPath) > 0 Then

    'Full path of the created shortcut
    Shortcut0 = p_GetSpecialFolder(8) & "\" & _
        p_File_Folder(TargetPath) & ".lnk"

    'Waiting for the end of the creation.
    t = Now()
    Do Until (Dir$(Shortcut0) <> "")
    
    If (Now() - t) > 0.00006 Then 'wait 5 seconds
        'If MsgBox("Attendre encore la cr‚­ïion du raccourci ?", _
        '    vbQuestion + vbOKCancel, "Raccourci") <> vbOK Then
            'MsgBox "ERROR"
            'Exit Sub
        'Else
            t = Now()
        'End If
    End If
    
    Loop

    'Open the temporary shortcut file in read mode.
    n0 = FreeFile()
    Open Shortcut0 For Binary Access Read As #n0
    'Wait for the file is correctly feed.
    Do Until LOF(n0) > 0
    Loop
    l0 = LOF(n0)

    'Open the shortcut file to create
    'Shortcut1 = ScFolder & "\" & ScCaption & ".lnk"
    If Dir(Filename) <> "" Then Kill Filename
    n1 = FreeFile()
    Open Filename For Binary Access Write As #n1

    'Look for the last byte to get
    P = (l0 - 4)
    y0 = ""
    Do Until (P <= 0) Or (y0 = vbNullChar & vbNullChar)
        Get #n0, P, y0
        P = P - 1
    Loop
    l1 = P + 2

    'Copy bytes
    For P = 1 To l1

        Get #n0, P, x0

        Select Case P
        Case 21 'path for icon, startup, parameters
            i = 3
            If StartFolder <> "" Then
                i = i + 16
            End If
            If ScParam <> "" Then
                i = i + 32
            End If
            If (IcoPath <> "") Or (IcoNum > 0) Then
                i = i + 64
            End If
            x1 = Chr$(i)
        Case 57 'Icon index
            x1 = Chr$(IcoNum)
        Case 61 'Window mode
            x1 = Chr$(WindowMode)
        Case Else
            x1 = x0
        End Select

        Put #n1, P, x1

    Next P

    'Close and delete the temporary shorcut
    Close #n0
    Kill Shortcut0

    'Add the Start folder, parameters and icon file
    x = ""
    If StartFolder <> "" Then
        x = x & Chr$(Len(StartFolder)) & vbNullChar & StartFolder
    End If
    If ScParam <> "" Then
        x = x & Chr$(Len(ScParam)) & vbNullChar & ScParam
    End If
    If IcoPath = "" Then
        If IcoNum > 0 Then
            x = x & Chr$(Len(TargetPath)) & vbNullChar _
               & TargetPath
        End If
    Else
        x = x & Chr$(Len(IcoPath)) & vbNullChar & IcoPath
    End If
    x = x & String(4, vbNullChar)
    Put #n1, l1 + 1, x

    Close #n1

Else

    MsgBox "Error when creating the shortcut.", _
          vbCritical + vbMsgBoxSetForeground, "Shortcut"

End If

End Sub

Private Function p_GetSpecialFolder(CSIDL As Long) As String

'Returns the full path of the folder corresponding to the
'Windows's id system folder.

Dim r     As Long
Dim pidl  As Long
Dim sPath As String

r = api_SHGetSpecialFolderLocation(ScanPath.hWnd, CSIDL, pidl)

If r = 0 Then

    sPath = Space$(260)
    r = api_SHGetPathFromIDList(ByVal pidl, ByVal sPath)
    If r Then
        p_GetSpecialFolder = Left$(sPath, _
           InStr(sPath, Chr$(0)) - 1)
    End If

End If

End Function

Private Function p_File_Folder(FullPath As String) As String
'Returns the name of the file alone.

Dim i As Integer

p_File_Folder = FullPath
i = Len(FullPath)
Do Until i = 0
    If Mid$(FullPath, i, 1) = "\" Then
        p_File_Folder = Mid$(FullPath, i + 1)
        i = 0
    Else
        i = i - 1
    End If
Loop

End Function


Sub MODIFY_BATCH_SHORTCUTS()







End Sub



'Put in Sub Of Code

Sub FIND_REPLACE()
Dim XXD, WE, TTD, A1, B1, G1, Iext, UU, YY, OO, YYXLOGG
    ScanPath.TxtPath = "D:\VB6\VB-NT"


    'MOST HARD CODED
    'SEARCH STRINGS
    '#1 FIND
    '#2 REPLACE
    'ScanPath.Text1 = "    Loop" + vbCrLf + "    " + vbCrLf + "    " + "FindClose lResult"
    'ScanPath.Text2 = "        FindClose lResult" + vbCrLf + "    " + vbCrLf + "    " + "Loop"

    ScanPath.Text1 = "End" + vbCrLf
    ScanPath.Text2 = "End'" + vbCrLf



ScanPath.ListView1.ListItems.Clear
'ScanPath.txtPath.Text = "M:\01 Installations\Games Installations\"


'ScanPath.cboMask.Text = "ScanPath.cls"
ScanPath.cboMask.Text = "*.BAS;*.FRM"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click


ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"


XXD = ScanPath.TxtPath + "\LOGG OF CODE SEARCH REPLACE LOGG.TXT"
If Dir(XXD) <> "" Then Kill XXD

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    Iext = UCase(Mid(B1, Len(B1) - 2))

    'Debug.Print B1

    UU = FreeFile
    
    Open A1 + B1 For Input As #UU
        YY = Input(LOF(UU), UU)
    Close #UU
    
    StartX = 1
    If Iext = "FRM" Then
        StartX = Len(YY)
        TX5 = InStr(YY, "Attribute VB_Name")
        If TX5 > 0 Then
            StartX = TX5
        End If
    End If
    
    
    OO = YY
    If InStr(YY, ScanPath.Text1) > 0 Then
        TdD = TdD + 1
        ScanPath.lblCount3 = Trim(Str(TdD)) + " Changes"
        DoEvents
        
        YYXLOGG = StartX
        Do
            DoEvents
            YYXLOGG = InStr(YYXLOGG, YY, ScanPath.Text1)
            If YYXLOGG = 0 Then Exit Do
            If YYXLOGG > 0 Then
                UU = FreeFile
                Open XXD For Append As #UU
                    
                    Print #UU, "#-- TEST CHUNK---" + A1 + B1
                    
                    Print #UU, Mid(YY, YYXLOGG - 20, 40)
                    
                    Print #UU, "#-- END TEST CHUNK---"
                Close #UU
            End If
            YYXLOGG = YYXLOGG + 1
        Loop Until YYXLOGG = 0
            
                
            
        YY = Replace(YY, ScanPath.Text1, ScanPath.Text2, StartX)
    

        

'        UU = FreeFile
'        Open A1 + B1 For Output As #UU
'            Print #UU, yy;
'        Close #UU

    End If
    
    
    
    

        
        
Next

Shell "notepad " + XXD
    

'DEMO RUN PUT BACK WHEN FINISH
'UU = FreeFile
'Open A1 + B1 For Output As #UU
'    Print #UU, yy;
'Close #UU


ScanPath.lblCount6 = "** DONE **"

ScanPath.WindowState = 0



End Sub


Sub GRAB_STRING()
Dim B1, FN, WE, TTD, A1, G1, RXRIME, oRXRIME, UU, HHG, UU2, YY, OO, Tx1, IIX, PIO, JJ
Dim HH, IIJ, HH1, IIX2

ScanPath.TxtPath = "M:\00 Loggers Text\#00 EMAIL DATA POP3\"

'    'MOST HARD CODED
'    ScanPath.Text1 = "    Loop" + vbCrLf + "    " + vbCrLf + "    " + "FindClose lResult"
'    ScanPath.Text2 = "        FindClose lResult" + vbCrLf + "    " + vbCrLf + "    " + "Loop"


ScanPath.Text2 = """>"
ScanPath.Text1 = "http:"

ScanPath.ListView1.ListItems.Clear

ScanPath.cboMask.Text = "*.eml"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"
    
B1 = ScanPath.ListView1.ListItems.Item(1)
RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
FN = ScanPath.TxtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    If WE = 1 Then
    Do
        RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
        FN = ScanPath.TxtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"
        If FSO.FileExists(FN) = True Then WE = WE + 1
        If WE >= ScanPath.ListView1.ListItems.Count Then End
    Loop Until FSO.FileExists(FN) = False
    
    End If
    
    
    
    RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
    If RXRIME <> oRXRIME Then
        oRXRIME = RXRIME
        FN = ScanPath.TxtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"
        Close #UU
        If Dir(FN) <> "" Then Kill FN
        UU = FreeFile
        Open FN For Append As #UU
    End If
    
    HHG = HHG + "# " + B1

    UU2 = FreeFile
    Open A1 + B1 For Binary As #UU2
        YY = Input(LOF(UU2), UU2)
    Close #UU2
    
    OO = YY
    Tx1 = 0
    
    IIX = 0
    If InStr(LCase(YY), ScanPath.Text1) > 0 Then
        Do
        Tx1 = InStr(Tx1 + 1, YY, ScanPath.Text1)
        If Tx1 > 0 Then
        
        
            TdD = TdD + 1
            ScanPath.lblCount3 = Trim(Str(TdD)) + " Links Found"
                
                PIO = 0
                If InStr(Tx1, YY, ScanPath.Text2) > 0 Then JJ = InStr(Tx1, YY, ScanPath.Text2) - Tx1: PIO = 1
                If PIO = 0 And InStr(Tx1, YY, vbCr) > 0 Then JJ = InStr(Tx1, YY, vbCr) - Tx1
                
                
                HH = Trim(Mid(YY, Tx1, JJ))
                If InStr(HH, """" + vbCr) > 0 Then
                    'A = A
                    HH = Mid(HH, 1, InStr(HH, """" + vbCr) - 1)
                End If
                
                HH = Trim(Mid(YY, Tx1, JJ))
                If InStr(HH, "<") > 0 Then
                    HH = Mid(HH, 1, InStr(HH, "<") - 1)
                End If
                
                If InStr(HH, vbCr) > 0 Then
                    HH = Mid(HH, 1, InStr(HH, vbCr) - 1)
                End If
                
                IIJ = 0
                If InStr(HH, "doubleclick") > 0 Then IIJ = 1
                
                HH = Trim(Mid(YY, Tx1, JJ))
                If InStr(HH, """ ") > 0 And IIJ = 0 Then
                    'A = A
                    HH = Mid(HH, 1, InStr(HH, """ ") - 1)
                End If
                
                    
                If IIJ = 0 Then HH1 = HH1 + HH + vbCrLf: IIX2 = 1
            
            End If
        Loop Until Tx1 = 0
    
        
        If IIX2 = 1 Then
        
            Print #UU, vbCrLf + HHG + vbCrLf
            
            HHG = "# -- " + Trim(Str(TdD)) + " Links Count"
            HHG = HHG + "# -- " + Trim(Str(WE)) + " Email Count"
        
            Print #UU, HHG + vbCrLf
        
            
            Print #UU, HH1
    
        End If
    End If
    
Next
Close #UU

Shell "notepad " + FN

ScanPath.lblCount6 = "** DONE **"

ScanPath.WindowState = 0




End Sub



Sub FilesToTextList()
'F:\00-D-Drive

'TEST SMALL
'ScanPath.txtPath = "M:\0 00 Art\0 MY DEVICES - ALL"

ScanPath.cboMask.Text = "*.JPG"


ScanPath.TxtPath = "M:\0 00 Art"



TG = ScanPath.TxtPath + "\FileList Auto Text -- " + Mid(ScanPath.TxtPath, InStrRev(ScanPath.TxtPath, "\") + 1) + ".txt"
'tg1$ = "FileList Auto Text -- " + Mid(ScanPath.txtPath, InStrRev(ScanPath.txtPath, "\") + 1) + ".txt"


If Dir(TG) <> "" Then Kill TG
Open TG For Append As #1

Print #1, "------------------------------------------------"
Print #1, "Date " + Format(Now, "DDD DD MMM YYYY HH:MM:SS")
Print #1, "------------------------------------------------"
Print #1, "Dir Scanned = " + ScanPath.TxtPath
Print #1, "------------------------------------------------"
Print #1, "Dir Count #" + Format(Val(ScanPath.lblCount2), "0000000000")
Print #1, "------------------------------------------------"
Print #1, "Files Count #" + Format(Val(ScanPath.lblCount1), "0000000000")
Print #1, "------------------------------------------------"




ScanPath.chkSubFolders = vbChecked
ScanPath.ListView1.ListItems.Clear


'MANUAL
ScanPath.TxtPath = "M:\0 00 Art"
Call ScanPath.CMDScan_NO_LIST_Click

ScanPath.TxtPath = "M:\0 00 Art1"
Call ScanPath.CMDScan_NO_LIST_Click



'Call ScanPath.cmdScan_NO_LIST_Click




Close #1

ScanPath.TxtPath = "M:\0 00 Art"

TG = ScanPath.TxtPath + "\FileList Auto Text -- " + Mid(ScanPath.TxtPath, InStrRev(ScanPath.TxtPath, "\") + 1) + ".txt"
'FR = FreeFile
Open TG For Binary As #1
'Seek #1, 1


Dim TH As String

TH = "------------------------------------------------" + vbCrLf
TH = TH + "Date " + Format(Now, "DDD DD MMM YYYY HH:MM:SS") + vbCrLf
TH = TH + "------------------------------------------------" + vbCrLf
TH = TH + "Dir Scanned = " + ScanPath.TxtPath + vbCrLf
TH = TH + "------------------------------------------------" + vbCrLf
TH = TH + "Dir Count #" + Format(Val(ScanPath.lblCount2), "0000000000") + vbCrLf
TH = TH + "------------------------------------------------" + vbCrLf
TH = TH + "Files Count #" + Format(Val(ScanPath.lblCount1), "0000000000") + vbCrLf
TH = TH + "------------------------------------------------" + vbCrLf

Put #1, 1, TH
Close #1



TG = ScanPath.TxtPath + "\FileList Auto Text -- " + Mid(ScanPath.TxtPath, InStrRev(ScanPath.TxtPath, "\") + 1) + "- TEXT INFO.txt"
Open TG For Output As #1
    Print #1, TH
Close #1




ScanPath.WindowState = 0

MsgBox "WORK ART LIST DONE", vbMsgBoxSetForeground


Exit Sub

Dim WE, A1$, B1$, YY$
'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"


For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    YY$ = A1$ + B1$
    Print #1, YY$
Next

Close #1




End Sub


Sub DelBoy()
Dim WE
'ScanPath.ListView1.ListItems.Clear

'ScanPath.chkSubFolders = vbChecked

'ScanPath.txtPath = "F:\1"


'Call ScanPath.cmdScan_Click
'Call ScanPath.cmdScanDir_Click


ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"


For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    
    
    On Error Resume Next
    Err.Clear
    'Kill A1 + G1
    FSO.DeleteFile A1 + G1
    
    'FsO.deletefolder A1
    
    
    If Err.Number > 0 Then AG = AG + 1
        ScanPath.lblCount3 = Trim(Str(AG)) + " Del Errors"
    Next

ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    

End Sub



Sub DelEMPTYS()

'ScanPath.ListView1.ListItems.Clear

'ScanPath.chkSubFolders = vbChecked

'ScanPath.txtPath = "F:\1"


'Call ScanPath.cmdScan_Click
'Call ScanPath.cmdScanDir_Click

'Call ScanPath.cmdScanDir_Click

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FOLDERS"


For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Folders"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    
    
    On Error Resume Next
    Err.Clear
    
    Set F = FSO.GetFolder(A1 + B1)
    If F.Size = 0 Then
        F.DELETE
    End If
    
    If Err.Number > 0 Then AG = AG + 1
    ScanPath.lblCount3 = Trim(Str(AG)) + " Del Errors"
    
    Next

ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    

End Sub



Sub MOVE_ANY_WHERE()

ScanPath.TxtPath = "M:\0 00 Art\00 My Pictures\z3 75000 Photos\PHOTOS\"
destipath = "M:\0 00 Art\00 My Pictures\z3 75000 Photos\PHOTOS2\"
If FSO.FolderExists(destipath) = False Then MkDir destipath

ScanPath.cboMask.Text = "*.JPG"

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = destipath + Mid(A1, Len(ScanPath.TxtPath) + 1)
    If FSO.FolderExists(D1) = False Then
        Err.Clear
        MkDir D1
        If Err.Number > 0 Then 'Stop
            D2 = D1
            Do
                D2 = Mid(D2, 1, InStrRev(D2, "\", Len(D2) - 1))
                Err.Clear
                MkDir D2
            'If Err.Number > 0 Then Stop
            Loop Until Err.Number = 0 Or D2 < Len(destipath)
            If Err.Number > 0 Then Stop
            
            
            
        End If
    
    End If
    
    FSO.MoveFile A1 + B1, D1 + B1
    
    
    On Error Resume Next
    Err.Clear
    
    If Err.Number > 0 Then AG = AG + 1
        'ScanPath.lblCount3 = Trim(Str(TTD)) + " Del Errors"
    Next

ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    

End Sub




Sub MAKE_PATH_DIFF_PLACE()

'ScanPath.txtPath = ScanPath.txtPath

ScanPath.TxtPath = "M:\00 Loggers Text\01_Drive_Lists"

destipath = ScanPath.txtPath02



If FSO.FolderExists(destipath) = False Then MkDir destipath

ScanPath.cboMask.Text = "*.*"

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScanDir_Click
On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL DIRS"

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process DIRS"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = destipath + Mid(A1, Len(ScanPath.TxtPath) + 1)
    If FSO.FolderExists(D1) = False Then
        Err.Clear
        MkDir D1
        If Err.Number > 0 Then 'Stop
            D2 = D1
            Do
                D2 = Mid(D2, 1, InStrRev(D2, "\", Len(D2) - 1))
                Err.Clear
                MkDir D2
            'If Err.Number > 0 Then Stop
            Loop Until Err.Number = 0 Or D2 < Len(destipath)
            If Err.Number > 0 Then Stop
            
            
            
        End If
    
    End If
    
    FSO.MoveFile A1 + B1, D1 + B1
    
    
    On Error Resume Next
    Err.Clear
    
    If Err.Number > 0 Then AG = AG + 1
        'ScanPath.lblCount3 = Trim(Str(TTD)) + " Del Errors"
    Next

ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    

End Sub




Sub WINRAR_EXTRACT()

'THIS EXTRACT RAR'S IN MULTI PATH FOLDERS



ScanPath.ListView1.ListItems.Clear
'ScanPath.txtPath.Text = "M:\01 Installations\Games Installations\"

ScanPath.cboMask.Text = "*.RAR"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL RARS"

'ScanPath.lblCount3 = "Ones Left ERR EXECUTE"

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " EXTRACT RARS"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    A1 = A1
    
    ChDrive A1
    ChDir A1
    
    'Shell "C:\Program Files\WinRAR\WinRAR.exe E """ + A1 + "\*.rar""", vbMinimizedNoFocus
    'Var = ExecCmd("C:\Program Files\WinRAR\WinRAR.exe E """ + A1 + "\*.rar""")
''''''    Var = ShellAndWait("C:\Program Files\WinRAR\WinRAR.exe E -y """ + A1 + B1 + """")
    'Var = ExecCmd("C:\Program Files\WinRAR\WinRAR.exe E -y """ + A1 + B1 + """")
    
    'If Var <> 0 Then MsgBox "Error This One -- " + vbCrLf + A1 + B1
    
    If VAR <> False And VAR <> True Then
        TTH = TTH + 1: ScanPath.lblCount3 = Trim(Str(TTH)) + " ERR EXECUTE"
    End If
    
    
    Kill A1 + B1
    ScanPath.ListView1.ListItems.Remove (WE)
    WE = WE - 1
    
    

Next

'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " CANT DEL"
ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0


End Sub


Sub CRCDUPE()

'CRC DUPE THAT DONT COUNT HEADER OF MP3 NEEDED



ScanPath.Show

ScanPath.lblCount1 = ""
ScanPath.lblCount2 = ""
ScanPath.lblCount3 = ""
ScanPath.lblCount4 = ""
ScanPath.lblCount5 = ""
ScanPath.lblCount6 = ""


Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32



ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.jpg;*.png;*.gif;*.ico;*.MP3;*.RAR;*.ZIP;*.3GP;*.MPG;*.txt"
ScanPath.cboMask.Text = "*.*"


ScanPath.cboSize.ListIndex = 2 'Bigger Than
ScanPath.cboSize.ListIndex = 0 ' Not Used

'ScanPath.cboSizeType(lCount).ListIndex = 0 'Byte
'ScanPath.cboSizeType(lCount).ListIndex = 1 'K
'ScanPath.cboSizeType(lCount).ListIndex = 2 'Meg
ScanPath.cboSizeType(0).ListIndex = 0

ScanPath.txtSize(0) = 1







If FSO.FolderExists(ScanPath.TxtPath) = False Then Exit Sub

'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count))

'ScanPath.txtPath = "M\0 00 Art Loggers\Temp Inet Files Gif Ico Png\"

'ScanPath.txtPath = "F:\##ACER-D2\"
Call ScanPath.cmdScan_Click

'ScanPath.txtPath = "\\55-88-happy\MY BOOK (M)\##ASUS-D\"
'Call ScanPath.cmdScan_Click

'sort of FILE then PATH ' reverse - most important at front
ScanPath.ListView1.SortOrder = lvwDescending
ScanPath.ListView1.SortKey = 1
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.SortKey = 0
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

'Sort On Size
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
    
wx8 = Val(ScanPath.ListView1.ListItems.Count)
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " Total"

'SORT SAME SIZE FILES AND DEL IF NOT SAME
For WE = ScanPath.ListView1.ListItems.Count To 2 Step -1
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
    ScanPath.lblCount3 = Trim(Str(WE))
    
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    fsi2 = ScanPath.ListView1.ListItems.Item(WE - 1).SubItems(3)
    'Debug.Print Str(WE) + " --- " + FSI1 + " ---- " + fsi2
    If Val(FSI1) <> Val(fsi2) Then
        '## YOU CHOOSE WANT TO DEL
        ScanPath.ListView1.ListItems.Remove (WE)
    End If
    If Val(FSI1) = Val(fsi2) Then
        WE = WE - 1
    End If
    If Len(FSI1) > TINK Then TINK = Len(FSI1)
Next

'PAD ZEROS IN FRONT BASD ON BIZZEST SIZE
For WE = 1 To ScanPath.ListView1.ListItems.Count
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    TY = String(TINK - Len(FSI1), "0") + FSI1
    
    ScanPath.ListView1.ListItems.Item(WE).SubItems(3) = TY
Next

'Exit Sub

    
'Sort On Size - RESORT AFTER PADDED
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

ReDim BLACKL(100)

'DEL ALL BLACK LIST
TS = "M:\04 Music ---\04 My Music\##00 No Dupe Chking List.txt"
If Dir(TS) <> "" Then
    FR1 = FreeFile
    Open TS For Input As #FR1
    BL = 1
    Do
        Line Input #FR1, LL
        BLACKL(BL) = LL
        BL = BL + 1
    Loop Until EOF(FR1)
    Close #FR1
    ReDim Preserve BLACKL(BL - 1)
    'BLACKL(UBOUND(BLACKL))

    For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
        DoEvents
        ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
        ScanPath.lblCount3 = Trim(Str(WE))
    
        G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
        A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(WE)
    
    
        For r = 1 To UBound(BLACKL)
            If InStr(A1, BLACKL(r)) > 0 Then
                ScanPath.ListView1.ListItems.Remove (WE)
            End If
        Next
    Next
End If



'DEL ZERO SIZE FILES AND SMALL
For WE = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
    ScanPath.lblCount3 = Trim(Str(WE))
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    'If Val(FSI1) > 70 Then Stop
    If Val(FSI1) >= 70 Then Exit For
    
    G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    
'    FileExt = Mid$(G1, Len(G1) - 3)
        
    we8 = we8 + 1
    'ScanPath.lblCount5 = Trim(Str(we8)) + " Del as Too Small <70 Bytes"
    
    If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then
        FSO.DeleteFile A1$ + G1$, True
    End If
    
    ScanPath.ListView1.ListItems.Remove (WE)
    WE = WE - 1

Next



'Exit Sub



'## MORE FILTER DEL ALL RAR
'## AND MORE YOU CHOOSE
ScanPath.Show
Reset
YY = "--"
    
'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count))

For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
    ScanPath.lblCount3 = Trim(Str(WE))
    
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    
    If InStr(B1$, ".rar") > 0 Then
        ScanPath.ListView1.ListItems.Remove (WE)
    End If
    
    '## YOU CHOOSE
    If InStr(B1$, "_New") = 0 Then
        'ScanPath.ListView1.ListItems.Remove (WE)
    End If

Next

we4 = ScanPath.ListView1.ListItems.Count
ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " After Same Size"

WE2 = 0

If wx8 > 0 Then ScanPath.lblCount4 = Trim(Str(WE2)) + " Del - Dupes = " + Format((WE2 / wx8) * 100, "0.00") + "%"

WE = 1
WE2 = ScanPath.ListView1.ListItems.Count
On Error Resume Next


'STARTJOB DO IN REVERSE BETTER MOST IMPORT AT FRONT USUAL LIKE MP3'S
Do

'For we = 1 To ScanPath.ListView1.ListItems.Count
    
    WE2 = WE2 - 1
    
    ScanPath.lblCount3 = Trim(Str(WE2))
    
    FSI1 = ScanPath.ListView1.ListItems.Item(1).SubItems(3)
    If FSI1 > 5000 Then ScanPath.lblCount3.Refresh
    If FSI1 > 50000 Then DoEvents
    
    'DoEvents
    
    If WE2 Mod 10 = 0 Then ScanPath.lblCount3.Refresh
    If WE2 Mod 50 = 0 Then DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(1)
    
    
    
    If FSI1 <> OFSI1 Then
        YY = ""
    End If
    OFSI1 = FSI1

    ScanPath.ListView1.ListItems.Remove (WE)


    If InStr(B1$, ".mp3") = 0 Then
        Err.Clear
        WXHEX$ = Hex(m_CRC.CalculateFile(A1$ + G1$))
    
        If InStr(YY, WXHEX$) > 0 And Err.Number = 0 Then
            we3 = we3 + 1
            ScanPath.lblCount4 = Trim(Str(we3)) + " Del - Dupes = " + Format((we3 / we4) * 100, "0.00") + "%"
            If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then FSO.DeleteFile A1$ + G1$
        End If
    Else
        Err.Clear
        'WxHex$ = Hex(m_CRC.CalculateFile(A1$ + G1$))
        Set F = FSO.GetFile(A1 + G1)
        A22 = Space(F.Size - &H410)
        
        fr5 = FreeFile
        Open A1$ + G1$ For Binary As #fr5
             Get #fr5, &H410, A22
        Close #fr5
        'len(a22)+&h410
        
        WXHEX$ = Hex(m_CRC.CalculateString(A22))
        'A22 = ""

        If InStr(YY, WXHEX$) > 0 And Err.Number = 0 Then
            we3 = we3 + 1
            ScanPath.lblCount4 = Trim(Str(we3)) + " Del - Dupes = " + Format((we3 / we4) * 100, "0.00") + "%"
            If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then
                
                
                'FsO.DeleteFile A1$ + G1$
                Call ShellFileDelete(A1$ + G1$, True)
            End If
        End If
    End If
    
    
    YY = YY + " - " + WXHEX$
   
   
    
'Next
Loop Until ScanPath.ListView1.ListItems.Count = 0


ScanPath.lblCount6 = "DONE *"
ScanPath.WindowState = 0




End Sub




Sub CRCDUPE_MULTI_MATRIX_ON_SMALLEST_FILE()

ScanPath.Show

ScanPath.lblCount1 = ""
ScanPath.lblCount2 = ""
ScanPath.lblCount3 = ""
ScanPath.lblCount4 = ""
ScanPath.lblCount5 = ""
ScanPath.lblCount6 = ""


Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32

'ScanPath.txtPath = "D:\# W880i\2009 VIDEO\"


ScanPath.ListView1.ListItems.Clear
'ScanPath.cboMask.Text = "*.jpg;*.png;*.gif;*.ico;*.MP3;*.RAR;*.ZIP;*.3GP;*.MPG"
ScanPath.cboMask.Text = "*.*"
ScanPath.cboSize.ListIndex = 2 'Bigger Than
'ScanPath.cboSizeType(lCount).ListIndex = 0 'Byte
'ScanPath.cboSizeType(lCount).ListIndex = 1 'K
'ScanPath.cboSizeType(lCount).ListIndex = 2 'Meg
ScanPath.cboSizeType(0).ListIndex = 0
ScanPath.txtSize(0) = 1


If FSO.FolderExists(ScanPath.TxtPath) = False Then Exit Sub


'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count))
'ScanPath.txtPath = "M\0 00 Art Loggers\Temp Inet Files Gif Ico Png\"
Call ScanPath.cmdScan_Click


'sort of FILE then PATH ' reverse - most important at front
ScanPath.ListView1.SortOrder = lvwDescending
ScanPath.ListView1.SortKey = 1
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.SortKey = 0
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

'Sort On Size
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
    
wx8 = Val(ScanPath.ListView1.ListItems.Count)
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " Total"

'SORT SAME SIZE FILES AND DEL IF NOT SAME
For WE = ScanPath.ListView1.ListItems.Count To 2 Step -1
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
    ScanPath.lblCount3 = Trim(Str(WE))
    
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    fsi2 = ScanPath.ListView1.ListItems.Item(WE - 1).SubItems(3)
    'Debug.Print Str(WE) + " --- " + FSI1 + " ---- " + fsi2
    If Val(FSI1) <> Val(fsi2) Then
        '## YOU CHOOSE WANT TO DEL
        'ScanPath.ListView1.ListItems.Remove (WE)
    End If
    If Val(FSI1) = Val(fsi2) Then
        WE = WE - 1
    End If
    If Len(FSI1) > TINK Then TINK = Len(FSI1)
Next

'PAD ZEROS IN FRONT BASED ON BIGGEST SIZE
For WE = 1 To ScanPath.ListView1.ListItems.Count
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    TY = String(TINK - Len(FSI1), "0") + FSI1
    
    ScanPath.ListView1.ListItems.Item(WE).SubItems(3) = TY
Next

'Exit Sub

    
'Sort On Size - RESORT AFTER PADDED
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

ReDim BLACKL(100)



'DEL ZERO SIZE FILES AND SMALL
For WE = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
    ScanPath.lblCount3 = Trim(Str(WE))
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    'If Val(FSI1) > 70 Then Stop
    If Val(FSI1) >= 70 Then Exit For
    
    G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    
'    FileExt = Mid$(G1, Len(G1) - 3)
        
    we8 = we8 + 1
    ScanPath.lblCount5 = Trim(Str(we8)) + " Del as Too Small <70 Bytes"
    
    'If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then Kill A1$ + G1$
    
    ScanPath.ListView1.ListItems.Remove (WE)
    WE = WE - 1

Next



'Exit Sub



'## MORE FILTER DEL ALL RAR
'## AND MORE YOU CHOOSE
ScanPath.Show
Reset
YY = "--"
    
'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count))

For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
    ScanPath.lblCount3 = Trim(Str(WE))
    
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    
    If InStr(B1$, ".rar") > 0 Then
        ScanPath.ListView1.ListItems.Remove (WE)
    End If
    
    '## YOU CHOOSE
    If InStr(B1$, "_New") = 0 Then
        ScanPath.ListView1.ListItems.Remove (WE)
    End If

Next

we4 = ScanPath.ListView1.ListItems.Count
ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " After Same Size"

WE2 = 0

If wx8 > 0 Then ScanPath.lblCount4 = Trim(Str(WE2)) + " Del - Dupes = " + Format((WE2 / wx8) * 100, "0.00") + "%"

WE = 1
WE2 = ScanPath.ListView1.ListItems.Count
On Error Resume Next



SMALLSIZE = 1E+14!
'## GET SMALLEST
For WE = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
    ScanPath.lblCount3 = Trim(Str(WE))
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    'If Val(FSI1) > 70 Then Stop
    If Val(FSI1) >= 100 Then
        If Val(FSI1) < SMALLSIZE Then SMALLSIZE = Val(FSI1)
    End If


Next



'CRC ALL FILES BASED ON SMALLEST ONE

For WE = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
    ScanPath.lblCount3 = Trim(Str(WE))
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    
    G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    
    
    
    TTH = ""
    Open A1 + G1 For Binary As #1
    TTH = Input(SMALLSIZE, 1)
    Close #1
    

    

    WXHEX$ = Hex(m_CRC.CalculateString(TTH))


    ScanPath.ListView1.ListItems.Item(WE).SubItems(5) = WXHEX$





Next




'Sort On Size - RESORT AFTER PADDED
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False


'Sort On CRC - KNOW WHAT TO DEL
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 5
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False




'THIS JOB DEL DUPE CRCS
'STARTJOB DO IN REVERSE BETTER MOST IMPORT AT FRONT USUAL LIKE
'DEL JOB FROM LIST WHAT REQUIRED

'Exit Sub


WE2 = ScanPath.ListView1.ListItems.Count + 1

On Error GoTo 0

Do

'For we = 1 To ScanPath.ListView1.ListItems.Count
    
    WE2 = WE2 - 1
    
    ScanPath.lblCount3 = Trim(Str(WE2))
    
    FSI1 = ScanPath.ListView1.ListItems.Item(WE2).SubItems(5)
    'If FSI1 > 5000 Then ScanPath.lblCount3.Refresh
    'If FSI1 > 50000 Then DoEvents
    
    'DoEvents
    
    ScanPath.lblCount3.Refresh
    DoEvents
    
    'If we2 Mod 10 = 0 Then ScanPath.lblCount3.Refresh
    'If we2 Mod 50 = 0 Then DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(WE2).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(WE2).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE2)
    
    If InStr(YY, FSI1) > 0 And YY <> "" Then
        ScanPath.ListView1.ListItems.Remove (WE2)
        ShellFileDelete A1 + B1
    End If
    
    If FSI1 <> OFSI1 Then
        YY = ""
    End If
    OFSI1 = FSI1
    
    YY = YY + " - " + FSI1
    
'Next
Loop Until WE2 <= 1


ScanPath.lblCount6 = "DONE *"
ScanPath.WindowState = 0




End Sub




Public Function ShellFileDelete(src As String, Optional NoConfirm As Boolean = False) As Boolean

lFlags = FOF_ALLOWUNDO
If NoConfirm Then lFlags = lFlags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_DELETE
.pFrom = src
.fFlags = lFlags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileDelete = (lRet = 0)

End Function



