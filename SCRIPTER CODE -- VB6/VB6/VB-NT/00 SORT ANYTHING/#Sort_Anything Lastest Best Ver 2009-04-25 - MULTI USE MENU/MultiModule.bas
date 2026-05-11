Attribute VB_Name = "MULTI_MODULE"
Public BS, VD, VS, DIR_TS, DIR_TD, i
Public L2, XPAUSE
Public R
Public OCNT
Public RCOUNT_CHUNK_STEP, RCOUNT_CHUNK, COPYSET_COUNT, RCOUNT_CHUNK_MARK
Public CNTERR, CNTERRPR
Public DESTIFOLDERPATH, SOURCEFOLDERPATH
Public WE, VC1, VC2, VALR, BLAST
Public LARGESIZECOPY_NOT_ERROR, LARGESIZECOPY_WITH_ERROR
Public ARRAY_DELETE_COPY()
Public ARRAY_DELETE_ALL_EMPTY_FOLDER()


Public FSPD, OA1, LARGESIZE_NOT_ERROR
Dim TX As Boolean
Public TTDTOT
Public AGTOT
Public EXITVAR
Public XX ', mFileCount2
Public TG, DESTIPATH

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
Dim lflags As Long


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

Dim T As Double
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
    T = Now()
    Do Until (Dir$(Shortcut0) <> "")
    
    If (Now() - T) > 0.00006 Then 'wait 5 seconds
        'If MsgBox("Attendre encore la cr‚­ïion du raccourci ?", _
        '    vbQuestion + vbOKCancel, "Raccourci") <> vbOK Then
            'MsgBox "ERROR"
            'Exit Sub
        'Else
            T = Now()
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
          vbCritical, "Shortcut"

End If

End Sub

Private Function p_GetSpecialFolder(CsIdl As Long) As String

'Returns the full path of the folder corresponding to the
'Windows's id system folder.

Dim R     As Long
Dim pidl  As Long
Dim sPath As String

R = api_SHGetSpecialFolderLocation(ScanPath.hWnd, CsIdl, pidl)

If R = 0 Then

    sPath = Space$(260)
    R = api_SHGetPathFromIDList(ByVal pidl, ByVal sPath)
    If R Then
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

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"


XXD = ScanPath.TxtPath + "\LOGG OF CODE SEARCH REPLACE LOGG.TXT"
If Dir(XXD) <> "" Then Kill XXD

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    Iext = UCase(Mid(B1, Len(B1) - 2))

    Debug.Print B1

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

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"
    
B1 = ScanPath.ListView1.ListItems.Item(1)
RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
FN = ScanPath.TxtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
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
        If Fs.FileExists(FN) = True Then WE = WE + 1
        If WE >= ScanPath.ListView1.ListItems.COUNT Then End
    Loop Until Fs.FileExists(FN) = False
    
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
    
    hhg = hhg + "# " + B1

    uu2 = FreeFile
    Open A1 + B1 For Binary As #uu2
        YY = Input(LOF(uu2), uu2)
    Close #uu2
    
    OO = YY
    TX1 = 0
    
    iix = 0
    If InStr(LCase(YY), ScanPath.Text1) > 0 Then
        Do
        TX1 = InStr(TX1 + 1, YY, ScanPath.Text1)
        If TX1 > 0 Then
        
        
            TdD = TdD + 1
            ScanPath.lblCount3 = Trim(Str(TdD)) + " Links Found"
                
                poi = 0
                If InStr(TX1, YY, ScanPath.Text2) > 0 Then JJ = InStr(TX1, YY, ScanPath.Text2) - TX1: pio = 1
                If pio = 0 And InStr(TX1, YY, vbCr) > 0 Then JJ = InStr(TX1, YY, vbCr) - TX1
                
                
                hh = Trim(Mid(YY, TX1, JJ))
                If InStr(hh, """" + vbCr) > 0 Then
                    A = A
                    hh = Mid(hh, 1, InStr(hh, """" + vbCr) - 1)
                End If
                
                hh = Trim(Mid(YY, TX1, JJ))
                If InStr(hh, "<") > 0 Then
                    hh = Mid(hh, 1, InStr(hh, "<") - 1)
                End If
                
                If InStr(hh, vbCr) > 0 Then
                    hh = Mid(hh, 1, InStr(hh, vbCr) - 1)
                End If
                
                iij = 0
                If InStr(hh, "doubleclick") > 0 Then iij = 1
                
                hh = Trim(Mid(YY, TX1, JJ))
                If InStr(hh, """ ") > 0 And iij = 0 Then
                    A = A
                    hh = Mid(hh, 1, InStr(hh, """ ") - 1)
                End If
                
                    
                If iij = 0 Then hh1 = hh1 + hh + vbCrLf: iix2 = 1
            
            End If
        Loop Until TX1 = 0
    
        
        If iix2 = 1 Then
        
            Print #UU, vbCrLf + hhg + vbCrLf
            
            hhg = "# -- " + Trim(Str(TdD)) + " Links Count"
            hhg = hhg + "# -- " + Trim(Str(WE)) + " Email Count"
        
            Print #UU, hhg + vbCrLf
        
            
            Print #UU, hh1
    
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
Call ScanPath.cmdScan_NO_LIST_Click

ScanPath.TxtPath = "M:\0 00 Art1"
Call ScanPath.cmdScan_NO_LIST_Click



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


'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"


For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    YY$ = A1$ + B1$
    Print #1, YY$
Next

Close #1




End Sub


Sub DelBoy()

'ScanPath.ListView1.ListItems.Clear

'ScanPath.chkSubFolders = vbChecked

'ScanPath.txtPath = "F:\1"


'Call ScanPath.cmdScan_Click
'Call ScanPath.cmdScanDir_Click


ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"


For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
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
    Fs.DeleteFile A1 + G1
    
    'Fs.deletefolder A1
    
    
    If Err.Number > 0 Then AG = AG + 1
        ScanPath.lblCount3 = Trim(Str(AG)) + " Del Errors"
    Next

ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    

End Sub



Sub DelEMPTYS()

'ScanPath.chkSubFolders = vbChecked
'Call ScanPath.cmdScan_Click
'Call ScanPath.cmdScanDir_Click

ScanPath.ListView1.ListItems.Clear
ScanPath.TxtPath = "H:\#ACER"
Call ScanPath.cmdScanDir_Click

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FOLDERS"

TTD4 = 0
For WE = ScanPath.ListView1.ListItems.COUNT To 1 Step -1
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD4 = TTD4 + 1
    
    ScanPath.lblCount2 = Trim(Str(TTD4)) + " Process Folders"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    
    On Error Resume Next
    Err.Clear
    
    VART = -10
    Set F = Fs.GetFolder(A1 + B1)
    Set FC = F.Files
    'If FC.COUNT <= 3 Then Stop
    'For Each f1 In FC
    '    f1.Name
    'Next
    'NOT HAVE . AND ..
    VART = FC.COUNT

    If VART = 0 Then
        Set F = Fs.GetFolder(A1 + B1)
        VART = -10
        VART = F.Size
        If VART = 0 Then
            Err.Clear
            Fs.DeleteFOLDER A1 + B1, False
        End If
    End If
    
    If Err.Number > 0 Then AG = AG + 1
    ScanPath.lblCount3 = Trim(Str(AG)) + " Del Errors"
    Next

    ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    
End Sub

Sub PATH_TREE_ANOTHER_PLACE()


ScanPath.TxtPath = "M:\01 INSTALLATIONS\Installation programs\"
DESTIPATHX = "M:\01 INSTALLATIONS\Installation programs\# 00 INSTALL LINK TREE"
DESTIPATH = "M:\01 INSTALLATIONS\Installation programs\# 00 INSTALL LINK TREE\"


'PUT END \
If Fs.FolderExists(DESTIPATH) = False Then MkDir DESTIPATH

ScanPath.cboMask.Text = "*.*"

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked
'Call ScanPath.cmdScan_Click
Call ScanPath.cmdScanDir_Click
On Error Resume Next

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"


For WE = ScanPath.ListView1.ListItems.COUNT To 1 Step -1
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)

    If InStr(DESTIPATHX, A1 + B1) > 0 Then
        Stop
    End If

Next


'DESTIPATHX

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = DESTIPATH + Mid(A1, Len(ScanPath.TxtPath) + 1)
    If Fs.FolderExists(D1) = False Then
        Err.Clear
        MkDir D1
        If Err.Number > 0 Then 'Stop
            d2 = D1
            Do
                d2 = Mid(d2, 1, InStrRev(d2, "\", Len(d2) - 1))
                Err.Clear
                MkDir d2
            'If Err.Number > 0 Then Stop
            Loop Until Err.Number = 0 Or d2 < Len(DESTIPATH)
            If Err.Number > 0 Then Stop
            
            
            
        End If
    
    End If
    
    'Fs.COPYFILE A1 + B1, D1 + B1
    
    
    On Error Resume Next
    Err.Clear
    
    If Err.Number > 0 Then AG = AG + 1
        'ScanPath.lblCount3 = Trim(Str(TTD)) + " Del Errors"
    Next

ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    

End Sub


Sub COPY_CDROM_ANYWHERE()



ScanPath.TxtPath = "G:\"
DESTIPATH = "M:\Videos\00 Vid Films\VOBS\"
'PUT END \
If Fs.FolderExists(DESTIPATH) = False Then MkDir DESTIPATH

ScanPath.cboMask.Text = "*.*"

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = DESTIPATH + Mid(A1, Len(ScanPath.TxtPath) + 1)
    If Fs.FolderExists(D1) = False Then
        Err.Clear
        MkDir D1
        If Err.Number > 0 Then 'Stop
            d2 = D1
            Do
                d2 = Mid(d2, 1, InStrRev(d2, "\", Len(d2) - 1))
                Err.Clear
                MkDir d2
            'If Err.Number > 0 Then Stop
            Loop Until Err.Number = 0 Or d2 < Len(DESTIPATH)
            If Err.Number > 0 Then Stop
            
            
            
        End If
    
    End If
    
    Fs.COPYFILE A1 + B1, D1 + B1
    
    
    On Error Resume Next
    Err.Clear
    
    If Err.Number > 0 Then AG = AG + 1
        'ScanPath.lblCount3 = Trim(Str(TTD)) + " Del Errors"
    Next

ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    

End Sub



Function FREESPDRIVED(DRI) As Double
On Error Resume Next
Set G = Fs.GetDrive(Mid(DRI, 1, 1))
If G.ISREADY = False Then Exit Function

If Err.Number > 0 Then Exit Function

FREESPDRIVED = G.freespace

End Function

Function FREESPDRIVES(DRI) As Double

On Error Resume Next
Set G = Fs.GetDrive(Mid(DRI, 1, 1))
If G.ISREADY = False Then Exit Function

If Err.Number > 0 Then Exit Function

FREESPDRIVES = G.freespace

End Function

Sub CALCTIME()

    wd = DateDiff("s", BLAST, Now)
    fday = Int(wd / 86400)
    f6 = Int(wd / 3600)
    F7 = Int(wd / 60)
    F8 = wd Mod 60
    lab2 = " " + Format$(TimeSerial(0, F7, F8), "HH:MM:SS")
    lab2 = " " + Replace(lab2, " 00:0", "")
    lab2 = " " + Replace(lab2, " 00:", "")
    lab2 = " " + Replace(lab2, " 0", "")
    lab2 = " " + Replace(lab2, " :", "")
    lab2 = Trim(Str(fday) + " Day " + lab2)
    
    VC5 = VC2
    If VC1 > VC5 Then VC5 = VC1
    VALR = " + TIME DURATION = " + lab2
    VALR = Replace(VALR, "  ", " ")
    VALR = Replace(VALR, "  ", " ")
    VALR = Replace(VALR, "  ", " ")
End Sub


Sub COPY_ANYWHERE_PROPER()

Dim VC, D1 As String


If ARRAY_DELETE_ALL_EMPTY_FOLDER(COPYSET_COUNT) = True Then
    ScanPath.lblCount6 = "** WORKING FILE SYS WORK -- KILL EMPTY ALL FOLDER ***"
    Call DelEMPTYS
End If

'Loop Until RCOUNT_CHUNK_MARK < 0
RCOUNT_CHUNK_STEP = 20000

Do

    On Error Resume Next
    Err.Clear
    If Fs.FolderExists(DESTIPATH) = False Then MkDir DESTIPATH
    
    If Err.Number > 0 Then
    '    MsgBox Err.Description + vbCrLf + " DESTI DRIVE PATH", vbCritical Or vbMsgBoxSetForeground
        Exit Sub
    End If
    
    
    ASIZEMAX = 0
    ScanPath.lblCount6 = "** WORKING FILE SYS WORK -- SCAN ***"
    
    ScanPath.chkCopyMemory.Value = vbChecked
    
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files / TOT = " + Trim(Str(TTDTOT))
    ScanPath.lblCount3 = Trim(Str(AGTOT)) + " TOT COPY Errors"
    
    
    'ScanPath.TxtPath = "G:\"
    'destipath = "M:\Videos\00 Vid Films\VOBS\"
    'PUT END \
    Dim FSPD As Double
    DESTIPATH = ScanPath.txtPath02
        
    DESTIFOLDERPATH = DESTIPATH
    SOURCEFOLDERPATH = ScanPath.TxtPath
    Debug.Print SOURCEFOLDERPATH
    Debug.Print DESTIFOLDERPATH
    FSPD = FREESPDRIVED(DESTIPATH)
    
    'NOT DRIVE SPACE END
    If FSPD < 100 * 1024 ^ 2 Then Exit Sub
    
    
    
    ScanPath.cboMask.Text = "*.*"
    ScanPath.chkSubFolders = vbChecked
    
    'If DUPESOUCRE = False Then
'        ScanPath.ListView1.ListItems.Clear
'        Call ScanPath.cmdScan_Click
    'End If
            
    
    ScanPath.ListView1.ListItems.Clear
    Call ScanPath.cmdScan_Click


    If ScanPath.ListView1.ListItems.COUNT = 0 Then EXIT_VAR = True: Exit Do

    
    'For WE = ScanPath.ListView1.ListItems.COUNT To 1 Step -1
    '    B1 = ScanPath.ListView1.ListItems.Item(WE)
    '    If B1 = "" Then
    '       'ScanPath.ListView1.ListItems.Remove (WE)
    '    End If
    'Next
    
    ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"
    DoEvents
    FR5 = FreeFile
    Open "D:\TEMP\LOGG MULTI COPY_SMALL.TXT" For Append As #FR5
        Print #FR5, "--BEGIN -------------------------------------------------------------------------"
        Print #FR5, Format(Now, "DDD DD-MM-YY HH-MM:SSA/P")
        Print #FR5, "count ##-- " + ScanPath.lblCount8
        Print #FR5, "--------------------------------------------------------------------------------"
    Close #FR5
    
    FR4 = FreeFile
    Open "D:\TEMP\LOGG MULTI COPY.TXT" For Append As #FR4
    Print #FR4, "--------------------------------------------------------------------------------"
    Print #FR4, "#### COPY ----";
    Print #FR4, ScanPath.txtPath02
    Print #FR4, "To"
    Print #FR4, ScanPath.TxtPath
    
    Print #FR4, "--------------------------------------------------------------------------------"
    Print #FR4, "TO DO COUNT ="; Str(ScanPath.ListView1.ListItems.COUNT)
    Print #FR4, "--------------------------------------------------------------------------------"
    Print #FR4, Format(Now, "DDD DD-MM-YY HH-MM:SSA/P")
    Print #FR4, "--------------------------------------------------------------------------------"
    
    
    
    OUTD1 = 0
    OUTD2 = 0
    OUTD3 = 0
    OUTD4 = 0
    AG = 0
    LARGESIZE = 0
    

    
    '------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    ScanPath.lblCount6 = "** WORKING FILE SYS WORK REMOVE ***"
    '------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    
    VC1 = ScanPath.ListView1.ListItems.COUNT
    VC2 = ScanPath.ListView1.ListItems.COUNT
    
    For WE = ScanPath.ListView1.ListItems.COUNT To 0 Step -1
        
        If EXITVAR = True Then Exit For

        
        A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
        B1 = ScanPath.ListView1.ListItems.Item(WE)
        size1 = Val(ScanPath.ListView1.ListItems.Item(WE).SubItems(3))
        
        G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
        
        If WE Mod 50 = 0 Then
            DoEvents
            
            ScanPath.lblCount7 = A1 + B1
            ScanPath.lblCount1 = Trim(Str(VC1)) + " TOTAL FILES + TO DO AFTER REMOVE -- " + Trim(Str(VC2))
'            ScanPath.PBar1.Value = 100 - ((WE / VC1) * 100)
            
'            Call CALCTIME
'            VALR = Format(100 - (WE / VC2) * 100, "0.00") + "%" + VALR
'            ScanPath.lblCount4 = VALR
            ScanPath.ListView1.ListItems(WE).EnsureVisible
            
            Err.Clear
        End If
            
        Call CALCTIME
        VALR = Format(100 - (WE / VC2) * 100, "0.0000") + "%" + VALR
        ScanPath.lblCount4 = VALR
        ScanPath.PBar1.Value = 100 - ((WE / VC1) * 100)
        DoEvents
        If OA1 <> A1 Then
            OA1 = A1
            
            'MKDIR TAKEN OUT BUT REMOTE PATH STILL WANT
            On Error Resume Next
            D1 = DESTIPATH + Mid(A1, Len(ScanPath.TxtPath) + 1)
            If Fs.FolderExists(D1) = False Then
                Err.Clear
                'MkDir D1
                If Err.Number > 0 Then 'Stop
                    d2 = D1
                    Do
                        d2 = Mid(d2, 1, InStrRev(d2, "\", Len(d2) - 1))
                        Err.Clear
                        'If d2 <> "" Then MkDir d2
                        'If d2 = "" Then Stop
                    'If Err.Number > 0 Then Stop
                    Loop Until Err.Number = 0 Or d2 < Len(DESTIPATH)
                    If Err.Number > 0 Then Stop
                End If
            End If
        End If
        
        TAGDEL = False
        
        FLAG_TO_GO = False
        If B1 <> "" Then
            GO8 = 0: GO7 = 0
            If Fs.FileExists(D1 + B1) = True Then
                GO7 = 1
                Set F = Fs.getfile(A1 + B1)
                Set D = Fs.getfile(D1 + B1)
                If (F.Size < D.Size) And _
                    (F.datelastmodified = D.datelastmodified) Then GO8 = True
            End If
            
            'DEL'S LIKE "meta.fm"
            
            If InStr(A1, "RECYCLER") > 0 Then B1 = ""
            If InStr(A1, "System Volume Information") > 0 Then B1 = ""
            If InStr(A1, "Recycled") > 0 Then B1 = ""
            
            'FILE MAY EXIST + SIZE DONT MATCH
            If GO8 = True Or B1 = "" Then
                ScanPath.ListView1.ListItems.Remove (WE)
                FLAG_TO_GO = True
                VC2 = ScanPath.ListView1.ListItems.COUNT
                'ARRAY SET TO DEL AFTER COPY
                    
                If ARRAY_DELETE_COPY(COPYSET_COUNT) = True And B1 <> "" Then
                    Fs.DeleteFile (A1 + B1), True
                    A = A
                End If
            End If
        End If
        
        If FLAG_TO_GO And False And B1 = "" Then
            ScanPath.ListView1.ListItems.Remove (WE)
            VC2 = ScanPath.ListView1.ListItems.COUNT
        End If
    
    'If WE > ScanPath.ListView1.ListItems.COUNT Then Exit For
    Debug.Print WE
    Next
    
    
    
    
    
    VC1 = ScanPath.ListView1.ListItems.COUNT
    ScanPath.lblCount1 = Trim(Str(VC1)) + " TOTAL FILES"
    
    
    '------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    ScanPath.lblCount6 = "** WORKING FILE SYS - GO - WORK COPY ***"
    '------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------
    
    
 
    
    VC = ScanPath.ListView1.ListItems.COUNT
    VC1 = ScanPath.ListView1.ListItems.COUNT
    VC2 = ScanPath.ListView1.ListItems.COUNT
    WE = 0
    
    
    
    For WE = 1 To ScanPath.ListView1.ListItems.COUNT
        
        If EXITVAR = True Then Exit For
        'DONT BEGIN COPY WHEN EXIT VAR FROM REMOVE
        
        'WE = WE + 1
        
        
        A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
        B1 = ScanPath.ListView1.ListItems.Item(WE)
        size1 = Val(ScanPath.ListView1.ListItems.Item(WE).SubItems(3))
        
        G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
        
        
        'If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
        
        If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
        'If WE > VC Then Exit Do
        
        
        
        Call CALCTIME
        VALR = Format((WE / VC2) * 100, "0.00") + "%" + VALR
        ScanPath.lblCount4 = VALR
        
        
        ScanPath.PBar1.Value = (WE / VC2) * 100
        
        TTD = TTD + 1
        TTDTOT = TTDTOT + 1
        ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files / TOT = " + Trim(Str(TTDTOT))
        DoEvents
        
        'If EXITVAR = True Then Exit For
        'If EXITVAR = True Then Exit Do
        
        'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    
        'If WE > VC Then Exit For
        
        G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
        
        ScanPath.ListView1.ListItems(WE).EnsureVisible
        DoEvents
        ScanPath.lblCount7 = A1 + B1
        
        On Error Resume Next
        D1 = DESTIPATH + Mid(A1, Len(ScanPath.TxtPath) + 1)
        If OA1 <> A1 Then
            OA1 = A1
            If Fs.FolderExists(D1) = False Then
                Err.Clear
                MkDir D1
                If Err.Number > 0 Then 'Stop
                    d2 = D1
                    Do
                        d2 = Mid(d2, 1, InStrRev(d2, "\", Len(d2) - 1))
                        Err.Clear
                        If d2 <> "" Then MkDir d2
                        'If d2 = "" Then Stop
                    'If Err.Number > 0 Then Stop
                    Loop Until Err.Number = 0 Or d2 < Len(DESTIPATH)
                If Err.Number > 0 Then Stop
                End If
            End If
        End If
        
        DESTIFOLDERPATH = D1
        SOURCEFOLDERPATH = A1
        
        Set F = Fs.getfile(A1 + G1)
        LARGESIZE = LARGESIZE + F.Size
        FILESIZE = F.Size
        
                    If FSPD <= size1 + 1024 ^ 2 Then Exit For
        
        'NO EXISTS DESTI
    
        Err.Clear
        DoEvents
                    
        TX = True
        SETERROR = False
    
        On Error Resume Next
        'T5 = ShellFileCopy(A1 + G1, D1 + B1, TX)
        'T5 = ShellFileCopy(A1 + G1, D1, TX)
        Err.Clear
        Set F = Fs.getfile(A1 + G1)
        F.Copy D1, True
          
        'If Fs.FileExists(D1 + B1) = False Then Stop
        'If Fs.FileExists(D1 + G1) = False Then Stop
            
        'LIMIT WILL NOT CHECK FOR REMOTE ALTERNATIVE NAME
        If Fs.FileExists(D1 + B1) = False Then
            AG = AG + 1: AGTOT = AGTOT + 1
        End If
        
        'ARRAY SET TO KILL AFTER COPY
        'KILL FROM SOURCE
        If ARRAY_DELETE_COPY(COPYSET_COUNT) = True _
            And Fs.FileExists(D1 + B1) = True Then
            Fs.DeleteFile (A1 + B1), True
        End If
        
        SETERROR = False
        If Err.Number > 0 Then
            SETERROR = True
            CNTERR = CNTERR + 1
            CNTERRPR = CNTERRPR + 1
        End If
    
        If SETERROR = False Then
            LARGESIZECOPY_NOT_ERROR = LARGESIZECOPY_NOT_ERROR + FILESIZE
        Else
            LARGESIZECOPY_WITH_ERROR = LARGESIZECOPY_WITH_ERROR + FILESIZE
        End If
        
        LTX = Format(LARGESIZECOPY_NOT_ERROR / (1024 ^ 2), "###,###,##0.00") + " MB - EF COPY "
        LTX = LTX + "- " + Format(LARGESIZE / (1024 ^ 2), "###,###,##0.00") + " MB - WE "
        LTX = Replace(LTX, "- ,", "")
    
        
        If SETERROR = False Then
            AG = AG + 1
        Else
            AGTOT = AGTOT + 1
        End If
        
        ScanPath.lblCount3 = Trim(Str(AG)) + " TOT PER SESSION -- " + Trim(Str(AGTOT)) + " TOT COPY Errors"
    
        If Err.Number = 0 Then
            'ScanPath.ListView1.ListItems.Remove (1) ': ScanPath.ListView1.Refresh
            DoEvents
        End If
        
        DoEvents
        
        Print #FR4, "SOURCE ##-- " + A1 + B1
        Print #FR4, "DESTI      ##-- " + D1 + B1; "    ----    ";
        
        If GO8 = 0 And Err.Number = 0 Then OUTR = "PROPER ": OUTD1 = OUTD1 + 1
        If GO8 = 0 And Err.Number > 0 Then OUTR = "FAIL ": OUTD2 = OUTD2 + 1
        If GO8 <> 0 Then OUTR = "EXISTS ": OUTD3 = OUTD3 + 1
        If GO8 <> 0 And GO7 <> 0 Then OUTR = "UPDATE DESTI ": OUTD4 = OUTD4 + 1
        
        'MOD ONLY IF ERROR
        If GO8 = 0 And Err.Number > 0 Then
            Print #FR4, OUTR + " -- " + Format(Now, "HH:MM:SS")
        End If
    
        FSPD = FREESPDRIVES(DESTIPATH)
        FREESPDRD = Format$(FSPD / 1024 ^ 3, "0.0000")
        
        'FSPS =
        FREESPDRS = Format$(FREESPDRIVED(SOURCEFOLDERPATH) / 1024 ^ 3, "0.0000")
    
        ScanPath.lblCount10 = LTX + " - Free GB D- " + FREESPDRD + " GB --- S- " + FREESPDRS + " GB"
        
        
        If EXITVAR = True Then Exit For
    Next
    
    
    
    
    FR5 = FreeFile
    Open "D:\TEMP\LOGG MULTI COPY_SMALL.TXT" For Append As #FR5
    
        Print #FR5, "--------------------------------------------------------------------------------"
        Print #FR5, Format(Now, "DDD DD-MM-YY HH-MM:SSA/P")
        Print #FR5, "count ##-- " + lblCount8
        
        Print #FR5, "count ##-- " + lblCount8
        
        Print #FR5, ScanPath.txtPath02
        Print #FR5, "To"
        Print #FR5, ScanPath.TxtPath
        
        Print #FR5, lblCount8 '  DRIVE NAME
        Print #FR5, ScanPath.lblCount4 '; Time; DURATION
        Print #FR5, "--------------------------------------------------------------------------------"
        
        On Error Resume Next
        
        If FSPD <= size1 Then
           Print #FR4, "--------------------------------------------------------------------------------"
           Print #FR4, "END DUE TO DRIVE SPACE"
           Print #FR4, "--------------------------------------------------------------------------------"
        End If
        
        Print #FR4, "--------------------------------------------------------------------------------"
        Print #FR4, "RESULT COUNT COPY"
        Print #FR4, "--------------------------------------------------------------------------------"
        Print #FR4, "PROPER " + OUTD1
        Print #FR4, "FAIL " + OUTD2
        Print #FR4, "EXISTS " + OUTD3
        Print #FR4, "UPDATE DESTI" + OUTD4
        Print #FR4, "--------------------------------------------------------------------------------"
        Print #FR4, "COPY BYTES " + LARGESIZE
        
        Print #FR4, "--------------------------------------------------------------------------------"
        
        EF = "EXIT END PROPER"
        If EXITVAR = True Then EF = "ABORTED"
        
        Print #FR4, "DONE #### EXIT RESULT = " + EF
        Print #FR4, "--------------------------------------------------------------------------------"
        Print #FR4, Format(Now, "DDD DD-MM-YY HH-MM:SSA/P")
        Print #FR4, "--------------------------------------------------------------------------------"
    Close #FR4, #FR5
    
    If EXITVAR = True Then Exit Do

Loop Until RCOUNT_CHUNK_MARK < 0

'Public RCOUNT_CHUNK_STEP, RCOUNT_CHUNK


ScanPath.lblCount6 = "** DONE **"

End Sub

Sub MOVE_ANYWHERE_PROPER()


ScanPath.lblCount6 = "** WAIT FILE SYS ***"
ScanPath.lblCount2 = "0--- Process Files"
ScanPath.lblCount3 = "0--- MOVE Errors"
ScanPath.lblCount4 = "0--- MOVE DEL Errors"

'ScanPath.TxtPath = "G:\"
'destipath = "M:\Videos\00 Vid Films\VOBS\"
'PUT END \
DESTIPATH = ScanPath.txtPath02
    
    DESTIFOLDERPATH = DESTIPATH
    SOURCEFOLDERPATH = ScanPath.TxtPath

If Fs.FolderExists(DESTIPATH) = False Then MkDir DESTIPATH

ScanPath.cboMask.Text = "*.*"

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
On Error Resume Next

ScanPath.lblCount6 = "** WORKING FILE SYS ***"

For WE = ScanPath.ListView1.ListItems.COUNT To 1 Step -1
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    If B1 = "" Then
        ScanPath.ListView1.ListItems.Remove (WE)
    End If
Next



ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"

FR4 = FreeFile
Open "D:\TEMP\LOGG MULTI COPY.TXT" For Append As #FR4

Print #FR4, "--------------------------------------------------------------------------------"

Print #FR4, "#### MOVE ----";
Print #FR4, ScanPath.txtPath02
Print #FR4, "To"
Print #FR4, ScanPath.TxtPath
Print #FR4, "--------------------------------------------------------------------------------"


Print #FR4, "TO DO COUNT ="; Str(ScanPath.ListView1.ListItems.COUNT)
Print #FR4, "--------------------------------------------------------------------------------"
Print #FR4, Format(Now, "DDD DD-MM-YY HH-MM:SSA/P")
Print #FR4, "--------------------------------------------------------------------------------"


OUTD1 = 0
OUTD2 = 0
OUTD3 = 0
OUTD4 = 0
AG = 0
AGD = 0
LARGESIZE = 0

ScanPath.lblCount6 = "** WORKING FILE SYS ***"
For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    If EXITVAR = True Then Exit For
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    ScanPath.ListView1.ListItems(WE).EnsureVisible
    ScanPath.lblCount7 = A1 + B1
    
    D1 = DESTIPATH + Mid(A1, Len(ScanPath.TxtPath) + 1)
    
    DESTIFOLDERPATH = D1
    SOURCEFOLDERPATH = A1
    
    If Fs.FolderExists(D1) = False Then
        Err.Clear
        MkDir D1
        If Err.Number > 0 Then 'Stop
            d2 = D1
            Do
                d2 = Mid(d2, 1, InStrRev(d2, "\", Len(d2) - 1))
                Err.Clear
                MkDir d2
            'If Err.Number > 0 Then Stop
            Loop Until Err.Number = 0 Or d2 < Len(DESTIPATH)
            If Err.Number > 0 Then Stop
        End If
    End If
    
    
    Err.Clear
    DoEvents
    'Fs.COPYFILE A1 + B1, D1 + B1
    
    GO8 = 0: GO7 = 0
    If Fs.FileExists(D1 + B1) = True Then
        GO7 = 1
        Set F = Fs.getfile(A1 + B1)
        Set D = Fs.getfile(D1 + B1)
        If (F.Size = D.Size) And _
            (F.datelastmodified = D.datelastmodified) Then GO8 = 1
    End If
    
    'EXISTS UPDATE ########
    '------------------
    'DEL ON MOVE
    
    LARGESIZE = LARGESIZE + F.Size
    
    If F.Size > 50000 Then
        TX = False
        Else: TX = True
    End If
    
    '### TOP DONT DEL IF ERR
    
    'EXISTS DESTI UPDATE DEL
    'NO EXISTS DESTI MOVE
    If GO8 = 0 Then
        T5 = ShellFileMove(A1 + B1, D1 + B1, TX)
        If Err.Number > 0 Or T5 = False Then AG = AG + 1
        ScanPath.lblCount3 = Trim(Str(AG)) + " MOVE Errors"
        ScanPath.lblCount4 = Trim(Str(AG)) + " MOVE DEL Errors"
    End If
    
    'EXISTS DESTI DUPE DEL
    If GO8 <> 0 And Err.Number = 0 Then
        Err.Clear
        Fs.DeleteFile (D1 + B1)
        If Err.Number > 0 Then AGD = AGD + 1
        ScanPath.lblCount3 = Trim(Str(AG)) + " MOVE Errors"
        ScanPath.lblCount4 = Trim(Str(AGD)) + " MOVE DEL Errors"
    End If
    
    
    
    DoEvents
    
    Print #FR4, "SOURCE ##-- " + A1 + B1
    Print #FR4, "DESTI      ##-- " + D1 + B1; "    ----    ";
    
    If GO8 = 0 And Err.Number = 0 Then OUTR = "PROPER ": OUTD1 = OUTD1 + 1
    If GO8 = 0 And Err.Number > 0 Then OUTR = "FAIL ": OUTD2 = OUTD2 + 1
    If GO8 <> 0 Then OUTR = "EXISTS ": OUTD3 = OUTD3 + 1
    If GO8 <> 0 And GO7 <> 0 Then OUTR = "UPDATE DESTI ": OUTD4 = OUTD4 + 1
    
    Print #FR4, OUTR + " -- " + Format(Now, "HH:MM:SS")
    
    
    
    
Next

Print #FR4, "--------------------------------------------------------------------------------"
Print #FR4, "RESULT COUNT MOVE"
Print #FR4, "--------------------------------------------------------------------------------"
Print #FR4, "PROPER " + OUTD1
Print #FR4, "FAIL " + OUTD2
Print #FR4, "EXISTS " + OUTD3
Print #FR4, "UPDATE DESTI" + OUTD4
Print #FR4, "--------------------------------------------------------------------------------"
Print #FR4, "COPY BYTES " + LARGESIZE

Print #FR4, "--------------------------------------------------------------------------------"
EF = "EXIT END PROPER"
If EXITVAR = True Then EF = "ABORTED"
Print #FR4, "DONE #### EXIT RESULT = " + EF
Print #FR4, "--------------------------------------------------------------------------------"
Print #FR4, Format(Now, "DDD DD-MM-YY HH-MM:SSA/P")
Print #FR4, "--------------------------------------------------------------------------------"

Close #FR4


ScanPath.lblCount6 = "** DONE **"
'ScanPath.WindowState = 0
    
End Sub


Sub COPY_ANYWHERE_ONE_FOLDER()

'ScanPath.TxtPath = "G:\"
'destipath = "M:\Videos\00 Vid Films\VOBS\"
'PUT END \


ScanPath.TxtPath = "M:\00 HTTrack\00 ICONS"
ScanPath.txtPath02 = "D:\VB6 Icons\HTTRACK"
'DESTIPATH =
DoEvents

If Fs.FolderExists(DESTIPATH) = False Then MkDir DESTIPATH

ScanPath.cboMask.Text = "*.*"
ScanPath.cboMask.Text = "*.JPG;*.PNG;*.GIF;*.ICO"

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"

ScanPath.lblCount6 = "** WORKING FILE SYS ***"

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
 
    
    Fs.COPYFILE A1 + B1, DESTIPATH + "\" + B1
    
    
    On Error Resume Next
    Err.Clear
    
    If Err.Number > 0 Then AG = AG + 1
        'ScanPath.lblCount3 = Trim(Str(TTD)) + " Del Errors"
    Next

ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    

End Sub




Sub MOVE_ANY_WHERE()

ScanPath.TxtPath = "M:\0 00 Art\00 My Pictures\z3 75000 Photos\PHOTOS\"
DESTIPATH = "M:\0 00 Art\00 My Pictures\z3 75000 Photos\PHOTOS2\"
If Fs.FolderExists(DESTIPATH) = False Then MkDir DESTIPATH

ScanPath.cboMask.Text = "*.JPG"

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = DESTIPATH + Mid(A1, Len(ScanPath.TxtPath) + 1)
    If Fs.FolderExists(D1) = False Then
        Err.Clear
        MkDir D1
        If Err.Number > 0 Then 'Stop
            d2 = D1
            Do
                d2 = Mid(d2, 1, InStrRev(d2, "\", Len(d2) - 1))
                Err.Clear
                MkDir d2
            'If Err.Number > 0 Then Stop
            Loop Until Err.Number = 0 Or d2 < Len(DESTIPATH)
            If Err.Number > 0 Then Stop
            
            
            
        End If
    
    End If
    
    Fs.MOVEFILE A1 + B1, D1 + B1
    
    
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

DESTIPATH = ScanPath.txtPath02



If Fs.FolderExists(DESTIPATH) = False Then MkDir DESTIPATH

ScanPath.cboMask.Text = "*.*"

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScanDir_Click
On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL DIRS"

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process DIRS"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = DESTIPATH + Mid(A1, Len(ScanPath.TxtPath) + 1)
    If Fs.FolderExists(D1) = False Then
        Err.Clear
        MkDir D1
        If Err.Number > 0 Then 'Stop
            d2 = D1
            Do
                d2 = Mid(d2, 1, InStrRev(d2, "\", Len(d2) - 1))
                Err.Clear
                MkDir d2
            'If Err.Number > 0 Then Stop
            Loop Until Err.Number = 0 Or d2 < Len(DESTIPATH)
            If Err.Number > 0 Then Stop
            
            
            
        End If
    
    End If
    
    Fs.MOVEFILE A1 + B1, D1 + B1
    
    
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

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL RARS"

'ScanPath.lblCount3 = "Ones Left ERR EXECUTE"

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
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
    Var = ShellAndWait("C:\Program Files\WinRAR\WinRAR.exe E -y """ + A1 + B1 + """")
    'Var = ExecCmd("C:\Program Files\WinRAR\WinRAR.exe E -y """ + A1 + B1 + """")
    
    'If Var <> 0 Then MsgBox "Error This One -- " + vbCrLf + A1 + B1
    
    If Var <> False And Var <> True Then
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



Sub EXECUTE_ALL_LINK_FILE_FOLDER()

'THIS EXTRACT RAR'S IN MULTI PATH FOLDERS

ScanPath.TxtPath = "I:\AGENT 2011-04-21\NOW"

ScanPath.ListView1.ListItems.Clear
'ScanPath.txtPath.Text = "M:\01 Installations\Games Installations\"

ScanPath.cboMask.Text = "*.URL"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL RARS"

'ScanPath.lblCount3 = "Ones Left ERR EXECUTE"

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " EXECUTE LINK FILE FOLDER"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
'    A1 = A1
    
'    ChDrive A1
'    ChDir A1
    
    'Shell "C:\Program Files\WinRAR\WinRAR.exe E """ + A1 + "\*.rar""", vbMinimizedNoFocus
    'Var = ExecCmd("C:\Program Files\WinRAR\WinRAR.exe E """ + A1 + "\*.rar""")
    TT = 1
    If B1 = "" Then TT = 0
    'If InStr(A1, "5+") Then TT = 0
    'If InStr(LCase(A1), LCase("SIXTH")) Then TT = 0
    If InStr(A1, "Hove Lets") Then TT = 0
    If InStr(A1, "# DATE") Then TT = 0
    If InStr(A1, "# NOT RESPOND") Then TT = 0
    If InStr(A1, "# NOT ON MARKET") Then TT = 0
    
    
    
    If TT = 1 Then
        Var = ShellAndWait("C:\Program Files\Mozilla Firefox\firefox.exe """ + A1 + B1 + """")
    End If
    'Var = ExecCmd("C:\Program Files\WinRAR\WinRAR.exe E -y """ + A1 + B1 + """")
    
    'If Var <> 0 Then MsgBox "Error This One -- " + vbCrLf + A1 + B1
    
    If Var <> False And Var <> True Then
        TTH = TTH + 1: ScanPath.lblCount3 = Trim(Str(TTH)) + " ERR EXECUTE"
    End If
    
    
'    Kill A1 + B1
'    ScanPath.ListView1.ListItems.Remove (WE)
'    WE = WE - 1
    
    

Next
End
'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " CANT DEL"
ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0


End Sub



Sub CRCDUPE_X()
If Command$ = "BOOT" Then ScanPath.WindowState = 1
'CRC DUPE THAT DONT COUNT HEADER OF MP3 NEEDED
'SEARCH_CRC_DUPE_###############################
'J:\Kat MP3 Recorder
ScanPath.chkRefreshListView.Value = vbChecked

'ScanPath.TxtPath = "J:\Kat MP3 Recorder"
'ScanPath.TxtPath = "J:\Kat MP3 Recorder"

For RSX = 1 To 1

'If RSX = 1 Then ScanPath.TxtPath = "M:\#0-D-DR"
'If RSX = 2 Then ScanPath.TxtPath = "M:\#0-C-DR"
'If RSX = 3 Then ScanPath.TxtPath = "K:\#ACER"
'If RSX = 4 Then ScanPath.TxtPath = "H:\#ACER"

'''
''''ONE JOB WORK
''''--<><><>
''''-----------------<><><>
'''VS = "My_Book_SOUCE"
'''VD = "NAS"
'''DIR_TS = ":\#ACER"
''''-----------------<><><>
'''Call DIR_PARAMETERS_FOR_ARRAY
'''ARRAY_DELETE_COPY(i) = True
''''ARRAY_DELETE_ALL_EMPTY_FOLDER(i) = True
''''-----------------<><><>
''''--<><><>
'''



ScanPath.Show
DoEvents

ScanPath.chkSubFolders = vbChecked
'Chk_SortOnDate.Visible = False
ScanPath.chkCopyMemory.Value = vbChecked
'lblSize(1).Visible = False
ScanPath.chkSubFolders.Value = vbChecked
'chkPatternMatching.Visible = False

ScanPath.lblCount1 = ""
ScanPath.lblCount2 = ""
ScanPath.lblCount3 = ""
ScanPath.lblCount4 = ""
ScanPath.lblCount5 = ""
ScanPath.lblCount6 = ""


Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32



ScanPath.ListView1.ListItems.Clear
'ScanPath.cboMask.Text = "*.jpg;*.png;*.gif;*.ico;*.MP3;*.RAR;*.ZIP;*.3GP;*.MPG;*.txt"
ScanPath.cboMask.Text = "*.jpg;*.png;*.gif;*.ico;*.MP3;*.RAR;*.ZIP;*.3GP;*.MPG;*.WMA;*.AVI;*.MP4"


ScanPath.cboSize.ListIndex = 2 'Bigger Than
ScanPath.cboSize.ListIndex = 0 ' Not Used

'ScanPath.cboSizeType(lCount).ListIndex = 0 'Byte
'ScanPath.cboSizeType(lCount).ListIndex = 1 'K
'ScanPath.cboSizeType(lCount).ListIndex = 2 'Meg
ScanPath.cboSizeType(0).ListIndex = 0

ScanPath.txtSize(0) = 1







If Fs.FolderExists(ScanPath.TxtPath) = False Then Exit Sub

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
    
wx8 = Val(ScanPath.ListView1.ListItems.COUNT)
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " Total"

'SORT SAME SIZE FILES AND DEL IF NOT SAME
For WE = ScanPath.ListView1.ListItems.COUNT To 2 Step -1
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT))
    ScanPath.lblCount3 = Trim(Str(WE))
    
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    fsi2 = ScanPath.ListView1.ListItems.Item(WE - 1).SubItems(3)
        Do While XPAUSE = True
        DoEvents
        Sleep 500
        Loop
        
    
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
For WE = 1 To ScanPath.ListView1.ListItems.COUNT
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

''DEL ALL BLACK LIST
'TS = "M:\04 Music ---\04 My Music\##00 No Dupe Chking List.txt"
'If Dir(TS) <> "" Then
'    FR1 = FreeFile
'    Open TS For Input As #FR1
'    BL = 1
'    Do
'        Line Input #FR1, LL
'        BLACKL(BL) = LL
'        BL = BL + 1
'    Loop Until EOF(FR1)
'    Close #FR1
'    ReDim Preserve BLACKL(BL - 1)
'    'BLACKL(UBOUND(BLACKL))
'
'    For WE = ScanPath.ListView1.ListItems.COUNT To 1 Step -1
'        DoEvents
'        ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT))
'        ScanPath.lblCount3 = Trim(Str(WE))
'
'        G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
'        A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
'        B1$ = ScanPath.ListView1.ListItems.Item(WE)
'
'
'        For R = 1 To UBound(BLACKL)
'            If InStr(A1, BLACKL(R)) > 0 Then
'                ScanPath.ListView1.ListItems.Remove (WE)
'            End If
'        Next
'    Next
'End If
'


'BEGIN DELETE

'DEL ZERO SIZE FILES AND SMALL
'For WE = 1 To ScanPath.ListView1.ListItems.COUNT
'    DoEvents
'    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT))
'    ScanPath.lblCount3 = Trim(Str(WE))
'    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
'    'If Val(FSI1) > 70 Then Stop
'    If Val(FSI1) >= 70 Then Exit For
'    'FSI1 = ""
'    If IsNumeric(FSI1) = False Then MsgBox "NOT A VALUE FOR SIZE FILE"
'    G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
'    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
'    B1$ = ScanPath.ListView1.ListItems.Item(WE)
'
''    FileExt = Mid$(G1, Len(G1) - 3)
'
'    we8 = we8 + 1
'    'ScanPath.lblCount5 = Trim(Str(we8)) + " Del as Too Small <70 Bytes"
'
'    If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then
'        Fs.DeleteFile A1$ + G1$, True
'    End If
'
'    ScanPath.ListView1.ListItems.Remove (WE)
'    WE = WE - 1
'
'Next



'Exit Sub



'## MORE FILTER DEL ALL RAR
'## AND MORE YOU CHOOSE
ScanPath.Show
Reset
YY = "--"
    
'If ScanPath.ListView1.ListItems.COUNT = 0 Then
'
'    For WE = ScanPath.ListView1.ListItems.COUNT To 1 Step -1
'        ScanPath.lblCount3 = Trim(Str(WE))
'
'        B1$ = ScanPath.ListView1.ListItems.Item(WE)
'
'        If InStr(B1$, ".rar") > 0 Then
'            ScanPath.ListView1.ListItems.Remove (WE)
'        End If
'
'        '## YOU CHOOSE
'        If InStr(B1$, "_New") = 0 Then
'            ScanPath.ListView1.ListItems.Remove (WE)
'        End If
'
'    Next
'End If
'
'



we4 = ScanPath.ListView1.ListItems.COUNT
ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " After Same Size"

WE2 = 0

If wx8 > 0 Then ScanPath.lblCount4 = Trim(Str(WE2)) + " Del - Dupes = " + Format((WE2 / wx8) * 100, "0.00") + "%"

WE = 1
WE2 = ScanPath.ListView1.ListItems.COUNT
we5 = ScanPath.ListView1.ListItems.COUNT
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
    If WE2 Mod 10 = 0 Then
        DoEvents
        ScanPath.ListView1.ListItems(WE).EnsureVisible
    End If
    
    G1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(1)
    
    
    
    Do While XPAUSE = True
        DoEvents
        Sleep 500
    Loop
    

    
    
    If FSI1 <> OFSI1 Then
        YY = ""
    End If
    OFSI1 = FSI1

    ScanPath.ListView1.ListItems.Remove (WE)


    If 1 = 1 Then 'InStr(UCase(B1$), ".MP3") <> -200 Then
        'DONT HAVE ELSE IF
        
        
        Err.Clear
        WXHEX$ = Hex(m_CRC.CalculateFile(A1$ + G1$))
    
        If InStr(YY, WXHEX$) > 0 And Err.Number = 0 Then
            we3 = we3 + 1
            ScanPath.lblCount4 = Trim(Str(we3)) + " Del - Dupes = " + Format((we3 / we4) * 100, "0.00") + "%  -- PROGRESS  = " + Format(100 - ((WE2 / we5) * 100), "0.00") + "%"
            ScanPath.lblCount6 = WXHEX$ + " -- CRC DUPE CURRENT FILE"
            If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then
                Fs.DeleteFile A1$ + G1$
            End If
        End If
    
    
    Else
        
        'XXXX ---- SCAN FIRST CHUNCK OF STRING FROM FILE FOR OPTION SECOND
        'HERE IS IF MP3 THEN SCAN AFTER  HEADER
        
        Err.Clear
        'WxHex$ = Hex(m_CRC.CalculateFile(A1$ + G1$))
        Set F = Fs.getfile(A1 + G1)
        A22 = Space(F.Size - &H410)
        
        FR5 = FreeFile
        Open A1$ + G1$ For Binary As #FR5
             Get #FR5, &H410, A22
        Close #FR5
        'len(a22)+&h410
        
        WXHEX$ = Hex(m_CRC.CalculateString(A22))
        'A22 = ""

        If InStr(YY, WXHEX$) > 0 And Err.Number = 0 Then
            we3 = we3 + 1
'            ScanPath.lblCount4 = Trim(Str(we3)) + " Del - Dupes = " + Format((we3 / we4) * 100, "0.00") + "%"
            ScanPath.lblCount4 = Trim(Str(we3)) + " Del - Dupes = " + Format((we3 / we4) * 100, "0.00") + "%  -- PROGRESS  = " + Format(100 - ((WE2 / we5) * 100), "0.00") + "%"
            
            
            'LABLE DONT DEL FIRST RUN - SET LABLE EXPLAIN
            If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then
                
                'DEL TO RECYCLE
                'Fs.DeleteFile A1$ + G1$
                Call ShellFileDelete(A1$ + G1$, True)
            End If
        End If
    End If
    
    
    YY = YY + " - " + WXHEX$
   
   
    
'Next
Loop Until ScanPath.ListView1.ListItems.COUNT = 0


ScanPath.lblCount6 = "DONE *"
ScanPath.WindowState = 0


Next


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


If Fs.FolderExists(ScanPath.TxtPath) = False Then Exit Sub


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
    
wx8 = Val(ScanPath.ListView1.ListItems.COUNT)
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " Total"

'SORT SAME SIZE FILES AND DEL IF NOT SAME
For WE = ScanPath.ListView1.ListItems.COUNT To 2 Step -1
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT))
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
For WE = 1 To ScanPath.ListView1.ListItems.COUNT
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
For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT))
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

For WE = ScanPath.ListView1.ListItems.COUNT To 1 Step -1
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

we4 = ScanPath.ListView1.ListItems.COUNT
ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " After Same Size"

WE2 = 0

If wx8 > 0 Then ScanPath.lblCount4 = Trim(Str(WE2)) + " Del - Dupes = " + Format((WE2 / wx8) * 100, "0.00") + "%"

WE = 1
WE2 = ScanPath.ListView1.ListItems.COUNT
On Error Resume Next



SMALLSIZE = 1E+14!
'## GET SMALLEST
For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT))
    ScanPath.lblCount3 = Trim(Str(WE))
    FSI1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    'If Val(FSI1) > 70 Then Stop
    If Val(FSI1) >= 100 Then
        If Val(FSI1) < SMALLSIZE Then SMALLSIZE = Val(FSI1)
    End If


Next



'CRC ALL FILES BASED ON SMALLEST ONE

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.COUNT))
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


WE2 = ScanPath.ListView1.ListItems.COUNT + 1

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

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_DELETE
.pFrom = src
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileDelete = (lRet = 0)

End Function



