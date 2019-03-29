Attribute VB_Name = "BASE_MoD1"
Public TAP, PROCESSBEGIN
Public COUNTDIR

Public FILE22
Public FILE12 ' OUR FILE USED WITHER _ #FF02
Public TIME_SLICER
Public mCancelScan2

' ----------  STORE BASE REDUNDANT CODE
Public IRFAN_TO_DO_FROM_SEND_TO
Public UNLOAD_WITH_IRFAN_KILL

Public RUN_FROM_MEDIA_PLAYER_CLASSIC



'----
'VBA FIND SCREEN WIDTH API - Google Search
'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=VBA+FIND+SCREEN+WIDTH+API&spell=1&sa=X&ved=0ahUKEwivv5OOmr3WAhVHmbQKHTKNBfIQvwUIJSgA&biw=1536&bih=634
'--------
'vba - Screen Dimensions in Visual Basic - Stack Overflow
'https://stackoverflow.com/questions/7844171/screen-dimensions-in-visual-basic
'----
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
'----
'MsgBox GetSystemMetrics(SM_CXSCREEN) & "x" & GetSystemMetrics(SM_CYSCREEN)



Private Sub SP_FileMatch(Filename As String, DFilename As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    Dim ASIZE1 As Long
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
'    Print #FF01, Path + Filename
    
    If MNU_IRFAN_MODE_SET = True Then
        Print #FF01, Path + Filename
        
        Call Form_SEND_TO.XSCRIPT(Path + Filename)
    
    
        Exit Sub
    End If
    
'CRC
'If FILECOMPLEX = True Then
'Dim AL As Long
'CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)

'AL = uWIN32.cAlternate
'AL = uWIN32.ftLastWriteTime
'AL = uWIN32.ftLastAccessTime
'FormatFileDate(uWIN32.ftLastAccessTime , "YY-M-DD HH:mm:ss")

    Set F = FS.GetFile(Path + DFilename)
    
    STRING2 = "  Date Created:-"
    DAT1 = STRING2 + Format(F.DateCreated, "YYYY-MM-DD HH-MM-SS")
    
    STRING2 = "  Date Last Modified:-"
    DAT2 = STRING2 + Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
    
    STRING2 = "  Date Last Accessed:-"
    DAT3 = STRING2 + Format(F.DateLastAccessed, "YYYY-MM-DD HH-MM-SS")
            
    
    DAT4 = "  Size - " + Format(F.Size, "000,000,000,000,000,000") + " - "
    DAT4 = Replace(DAT4, "000,", "    ")
    DAT4 = Replace(DAT4, " 00,", "    ")
    DAT4 = Replace(DAT4, "  0,", "    ")

    i = Space(12)
    LSet i = DFilename
    DAT5 = i + " -- "
    
    If DFilename = "" Then Stop
    
'    CRC
'    STRING1 = String(10, ":")
'    STRING2 = DFilename
'    Mid(STRING1, 1, Len(STRING2)) = STRING2
'    STRING2 = Replace(STRING1, ":", " ") + " "
'
    'Print #FF01, DAT7; 'CRC
    
    
    
    
    DATN0 = Format(mDIRCount2 + mFILECount2, "000,000,000,000,000")
'    DATN0 = Replace(DATN0, "000,", "    ")
'    DATN0 = Replace(DATN0, " 00,", "    ")
'    DATN0 = Replace(DATN0, "  0,", "    ")

'    DATN1 = Format(mFILECount2, "000,000,000,000,000") + " -- "
'    DATN1 = Replace(DATN1, "000,", "    ")
'    DATN1 = Replace(DATN1, " 00,", "    ")
'    DATN1 = Replace(DATN1, "  0,", "    ")
    
    Print #FF01, "-- " + DATN0;
'    Print #FF01, DATN1;
    Print #FF01, DAT1;
    Print #FF01, DAT2;
    Print #FF01, DAT3;
    Print #FF01, DAT4;
    Print #FF01, DAT5;
    'Print #FF01, DAT7; 'CRC
    
    Print #FF01, Path + Filename
    
    Print #FF02, Path + Filename
    

    
Exit Sub
    
    
    
    If Path = "" Then Exit Sub
    If Filename = "" Then Exit Sub

'    If chk_LIST_VIEW_SHORT_5 = vbChecked Then
            'lblCount7 = Path + Filename

        Set F = FS.GetFile(Path + DFilename)
'        If XdATE2 < F.datelastmodified Then
'
'
'            XdATE2 = F.datelastmodified
'
'
'
'        End If
'                   lblCount7 = Path + Filename
                   
        With ListView1
        Set LV = .ListItems.Add(, , Filename)
        LV.SubItems(1) = Path
        'mod by matt l
        LV.SubItems(2) = DFilename
        
        LV.SubItems(3) = F.Size
        LV.SubItems(4) = Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
        
        
        
        
        End With
        
        TSTRING = Path + Filename
        Print #FF01, TSTRING
    
    
    'FF01
'    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
'    B1$ = ScanPath.ListView1.ListItems.Item(WE)
'    YY$ = A1$ + B1$
'    Print #1, YY$
'Next
'
'Close #1
        
        
        
        
        Exit Sub
    
'    End If
    



'    If XX = 58 Then
'
'        'lblCount1.Caption = mFileCount2
'        Print #1, "*" + Path + Filename
'
'    Exit Sub
'    End If
    
    
    
    
    
    With ListView1
        Set LV = .ListItems.Add(, , Filename)
        LV.SubItems(1) = Path
        'mod by matt l
        LV.SubItems(2) = DFilename
        
        
        
        
        
        
        
        
        
        
        If chkCopyMemory.Value Then
            'VB does not allows UDT's to be passed directly from a Class but we can access the structure
            'using pointers. We use CopyMemory to place the WIN32_FIND_DATA variable from the Class into
            'the uWIN32 declared in this Sub.
            
            'This is primarily a speed trick - we could retrieve the date & time information "ourselves"
            'using the standard VB functions or an API call but since we already have it...
            
            'CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
            'LV.SubItems(3) = uWIN32.nFileSizeLow
            'LV.SubItems(4) = FormatFileDate(uWIN32.ftLastWriteTime)
        
                    'Remember Public FS
            If Filename <> "" Then
                Set F = FS.GetFile(Path + DFilename)
                ADATE1 = F.DateLastModified
                ASIZE1 = F.Size
'                If Val(ASIZE1) = 647732 Then Stop
            Else  ' IF FOLDER
                
                
                
'                Set F = FS.GetFolder(Path)
'                ADATE1 = F.datelastmodified
'
'                'FOLDER SIZE BUT NOT IF REQUEST AND WHEN MULTI SAME SIZE
'
'
'                'We May want this later Slow Down is Massive
'                If OldFolder <> Path Then
'                    ASIZE1 = F.Size
'                Else
'                    ASIZE1 = OldSize
'                End If
'                OldFolder = Path
'                OldSize = ASIZE1
            End If
'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastWriteTime)
            LV.SubItems(3) = ASIZE1
            LV.SubItems(4) = Format(ADATE1, "YYYY-MM-DD HH-MM-SS")


        
        End If
        
        'Refresh Count & Scroll the list every 10th file
        If (LV.index Mod 10) = 0 Then
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
            End If
            
            'lblCount1.Caption = LV.Index
        End If
    End With
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical
End Sub






Private Sub SP_DirMatch(Directory As String, DDirectory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
    
    If MNU_IRFAN_MODE_SET = True Then Exit Sub
    
    Set F = FS.GetFolder(Path + DDirectory)
    
    
    STRING2 = "  Date Created:-"
    DAT1 = STRING2 + Format(F.DateCreated, "YYYY-MM-DD HH-MM-SS")
    
    STRING2 = "  Date Last Modified:-"
    DAT2 = STRING2 + Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
    
    STRING2 = "  Date Last Accessed:-"
    DAT3 = STRING2 + Format(F.DateLastAccessed, "YYYY-MM-DD HH-MM-SS")
            
    
    DAT4 = "  Size - " + Format(F.Size, "000,000,000,000,000,000") + " - "
    DAT4 = Replace(DAT4, "000,", "    ")
    DAT4 = Replace(DAT4, " 00,", "    ")
    DAT4 = Replace(DAT4, "  0,", "    ")
    

    'DAT5 = " - " + DDirectory + " -- "
    i = Space(12)
    LSet i = DDirectory
    DAT5 = "  " + i + " -- "
    
    If DDirectory = "" Then Stop
    
    
    
'    CRC
'    STRING1 = String(10, ":")
'    STRING2 = DFilename
'    Mid(STRING1, 1, Len(STRING2)) = STRING2
'    STRING2 = Replace(STRING1, ":", " ") + " "
'
    'Print #FF01, DAT7; 'CRC
    
    
    
    
    DATN0 = Format(mDIRCount2 + mFILECount2, "000,000,000,000,000")
'    DATN0 = Replace(DATN0, "000,", "    ")
'    DATN0 = Replace(DATN0, " 00,", "    ")
'    DATN0 = Replace(DATN0, "  0,", "    ")

'    DATN1 = Format(mFILECount2, "000,000,000,000,000") + " -- "
'    DATN1 = Replace(DATN1, "000,", "    ")
'    DATN1 = Replace(DATN1, " 00,", "    ")
'    DATN1 = Replace(DATN1, "  0,", "    ")
    
    Print #FF01, " - " + DATN0;
'    Print #FF01, DATN1;
    Print #FF01, DAT1;
    Print #FF01, DAT2;
    Print #FF01, DAT3;
    Print #FF01, DAT4;
    Print #FF01, DAT5;
    'Print #FF01, DAT7; 'CRC
    
    Print #FF01, Path + Directory
    
'    Print #FF02, Path + Directory
    
    Print #FF03, "D- " + DATN0;
    Print #FF03, DATN1;
    Print #FF03, DAT1;
    Print #FF03, DAT2;
    Print #FF03, DAT3;
    Print #FF03, DAT4;
    Print #FF03, DAT5;
    'Print #FF01, DAT7; 'CRC
    Print #FF03, Path + Directory
    
    Print #FF04, Path + Directory
    
    
    Exit Sub
    
    If Path = "" Then Exit Sub
    
    
    Dim LV As ListItem
    Set F = FS.GetFolder(Path + DDirectory)
    With ListView1
    Set LV = .ListItems.Add(, , "")
    LV.SubItems(1) = Path + Directory


    'mod by matt l
    LV.SubItems(2) = DDirectory
    
'    LV.SubItems(3) = F.Size
    LV.SubItems(4) = Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
    'TSTRING=Path + Directory
    
    
    
    'FF01
'    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
'    B1$ = ScanPath.ListView1.ListItems.Item(WE)
'    YY$ = A1$ + B1$
'    Print #1, YY$
'Next
'
'Close #1

    
    End With
    Exit Sub
    'OPATH = Path

'    TL = ScanPath.ListView1.ListItems.Count
'    If TL < 1 Then Exit Sub
'    ScanPath.ListView1.ListItems(TL).EnsureVisible
Exit Sub

'--------------


'sort of PATH then FILE
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.Sorted = False
    
    
For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    YY$ = A1 + B1
    Print #1, YY$
Next


ScanPath.ListView1.ListItems.Clear


End Sub




