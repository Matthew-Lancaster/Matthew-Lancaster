Attribute VB_Name = "Module1"
'---------------------
'Count = 090 -- Wed 06-Sep-2017 20:52:32
'---------------------
'Form FindWindow ---
'FreeVBCode code snippet: Recursively Delete Empty Folders - Google Chrome
'---------------------
'----
'FreeVBCode code snippet: Recursively Delete Empty Folders
'http://www.freevbcode.com/ShowCode.asp?ID=7821
'----
'---------------------

Dim i_2

Dim DEBUG_TEXT

Dim DELETE_COUNT


Sub Main()

Dim i As String

Debug.Print ""
Debug.Print ""

If Command$ <> "" Then Command_ALT = Command$

If Command_ALT = "" And IsIDE = False Then
    MsgBox "ENTER SOMETHING PATH FOR THE COMMAND LINE", vbMsgBoxSetForeground
    If IsIDE = False Then End
End If

If IsIDE = True And Command_ALT = "" Then Command_ALT = "F:\"
'If IsIDE = True Then Command_ALT = "F:\"


i = "J:\"
'i = "\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\# MY DOCS\# 01 My Documents\My Pictures2\"
'i = "Y:\0 00 ART\# 00 My Pictures\01 All Psycho\2008 Internet My Website Skull Bak"
i = "\\4-asus-gl522vw\4_asus_gl522vw_02_d_drive\DSC\"
i = "\\5-asus-p2520la\5_asus_p2520la_02_d_drive\DSC"
i = "I:\"
'i = "D:\DSC 0 00 VIDEO CAMERA - ME"
i = "D:\VIDEO"
i = "D:\VB6-EXE"
i = "D:\VB6-EXE"

Call FindFiles(i, "*.EXE", True)
End

Call DeleteEmptyFolders(i)
i = "\\1-asus-x5dij\1_asus_x5dij_02_d_drive\VB6-EXE"
Call DeleteEmptyFolders(i)
i = "\\2-asus-eee\2_asus_eee_02_d_drive\VB6-EXE"
Call DeleteEmptyFolders(i)
i = "\\3-linda-pc\3_linda_pc_02_d_drive\VB6-EXE"
Call DeleteEmptyFolders(i)
i = "\\4-asus-gl522vw\4_asus_gl522vw_02_d_drive\VB6-EXE"
Call DeleteEmptyFolders(i)
i = "\\5-asus-p2520la\5_asus_p2520la_02_d_drive\VB6-EXE"
Call DeleteEmptyFolders(i)
i = "\\8-msi-gp62m-7rd\8_msi_gp62m_7rd_02_d_drive\VB6-EXE"
Call DeleteEmptyFolders(i)
End

If IsIDE = True Then Command_ALT = i


'--------------------------------------------------
'FILES TO KEEP
'DELETE ALL FILE BUT ONE CETAIN EXTENSION TYPE .EXE
'FOR THE VB6-EXE FOLDER
'GOT MASHED WITH OTHER SET
'--------------------------------------------------
'Call DeleteFileExtension(Command_ALT, ".EXE")
'End


i = MsgBox("WORK TO DO ASK DELETE EMPTY FOLDER BY COMMAND$ _ HERE VERFY IF WANT TO" + vbCrLf + "----" + vbCrLf + Command_ALT, vbYesNo + vbMsgBoxSetForeground)
If i <> vbYes Then End

Do
    DELETE_COUNT = 0
    Call DeleteEmptyFolders(Command_ALT)
    If DELETE_COUNT = OLD_DELETE_COUNT Or DELETE_COUNT = 0 Then Exit Do
    OLD_DELETE_COUNT = DELETE_COUNT
Loop Until False = True
MsgBox "DELETE COUNT =" + Str(DELETE_COUNT)


End Sub



Public Sub DeleteEmptyFolders(ByVal strFolderPath As String)

    Dim fsoSubFolders As Folders
    Dim fsoFolder As Folder
    Dim fsoSubFolder As Folder
    
    Dim strPaths()
    Dim lngFolder As Long
    Dim lngSubFolder As Long
      
    DoEvents
      
    Set m_fsoObject = New FileSystemObject
    If Not m_fsoObject.FolderExists(strFolderPath) Then
        MsgBox "NOT A PROPER FOLDER EXISTER", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    Set fsoFolder = m_fsoObject.GetFolder(strFolderPath)
    'If strFolderPath = "I:\DSC\2015 2016\2015 Sony CyberShot DSC-HX60V\DCIM\20150717\NOT ONLINE DOC" Then Stop
      
    On Error Resume Next
    Err.Clear
    
    'Has sub-folders
    XC = fsoFolder.SubFolders.Count
    If Err.Number = 0 Then
        If fsoFolder.SubFolders.Count > 0 Then
            lngFolder = 1
            ReDim strPaths(1 To fsoFolder.SubFolders.Count)
            'Get each sub-folders path and add to an array
            For Each fsoSubFolder In fsoFolder.SubFolders
                'If fsoSubFolder.Path = "I:\DSC\2015 2016\2015 Sony CyberShot DSC-HX60V\DCIM\20150717\NOT ONLINE DOC" Then Stop
    
                
                XCF = fsoSubFolder.Path
                If Err.Number = 0 Then
                    strPaths(lngFolder) = fsoSubFolder.Path
                    lngFolder = lngFolder + 1
                End If
            Next fsoSubFolder
            
            lngSubFolder = 1
            'Recursively call the function for each sub-folder
            Do While lngSubFolder < lngFolder
               Call DeleteEmptyFolders(strPaths(lngSubFolder))
                'If strPaths(lngSubFolder) = "I:\DSC\2015 2016\2015 Sony CyberShot DSC-HX60V\DCIM\20150717\NOT ONLINE DOC" Then Stop
               lngSubFolder = lngSubFolder + 1
            Loop
        End If
    End If
    
    If Err.Number = 0 Then
        DELETE_DONE = False
        
        
        'WITH FILES CHECK THINGS LIKE Thumbs.db DON'T EXIST HIDEN SEARCH AND DEL REQUIRED
        '--------------------------------------------------------------------------------
        'DEL I:\Thumbs.db /S /Q /F /AH
        'DEL I:\DESKTOP.INI /S /Q /F /AH
        'DEL I:\.picasa.ini /S /Q /F /AH
        'DEL i:\*.FFmpeg-Verify.txt /S /Q /F
        '--------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------
        'No sub-folders or files
        If fsoFolder.Files.Count = 0 And fsoFolder.SubFolders.Count = 0 Then
            DELETE_DONE = True
            DELETE_COUNT = DELETE_COUNT + 1
            Debug.Print Trim(Str(DELETE_COUNT)) + "--" + fsoFolder.Path
            Err.Clear
            fsoFolder.Delete True
            If Err.Number > 0 Then DELETE_DONE = False: DELETE_COUNT = DELETE_COUNT - 1
            
        End If
        
        DELETE_DONE = True
        If DELETE_DONE = False Then
            Err.Clear
            '----------------------------------------------------
            ' DON'T RUN WITHOUT FILESIZE TO MAKE RUN TIME QUICKER
            ' MAKES AN ERROR FOR YOU HERE
            ' YOUR DELETE WHOLE FOLDERS AT A GO
            ' DON'T RUN THIS WHOLE ROUTINE IF THAT THE ANSWER
            ' IT'S A DOUBLE CHECKER
            '----------------------------------------------------
            XC = fsoFolder.Size
            If Err.Number = 0 Then
                If XC = 0 And fsoFolder.SubFolders.Count = 0 Then
                    
                    
                    DELETE_COUNT = DELETE_COUNT + 1
                    Debug.Print Trim(Str(DELETE_COUNT)) + "--" + fsoFolder.Path
                    Err.Clear
                    fsoFolder.Delete True
                    If Err.Number > 0 Then
                        Debug.Print Trim(Str(DELETE_COUNT)) + "-- ERROR DELETE --" + Str(fsoFolder.Files.Count) + " FILES"
                        For Each fsoSubFILE In fsoFolder.Files
                            T = fsoSubFILE.Name
                            Debug.Print T
                        Next
                    
                    End If
                End If
            End If
        End If
    End If

End Sub


Public Sub DeleteFileExtension(ByVal strFolderPath As String, ByVal str_Extension As String)

    Dim fsoSubFolders As Folders
    Dim fsoFolder As Folder
    Dim fsoSubFolder As Folder
    
    Dim strPaths()
    Dim lngFolder As Long
    Dim lngSubFolder As Long
      
    DoEvents
      
    Set m_fsoObject = New FileSystemObject
    If Not m_fsoObject.FolderExists(strFolderPath) Then
        MsgBox "NOT A PROPER FOLDER EXISTER", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    Set fsoFolder = m_fsoObject.GetFolder(strFolderPath)
    'If strFolderPath = "I:\DSC\2015 2016\2015 Sony CyberShot DSC-HX60V\DCIM\20150717\NOT ONLINE DOC" Then Stop
      
    On Error Resume Next
    Err.Clear
    
    'Has sub-folders
    XC = fsoFolder.SubFolders.Count
    If Err.Number = 0 Then
        If fsoFolder.SubFolders.Count > 0 Then
            lngFolder = 1
            ReDim strPaths(1 To fsoFolder.SubFolders.Count)
            'Get each sub-folders path and add to an array
            For Each fsoSubFolder In fsoFolder.SubFolders
                'If fsoSubFolder.Path = "I:\DSC\2015 2016\2015 Sony CyberShot DSC-HX60V\DCIM\20150717\NOT ONLINE DOC" Then Stop
    
                
                XCF = fsoSubFolder.Path
                If Err.Number = 0 Then
                    strPaths(lngFolder) = fsoSubFolder.Path
                    lngFolder = lngFolder + 1
                End If
            Next fsoSubFolder
            
            lngSubFolder = 1
            'Recursively call the function for each sub-folder
            Do While lngSubFolder < lngFolder
               Call DeleteFileExtension(strPaths(lngSubFolder), str_Extension)
                'If strPaths(lngSubFolder) = "I:\DSC\2015 2016\2015 Sony CyberShot DSC-HX60V\DCIM\20150717\NOT ONLINE DOC" Then Stop
               lngSubFolder = lngSubFolder + 1
            Loop
        End If
    End If
    
    If Err.Number = 0 Then
        
        
        'WITH FILES CHECK THINGS LIKE Thumbs.db DON'T EXIST HIDEN SEARCH AND DEL REQUIRED
        '--------------------------------------------------------------------------------
        'DEL I:\Thumbs.db /S /Q /F /AH
        'DEL I:\DESKTOP.INI /S /Q /F /AH
        'DEL I:\.picasa.ini /S /Q /F /AH
        'DEL i:\*.FFmpeg-Verify.txt /S /Q /F
        '--------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------
        
        For Each File In fsoFolder.Files
            'Debug.Print File.Name
            
            XY = ""
            If InStrRev(File.Name, ".") > 0 Then
                XY = UCase(Mid(File.Name, InStrRev(File.Name, ".")))
            End If
            
            ' LIST ALL EXTENSIONS
            If XY <> "" And InStr(UCase(DEBUG_TEXT), XY) = 0 Then
                DEBUG_TEXT = DEBUG_TEXT + XY
                Debug.Print DEBUG_TEXT
            End If
            
            If XY <> "" And InStr(str_Extension, XY) = 0 Then
                DELETE_COUNT = DELETE_COUNT + 1
                Debug.Print Trim(Str(DELETE_COUNT)) + "--" + File.Name
                DELETE_DONE = True
                Err.Clear
                File.Delete True
                If Err.Number > 0 Then DELETE_DONE = False: DELETE_COUNT = DELETE_COUNT - 1
                If DELETE_DONE = False Then
                    MsgBox "FAILED DELETE"
                End If
            End If
        Next
        
        
    End If

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


