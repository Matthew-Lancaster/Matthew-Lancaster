VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CHECK HERE AFTER COMPARE CONFLICT
'Dim FLAG_DO_IN_CHUNK_OUT_OF_MEMORY
Dim FILE_BEFORE_OUT_OF_MEMORY
Dim FILE_AT_OUT_OF_MEMORY
Dim FSIZE_XY


'FRI MAY 12 PROBLEM


'CRC Checksum Class
'------------------------------------
'
'A very fast solution to calculate the
'CRC Checksum (at the moment CRC16 and
'CRC32 values) with the help of some
'pre-compiled assembler code
'
'(c) 2000, Fredrik Qvarfort
'


Option Explicit

Public VarCRCMSGBOXFLAG

Public Enum CRCAlgorithms
  CRC16
  CRC32
End Enum
Private m_Algorithm As Boolean

Private m_CRC16 As Long
Private m_CRC16Asm() As Byte
Private m_CRC16Init As Boolean
Private m_CRC16Table(0 To 255) As Long

Private m_CRC32 As Long
Private m_CRC32Asm() As Byte
Private m_CRC32Init As Boolean
Private m_CRC32Table(0 To 255) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



'-------------------------------------------------
'READ IN CHUNK FOR A WHOLE FILE BYTE ARRAY
'-------------------------------------------------
'-------------------------------------------------
'BLOCK READ BYTE ARRAY 01 OF 09 -- THE DECLARE
'-------------------------------------------------

'Option Explicit

Const MOVEFILE_REPLACE_EXISTING = &H1
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_BEGIN = 0
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_NEW = 1
Const OPEN_EXISTING = 3
Const OPEN_ALLWAYS = 4
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000

'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

Public File_Num As Long
Public File_Len As Currency

'-------------------------------------------------------
'BLOCK READ BYTE ARRAY 01 OF 09 -- THE DECLARE BLOCK END
'-------------------------------------------------------



'-------------------------------------------------------
'BLOCK READ BYTE ARRAY 01 OF 02 -- THE DECLARE BLOCK
'-------------------------------------------------------

'Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
'Private Declare Function LCMapStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByVal lpDestStr As Long, ByVal cchDest As Long) As Long
'
'Private Declare Function CreateFile Lib "KERNEL32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function LCMapString Lib "KERNEL32.dll" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
'Private Declare Function lstrlen Lib "KERNEL32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
'
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'
'Private Const INVALID_HANDLE_VALUE = -1
'
'Private Const GENERIC_READ = &H80000000
'Private Const GENERIC_WRITE = &H40000000
'
'Private Const FILE_SHARE_READ = &H1
'Private Const FILE_SHARE_WRITE = &H2
'
'Private Const CREATE_NEW = 1
'Private Const CREATE_ALWAYS = 2
'Private Const OPEN_EXISTING = 3
'Private Const OPEN_ALWAYS = 4
'Private Const TRUNCATE_EXISTING = 5
'
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'
'Private Const FILE_BEGIN = 0
'Private Const FILE_CURRENT = 1
'Private Const FILE_END = 2
'-------------------------------------------------------
'BLOCK READ BYTE ARRAY 01 OF 02 -- THE DECLARE BLOCK END
'-------------------------------------------------------
'
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Function CopyArray(Source1() As Byte, Source2() As Byte) As Byte()

    Dim Size1 As Long
    Dim Size2 As Long
    Dim Result() As Byte
    
    'number of bytes in each array
    'note that a Long uses 4 bytes
    Size1 = (UBound(Source1) - LBound(Source1) + 1) '* 4
    Size2 = (UBound(Source2) - LBound(Source2) + 1) ' * 4
    
    'set the size of the resulting array
    ReDim Result(UBound(Source1) + UBound(Source2))
    
'    CopyMemory Result(LBound(Result)), Source1(LBound(Source1)), Size1
    CopyMemory Result(UBound(Source1)), Source2(LBound(Source2)), Size2
    
    CopyArray = Result

End Function


Public Function AddBytes(ByteArray() As Byte) As Variant

  Dim ByteSize As Long
  
  'We need to add a simple error trapping
  'here because if the bytearray is not
  'dimensioned we want it to just skip
  'the assembler code call below
  On Local Error GoTo NoData
  
  'Precalculate the size of the byte array
  ByteSize = UBound(ByteArray) - LBound(ByteArray) + 1
  
  'No error trapping needed, if something
  'goes bad below something is definitely
  'fishy with your computer
  On Local Error GoTo 0
  
  'Run the pre-compiled assembler code
  'for the current selected algorithm
  Select Case m_Algorithm
  Case CRC16
    Call CallWindowProc(VarPtr(m_CRC16Asm(0)), VarPtr(m_CRC16), VarPtr(ByteArray(LBound(ByteArray))), VarPtr(m_CRC16Table(0)), ByteSize)
  Case CRC32
    Call CallWindowProc(VarPtr(m_CRC32Asm(0)), VarPtr(m_CRC32), VarPtr(ByteArray(LBound(ByteArray))), VarPtr(m_CRC32Table(0)), ByteSize)
  End Select
  
NoData:
  'Return the current CRC value
  AddBytes = Value
  
End Function
Public Function AddString(Text As String) As Variant

  'Convert the string into a byte array
  'and send it to the function that can
  'handle bytearrays
  AddString = AddBytes(StrConv(Text, vbFromUnicode))
  
End Function
Public Property Let Algorithm(New_Value As CRCAlgorithms)

  'Set the new algorithm
  m_Algorithm = New_Value

  'Make sure we have initialized the
  'current selected algorithm
  Select Case m_Algorithm
  Case CRC16
    If (Not m_CRC16Init) Then Call InitializeCRC16
  Case CRC32
    If (Not m_CRC32Init) Then Call InitializeCRC32
  End Select

  'Make sure we reset the data of the
  'current selected algorithm
  Call Clear
  
End Property
Public Property Get Algorithm() As CRCAlgorithms

  Algorithm = m_Algorithm
  
End Property

Public Function CalculateBytes(ByteArray() As Byte) As Variant

  'Reset the current CRC calculation
  Call Clear
  
  'Calculate the CRC from the bytearray
  'and return the current CRC value
  CalculateBytes = AddBytes(ByteArray)
  
End Function


Function FileLen(FileName)
Dim Fs, F
Set Fs = CreateObject("Scripting.FileSystemObject")
Set F = Fs.getfile(FileName)
FileLen = F.Size
  

End Function

Public Function CalculateFile_Part(FileName As String, MP3_HEADER_TAG_SIZE) As Variant

  Dim Filenr As Integer
  Dim ByteArray() As Byte
  
  'Make sure the file contains data
  'to avoid errors later below
  If (FileLen(FileName) = 0) Then Exit Function
  
  'Open the file in binary mode, read
  'the data into a bytearray and then
  'close the file
  ReDim ByteArray(0 To 0)
'  ByteArray = Nothing
  On Error GoTo CalcErrHandler
  Filenr = FreeFile
  Open FileName For Binary As #Filenr
    ReDim ByteArray(0 To (LOF(Filenr) - 1) - MP3_HEADER_TAG_SIZE)
    Get #Filenr, MP3_HEADER_TAG_SIZE, ByteArray()
  Close #Filenr
  
  'Now send the bytearray to the function
  'that can calculate a CRC from it
  CalculateFile_Part = CalculateBytes(ByteArray)
  Exit Function
  
CalcErrHandler:
  If Err.Number > 0 Then
    If VarCRCMSGBOXFLAG = False Then
        MsgBox "Error Read File in CRC Check" + vbCrLf + FileName + vbCrLf + Err.Description
    End If
    If IsIDE And VarCRCMSGBOXFLAG = False Then
        Stop
'        Reset
'        Close
        Resume Next
    Else
        If VarCRCMSGBOXFLAG = False Then
            Dim i
            i = MsgBox("Do You Want to Unload" + vbCrLf + "Cancel For No and Ignore These Messages", vbYesNoCancel)
            If i = vbYes Then Unload ScanPath: Exit Function
            If i = vbCancel Then VarCRCMSGBOXFLAG = True
        End If
    End If
  End If
  CalculateFile_Part = "00000000"
End Function

Public Function DO_CHUNCK_READ_CalculateFile(FileName As String) As String

    'ORGINALLY USED TO BE
    'Public Function DO_CHUNCK_READ_CalculateFile(FileName As String) As Variant
    'BUT UNABLE TO ADD STRING TOGETHER
    'SORT STRING ADD SOLUTION BY USE HEX STRING WAS NEED PASS BACK TO ROUTINE TO DO
    'ANYHOW
    
    'Dim DO_CHUNCK_READ_CalculateFile As String
    
    
    DO_CHUNCK_READ_CalculateFile = "00000000"
    
    Dim Filenr As Integer
    Dim ByteArray() As Byte
    Dim NEST_LOOP, NEST_LOOP_CHUNK
    Dim VAR_CRC
    Dim CHUNK_SIZE
    
    'Make sure the file contains data
    'to avoid errors later below
    FSIZE_XY = FileLen(FileName)
    If FSIZE_XY = 0 Then Exit Function

    
    'Open the file in binary mode, read
    'the data into a bytearray and then
    'close the file
    
    '-------------------------------------------------------------
    'SAT MAY 14 WORK TO CORRECT THIS MEMORY OVERFLOW SOLUTION
    'GET MEMORY ERROR HERE LOAD LINE -- Get #Filenr, , ByteArray()
    'ABOUT 3000 MEG FILE STUCK AT
    '-------------------------------------------------------------
    
      
        
    On Error GoTo CalcErrHandler
    Filenr = FreeFile
    Open FileName For Binary As #Filenr
        
        'MEG = 10*1024 ^ 2
        CHUNK_SIZE = 10 * 1024 ^ 2 '10 MEG
        
        For NEST_LOOP = 0 To LOF(Filenr) - 1 Step CHUNK_SIZE
                    
            NEST_LOOP_CHUNK = NEST_LOOP + CHUNK_SIZE
            If NEST_LOOP_CHUNK > LOF(Filenr) - 1 Then
                NEST_LOOP_CHUNK = LOF(Filenr) - 1
            End If
            
            ReDim ByteArray(NEST_LOOP To NEST_LOOP_CHUNK)
            Get #Filenr, , ByteArray()
            
            'Now send the bytearray to the function
            'that can calculate a CRC from it
            
            VAR_CRC = CalculateBytes(ByteArray)
            
            DO_CHUNCK_READ_CalculateFile = DO_CHUNCK_READ_CalculateFile + "-" + Hex(VAR_CRC)
            'LEN(DO_CHUNCK_READ_CalculateFile)
    
            'ScanPath.lblCount12_MEMORY_OVERFLOW2 = Format(FSIZE_XY / 1024 ^ 2, , "###,###,###,###.0") + " FSIZE After Memory Overflow"
            ScanPath.lblCount12_MEMORY_OVERFLOW2 = Format(FSIZE_XY, "###,###,###,###") + " Byte @ Memory Chunk MODE"
            ScanPath.lblCount12_MEMORY_OVERFLOW3 = "Reading CRC In Chunk Mode Memory Limit"

    
        Next
    
    Close #Filenr
    
    Exit Function
  
CalcErrHandler:
    If Err.Number > 0 Then
        DO_CHUNCK_READ_CalculateFile = "Error - " + Err.Description
      
        If VarCRCMSGBOXFLAG = False Then
            MsgBox "Error Read File in CRC Check" + vbCrLf + FileName + vbCrLf + Err.Description
            If Err.Description = "Out of memory" Then
                MsgBox "PROBLEM STILL WITH MEMORY IN CHUNCK MODE" + vbCrLf + "Err.Description = Out of memory"
            End If
        End If
        
        If IsIDE And VarCRCMSGBOXFLAG = False Then
            Stop
            Resume Next
        Else
            If VarCRCMSGBOXFLAG = False Then
                Dim i
                i = MsgBox("Do You Want to Unload" + vbCrLf + "Cancel For No and Ignore These Messages", vbYesNoCancel)
                If i = vbYes Then Unload ScanPath: Exit Function
                If i = vbCancel Then VarCRCMSGBOXFLAG = True
            End If
        End If
    End If
    
    DO_CHUNCK_READ_CalculateFile = "00000000"

End Function


'----------------------------------------------------------


Public Function CalculateFile(FileName As String) As String

    'ORGINALLY USED TO BE
    'Public Function CalculateFile(FileName As String) As Variant
    'BUT UNABLE TO ADD STRING TOGETHER
  
    If FLAG_DO_IN_CHUNK_OUT_OF_MEMORY = True Then
        CalculateFile = DO_CHUNCK_READ_CalculateFile(FileName)
        Exit Function
    End If
    
    Dim Filenr As Integer
    Dim ByteArray() As Byte
    Dim ByteArray_CHUNK() As Byte
    Dim ByteArray_TEST() As Byte

    Dim NEST_LOOP, NEST_LOOP_CHUNK
    Dim VAR_CRC
    Dim CHUNK_SIZE
    Dim CHUNK_COUNTER
    Dim TEST_IN_IDE
    
    CalculateFile = "00000000"
    
    'Make sure the file contains data
    'to avoid errors later below
    Dim FSIZE_XY As Currency
    FSIZE_XY = FileLen(FileName)
    If FSIZE_XY = 0 Then Exit Function
    
    'Open the file in binary mode, read
    'the data into a bytearray and then
    'close the file
    
    '-------------------------------------------------------------
    'FRI MAY 12 PROBLEM
    'GET MEMORY ERROR HERE LOAD LINE -- Get #Filenr, , ByteArray()
    'ABOUT 1000 MEG FILE STUCK AT
    '-------------------------------------------------------------
'    On Error GoTo CalcErrHandler
    Filenr = FreeFile
    Open FileName For Binary As #Filenr
        ReDim ByteArray(0 To LOF(Filenr) - 1)
        Get #Filenr, , ByteArray()
    Close #Filenr
    'Now send the bytearray to the function
    'that can calculate a CRC from it
     VAR_CRC = CalculateBytes(ByteArray)
     TEST_IN_IDE = False
     If TEST_IN_IDE = True Then Debug.Print VAR_CRC
     If TEST_IN_IDE = True Then Debug.Print UBound(ByteArray)
'PROJECT


BACKONE:
    'MEG = 10*1024 ^ 2
    CHUNK_SIZE = 10 * 1024 ^ 2 '10 MEG
    ReDim ByteArray(0 To FSIZE_XY - 1)
    ReDim ByteArray(0)
    
    ReDim ByteArray_CHUNK(0 To CHUNK_SIZE)

    Filenr = FreeFile
    Open FileName For Binary As #Filenr
        For NEST_LOOP = 1 To FSIZE_XY Step CHUNK_SIZE
'            CHUNK_COUNTER = CHUNK_COUNTER + 1
            NEST_LOOP_CHUNK = NEST_LOOP + CHUNK_SIZE
            If NEST_LOOP_CHUNK > FSIZE_XY Then
                NEST_LOOP_CHUNK = FSIZE_XY - 1
                ReDim ByteArray_CHUNK(0 To (FSIZE_XY - NEST_LOOP))
                Else: Stop
            End If
            
            Get #Filenr, , ByteArray_CHUNK()
'            ByteArray(NEST_LOOP - 1) = ByteArray_CHUNK()
            ByteArray_TEST() = CopyArray(ByteArray, ByteArray_CHUNK)
'            ByteArray() = CopyArray(ByteArray, ByteArray_CHUNK)

        Next
    Close #Filenr
            
'    DO_CHUNCK_READ_CalculateFile = DO_CHUNCK_READ_CalculateFile + "-" + Hex(VAR_CRC)
    'LEN(DO_CHUNCK_READ_CalculateFile)

    'ScanPath.lblCount12_MEMORY_OVERFLOW2 = Format(FSIZE_XY / 1024 ^ 2, , "###,###,###,###.0") + " FSIZE After Memory Overflow"
    ScanPath.lblCount12_MEMORY_OVERFLOW2 = Format(FSIZE_XY, "###,###,###,###") + " Byte @ Memory Chunk MODE"
    ScanPath.lblCount12_MEMORY_OVERFLOW3 = "Reading CRC In Chunk Mode Memory Limit"

    VAR_CRC = Hex(CalculateBytes(ByteArray_CHUNK))
    If TEST_IN_IDE = True Then Debug.Print VAR_CRC
    If TEST_IN_IDE = True Then Debug.Print UBound(ByteArray_CHUNK)
    
    VAR_CRC = Hex(CalculateBytes(ByteArray_TEST))
    If TEST_IN_IDE = True Then Debug.Print VAR_CRC
    If TEST_IN_IDE = True Then Debug.Print UBound(ByteArray_TEST)
    
    'Now send the bytearray to the function
'    'that can calculate a CRC from it
'    CalculateFile = Hex(CalculateBytes(ByteArray))
'    If TEST_IN_IDE = True Then Debug.Print CalculateFile
'    If TEST_IN_IDE = True Then Debug.Print (UBound(ByteArray) - LBound(ByteArray) + 1)
     VAR_CRC = VAR_CRC
'    GoTo BACKONE


'    If Get(hFile, _
      Buffer(LBound(Buffer)), _
      UBound(Buffer) - LBound(Buffer) + 1, _
      ReadBytes, _
      0) Then

'      If ReadBytes = 0 Then
'         fEOF = True
'      ElseIf ReadBytes < UBound(Buffer) - LBound(Buffer) + 1 Then
'         ReDim Preserve Buffer(ReadBytes - 1)
'      End If
      
'        If ReadFile(hFile, _
'      Buffer(LBound(Buffer)), _
'      UBound(Buffer) - LBound(Buffer) + 1, _
'      ReadBytes, _
'      0) Then
'
'      If ReadBytes = 0 Then
'         fEOF = True
'      ElseIf ReadBytes < UBound(Buffer) - LBound(Buffer) + 1 Then
'         ReDim Preserve Buffer(ReadBytes - 1)
'      End If
      
      
'    Close #Filenr
'    NEED SMOOTHER READ BYTEARRAY IN CHUNCKS AND TOTAL IT
'    LEARN
    
'    Call API_ReadFile(FileName, 1, 1000, ByteArray())
'    Call API_OpenFile(FileName, 1, FSIZE_XY)
    'API_OpenFile(ByVal FileName As String, ByRef FileNumber As Long, ByRef FileSize As Currency)

    
    
    'Now send the bytearray to the function
    'that can calculate a CRC from it
'    CalculateFile = Hex(CalculateBytes(ByteArray))
    
    FILE_BEFORE_OUT_OF_MEMORY = FSIZE_XY
    
'    ScanPath.lblCount11_MEMORY_OVERFLOW1 = Format(FSIZE_XY / 1024 ^ 2, "###,###,###,###.0") + " MB Before Memory Overflow"
'    ScanPath.lblCount12_MEMORY_OVERFLOW2 = Format(FSIZE_XY / 1024 ^ 2, "###,###,###,###.0") + " MB After Memory Overflow -- Not Yet"
    ScanPath.lblCount11_MEMORY_OVERFLOW1 = Format(FSIZE_XY, "###,###,###,###") + " Byte @ Memory OverFlow"
    ScanPath.lblCount12_MEMORY_OVERFLOW2 = Format(FSIZE_XY, "###,###,###,###") + " Byte @ Memory Chunk MODE"
    ScanPath.lblCount12_MEMORY_OVERFLOW3 = "Reading CRC in Whole File Mode Memory Okay"

    Exit Function
  
CalcErrHandler:
    If Err.Number > 0 Then
        CalculateFile = "Err.Description -- " + Err.Description
        
        If VarCRCMSGBOXFLAG = False Then
            'MsgBox "Error Read File in CRC Check" + vbCrLf + FileName + vbCrLf + Err.Description
        
            If Err.Description = "Out of memory" Then
                Close #Filenr
                FILE_AT_OUT_OF_MEMORY = FSIZE_XY
                
                FLAG_DO_IN_CHUNK_OUT_OF_MEMORY = True
                
                '\\1-ASUS-X5DIJ\ASUS-X5DIJ D\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\MP_ROOT0
                If IsIDE = True Then
                    Stop
                Else
                    MsgBox "CHUNK MEMORY MODE HAS BEGAN -- BUT HAS FAULT AT MOMENT YOU MAY WANT TO STOP -- WILL BE PUT IN PAUSE" + vbCrLf + "FOR YOU TO INCIATE -- AN EXIT WHEN READY"
                
                    '--------------------------
                    ScanPath.Command2.BackColor = vbRed
                    If CommandPause = True Then
                        Do
                            DoEvents
                            Sleep 50
                            DoEvents
                        Loop Until CommandPause = False Or CommandExit = True
                    End If
                    'If CommandExit = True Then Exit For
                    If CommandExit = True Then Exit Function
                    '--------------------------
                
                
                
                End If
                
                
                
                CalculateFile = DO_CHUNCK_READ_CalculateFile(FileName)
            
            Else
            
                MsgBox "Error Read File in CRC Check" + vbCrLf + String(40, "-") + vbCrLf + FileName + vbCrLf + String(40, "-") + vbCrLf + "Err.Description -- " + Err.Description, , ScanPath.Caption
            
            End If
        End If
        
        If IsIDE And VarCRCMSGBOXFLAG = False Then
            If FLAG_DO_IN_CHUNK_OUT_OF_MEMORY = False Then
                Stop
            End If
            'Resume Next
        Else
            If VarCRCMSGBOXFLAG = False Then
                Dim i
                i = MsgBox("Do You Want to Unload" + vbCrLf + "Yes = Yes Unload - Cancel For No Continue and Ignore These Messages", vbYesNoCancel)
                If i = vbYes Then Unload ScanPath: Exit Function
                If i = vbCancel Then VarCRCMSGBOXFLAG = True
            End If
        End If
    End If

'    If FLAG_DO_IN_CHUNK_OUT_OF_MEMORY = False Then
'        CalculateFile = "00000000"
'    End If

End Function



Public Function CalculateString(Text As String)

  'Convert the string into a bytearray
  'and send it to the function that
  'calculates the CRC from a bytearray
  CalculateString = CalculateBytes(StrConv(Text, vbFromUnicode))
  
End Function
Public Property Get Value() As Variant

  Select Case m_Algorithm
  Case CRC16
    Value = (m_CRC16 And 65535)
  Case CRC32
    Value = (Not m_CRC32)
  End Select
  
End Property

Public Property Let Value(New_Value As Variant)

  Select Case m_Algorithm
  Case CRC16
    m_CRC16 = New_Value
  Case CRC32
    m_CRC32 = New_Value
  End Select
  
End Property

Private Sub InitializeCRC16()

  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim CRC As Long
  Dim sASM As String
  
  'Create the fixed lookup-table, this
  'is calculated because it won't take
  'long and is only done once
  For i = 0 To 255
    k = i * 256
    CRC = 0
    For j = 0 To 7
      If (((CRC Xor k) And 32768) = 32768) Then
        CRC = (CRC * 2) Xor &H1021
      Else
        CRC = (CRC * 2)
      End If
      k = k * 2
    Next
    m_CRC16Table(i) = CRC '(CRC And 65535)
  Next
  
  'Create a bytearray to hold the
  'precompiled assembler code
  sASM = "5589E55756505351528B45088B008B750C8B7D108B4D1431DB8A1E30E3668B149F30C66689D0464975EF25FFFF00008B4D0889015A595B585E5F89EC5DC21000"
  ReDim m_CRC16Asm(0 To Len(sASM) \ 2 - 1)
  For i = 1 To Len(sASM) Step 2
    m_CRC16Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
  Next
  
  'Mark the CRC16 algorithm as initialized
  m_CRC16Init = True
  
End Sub
Public Sub Clear()

  'Here can be sloppy and reset both
  'crc variables (this procedure will
  'be more advanced when adding more
  'checksums algorithms..)
  m_CRC16 = 0
  m_CRC32 = &HFFFFFFFF
  
End Sub

Private Sub InitializeCRC32()

  Dim i As Long
  Dim sASM As String
  
  m_CRC32Table(0) = &H0
  m_CRC32Table(1) = &H77073096
  m_CRC32Table(2) = &HEE0E612C
  m_CRC32Table(3) = &H990951BA
  m_CRC32Table(4) = &H76DC419
  m_CRC32Table(5) = &H706AF48F
  m_CRC32Table(6) = &HE963A535
  m_CRC32Table(7) = &H9E6495A3
  m_CRC32Table(8) = &HEDB8832
  m_CRC32Table(9) = &H79DCB8A4
  m_CRC32Table(10) = &HE0D5E91E
  m_CRC32Table(11) = &H97D2D988
  m_CRC32Table(12) = &H9B64C2B
  m_CRC32Table(13) = &H7EB17CBD
  m_CRC32Table(14) = &HE7B82D07
  m_CRC32Table(15) = &H90BF1D91
  m_CRC32Table(16) = &H1DB71064
  m_CRC32Table(17) = &H6AB020F2
  m_CRC32Table(18) = &HF3B97148
  m_CRC32Table(19) = &H84BE41DE
  m_CRC32Table(20) = &H1ADAD47D
  m_CRC32Table(21) = &H6DDDE4EB
  m_CRC32Table(22) = &HF4D4B551
  m_CRC32Table(23) = &H83D385C7
  m_CRC32Table(24) = &H136C9856
  m_CRC32Table(25) = &H646BA8C0
  m_CRC32Table(26) = &HFD62F97A
  m_CRC32Table(27) = &H8A65C9EC
  m_CRC32Table(28) = &H14015C4F
  m_CRC32Table(29) = &H63066CD9
  m_CRC32Table(30) = &HFA0F3D63
  m_CRC32Table(31) = &H8D080DF5
  m_CRC32Table(32) = &H3B6E20C8
  m_CRC32Table(33) = &H4C69105E
  m_CRC32Table(34) = &HD56041E4
  m_CRC32Table(35) = &HA2677172
  m_CRC32Table(36) = &H3C03E4D1
  m_CRC32Table(37) = &H4B04D447
  m_CRC32Table(38) = &HD20D85FD
  m_CRC32Table(39) = &HA50AB56B
  m_CRC32Table(40) = &H35B5A8FA
  m_CRC32Table(41) = &H42B2986C
  m_CRC32Table(42) = &HDBBBC9D6
  m_CRC32Table(43) = &HACBCF940
  m_CRC32Table(44) = &H32D86CE3
  m_CRC32Table(45) = &H45DF5C75
  m_CRC32Table(46) = &HDCD60DCF
  m_CRC32Table(47) = &HABD13D59
  m_CRC32Table(48) = &H26D930AC
  m_CRC32Table(49) = &H51DE003A
  m_CRC32Table(50) = &HC8D75180
  m_CRC32Table(51) = &HBFD06116
  m_CRC32Table(52) = &H21B4F4B5
  m_CRC32Table(53) = &H56B3C423
  m_CRC32Table(54) = &HCFBA9599
  m_CRC32Table(55) = &HB8BDA50F
  m_CRC32Table(56) = &H2802B89E
  m_CRC32Table(57) = &H5F058808
  m_CRC32Table(58) = &HC60CD9B2
  m_CRC32Table(59) = &HB10BE924
  m_CRC32Table(60) = &H2F6F7C87
  m_CRC32Table(61) = &H58684C11
  m_CRC32Table(62) = &HC1611DAB
  m_CRC32Table(63) = &HB6662D3D
  m_CRC32Table(64) = &H76DC4190
  m_CRC32Table(65) = &H1DB7106
  m_CRC32Table(66) = &H98D220BC
  m_CRC32Table(67) = &HEFD5102A
  m_CRC32Table(68) = &H71B18589
  m_CRC32Table(69) = &H6B6B51F
  m_CRC32Table(70) = &H9FBFE4A5
  m_CRC32Table(71) = &HE8B8D433
  m_CRC32Table(72) = &H7807C9A2
  m_CRC32Table(73) = &HF00F934
  m_CRC32Table(74) = &H9609A88E
  m_CRC32Table(75) = &HE10E9818
  m_CRC32Table(76) = &H7F6A0DBB
  m_CRC32Table(77) = &H86D3D2D
  m_CRC32Table(78) = &H91646C97
  m_CRC32Table(79) = &HE6635C01
  m_CRC32Table(80) = &H6B6B51F4
  m_CRC32Table(81) = &H1C6C6162
  m_CRC32Table(82) = &H856530D8
  m_CRC32Table(83) = &HF262004E
  m_CRC32Table(84) = &H6C0695ED
  m_CRC32Table(85) = &H1B01A57B
  m_CRC32Table(86) = &H8208F4C1
  m_CRC32Table(87) = &HF50FC457
  m_CRC32Table(88) = &H65B0D9C6
  m_CRC32Table(89) = &H12B7E950
  m_CRC32Table(90) = &H8BBEB8EA
  m_CRC32Table(91) = &HFCB9887C
  m_CRC32Table(92) = &H62DD1DDF
  m_CRC32Table(93) = &H15DA2D49
  m_CRC32Table(94) = &H8CD37CF3
  m_CRC32Table(95) = &HFBD44C65
  m_CRC32Table(96) = &H4DB26158
  m_CRC32Table(97) = &H3AB551CE
  m_CRC32Table(98) = &HA3BC0074
  m_CRC32Table(99) = &HD4BB30E2
  m_CRC32Table(100) = &H4ADFA541
  m_CRC32Table(101) = &H3DD895D7
  m_CRC32Table(102) = &HA4D1C46D
  m_CRC32Table(103) = &HD3D6F4FB
  m_CRC32Table(104) = &H4369E96A
  m_CRC32Table(105) = &H346ED9FC
  m_CRC32Table(106) = &HAD678846
  m_CRC32Table(107) = &HDA60B8D0
  m_CRC32Table(108) = &H44042D73
  m_CRC32Table(109) = &H33031DE5
  m_CRC32Table(110) = &HAA0A4C5F
  m_CRC32Table(111) = &HDD0D7CC9
  m_CRC32Table(112) = &H5005713C
  m_CRC32Table(113) = &H270241AA
  m_CRC32Table(114) = &HBE0B1010
  m_CRC32Table(115) = &HC90C2086
  m_CRC32Table(116) = &H5768B525
  m_CRC32Table(117) = &H206F85B3
  m_CRC32Table(118) = &HB966D409
  m_CRC32Table(119) = &HCE61E49F
  m_CRC32Table(120) = &H5EDEF90E
  m_CRC32Table(121) = &H29D9C998
  m_CRC32Table(122) = &HB0D09822
  m_CRC32Table(123) = &HC7D7A8B4
  m_CRC32Table(124) = &H59B33D17
  m_CRC32Table(125) = &H2EB40D81
  m_CRC32Table(126) = &HB7BD5C3B
  m_CRC32Table(127) = &HC0BA6CAD
  m_CRC32Table(128) = &HEDB88320
  m_CRC32Table(129) = &H9ABFB3B6
  m_CRC32Table(130) = &H3B6E20C
  m_CRC32Table(131) = &H74B1D29A
  m_CRC32Table(132) = &HEAD54739
  m_CRC32Table(133) = &H9DD277AF
  m_CRC32Table(134) = &H4DB2615
  m_CRC32Table(135) = &H73DC1683
  m_CRC32Table(136) = &HE3630B12
  m_CRC32Table(137) = &H94643B84
  m_CRC32Table(138) = &HD6D6A3E
  m_CRC32Table(139) = &H7A6A5AA8
  m_CRC32Table(140) = &HE40ECF0B
  m_CRC32Table(141) = &H9309FF9D
  m_CRC32Table(142) = &HA00AE27
  m_CRC32Table(143) = &H7D079EB1
  m_CRC32Table(144) = &HF00F9344
  m_CRC32Table(145) = &H8708A3D2
  m_CRC32Table(146) = &H1E01F268
  m_CRC32Table(147) = &H6906C2FE
  m_CRC32Table(148) = &HF762575D
  m_CRC32Table(149) = &H806567CB
  m_CRC32Table(150) = &H196C3671
  m_CRC32Table(151) = &H6E6B06E7
  m_CRC32Table(152) = &HFED41B76
  m_CRC32Table(153) = &H89D32BE0
  m_CRC32Table(154) = &H10DA7A5A
  m_CRC32Table(155) = &H67DD4ACC
  m_CRC32Table(156) = &HF9B9DF6F
  m_CRC32Table(157) = &H8EBEEFF9
  m_CRC32Table(158) = &H17B7BE43
  m_CRC32Table(159) = &H60B08ED5
  m_CRC32Table(160) = &HD6D6A3E8
  m_CRC32Table(161) = &HA1D1937E
  m_CRC32Table(162) = &H38D8C2C4
  m_CRC32Table(163) = &H4FDFF252
  m_CRC32Table(164) = &HD1BB67F1
  m_CRC32Table(165) = &HA6BC5767
  m_CRC32Table(166) = &H3FB506DD
  m_CRC32Table(167) = &H48B2364B
  m_CRC32Table(168) = &HD80D2BDA
  m_CRC32Table(169) = &HAF0A1B4C
  m_CRC32Table(170) = &H36034AF6
  m_CRC32Table(171) = &H41047A60
  m_CRC32Table(172) = &HDF60EFC3
  m_CRC32Table(173) = &HA867DF55
  m_CRC32Table(174) = &H316E8EEF
  m_CRC32Table(175) = &H4669BE79
  m_CRC32Table(176) = &HCB61B38C
  m_CRC32Table(177) = &HBC66831A
  m_CRC32Table(178) = &H256FD2A0
  m_CRC32Table(179) = &H5268E236
  m_CRC32Table(180) = &HCC0C7795
  m_CRC32Table(181) = &HBB0B4703
  m_CRC32Table(182) = &H220216B9
  m_CRC32Table(183) = &H5505262F
  m_CRC32Table(184) = &HC5BA3BBE
  m_CRC32Table(185) = &HB2BD0B28
  m_CRC32Table(186) = &H2BB45A92
  m_CRC32Table(187) = &H5CB36A04
  m_CRC32Table(188) = &HC2D7FFA7
  m_CRC32Table(189) = &HB5D0CF31
  m_CRC32Table(190) = &H2CD99E8B
  m_CRC32Table(191) = &H5BDEAE1D
  m_CRC32Table(192) = &H9B64C2B0
  m_CRC32Table(193) = &HEC63F226
  m_CRC32Table(194) = &H756AA39C
  m_CRC32Table(195) = &H26D930A
  m_CRC32Table(196) = &H9C0906A9
  m_CRC32Table(197) = &HEB0E363F
  m_CRC32Table(198) = &H72076785
  m_CRC32Table(199) = &H5005713
  m_CRC32Table(200) = &H95BF4A82
  m_CRC32Table(201) = &HE2B87A14
  m_CRC32Table(202) = &H7BB12BAE
  m_CRC32Table(203) = &HCB61B38
  m_CRC32Table(204) = &H92D28E9B
  m_CRC32Table(205) = &HE5D5BE0D
  m_CRC32Table(206) = &H7CDCEFB7
  m_CRC32Table(207) = &HBDBDF21
  m_CRC32Table(208) = &H86D3D2D4
  m_CRC32Table(209) = &HF1D4E242
  m_CRC32Table(210) = &H68DDB3F8
  m_CRC32Table(211) = &H1FDA836E
  m_CRC32Table(212) = &H81BE16CD
  m_CRC32Table(213) = &HF6B9265B
  m_CRC32Table(214) = &H6FB077E1
  m_CRC32Table(215) = &H18B74777
  m_CRC32Table(216) = &H88085AE6
  m_CRC32Table(217) = &HFF0F6A70
  m_CRC32Table(218) = &H66063BCA
  m_CRC32Table(219) = &H11010B5C
  m_CRC32Table(220) = &H8F659EFF
  m_CRC32Table(221) = &HF862AE69
  m_CRC32Table(222) = &H616BFFD3
  m_CRC32Table(223) = &H166CCF45
  m_CRC32Table(224) = &HA00AE278
  m_CRC32Table(225) = &HD70DD2EE
  m_CRC32Table(226) = &H4E048354
  m_CRC32Table(227) = &H3903B3C2
  m_CRC32Table(228) = &HA7672661
  m_CRC32Table(229) = &HD06016F7
  m_CRC32Table(230) = &H4969474D
  m_CRC32Table(231) = &H3E6E77DB
  m_CRC32Table(232) = &HAED16A4A
  m_CRC32Table(233) = &HD9D65ADC
  m_CRC32Table(234) = &H40DF0B66
  m_CRC32Table(235) = &H37D83BF0
  m_CRC32Table(236) = &HA9BCAE53
  m_CRC32Table(237) = &HDEBB9EC5
  m_CRC32Table(238) = &H47B2CF7F
  m_CRC32Table(239) = &H30B5FFE9
  m_CRC32Table(240) = &HBDBDF21C
  m_CRC32Table(241) = &HCABAC28A
  m_CRC32Table(242) = &H53B39330
  m_CRC32Table(243) = &H24B4A3A6
  m_CRC32Table(244) = &HBAD03605
  m_CRC32Table(245) = &HCDD70693
  m_CRC32Table(246) = &H54DE5729
  m_CRC32Table(247) = &H23D967BF
  m_CRC32Table(248) = &HB3667A2E
  m_CRC32Table(249) = &HC4614AB8
  m_CRC32Table(250) = &H5D681B02
  m_CRC32Table(251) = &H2A6F2B94
  m_CRC32Table(252) = &HB40BBE37
  m_CRC32Table(253) = &HC30C8EA1
  m_CRC32Table(254) = &H5A05DF1B
  m_CRC32Table(255) = &H2D02EF8D

  'Create a bytearray to hold the
  'precompiled assembler code
  sASM = "5589E557565053518B45088B008B750C8B7D108B4D1431DB8A1E30C3C1E80833049F464975F28B4D088901595B585E5F89EC5DC21000"
  ReDim m_CRC32Asm(0 To Len(sASM) \ 2 - 1)
  For i = 1 To Len(sASM) Step 2
    m_CRC32Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
  Next
  
  'Mark the CRC32 algorithm as initialized
  m_CRC32Init = True

End Sub

Private Sub Class_Initialize()

  'The default algorithm is CRC32
  Algorithm = CRC32
  
End Sub



'----
'VB6 - Huge (>2GB) File I/O Class-VBForums
'http://www.vbforums.com/showthread.php?531321-VB6-Huge-(-gt-2GB)-File-I-O-Class
'----
'-------------------------------------------------
'READ IN CHUNK FOR A WHOLE FILE BYTE ARRAY
'-------------------------------------------------
'AFTER ALL THE DOCUMENTAION
'IT IS NOT A READ CHUNK CODE
'---------------------------

'-------------------------------------------------
'BLOCK READ BYTE ARRAY 01 OF 07 -- THE DECLARE
'-------------------------------------------------

'Option Explicit
'
'Const MOVEFILE_REPLACE_EXISTING = &H1
'Const FILE_ATTRIBUTE_TEMPORARY = &H100
'Const FILE_BEGIN = 0
'Const FILE_SHARE_READ = &H1
'Const FILE_SHARE_WRITE = &H2
'Const CREATE_NEW = 1
'Const OPEN_EXISTING = 3
'Const OPEN_ALLWAYS = 4
'Const GENERIC_READ = &H80000000
'Const GENERIC_WRITE = &H40000000
'
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
'Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
'
'Public File_Num As Long
'Public File_Len As Currency

'--------------------------------------------------------------------------
'BLOCK READ BYTE ARRAY 02 OF 07 -- OPENING FILE ROUTINE
'THE RUDE ROUTINE
'--------------------------------------------------------------------------

Public Sub API_OpenFile(ByVal FileName As String, ByRef FileNumber As Long, ByRef FileSize As Currency)
Dim FileH As Long
Dim Ret As Long
On Error Resume Next
FileH = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_ALLWAYS, 0, 0)
If Err.Number > 0 Then
    Err.Clear
    FileNumber = -1
Else
    FileNumber = FileH
    Ret = SetFilePointer(FileH, 0, 0, FILE_BEGIN)
    API_FileSize FileH, FileSize
End If
On Error GoTo 0
End Sub

'------------------------------------------------------
'BLOCK READ BYTE ARRAY 03 OF 07 --
'------------------------------------------------------

Public Sub API_FileSize(ByVal FileNumber As Long, ByRef FileSize As Currency)
    Dim FileSizeL As Long
    Dim FileSizeH As Long
    FileSizeH = 0
    FileSizeL = GetFileSize(FileNumber, FileSizeH)
    Long2Size FileSizeL, FileSizeH, FileSize
End Sub

'------------------------------------------------------
'BLOCK READ BYTE ARRAY 04 OF 07
'------------------------------------------------------

Public Sub API_ReadFile(ByVal FileNumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef Data() As Byte)
Dim PosL As Long
Dim PosH As Long
Dim SizeRead As Long
Dim Ret As Long
Size2Long Position, PosL, PosH
Ret = SetFilePointer(FileNumber, PosL, PosH, FILE_BEGIN)
Ret = ReadFile(FileNumber, Data(0), BlockSize, SizeRead, 0&)



BlockSize = SizeRead
End Sub

'------------------------------------------------------
'BLOCK READ BYTE ARRAY 05 OF 07 --
'------------------------------------------------------

Public Sub API_CloseFile(ByVal FileNumber As Long)
Dim Ret As Long
Ret = CloseHandle(FileNumber)
End Sub

Public Sub API_WriteFile(ByVal FileNumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef Data() As Byte)
Dim PosL As Long
Dim PosH As Long
Dim SizeWrit As Long
Dim Ret As Long
Size2Long Position - 1, PosL, PosH
Ret = SetFilePointer(FileNumber, PosL, PosH, FILE_BEGIN)
Ret = WriteFile(FileNumber, Data(0), BlockSize, SizeWrit, 0&)
BlockSize = SizeWrit
End Sub

'------------------------------------------------------
'BLOCK READ BYTE ARRAY 0SIX OF 07 --
'------------------------------------------------------

Private Sub Size2Long(ByVal FileSize As Currency, ByRef LongLow As Long, ByRef LongHigh As Long)
'&HFFFFFFFF unsigned = 4294967295
Dim Cutoff As Currency
Cutoff = 2147483647
Cutoff = Cutoff + 2147483647
Cutoff = Cutoff + 1 ' now we hold the value of 4294967295 and not -1
LongHigh = 0
Do Until FileSize < Cutoff
    LongHigh = LongHigh + 1
    FileSize = FileSize - Cutoff
Loop
If FileSize > 2147483647 Then
    LongLow = -CLng(Cutoff - (FileSize - 1))
Else
    LongLow = CLng(FileSize)
End If
End Sub

'------------------------------------------------------
'BLOCK READ BYTE ARRAY 07 OF 07 --
'------------------------------------------------------

Private Sub Long2Size(ByVal LongLow As Long, ByVal LongHigh As Long, ByRef FileSize As Currency)
'&HFFFFFFFF unsigned = 4294967295
Dim Cutoff As Currency
Cutoff = 2147483647
Cutoff = Cutoff + 2147483647
Cutoff = Cutoff + 1 ' now we hold the value of 4294967295 and not -1
FileSize = Cutoff * LongHigh
If LongLow < 0 Then
    FileSize = FileSize + (Cutoff + (LongLow + 1))
Else
    FileSize = FileSize + LongLow
End If
End Sub

'-------------------------------------------------
'BLOCK READ BYTE ARRAY END
'-------------------------------------------------

'BUG IN HERE CODE
'BUG TO FIX IN ONE ABOVE
'------------------
'vbguy222 vbguy222 Is offline
'New Member
'
'Join Date
'Oct 2010
'Posts
'1
'Re: VB6 - Huge (>2GB) File I/O Class
'A quick note for anyone that might come through here and want to use the code posted above. The API_FileSize method incorrectly reports 0-byte file sizes as being 4294967295 bytes in length. To fix this, modify the last line in the method from
'Code:
'Long2Size FileSizeL, FileSizeH, FileSize
'to this
'
'Code:
'If FileSizeL = -1 Then
'   FileSize = 0
'Else
'   Long2Size FileSizeL, FileSizeH, FileSize
'End If
'Cheers.
'
'----------------



'----
'VB6 - Huge (>2GB) File I/O Class-VBForums
'http://www.vbforums.com/showthread.php?531321-VB6-Huge-(-gt-2GB)-File-I-O-Class
'----
'----
'NEXT PROJECT BLOCK
'Aug 4th, 2011, 08:06 PM #14
'Jonney Jonney Is offline
'Frenzied Member
'
'Join Date
'Jan 2010
'Posts
'1,030


'-------------------------------------------------
'BLOCK READ BYTE ARRAY 01 OF 02 -- THE DECLARE
'-------------------------------------------------
'
'
'Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
'Private Declare Function LCMapStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByVal lpDestStr As Long, ByVal cchDest As Long) As Long
'
'Private Declare Function CreateFile Lib "KERNEL32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function LCMapString Lib "KERNEL32.dll" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
'Private Declare Function lstrlen Lib "KERNEL32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
'
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'
'Private Const INVALID_HANDLE_VALUE = -1
'
'Private Const GENERIC_READ = &H80000000
'Private Const GENERIC_WRITE = &H40000000
'
'Private Const FILE_SHARE_READ = &H1
'Private Const FILE_SHARE_WRITE = &H2
'
'Private Const CREATE_NEW = 1
'Private Const CREATE_ALWAYS = 2
'Private Const OPEN_EXISTING = 3
'Private Const OPEN_ALWAYS = 4
'Private Const TRUNCATE_EXISTING = 5
'
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'
'Private Const FILE_BEGIN = 0
'Private Const FILE_CURRENT = 1
'Private Const FILE_END = 2

'Public Function ReadBytes(ByRef Buffer() As Byte) As Long
'
'   RaiseErrorIfClosed
'
'   If ReadFile(hFile, _
'      Buffer(LBound(Buffer)), _
'      UBound(Buffer) - LBound(Buffer) + 1, _
'      ReadBytes, _
'      0) Then
'
'      If ReadBytes = 0 Then
'         fEOF = True
'      ElseIf ReadBytes < UBound(Buffer) - LBound(Buffer) + 1 Then
'         ReDim Preserve Buffer(ReadBytes - 1)
'      End If
'   Else
'      RaiseError HBF_READ_FAILURE
'   End If
'End Function







