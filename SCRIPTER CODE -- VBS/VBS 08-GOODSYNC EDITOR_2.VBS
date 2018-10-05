Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1, ForWriting = 2, ForAppending = 8

'--------------------------------------------------------
'AFTER THE EFFORT 
'VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
'IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
'HERE VBS DOES SINGLE 
'--------------------------------------------------------

Set File_1 = fso.OpenTextFile("C:\Users\MATT 01\AppData\Roaming\GoodSync\jobs-groups-options - Copy.tic", ForReading)
TT_EDITOR = File_1.ReadAll

X = 1
Do
    T1 = InStr(X, TT_EDITOR, "|M:0")
    
    If T1 > 0 Then
        
        T2 = InStr(T1, TT_EDITOR, "|i=")
        N3 = Mid(TT_EDITOR, T1, (T2 - T1) + Len("|i="))
        
        REASSEMBLE_VAR = "|M:0:|U=0|B:0:|J:0:|i="
        TT_EDITOR = Replace(TT_EDITOR, N3, REASSEMBLE_VAR)
        
        Counter_VAR = Counter_VAR + 1
    End If
    
    X = T1 + 4
Loop Until T1 = 0

Set File_2 = fso.OpenTextFile("C:\Users\MATT 01\AppData\Roaming\GoodSync\jobs-groups-options.tic", ForWriting)

File_2.Write TT_EDITOR

MsgBox "Done Total " & Counter_VAR & " Proecesses"


' Const vbFromUnicode = 128
' Const vbUnicode = 64
' Const adTypeBinary = 1
' Const adTypeText = 2
' Const adLongVarBinary = 205

Public Function StrConv(ByRef stringData, ByRef conversion)
    Dim Stream
    Set Stream = CreateObject("ADODB.Stream")        
	Const UnicodeCharaset = "Big5"
    Const BinaryCharset = "X-ANSI"
    Select Case conversion
    Case vbFromUnicode
        ' Converts a Unicode String To Ascii
        With Stream
            .Charset = UnicodeCharaset
            .Type = adTypeText
            .Open
            .WriteText stringData
            .Position = 0
            .Charset = BinaryCharset
            .Type = adTypeBinary
            StrConv = MidB(.Read, 1)
        End With
        
    Case vbUnicode
        ' Converts an Ascii String To Unicode
        Dim Length
        Dim Buffer
        
        If TypeName(stringData) = "Null" Then
            CStrU = ""
            Exit Function
        End If
        
        stringData = MidB(stringData, 1)
    
        Length = LenB(stringData)
        Dim Rs
        Set Rs = CreateObject("ADODB.Recordset")
        Call Rs.Fields.Append("BinaryData", adLongVarBinary, Length)
        Rs.Open
        Rs.AddNew
        Rs.Fields("BinaryData").AppendChunk (stringData & ChrB(0))
        Rs.Update
        Buffer = Rs.Fields("BinaryData").GetChunk(Length)
        Rs.Close
        Set Rs = Nothing
        With Stream
            .Charset = BinaryCharset
            .Type = adTypeBinary
            .Open
            Call .Write(Buffer)
            .Position = 0
            .Type = adTypeText
            .Charset = UnicodeCharaset
        End With
        
        StrConv = Stream.ReadText(-1)
    End Select
    Stream.Close
    Set Stream = Nothing
End Function
' ----
' StrConv Unicode / Ascii for VBScript - Delphi K.Top 討論區
' http://delphi.ktop.com.tw/board.php?cid=31&fid=77&tid=66001
' ----
' Fri 22 December 2017 20:12:54----------