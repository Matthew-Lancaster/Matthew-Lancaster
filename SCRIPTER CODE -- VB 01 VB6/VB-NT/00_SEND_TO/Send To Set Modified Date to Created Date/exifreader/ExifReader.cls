VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExifReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private ExifTemp() As Byte
Private Offset_to_IFD0 As Long
Private Offset_to_APP1 As Long
Private Offset_to_TIFF As Long
Private Length_of_APP1 As Long
Private Offset_to_Next_IFD As Long
Private IFDDirectory() As IFD_Data
Private Offset_to_ExifSubIFD As Long
Private m_Tag As Long
Private m_file As String
Private IsLoaded As Boolean

Private Enum EXIF_DATA_FORMAT
    m_BYTE = 1
    m_STRING = 2
    m_SHORT = 3
    m_LONG = 4
    m_RATIONAL = 5
    m_SBYTE = 6
    m_UNDEFINED = 7
    m_SSHORT = 8
    m_SLONG = 9
    m_SRATIONAL = 10
    m_SINGLE = 11
    m_DOUBLE = 12
End Enum

Private Type IFD_Data_Values
    BytVal As Byte
    StrVal As String
    IntVal As Integer
    LngVal As Long
    SngVal As Single
    DblVal As Double
End Type
 
Private Type IFD_Data
    Tag_No As EXIF_TAG
    MakerNote As Boolean
    Data_Format As EXIF_DATA_FORMAT
    Components As Long
    Offset_To_Value As Long
    Value As Variant
End Type

Public Enum EXIF_TAG
    'IFD0 Tags
        ImageDescription = &H10E&
        Make = &H10F&
        Model = &H110&
        Orientation = &H112&
        XResolution = &H11A&
        YResolution = &H11B&
        ResolutionUnit = &H128&
        Software = &H131&
        DateTime = &H132&
        WhitePoint = &H13E&
        PrimaryChromaticities = &H13F&
        YCbCrCoefficients = &H211&
        YCbCrPositioning = &H213&
        ReferenceBlackWhite = &H214&
        Copyright = &H8298&
        ExifOffset = &H8769&
    'ExifSubIFD Tags
        ExposureTime = &H829A&
        FNumber = &H829D&
        ExposureProgram = &H8822&
        ISOSpeedRatings = &H8827&
        ExifVersion = &H9000&
        DateTimeOriginal = &H9003&
        DateTimeDigitized = &H9004&
        ComponentsConfiguration = &H9101&
        CompressedBitsPerPixel = &H9102&
        ShutterSpeedValue = &H9201&
        ApertureValue = &H9202&
        BrightnessValue = &H9203&
        ExposureBiasValue = &H9204&
        MaxApertureValue = &H9205&
        SubjectDistance = &H9206&
        MeteringMode = &H9207&
        LightSource = &H9208&
        Flash = &H9209&
        FocalLength = &H920A&
        MakerNote = &H927C&
        UserComment = &H9286&
        SubsecTime = &H9290&
        SubsecTimeOriginal = &H9291&
        SubsecTimeDigitized = &H9292&
        FlashPixVersion = &HA000&
        ColorSpace = &HA001&
        ExifImageWidth = &HA002&
        ExifImageHeight = &HA003&
        RelatedSoundFile = &HA004&
        ExifInteroperabilityOffset = &HA005&
        FocalPlaneXResolution = &HA20E&
        FocalPlaneYResolution = &HA20F&
        FocalPlaneResolutionUnit = &HA210&
        ExposureIndex = &HA215&
        SensingMethod = &HA217&
        FileSource = &HA300&
        SceneType = &HA301&
        CFAPattern = &HA302&
    'Interoperability IFD Tags
        InteroperabilityIndex = &H1&
        InteroperabilityVersion = &H2&
        RelatedImageFileFormat = &H1000&
        RelatedImageWidth = &H1001&
        RelatedImageLength = &H1002&
    'IFD1 Tags
        ImageWidth = &H100&
        ImageHeight = &H101&
        BitsPerSample = &H102&
        Compression = &H103&
        PhotometricInterpretation = &H106&
        StripOffsets = &H111&
        SamplePerPixel = &H115&
        RowsPerStrip = &H116&
        StripByteCounts = &H117&
        XResolution2 = &H11A&
        YResolution2 = &H11B&
        PlanarConfiguration = &H11C&
        ResolutionUnit2 = &H128&
        JPEGInterchangeFormat = &H201&
        JPEGInterchangeFormatLength = &H202&
        YCbCrCoeffecients = &H211&
        YCbCrSubSampling = &H212&
        YCbCrPositioning2 = &H213&
        ReferenceBlackWhite2 = &H214&
    'Misc Tags
        NewSubfileType = &HFE&
        SubfileType = &HFF&
        TransferFunction = &H12D&
        Artist = &H13B&
        Predictor = &H13D&
        TileWidth = &H142&
        TileLength = &H143&
        TileOffsets = &H144&
        TileByteCounts = &H145&
        SubIFDs = &H14A&
        JPEGTables = &H15B&
        CFARepeatPatternDim = &H828D&
        CFAPattern2 = &H828E&
        BatteryLevel = &H828F&
        IPTC_NAA = &H83BB&
        InterColorProfile = &H8773&
        SpectralSensitivity = &H8824&
        GPSInfo = &H8825&
        OECF = &H8828&
        Interlace = &H8829&
        TimeZoneOffset = &H882A&
        SelfTimerMode = &H882B&
        FlashEnergy = &H920B&
        SpatialFrequencyResponse = &H920C&
        Noise = &H920D&
        ImageNumber = &H9211&
        SecurityClassification = &H9212&
        ImageHistory = &H9213&
        SubjectLocation = &H9214&
        ExposureIndex2 = &H9215&
        TIFFEPStandardID = &H9216&
        FlashEnergy2 = &HA20B&
        SpatialFrequencyResponse2 = &HA20C&
        SubjectLocation2 = &HA214&
End Enum

Public Property Let picFile(picFile As String)
    m_file = picFile
End Property

Public Property Get MakerNoteTag(Optional ByVal MakerTag As Long) As Variant
    If IsLoaded = False Then Exit Property
    
    Dim i As Long
    
    For i = 1 To UBound(IFDDirectory)
        If IFDDirectory(i).Tag_No = MakerTag And IFDDirectory(i).MakerNote = True Then
            MakerNoteTag = IFDDirectory(i).Value
            Exit For
        End If
    Next
    
End Property

Public Property Get Tag(Optional ByVal ExifTag As EXIF_TAG) As Variant
    If IsLoaded = False And m_file <> "" Then
        Load (m_file)
    ElseIf IsLoaded = False And m_file = "" Then
        Exit Property
    End If
    
    If ExifTag = 0 Then
        On Error Resume Next
        Tag = UBound(IFDDirectory)
        On Error GoTo 0
        Exit Property
    End If
    
    Dim i As Long
    
    For i = 1 To UBound(IFDDirectory)
        If IFDDirectory(i).Tag_No = ExifTag Then
            Tag = IFDDirectory(i).Value
            Exit For
        End If
    Next

End Property

Public Sub Load(Optional ByVal picFile As String)
    If m_file = "" Then
        m_file = picFile
        If m_file = "" Then
            Exit Sub
        End If
    End If
    
    
    
    OpenJPGFile m_file
    If InspectJPGFile = False Then
        IsLoaded = False
        Exit Sub
    End If
    
    If IsIntel Then
        Offset_to_IFD0 = _
            ExifTemp(Offset_to_APP1 + 17) * 256& * 256& * 256& + _
            ExifTemp(Offset_to_APP1 + 16) * 256& * 256& + _
            ExifTemp(Offset_to_APP1 + 15) * 256& + _
            ExifTemp(Offset_to_APP1 + 14)
    Else
        Offset_to_IFD0 = _
            ExifTemp(Offset_to_APP1 + 14) * 256& * 256& * 256& + _
            ExifTemp(Offset_to_APP1 + 15) * 256& * 256& + _
            ExifTemp(Offset_to_APP1 + 16) * 256& + _
            ExifTemp(Offset_to_APP1 + 17)
    End If
    
    'Debug.Print "Offset_to_IFD0: " & Offset_to_IFD0
    IsLoaded = True
    GetDirectoryEntries Offset_to_TIFF + Offset_to_IFD0
    
End Sub

Private Function OpenJPGFile(ByVal inFile As String)

    Dim fFile As Integer
    
    fFile = FreeFile

    Open inFile For Binary As #fFile
        ReDim ExifTemp(LOF(fFile)) As Byte
        Get #fFile, , ExifTemp
    Close #fFile
        
End Function

Private Function InspectJPGFile() As Boolean
    
    Dim i As Long
    
    If ExifTemp(0) <> &HFF And ExifTemp(1) <> &HD8 Then
        InspectJPGFile = False
    Else
    
        For i = 2 To UBound(ExifTemp) - 1
            If ExifTemp(i) = &HFF And ExifTemp(i + 1) = &HE1 Then
                Offset_to_APP1 = i
                Exit For
            End If
        Next
        
        If Offset_to_APP1 = 0 Then
            InspectJPGFile = False
        End If
        
        Offset_to_TIFF = Offset_to_APP1 + 10
        
        Length_of_APP1 = _
            ExifTemp(Offset_to_APP1 + 2) * 256& + _
            ExifTemp(Offset_to_APP1 + 3)
        
        If Chr(ExifTemp(Offset_to_APP1 + 4)) & Chr(ExifTemp(Offset_to_APP1 + 5)) & _
            Chr(ExifTemp(Offset_to_APP1 + 6)) & Chr(ExifTemp(Offset_to_APP1 + 7)) <> "Exif" Then
            InspectJPGFile = False
            Exit Function
        End If
        
        InspectJPGFile = True
        
    End If
    
End Function

Private Function IsIntel() As Boolean

    If Hex(ExifTemp(Offset_to_TIFF)) = "49" Then
        IsIntel = True
    Else
        IsIntel = False
    End If
    
End Function

Private Function sPad(sSrc As String, nSize As Integer, sPadChar As String) As String

    If Len(sSrc) < nSize Then
        sPad = String(nSize - Len(sSrc), sPadChar) & sSrc
    Else
        sPad = sSrc
    End If
    
End Function

Private Sub GetDirectoryEntries(Offset As Long)
    
    Dim No_of_Entries As Long
    Dim Upper_IFDDirectory As Long
    Dim NewDimensions As Long
    Dim Processed_ExifSubIFD As Boolean
    Dim BytesPerComponent As Long
    Dim Offset_to_MakerNote As Long
    Dim i As Long, j As Long
    
    Do
        If IsIntel Then
            No_of_Entries = _
                ExifTemp(Offset + 1) * 256& + _
                ExifTemp(Offset + 0)
        Else
            No_of_Entries = _
                ExifTemp(Offset + 0) * 256& + _
                ExifTemp(Offset + 1)
        End If

        On Error Resume Next
        Upper_IFDDirectory = UBound(IFDDirectory)
        On Error GoTo 0
        
        NewDimensions = Upper_IFDDirectory + No_of_Entries
        
        ReDim Preserve IFDDirectory(1 To NewDimensions) As IFD_Data
        
        For i = 1 To No_of_Entries
        
            With IFDDirectory(Upper_IFDDirectory + i)
            
                If IsIntel Then

                    .Tag_No = _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 1) * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 0)
                    
                    .Data_Format = _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 3) * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 2)
                        
                    .Components = _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 7) * 256& * 256& * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 6) * 256& * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 5) * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 4)
                    
                    Select Case .Data_Format
                    
                        Case m_BYTE, m_SBYTE
                            BytesPerComponent = 1
                            If .Components * BytesPerComponent <= 4 Then
                                .Value = _
                                    sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)), 2, "0") & _
                                    sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 10)), 2, "0") & _
                                    sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 9)), 2, "0") & _
                                    sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)), 2, "0")
                            Else
                                .Offset_To_Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                                For j = 0 To .Components - 1
                                    .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                                Next
                            End If
                            
                        Case m_STRING, m_UNDEFINED
                            BytesPerComponent = 1
                            If .Components * BytesPerComponent <= 4 Then
                                .Value = _
                                    Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)) & _
                                    Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 10)) & _
                                    Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 9)) & _
                                    Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 8))
                            Else
                                .Offset_To_Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                                For j = 0 To .Components - 2
                                    .Value = .Value & Chr(ExifTemp(Offset_to_TIFF + .Offset_To_Value + j))
                                Next
                            End If
                            
                        Case m_SHORT, m_SSHORT
                            BytesPerComponent = 2
                            If .Components * BytesPerComponent <= 4 Then
                                .Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                                '.Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                            Else
                                .Offset_To_Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                                For j = 0 To .Components - 1
                                    .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                                Next
                            End If
                            
                        Case m_LONG, m_SLONG
                            BytesPerComponent = 4
                            If .Components * BytesPerComponent <= 4 Then
                                .Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                            Else
                                .Offset_To_Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                                For j = 0 To .Components - 1
                                    .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                                Next
                            End If
                            
                        Case m_RATIONAL, m_SRATIONAL
                            BytesPerComponent = 8
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                            .Value = _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 3) * 256& * 256& * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 2) * 256& * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 1) * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 0) & _
                                "/" & _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 7) * 256& * 256& * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 6) * 256& * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 5) * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 4)
                            
                    End Select
                    
                Else
                
                   .Tag_No = _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 0) * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 1)
                        
                    .Data_Format = _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 2) * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 3)
                        
                    .Components = _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 4) * 256& * 256& * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 5) * 256& * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 6) * 256& + _
                        ExifTemp((Offset + 2) + ((i - 1) * 12) + 7)
                        
                    Select Case .Data_Format
                    
                        Case m_BYTE, m_SBYTE
                            BytesPerComponent = 1
                            If .Components * BytesPerComponent <= 4 Then
                                .Value = _
                                    sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)), 2, "0") & _
                                    sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 9)), 2, "0") & _
                                    sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 10)), 2, "0") & _
                                    sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)), 2, "0")
                            Else
                                .Offset_To_Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                                For j = 0 To .Components - 1
                                    .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                                Next
                            End If
                            
                        Case m_STRING, m_UNDEFINED
                            BytesPerComponent = 1
                            If .Components * BytesPerComponent <= 4 Then
                                .Value = _
                                    Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)) & _
                                    Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 9)) & _
                                    Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 10)) & _
                                    Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 11))
                            Else
                                .Offset_To_Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                                For j = 0 To .Components - 1
                                    .Value = .Value & Chr(ExifTemp(Offset_to_TIFF + .Offset_To_Value + j))
                                Next
                            End If
                            
                        Case m_SHORT, m_SSHORT
                            BytesPerComponent = 2
                            If .Components * BytesPerComponent <= 4 Then
                                .Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                            Else
                                .Offset_To_Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                                For j = .Components - 1 To 0 Step -1
                                    .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                                Next
                            End If
                            
                        Case m_LONG, m_SLONG
                            BytesPerComponent = 4
                            If .Components * BytesPerComponent <= 4 Then
                                .Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                            Else
                                .Offset_To_Value = _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                                For j = 0 To .Components - 1
                                    .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                                Next
                            End If
                            
                        Case m_RATIONAL, m_SRATIONAL
                            BytesPerComponent = 8
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                            .Value = _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 0) * 256& * 256& * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 1) * 256& * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 2) * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 3) & _
                                "/" & _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 4) * 256& * 256& * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 5) * 256& * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 6) * 256& + _
                                ExifTemp(Offset_to_TIFF + .Offset_To_Value + 7)
                            
                    End Select
                    
                End If
                
                'Debug.Print Upper_IFDDirectory + i & ".Tag_No: " & .Tag_No & _
                    "; .Data_Format: " & .Data_Format & _
                    "; .Components: " & .Components & _
                    "; .Offset_To_Value: " & .Offset_To_Value & _
                    "; .Value: " & .Value
                If .Tag_No = MakerNote Then
                    Offset_to_MakerNote = .Offset_To_Value
                End If
                If .Tag_No = ExifOffset Then
                    Offset_to_ExifSubIFD = CLng(.Value)
                    'Debug.Print "Offset_to_ExifSubIFD: " & Offset_to_ExifSubIFD
                End If
                
            End With
            
        Next
        
        If IsIntel Then
            If Not Processed_ExifSubIFD Then
                Offset_to_Next_IFD = _
                    ExifTemp(Offset + 2 + (No_of_Entries * 12) + 3) * 256& * 256& * 256& + _
                    ExifTemp(Offset + 2 + (No_of_Entries * 12) + 2) * 256& * 256& + _
                    ExifTemp(Offset + 2 + (No_of_Entries * 12) + 1) * 256& + _
                    ExifTemp(Offset + 2 + (No_of_Entries * 12) + 0)
                'Debug.Print "Offset_to_Next_IFD: " & Offset_to_Next_IFD
            Else
                Offset_to_Next_IFD = 0
            End If
        Else
            If Not Processed_ExifSubIFD Then
                Offset_to_Next_IFD = _
                    ExifTemp(Offset + 2 + (No_of_Entries * 12) + 0) * 256& * 256& * 256& + _
                    ExifTemp(Offset + 2 + (No_of_Entries * 12) + 1) * 256& * 256& + _
                    ExifTemp(Offset + 2 + (No_of_Entries * 12) + 2) * 256& + _
                    ExifTemp(Offset + 2 + (No_of_Entries * 12) + 3)
                'Debug.Print "Offset_to_Next_IFD: " & Offset_to_Next_IFD
            Else
                Offset_to_Next_IFD = 0
            End If
        End If
        
        If Offset_to_Next_IFD = 0 And Processed_ExifSubIFD = False Then
            Offset_to_Next_IFD = Offset_to_ExifSubIFD
            Processed_ExifSubIFD = True
        End If
        
        Offset = Offset_to_TIFF + Offset_to_Next_IFD
 
    Loop While Offset_to_Next_IFD <> 0
    
    'If Offset_to_MakerNote <> 0 Then
        'ProcessMakerNote Offset_to_MakerNote + Offset_to_TIFF
    'End If
    
End Sub

Private Sub ProcessMakerNote(ByVal Offset As Long)
    Dim No_of_Entries As Long
    Dim Upper_IFDDirectory As Long
    Dim NewDimensions As Long
    Dim BytesPerComponent As Long
    Dim i As Long, j As Long
    
    If IsIntel Then
        No_of_Entries = _
            ExifTemp(Offset + 1) * 256& + _
            ExifTemp(Offset + 0)
    Else
        No_of_Entries = _
            ExifTemp(Offset + 0) * 256& + _
            ExifTemp(Offset + 1)
    End If

    On Error Resume Next
    Upper_IFDDirectory = UBound(IFDDirectory)
    On Error GoTo 0
    
    NewDimensions = Upper_IFDDirectory + No_of_Entries
    
    ReDim Preserve IFDDirectory(1 To NewDimensions) As IFD_Data
    
    For i = 1 To No_of_Entries
    
        With IFDDirectory(Upper_IFDDirectory + i)
        
            If IsIntel Then

                .Tag_No = _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 1) * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 0)
                    
                .MakerNote = True
                
                .Data_Format = _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 3) * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 2)
                    
                .Components = _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 7) * 256& * 256& * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 6) * 256& * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 5) * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 4)
                
                Select Case .Data_Format
                
                    Case m_BYTE, m_SBYTE
                        BytesPerComponent = 1
                        If .Components * BytesPerComponent <= 4 Then
                            .Value = _
                                sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)), 2, "0") & _
                                sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 10)), 2, "0") & _
                                sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 9)), 2, "0") & _
                                sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)), 2, "0")
                        Else
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                            For j = 0 To .Components - 1
                                .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                            Next
                        End If
                        
                    Case m_STRING, m_UNDEFINED
                        BytesPerComponent = 1
                        If .Components * BytesPerComponent <= 4 Then
                            .Value = _
                                Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)) & _
                                Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 10)) & _
                                Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 9)) & _
                                Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 8))
                        Else
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                            For j = 0 To .Components - 2
                                .Value = .Value & Chr(ExifTemp(Offset_to_TIFF + .Offset_To_Value + j))
                            Next
                        End If
                        
                    Case m_SHORT, m_SSHORT
                        BytesPerComponent = 2
                        If .Components * BytesPerComponent <= 4 Then
                            .Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                            '.Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                        Else
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                            For j = 0 To .Components - 1
                                .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                            Next
                        End If
                        
                    Case m_LONG, m_SLONG
                        BytesPerComponent = 4
                        If .Components * BytesPerComponent <= 4 Then
                            .Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                        Else
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                            For j = 0 To .Components - 1
                                .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                            Next
                        End If
                        
                    Case m_RATIONAL, m_SRATIONAL
                        BytesPerComponent = 8
                        .Offset_To_Value = _
                            ExifTemp((Offset + 2) + ((i - 1) * 12) + 11) * 256& * 256& * 256& + _
                            ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& * 256& + _
                            ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& + _
                            ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)
                        .Value = _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 3) * 256& * 256& * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 2) * 256& * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 1) * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 0) & _
                            "/" & _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 7) * 256& * 256& * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 6) * 256& * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 5) * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 4)
                        
                End Select
                
            Else
            
               .Tag_No = _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 0) * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 1)
                    
                .Data_Format = _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 2) * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 3)
                    
                .Components = _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 4) * 256& * 256& * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 5) * 256& * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 6) * 256& + _
                    ExifTemp((Offset + 2) + ((i - 1) * 12) + 7)
                    
                Select Case .Data_Format
                
                    Case m_BYTE, m_SBYTE
                        BytesPerComponent = 1
                        If .Components * BytesPerComponent <= 4 Then
                            .Value = _
                                sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)), 2, "0") & _
                                sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 9)), 2, "0") & _
                                sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 10)), 2, "0") & _
                                sPad(Hex(ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)), 2, "0")
                        Else
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                            For j = 0 To .Components - 1
                                .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                            Next
                        End If
                        
                    Case m_STRING, m_UNDEFINED
                        BytesPerComponent = 1
                        If .Components * BytesPerComponent <= 4 Then
                            .Value = _
                                Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 8)) & _
                                Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 9)) & _
                                Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 10)) & _
                                Chr(ExifTemp((Offset + 2) + ((i - 1) * 12) + 11))
                        Else
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                            For j = 0 To .Components - 1
                                .Value = .Value & Chr(ExifTemp(Offset_to_TIFF + .Offset_To_Value + j))
                            Next
                        End If
                        
                    Case m_SHORT, m_SSHORT
                        BytesPerComponent = 2
                        If .Components * BytesPerComponent <= 4 Then
                            .Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                        Else
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                            For j = .Components - 1 To 0 Step -1
                                .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                            Next
                        End If
                        
                    Case m_LONG, m_SLONG
                        BytesPerComponent = 4
                        If .Components * BytesPerComponent <= 4 Then
                            .Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                        Else
                            .Offset_To_Value = _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                                ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                            For j = 0 To .Components - 1
                                .Value = .Value & ExifTemp(Offset_to_TIFF + .Offset_To_Value + j)
                            Next
                        End If
                        
                    Case m_RATIONAL, m_SRATIONAL
                        BytesPerComponent = 8
                        .Offset_To_Value = _
                            ExifTemp((Offset + 2) + ((i - 1) * 12) + 8) * 256& * 256& * 256& + _
                            ExifTemp((Offset + 2) + ((i - 1) * 12) + 9) * 256& * 256& + _
                            ExifTemp((Offset + 2) + ((i - 1) * 12) + 10) * 256& + _
                            ExifTemp((Offset + 2) + ((i - 1) * 12) + 11)
                        .Value = _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 0) * 256& * 256& * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 1) * 256& * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 2) * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 3) & _
                            "/" & _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 4) * 256& * 256& * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 5) * 256& * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 6) * 256& + _
                            ExifTemp(Offset_to_TIFF + .Offset_To_Value + 7)
                        
                End Select
                
            End If
            
        End With
        
    Next
        
End Sub

