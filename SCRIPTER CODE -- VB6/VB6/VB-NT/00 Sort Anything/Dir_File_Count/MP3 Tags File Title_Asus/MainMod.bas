Attribute VB_Name = "MainMod"
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
    
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2

Public Power


Sub Jeepers()
    
Dim outInfo As MPEGInfo

ScanPath.List1.AddItem "Converting MP3 Tags..."
    ScanPath.List1.Refresh
    DoEvents

Dim InTag As ID3Tag

Dim Fs22

Set Fs22 = CreateObject("Scripting.FileSystemObject")


For we = 1 To ScanPath.ListView1.ListItems.Count

    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ScanPath.ListView1.ListItems.Item(we)

    Comment$ = ""
    If ReadID3v1(a1$ + b1$, InTag) = True Then
        Comment$ = "ID3v1 Previous Info " + Format$(Now, "DD-MM-YYYY HH:MM:SS") + vbCrLf
        If InTag.Title <> "" Then Comment$ = Comment$ + "Title-----" + vbCrLf + InTag.Title + vbCrLf
        If InTag.Artist <> "" Then Comment$ = Comment$ + "Artist----" + vbCrLf + InTag.Artist + vbCrLf
        If InTag.Album <> "" Then Comment$ = Comment$ + "Album-----" + vbCrLf + InTag.Album + vbCrLf
        If InTag.SongYear <> "" Then Comment$ = Comment$ + "SongYear--" + vbCrLf + InTag.SongYear + vbCrLf
        If InTag.Genre <> "" Then Comment$ = Comment$ + "Genre-----" + vbCrLf + InTag.Genre + vbCrLf
        If InTag.Comment <> "" Then Comment$ = Comment$ + "Comment----" + vbCrLf + InTag.Comment + vbCrLf
        If InTag.TrackNr <> "" Then Comment$ = Comment$ + "TrackNr----" + vbCrLf + InTag.TrackNr + vbCrLf
        wa = DeleteID3v1(a1$ + b1$)
    End If
        
    wa = ReadID3v2(a1$ + b1$, InTag)
    
    If InStr(InTag.Comment, "Matty Roid") > 0 And wa = True Then Skiproid = True Else Skiproid = False
    
    If InStr(InTag.Comment, "Matty Roid") = 0 Then
        If wa = False Then Comment$ = Comment$ + "ID3v2 " + Format$(Now, "DD-MM-YYYY HH:MM:SS") + vbCrLf
        If wa = True Then Comment$ = Comment$ + "ID3v2 Previous Info " + Format$(Now, "DD-MM-YYYY HH:MM:SS") + vbCrLf
        If InTag.Comment <> "" Then If wa = True Then Comment$ = Comment$ + "Comment----" + vbCrLf + InTag.Comment + vbCrLf
        If InTag.Title <> "" Then If wa = True Then Comment$ = Comment$ + "Title------" + vbCrLf + InTag.Title + vbCrLf
        Comment$ = Comment$ + "--- Matty Roid ---"
        InTag.Comment = Comment$
    End If
        
    InTag.Title = Mid$(b1$, 1, Len(b1$) - 4)
    

    mdlID3.ReadMPEGInfo a1$ + b1$, outInfo
        lblMPEG = outInfo.MPEGVersion
        lblLen = "Length: " & outInfo.Length \ 60 & ":" & Format$(outInfo.Length Mod 60, "00")
        lblkBit = "Bitrate: " & outInfo.Bitrate & " kbit / sec, " & IIf(outInfo.HasVBR, "variable bit rate", "constant bit rate")
        lblFreq = "Frequency: " & outInfo.Frequency & " hz, " & LCase$(outInfo.ChannelMode)
        lblCopy = "Copyrighted: " & IIf(outInfo.IsCopyrighted, "yes", "no")
        lblOrig = "Original: " & IIf(outInfo.IsOriginal, "yes", "no")
        lblCRC = "Uses checksums: " & IIf(outInfo.HasCRC, "yes", "no")
        lblEmphasis = "Uses emphasis: " & IIf(outInfo.HasEmphasis, "yes", "no")
        lblv1Vers = "ID3 v1 tag: " & IIf(outInfo.ID3v1Version = -1, "none", "version 1." & outInfo.ID3v1Version)
        lblv2Vers = "ID3 v2 tag: " & IIf(outInfo.ID3v2Version = -1, "none", "version 2." & outInfo.ID3v2Version)
    

    
    leng$ = Trim(Str$(outInfo.Length \ 60)) & ":" & Format$(outInfo.Length Mod 60, "00") + " Mins Lenght."
    kbps$ = """;""" + Trim(Str$(outInfo.Bitrate)) + " Kbps"
    c1$ = """" + b1$ + a1$
        
    
    InTag.Artist = leng$
    wa = WriteID3v2(a1$ + b1$, InTag)
    
    DoEvents
    
    If Skiproid = True Then c1$ = " --- Skip Comment" Else c1$ = ""
    
    ScanPath.List1.AddItem Format$(we, "000 ") + a1$ + b1$ + c1$
    ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
    ScanPath.List1.Refresh
    DoEvents

If Power = True Then Exit For

Next
    
If Power = True Then ScanPath.List1.AddItem "Aborted......."
ScanPath.List1.AddItem "Completed......."
ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
ScanPath.List1.Refresh


'Timer1.Enabled = True

End Sub

