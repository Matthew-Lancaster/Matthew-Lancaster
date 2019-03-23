VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MP3 Tag Info"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7605
   Icon            =   "MP3 Tags.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1935
      Top             =   3555
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   7605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
'Windows API/Global Declarations for :Dr
'     ive Type Finder
'**************************************






Private Sub Form_Activate()
'Load Mp3_Tag

Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")


Dim strDrive As String
Dim strMessage As String
Dim intCnt As Integer


For intCnt = 90 To 67 Step -1
    strDrive = Chr(intCnt)


    Select Case GetDriveType(strDrive + ":\")
        Case DRIVE_REMOVABLE
        rtn = "Floppy Drive"
        Case DRIVE_FIXED
        rtn = "Hard Drive"
        Case DRIVE_REMOTE
        rtn = "Network Drive"
        Case DRIVE_CDROM
        rtn = "CD-ROM Drive"
        Case DRIVE_RAMDISK
        rtn = "RAM Disk"
        Case Else
        rtn = ""
    End Select

    If rtn = "CD-ROM Drive" Then
    On Local Error Resume Next
    e$ = Dir$(strDrive + ":\*.*")
    Checkeddrv = Checkeddrv + 1
    If Err.Number > 0 And Checkeddrv = 2 Then MsgBox ("No Disk in Last Cd Rom Drive " + strDrive + ":\"): End
    If Err.Number = 0 Then Exit For
    End If

'    If rtn <> "" Then
'       strMessage = strMessage & vbCrLf & "Drive " & strDrive & " is type: " & rtn
'    End If
Next intCnt





Load ScanPath
ScanPath.txtPath = strDrive + ":\"
'ScanPath.txtPath = "E:\My Music\Talking Books"
ScanPath.cboMask.Text = "*.*"

Call ScanPath.cmdScan_Click


books = Val(ScanPath.lblCount.Caption)
List1.AddItem ("BookMarks = " + Str$(books))

freef1 = FreeFile
Open "C:\# MY DOCS\CD-ROM\Kmp3.txt" For Output As #freef1

If Dir(strDrive + ":\Z_CallerId\*.*") <> "" Then alps$ = "Z_CallerId": kil = 1
If Dir(strDrive + ":\zzCallerId\*.*") <> "" Then alps$ = "zzCallerId": kil = 2
If Dir(strDrive + ":\0Rm35\*.*") <> "" Then alps$ = "Z_CallerId": kil = 3
If Dir(strDrive + ":\0Rm34\*.*") <> "" Then alps$ = "Z_CallerId": kil = 3
If Dir(strDrive + ":\0Rm33\*.*") <> "" Then alps$ = "Z_CallerId": kil = 3

'len(alpS$)

For we = 1 To ScanPath.ListView1.ListItems.Count
    a1$ = ScanPath.ListView1.ListItems.Item(we)
    b1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)


    'c1$ = Mid$(b1$, Len(ScanPath.txtPath) + 2) + a1$

'    Call Mp3_Tag.smdpick(b1$ + a1$)
    Dim outInfo As MPEGInfo

    mp32 = 0
    If InStr(UCase$(a1$), ".MP3") Then mp32 = 1
    
    outInfo.Length = 0
    outInfo.Bitrate = 0
    
    
    If mp32 = 1 Then
        mdlID3.ReadMPEGInfo b1$ + a1$, outInfo
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
    End If

    
    leng$ = """;""" + Trim(Str$(outInfo.Length \ 60)) & ":" & Format$(outInfo.Length Mod 60, "00") + """"
    kbps$ = """;""" + Trim(Str$(outInfo.Bitrate)) + " Kbps"
    c1$ = """" + b1$ + a1$
    
    
    xyz = 0
    If InStr(UCase$(a1$), ".MP3") Then xyz = 1
    
    If xyz = 0 And InStr(b1$, alps$ + "\Apps") And kil = 1 Then
        
        qwd = Len(alps$ + "\Apps") + 5
        d1$ = Mid$(b1$, 1, 3) + "Applications-\" + Mid$(b1$, qwd, InStr(qwd, b1$, "\") - qwd)
        

        'If lastpath$ = "" Then lastpath$ = d1$
        If lastpath$ <> d1$ Then
            leng$ = """;""" + """"
            Set f = fs.GetFolder(b1$)
            s22 = f.Size
            kbps$ = """;""" + Format$(s22 / (1048576), "0.0") + " MByte"
            c1$ = """" + d1$
            xyz = 1
            lastpath$ = d1$
        End If
    End If
    
    If xyz = 0 And InStr(b1$, alps$) And (kil = 2 Or kil = 3) Then
        
        qwd = Len(alps$) + 3
        If kil = 3 Then qwd = qwd + 1
        
        If InStr(qwd + 2, b1$, "\") Then
        d1$ = Mid$(b1$, 1, 3) + "Applications-" + Mid$(b1$, qwd, InStr(qwd + 2, b1$, "\") - qwd)
        End If

        'If lastpath$ = "" Then lastpath$ = d1$
        kit = 0
        If InStr(b1$, "Andy Lett") Then kit = 1
        If InStr(b1$, "My Web") Then kit = 1
        
        If lastpath$ <> d1$ And kit = 0 Then
            leng$ = """;""" + """"
            Set f = fs.GetFolder(b1$)
            s22 = f.Size
            kbps$ = """;""" + Format$(s22 / (1048576), "0.0") + " MByte"
            c1$ = """" + d1$
            xyz = 1
            lastpath$ = d1$
        End If
    End If
    
    
    
    
    
    
    
    DoEvents
    
    If InStr(b1$, Mid$(b1$, 1, 3) + alps$ + "\My Web\") Then webc = webc + 1
    Qe2$ = UCase$(Mid$(c1$, Len(c1$) - 10))
    
    
    kj = 0
    If InStrRev(Qe2$, ".EXE") Then exec = exec + 1: kj = 1
    If InStrRev(Qe2$, ".MP3") Then mp3c = mp3c + 1: kj = 1
    If InStrRev(Qe2$, ".TXT") Then txtc = txtc + 1: kj = 1
    If InStrRev(Qe2$, ".JPG") Then jpgc = jpgc + 1: kj = 1
    If InStrRev(Qe2$, ".html") Then htmc = htmc + 1: kj = 1
    If InStrRev(Qe2$, ".html") Then htmc = htmc + 1: kj = 1
    
    qe3$ = UCase$(Mid$(c1$, Len(c1$) - 3))
    wg = InStr(qe4$, qe3$)
    If wg = 0 And kj = 0 Then qe4$ = qe4$ + qe3$
    
    
    If xyz = 1 And Not (outInfo.Length = 0 And mp32 = 1) Then
        List1.AddItem (c1$ + kbps$ + leng$)
        List1.ListIndex = List1.ListCount - 1
        Print #freef1, c1$ + kbps$ + leng$
    End If
Next

others$ = Trim$(Str$(Len(qe4$) / 4))
Set f = fs.GetFolder(Mid$(b1$, 1, 3) + alps$ + "\My Web\")
s22 = f.Size / 1048576
wew$ = Format$(s22, "0.0") + " MByte"
wewc$ = Trim$(Str(webc)) + " File Count"

List1.AddItem """Total Taken By Matt's Web http://MatthewLan.com = " + wewc$ + """;""" + wew$ + """;""" + """"
Print #freef1, """Total Taken By Matt's Web http://MatthewLan.com = " + wewc$ + """;""" + wew$ + """;""" + """"

List1.AddItem """Total File's On Disk =" + Str$(books) + """;""" + """;""" + """"
List1.AddItem """Total File's *.EXE=" + Str$(exec) + " *.MP3=" + Str$(mp3c) + " *.TXT=" + Str$(txtc) + " *.JPG=" + Str$(jpgc) + " *.html=" + Str$(htmc) + """;""" + """;""" + """"
List1.AddItem """Total File's Other=" + Str$(books - exec - mp3c - txtc - jpgc - htmc) + " Of " + others$ + " Types of Other Extensions" + """;""" + """;""" + """"

Print #freef1, """Total File's On Disk =" + Str$(books) + """;""" + """;""" + """"
Print #freef1, """Total File's *.EXE=" + Str$(exec) + " *.MP3=" + Str$(mp3c) + " *.TXT=" + Str$(txtc) + " *.JPG=" + Str$(jpgc) + " *.html=" + Str$(htmc) + """;""" + """;""" + """"
Print #freef1, """Total File's Other=" + Str$(books - exec - mp3c - txtc - jpgc - htmc) + " Of " + others$ + " Types of Other Extensions" + """;""" + """;""" + """"

Set f = fs.GetFolder(Mid$(b1$, 1, 3))
s22 = f.Size

List1.AddItem """Total Size of Disk = " + Format$(s22 / (1048576), "0.0") + " Mega Bytes" + """;""" + """;""" + """"
Print #freef1, """Total Size of Disk = " + Format$(s22 / (1048576), "0.0") + " Mega Bytes" + """;""" + """;""" + """"
Close #freef1

List1.AddItem "--------Finished----------"
List1.ListIndex = List1.ListCount - 1

'Timer1.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
End
End Sub
