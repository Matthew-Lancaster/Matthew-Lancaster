Attribute VB_Name = "SCANPATH_MOD"
'COMMON CODE



Sub GROUP_IN_500_JPG_MEDIA_PLAYER()


'<QUOTE>
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
' ---- SMALL PIECE OF CODE TO MOVE FILENAME IN BATCH OF 500 SET
' ---- FOR MEDIA PLAYERS   --- 2012-11-06
' ---- EXAMPLE AT UNQUOTE

' ---- GROUP IN 500'S
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
        Key = "PATH"
        'ScanPath.ListView1.ColumnHeaders , Item(Key), , 4100
        Set Item = ScanPath.ListView1.ColumnHeaders.Item(Key)
        Item.Width = 6000
        
        
        
ScanPath.Show
DoEvents

VV = "H:\00 ART WORK IN 500'S\"
ScanPath.TxtPath = VV
ScanPath.cmdScanDir_Click

For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
'    Debug.Print A1
    If A1 <> VV Then ScanPath.ListView1.ListItems.Remove WE
    
Next

Dim FL()
ReDim FL(ScanPath.ListView1.ListItems.Count)

For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    FL(WE) = A1 + B1
Next

ScanPath.ListView1.ListItems.Clear

For RR = 1 To UBound(FL)
    ScanPath.ListView1.ListItems.Clear
    ScanPath.TxtPath = FL(RR)
    ScanPath.cmdScan_Click
'    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    For WE = 1 To ScanPath.ListView1.ListItems.Count
        A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
        B1 = ScanPath.ListView1.ListItems.Item(WE)
'        Debug.Print Replace(Mid(FL(RR), InStrRev(FL(RR), "\", Len(FL(RR)) - 1)), "\", "")
'        Debug.Print Format((1 + Int((WE - 1) / 500)) * 500, "0000")
        D1 = Replace(Mid(FL(RR), InStrRev(FL(RR), "\", Len(FL(RR)) - 1)), "\", "")

        E1 = Format((1 + Int((WE - 1) / 500)) * 500, "0000")
        F1 = A1 + B1
        G1 = VV + D1 + "\" + E1 + " " + D1
        H1 = VV + D1 + "\" + E1 + " " + D1 + "\" + B1

        If OE1 <> E1 Then
            If FS.FolderExists(G1) = False Then
'                        Debug.Print G1
                MkDir G1
            End If
        End If
        OE1 = E1
        


        
        If FS.FileExists(H1) = False Then
'            Debug.Print H1
            FS.MOVEFILE F1, H1
        End If
Next
Next

'<UNQUOTE>
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'''' ---- SMALL PIECE OF CODE TO MOVE FILENAME IN BATCH OF 500 SET
'''' ---- FOR MEDIA PLAYERS   --- 2012-11-06
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------
'Source --SOLUTION - -SOLVABILITY
'---------------------------------------------------------------------------
'PATH :-:
'FILE :-:#00 File List DIR & SUB HERE -- 2012_11_06_21_41_06 -- NORMAL ---- H__00 ART WORK IN 500'S VOL - WD PASSPORT- DIRECTORY.txt
'---------------------------------------------------------------------------
'FOLDER DIRECTORY OBJECTIVE SCAN --
'## H:\00 ART WORK IN 500'S
'---------------------------------------------------------------------------
'DIR:                  ___,___,___,___,_66
'FILE:                 ___,___,___,_24,874
'PROCESS:              ___,___,___,_24,940
'---------------------------------------------------------------------------
'Volume Name:          WD PASSPORT
'Drive Type:           2
'Drive Type:           FIXED
'FILE System:          NTFS
'Serial Number:        -1741832190
'Serial Number &H:     98,2D,C0,02
'Share Name:           LOCAL SYSTEM DRIVE
'Total Size:           ___,_60,011,610,112 - GIGA RANGE
'Available:            ___,__7,679,459,328 - GIGA RANGE
'Free Space:           ___,__7,679,459,328 - GIGA RANGE
'Process Lenght Time:  0.217 Minute Divide Clock
'Process Lenght Time:  13 -- Seconds
'BEGIN TIME:           06 Nov 2012 21:41:06h
'AFTER TIME:           06 Nov 2012 21:41:19h - 13 Sec
'FORMAT BEGIN TIME:    Tue 06 Nov 2012 09:41:06p
'OBJECTIVE MODE BRIEF: SINGLE DRIVE SELECT VIA BUTTON
'---------------------------------------------------------------------------
'H:\00 ART WORK IN 500'S\Britannica 2000 Deluxe
'H:\00 ART WORK IN 500'S\DOS 6.22 & CDROM - IMAGES
'H:\00 ART WORK IN 500'S\DOS 6.22 & CDROM - WOMEN
'H:\00 ART WORK IN 500'S\Fractals
'H:\00 ART WORK IN 500'S\Great Images In Nasa
'H:\00 ART WORK IN 500'S\Internet 20 Days
'H:\00 ART WORK IN 500'S\Mental 00
'H:\00 ART WORK IN 500'S\Mental 01
'H:\00 ART WORK IN 500'S\Nasa Dryden
'H:\00 ART WORK IN 500'S\Nasa Earth OB
'H:\00 ART WORK IN 500'S\NIKON
'H:\00 ART WORK IN 500'S\RSS Feeds
'H:\00 ART WORK IN 500'S\Screen Shots
'H:\00 ART WORK IN 500'S\Britannica 2000 Deluxe\0500 Britannica 2000 Deluxe
'H:\00 ART WORK IN 500'S\Britannica 2000 Deluxe\1000 Britannica 2000 Deluxe
'H:\00 ART WORK IN 500'S\Britannica 2000 Deluxe\1500 Britannica 2000 Deluxe
'H:\00 ART WORK IN 500'S\Britannica 2000 Deluxe\2000 Britannica 2000 Deluxe
'H:\00 ART WORK IN 500'S\Britannica 2000 Deluxe\2500 Britannica 2000 Deluxe
'H:\00 ART WORK IN 500'S\DOS 6.22 & CDROM - IMAGES\0500 DOS 6.22 & CDROM - IMAGES
'H:\00 ART WORK IN 500'S\DOS 6.22 & CDROM - IMAGES\1000 DOS 6.22 & CDROM - IMAGES
'H:\00 ART WORK IN 500'S\DOS 6.22 & CDROM - IMAGES\1500 DOS 6.22 & CDROM - IMAGES
'H:\00 ART WORK IN 500'S\DOS 6.22 & CDROM - WOMEN\0500 DOS 6.22 & CDROM - WOMEN
'H:\00 ART WORK IN 500'S\DOS 6.22 & CDROM - WOMEN\1000 DOS 6.22 & CDROM - WOMEN
'H:\00 ART WORK IN 500'S\DOS 6.22 & CDROM - WOMEN\1500 DOS 6.22 & CDROM - WOMEN
'H:\00 ART WORK IN 500'S\Fractals\0500 Fractals
'H:\00 ART WORK IN 500'S\Fractals\1000 Fractals
'H:\00 ART WORK IN 500'S\Fractals\1500 Fractals
'H:\00 ART WORK IN 500'S\Fractals\2000 Fractals
'H:\00 ART WORK IN 500'S\Fractals\2500 Fractals
'H:\00 ART WORK IN 500'S\Great Images In Nasa\0500 Great Images In Nasa
'H:\00 ART WORK IN 500'S\Great Images In Nasa\1000 Great Images In Nasa
'H:\00 ART WORK IN 500'S\Great Images In Nasa\1500 Great Images In Nasa
'H:\00 ART WORK IN 500'S\Internet 20 Days\0500 Internet 20 Days
'H:\00 ART WORK IN 500'S\Internet 20 Days\1000 Internet 20 Days
'H:\00 ART WORK IN 500'S\Internet 20 Days\1500 Internet 20 Days
'H:\00 ART WORK IN 500'S\Internet 20 Days\2000 Internet 20 Days
'H:\00 ART WORK IN 500'S\Mental 00\0500 Mental 00
'H:\00 ART WORK IN 500'S\Mental 00\1000 Mental 00
'H:\00 ART WORK IN 500'S\Mental 00\1500 Mental 00
'H:\00 ART WORK IN 500'S\Mental 00\2000 Mental 00
'H:\00 ART WORK IN 500'S\Mental 00\2500 Mental 00
'H:\00 ART WORK IN 500'S\Mental 01\0500 Mental 01
'H:\00 ART WORK IN 500'S\Mental 01\1000 Mental 01
'H:\00 ART WORK IN 500'S\Mental 01\1500 Mental 01
'H:\00 ART WORK IN 500'S\Mental 01\2000 Mental 01
'H:\00 ART WORK IN 500'S\Nasa Dryden\0500 Nasa Dryden
'H:\00 ART WORK IN 500'S\Nasa Dryden\1000 Nasa Dryden
'H:\00 ART WORK IN 500'S\Nasa Dryden\1500 Nasa Dryden
'H:\00 ART WORK IN 500'S\Nasa Dryden\2000 Nasa Dryden
'H:\00 ART WORK IN 500'S\Nasa Dryden\2500 Nasa Dryden
'H:\00 ART WORK IN 500'S\Nasa Earth OB\0500 Nasa Earth OB
'H:\00 ART WORK IN 500'S\Nasa Earth OB\1000 Nasa Earth OB
'H:\00 ART WORK IN 500'S\Nasa Earth OB\1500 Nasa Earth OB
'H:\00 ART WORK IN 500'S\Nasa Earth OB\2000 Nasa Earth OB
'H:\00 ART WORK IN 500'S\Nasa Earth OB\2500 Nasa Earth OB
'H:\00 ART WORK IN 500'S\NIKON\0500 NIKON
'H:\00 ART WORK IN 500'S\NIKON\1000 NIKON
'H:\00 ART WORK IN 500'S\NIKON\1500 NIKON
'H:\00 ART WORK IN 500'S\NIKON\2000 NIKON
'H:\00 ART WORK IN 500'S\NIKON\2500 NIKON
'H:\00 ART WORK IN 500'S\RSS Feeds\0500 RSS Feeds
'H:\00 ART WORK IN 500'S\RSS Feeds\1000 RSS Feeds
'H:\00 ART WORK IN 500'S\RSS Feeds\1500 RSS Feeds
'H:\00 ART WORK IN 500'S\RSS Feeds\2000 RSS Feeds
'H:\00 ART WORK IN 500'S\RSS Feeds\2500 RSS Feeds
'H:\00 ART WORK IN 500'S\Screen Shots\0500 Screen Shots
End Sub
