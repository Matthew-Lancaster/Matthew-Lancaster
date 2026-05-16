Attribute VB_Name = "Module1"
Option Base 1
Private Type imgfiles
  fn As String   'filename
  fz As Long     'filesize
  wd As Long     'width
  ht As Long     'height
  dup As Integer 'possible duplicate
End Type

Private fary() As imgfiles, hary() As imgfiles
Private filenr As Integer, id As Long, sizes() As Integer
Private throwhtml As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub xstart(xpath As String)
    If Dir(xpath) = "" Then
        Call MsgBox("File not found.", vbExclamation)
        Exit Sub
    End If
  'open the file with the default program
    Call ShellExecute(hwnd, "Open", xpath, "", App.Path, 1)
End Sub

Public Sub all(place As String, ft As String, fd As Boolean)
  Dim FSA As New fcls, FSA2 As New fcls
  Dim i As Long, x As Long, min As Long, max As Long
  Dim fn As String, fts() As String
  Dim possdups As Boolean
  fts = Split(ft, ",")
  'find files add to array
  x = FSA.FindFiles(place, "*." & ft, , fd)
  'write html file if any files found
  If FSA.Count > 0 Then
    filenr = FreeFile
    Open App.Path & "\cfd.html" For Output As #filenr
    Print #filenr, "<html><head><title>check4duplicates - list all</title>"
    If Form1.Check2.Value = 1 Then
      Print #filenr, "<script type=""text/vbscript"" src=""cjh.vbs"">"
    Else
      Print #filenr, "<script type=""text/vbscript"" src=""jh.vbs"">"
    End If
    Print #filenr, "</script>"
    Print #filenr, "</head><body>"
    For i = 1 To FSA.Count
       Print #filenr, "<img id=""" & i & """ src=""" & FSA.Files(i) _
       & """ alt=""Click image to delete " & FSA.Files(i) & """ " & "onclick=""delhideimg(" & i & ")""/>"
    Next i
    Print #filenr, "</body></html>"
    Close #filenr
    xstart App.Path & "\cfd.html"
  Else
    MsgBox "no files of filetype """ & ft & """ found in location """ & place & """"
  End If
End Sub

Public Sub dups(place As String, ft As String, fd As Boolean)
  Dim FSA As New fcls, FSA2 As New fcls
  Dim i As Long, x As Long, min As Long, max As Long
  Dim fn As String
  Dim possdups As Boolean
  'find files add to array with filesize
  'also get max filesize
  x = FSA.FindFiles(place, "*." & ft, , fd)
If FSA.Count > 0 Then
  ReDim fary(FSA.Count) As imgfiles
  For i = 1 To FSA.Count
    fn = Right(FSA.Files(i), Len(FSA.Files(i)) - Len(place))
    x = FSA2.FindFiles(place, fn)
    fary(i).fn = FSA.Files(i)
    fary(i).fz = FSA2.Size
    If max < FSA2.Size Then max = FSA2.Size
  Next i
  'get min filesize
  min = max
  For i = 1 To UBound(fary)
    If min > fary(i).fz And fary(i).fz <> 0 Then min = fary(i).fz
  Next i
  'now find how many of each size
  ReDim sizes(max) As Integer
  For i = 1 To FSA.Count
    If fary(i).fz <> 0 Then sizes(fary(i).fz) = sizes(fary(i).fz) + 1
  Next i
  'check if possible duplicates exist
  possdups = False
  For i = min To max
    If sizes(i) > 1 Then
      possdups = True
      Exit For
    End If
  Next i
  
  'write html file if possible duplicates
  If possdups = True Then
    throwhtml = False
    filenr = FreeFile
    Open App.Path & "\cfd.html" For Output As #filenr
    Print #filenr, "<html><head><title>check4duplicates - list possible duplicates</title>"
    If Form1.Check2.Value = 1 Then
      Print #filenr, "<script type=""text/vbscript"" src=""cjh.vbs"">"
    Else
      Print #filenr, "<script type=""text/vbscript"" src=""jh.vbs"">"
    End If
    Print #filenr, "</script>"
    Print #filenr, "</head><body>"
    For i = min To max
      If sizes(i) > 1 Then
        SetImageAttrOfSize i
        PrintFiles
      End If
    Next i
    Print #filenr, "</body></html>"
    Close #filenr
    If throwhtml = True Then 'is triggered when identical filesize & width & height
      xstart App.Path & "\cfd.html"
    Else
      MsgBox "imagefiles of identical size found, but none of identical width/height"
    End If
  Else
    MsgBox "no files of equal size found, no possible duplicates"
  End If
Else
  MsgBox "no files of filetype """ & ft & """ found in location """ & place & """"
End If
  Erase fary
  Erase sizes
End Sub

Private Sub SetImageAttrOfSize(fz As Long)
Dim i As Long, y As Long, pic As Object
Dim pic2 As StdPicture
On Error Resume Next
  For i = 1 To UBound(fary)
    If fary(i).fz = fz Then
      Set pic = LoadPicture(fary(i).fn)
      fary(i).wd = pic.Width
      fary(i).ht = pic.Height
      Set pic = Nothing
      addtohary i 'helper collection
    End If
  Next i
  'also find more exact possible duplicates in helper collection
  For i = 1 To UBound(hary) - 1
    For y = i + 1 To UBound(hary)
      If hary(i).wd = hary(y).wd And hary(i).ht = hary(y).ht Then
        hary(i).dup = 1
        hary(y).dup = 1
        throwhtml = True
      End If
    Next y
  Next i
End Sub

Private Sub PrintFiles()
Dim i As Long
  If UBound(hary) > 0 Then
  For i = 1 To UBound(hary)
    If hary(i).dup = 1 Then
      id = id + 1
      Print #filenr, "<img id=""" & id & """ src=""" & hary(i).fn _
       & """ alt=""Click image to delete " & hary(i).fn & """ " & "onclick=""delhideimg(" & id & ")""/>"
    End If
  Next i
  For i = 1 To UBound(hary)
    If hary(i).dup = 1 Then
      Print #filenr, "<hr />"
      Exit For
    End If
  Next i
  End If
  Erase hary 'reset helper
End Sub

Public Sub addtohary(i As Long)
Dim x As Long
On Error GoTo tom
  x = UBound(hary) + 1
  ReDim Preserve hary(x) As imgfiles
  hary(x) = fary(i)
  Exit Sub
tom:
  ReDim hary(1) As imgfiles
  hary(1) = fary(i)
End Sub

