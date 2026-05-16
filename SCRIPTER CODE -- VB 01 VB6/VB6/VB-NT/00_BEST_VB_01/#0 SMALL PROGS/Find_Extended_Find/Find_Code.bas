Attribute VB_Name = "Find_Code"
Option Explicit
'© Copyright 2003 Roger Gilchrist
'rojagilkrist@hotmail.com
'This is a very slightly modified version of routines I have used in earlier versions
'Slight changes to workaround UserDocument limits
Public Enum EnumMsg
  Search
  Complete
  inComplete
  Missing
  Found
  replaced
End Enum
#If False Then  'Trick preserves Case of Enums when typing in IDE
Private Search, Complete, inComplete, Missing, Found, replaced
#End If
'Search Settings
Public Const LongLimit               As Single = 2147483647
'Prevent FlexGrid from over-flowing;
'very unlikely that anything could be found that often
'just a safety valve inherited from the old listbox/IntegerLimit
Public bWholeWordonly                As Boolean
Public bCaseSensitive                As Boolean
Public bNoComments                   As Boolean
Public bCommentsOnly                 As Boolean
Public bNoStrings                    As Boolean
Public bStringsOnly                  As Boolean
Public iCurCodePane                  As Integer
Public BPatternSearch                As Boolean
Public bPTmpWholeWordonly            As Boolean
Public bPTmpCaseSensitive            As Boolean
Public bPTmpNoComments               As Boolean
Public bPTmpCommentsOnly             As Boolean
Public bPTmpNoStrings                As Boolean
Public bPTmpStringsOnly              As Boolean
Public bPTmpFindSelectWholeLine      As Boolean
Public bGridlines                    As Boolean
Public bShowCompLineNo               As Boolean
Public bShowProcLineNo               As Boolean
Public bShowReplace                  As Boolean
Private bComplete                    As Boolean
'halt search before completion
Public bCancel                       As Boolean
Private Const Apostrophe             As String = "'"
Public bShowProject                  As Boolean
Public bShowComponent                As Boolean
Public bShowRoutine                  As Boolean
Public bFindSelectWholeLine          As Boolean
Private ReplaceCount                 As Long

Private Sub ApplySelectedTextRestriction(cde As String, _
                                         ByVal FStartR As Long, _
                                         ByVal FStartC As Long, _
                                         ByVal FEndR As Long, _
                                         ByVal FEndC As Long, _
                                         ByVal SStartR As Long, _
                                         ByVal SStartC As Long, _
                                         ByVal SEndR As Long, _
                                         ByVal SEndC As Long)

  'Code is needed in DoSearch and DoReplace so a separate Procedure
  
  Dim InSelRange                  As Boolean

  If iCurCodePane = 3 Then
    If Len(cde) Then
      InSelRange = False
      If BetweenLng(SStartR, FStartR, SEndR) Then
        If BetweenLng(SStartR, FEndR, SEndR) Then
          InSelRange = True
          'ver 2.2.1
          'Refinement of test; only checks column if in first or last line of selected text
          ' any other line must be in the range
          If SStartR = FStartR Then
            InSelRange = SStartC >= FStartC
           ElseIf SEndR = FEndR Then
            InSelRange = FEndC <= SEndC
          End If
          If FStartR = FEndR Then
            'special case single line selected
            InSelRange = SStartC >= FStartC And FEndC <= SEndC
          End If
        End If
      End If
      If Not InSelRange Then
        cde = vbNullString
      End If
    End If
  End If

End Sub

Private Sub ApplyStringCommentFilters(cde As String, _
                                      ByVal StrTarget As String)

  Dim Codepos As Long

  Codepos = InStr(1, cde, StrTarget, IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
  If bNoComments Then
    If InComment(cde, Codepos) Then
      cde = vbNullString
    End If
  End If
  If bCommentsOnly Then
    If Not InComment(cde, Codepos) Then
      cde = vbNullString
    End If
  End If
  If bStringsOnly Then
    If InQuotes(cde, Codepos) = False Then
      cde = vbNullString
    End If
  End If
  If bNoStrings Then
    If InQuotes(cde, Codepos) Then
      cde = vbNullString
    End If
  End If

End Sub

Public Sub AutoPatternOff(ByVal StrF As String)

  If BPatternSearch Then
    If instrAny(StrF, "*", "!", "[", "]", "\") = False Then
      BPatternSearch = False
      mobjDoc.ClearForPattern
    End If
  End If

End Sub

Private Sub AutoSelectInitialize(PrevRange As Long, _
                                 AutoRevert As Boolean)

  'Code is needed in DoSearch and DoReplace so a separate Procedure

  If bAutoSelectText Then
    If InStr(GetSelectedText(VBInstance), vbNewLine) > 0 Then
      PrevRange = iCurCodePane
      iCurCodePane = 3
      AutoRevert = True
      mobjDoc.ToggleButtonFaces
    End If
  End If

End Sub

Private Function BetweenLng(ByVal MinV As Long, _
                            ByVal Val As Long, _
                            ByVal MaxV As Long, _
                            Optional ByVal InClusive As Boolean = True) As Boolean

  If InClusive Then
    If Val >= MinV Then
      If Val <= MaxV Then
        BetweenLng = True
      End If
    End If
   Else
    If Val > MinV Then
      If Val < MaxV Then
        BetweenLng = True
      End If
    End If
  End If

End Function

Private Function CancelSearch(G_Rows As Long) As Boolean

  CancelSearch = bCancel Or EscPressed Or G_Rows = LongLimit

End Function

Public Sub ClearFGrid(Fgrd As MSFlexGrid)

  With Fgrd
    .Rows = 2
    .TextMatrix(.Row, 0) = vbNullString
    .TextMatrix(.Row, 1) = vbNullString
    .TextMatrix(.Row, 2) = vbNullString
    .TextMatrix(.Row, 3) = vbNullString
  End With

End Sub

Public Sub ComboBoxSavePreviousSearch(combo As ComboBox, _
                                      Optional ByVal AddWord As String = vbNullString, _
                                      Optional ByVal ListLimit As Long = 20)

  'add a new word to history
  'if already in history move to top of history

  On Error Resume Next
  If LenB(AddWord) Then
    'new this allows you to insert a Newline from Replace combo
    'only the first part of line will be inserted in Find combo
    'not perfect but will work
    If InStr(AddWord, vbNewLine) Then
      AddWord = Left$(AddWord, InStr(AddWord, vbNewLine) - 1)
    End If
    combo.Text = AddWord
   Else
    AddWord = combo.Text
  End If
  With combo
    If LenB(.Text) Then
      If PosInCombo(.Text, combo) = -1 Then
        .AddItem .Text, 0
        If .ListCount > ListLimit Then
          .RemoveItem ListLimit
        End If
       Else
        .RemoveItem PosInCombo(.Text, combo)
        .AddItem AddWord, 0
        .Text = AddWord
      End If
    End If
  End With
  On Error GoTo 0

End Sub

Private Function ConvertAsciiSearch(ByVal strConv As String) As String

  'this routine adds the ability to search for Characters referred to by ascii value
  'ONLY works if pattern search is ON
  'other wise you are searching for the literal string
  
  Dim StartAsc As Long
  Dim EndAsc   As Long
  Dim AscVal   As Long

  StartAsc = 1
  Do While InStr(StartAsc, strConv, "\ASC(")
    StartAsc = InStr(StartAsc, strConv, "\ASC(")
    EndAsc = InStr(StartAsc, strConv, ")")
    If EndAsc > StartAsc Then
      AscVal = Val(Mid$(strConv, StartAsc + 5, EndAsc - 1))
      If AscVal > -1 And AscVal < 256 Then
        strConv = Left$(strConv, StartAsc - 1) & Chr$(AscVal) & Mid$(strConv, EndAsc + 1)
       Else
        StartAsc = EndAsc
        'this will jump you past an ivalid asc value and check for other "\ASC(" triggers
      End If
     Else
      Exit Do
      'this jumps out of Loop if the end bracket is missing
    End If
  Loop
  ConvertAsciiSearch = strConv

End Function

Public Sub DOFind(Fgrd As MSFlexGrid, _
                  cmbS As ComboBox)

  Dim Proj           As VBProject
  Dim Comp           As VBComponent
  Dim Pane           As CodePane
  Dim StartText      As Long
  Dim EndText        As Long
  Dim StrProjName    As String
  Dim StrCompName    As String
  Dim StrTarget      As String
  Dim strRoutine     As String
  Dim strFound       As String
  Dim StartLine      As Long
  Dim StartCol       As Long
  Dim EndLine        As Long
  Dim EndCol         As Long
  Dim GotIT          As Boolean
  Dim ProcLineNo     As Long

  On Error Resume Next
ReTry:
  With Fgrd
    '<Project|<Component|Line|<Procedure|Line|<Code
    StrProjName = .TextMatrix(.Row, 0)
    StrCompName = .TextMatrix(.Row, 1)
    StartLine = CLng(.TextMatrix(.Row, 2))
    strRoutine = .TextMatrix(.Row, 3)
    ProcLineNo = CLng(.TextMatrix(.Row, 4))
    strFound = .TextMatrix(.Row, 5)
    'StartLine = .RowData(.Row)
  End With 'Fgrd
  StrTarget = cmbS.Text
  'this is the fast find used if no editing has been done
  For Each Proj In VBInstance.VBProjects
    If StrProjName = Proj.Name Then
      For Each Comp In Proj.VBComponents
        If Comp.Name = StrCompName Then
          If Comp.CodeModule.Find(strFound, StartLine, 1, -1, -1) Then
            GotIT = True
            Exit For
          End If
        End If
      Next Comp
      If GotIT Then
        Exit For
      End If
    End If
  Next Proj
  'this will do Find if the line number has been changed by editing
  If Not GotIT Then
    StartLine = 1
    For Each Proj In VBInstance.VBProjects
      If StrProjName = Proj.Name Then
        For Each Comp In Proj.VBComponents
          If Comp.Name = StrCompName Then
            StartLine = 1
            If Comp.CodeModule.Find(strFound, StartLine, 1, Comp.CodeModule.CountOfLines, -1) Then
              Do
                'this do loop takes care of the possibility of identical lines being present in different routines
                If strRoutine = GetProcName(Comp.CodeModule, StartLine) Then
                  GotIT = True
                  'reset the Line data so fast search will be used next time
                  Fgrd.TextMatrix(Fgrd.Row, 2) = StartLine
                  Fgrd.TextMatrix(Fgrd.Row, 4) = GetProcLineNumber(Comp.CodeModule, StartLine)
                  Exit Do
                End If
                StartLine = StartLine + 1
              Loop While Comp.CodeModule.Find(strFound, StartLine, 1, Comp.CodeModule.CountOfLines, -1)
            End If
          End If
          If GotIT Then
            Exit For
          End If
        Next Comp
        If GotIT Then
          Exit For
        End If
      End If
    Next Proj
  End If
  If GotIT Then
    If bFindSelectWholeLine Then
      'select the whole line
      StartText = InStr(1, Comp.CodeModule.Lines(StartLine, 1), strFound, vbTextCompare)
      EndText = StartText + Len(strFound)
     Else
      'select the search word
      If Not BPatternSearch Then
        StartText = InStr(1, Comp.CodeModule.Lines(StartLine, 1), StrTarget, vbTextCompare)
        EndText = StartText + Len(StrTarget)
       Else
        ' the string/comment filters cannot work on a PatternSearch StrFind
        'but you can get the actual string found and test that
        StartCol = 1
        EndLine = -1
        EndCol = -1
        If Comp.CodeModule.Find(StrTarget, StartLine, StartText, EndLine, EndText, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
        End If
      End If
    End If
    Set Pane = Comp.CodeModule.CodePane
    With Pane
      'when docked only the first instance selected in GrdFound got highlighted
      'until I added next line, no idea why it works.
      .Window.Visible = False
      .Show
      .Window.SetFocus
      ' .TopLine = Abs(Int(.CountOfVisibleLines / 2) - StartLine) + 1
    End With
    If StartText > 0 Then
      ReportAction Fgrd, Found
      With Pane
        .SetSelection StartLine, StartText, StartLine, EndText
      End With
      Set Pane = Nothing
      Exit Sub
    End If
   Else
    'Line is missing probably edited out
    DoSearch Fgrd, cmbS
    '    With Fgrd
    '      DelRow = .Row
    '      .RemoveItem DelRow
    '      If .Rows > 1 Then
    '        If DelRow > .Rows Then
    '          .Row = .Rows
    '        End If
    '        GoTo ReTry
    '      End If
    '    End With 'Fgrd
  End If
  On Error GoTo 0

End Sub

Public Sub DoReplace(Fgrd As MSFlexGrid, _
                     cmbF As ComboBox, _
                     cmbR As ComboBox)

  Dim MSg                       As String
  Dim code                      As String
  Dim CompMod                   As CodeModule
  Dim Comp                      As VBComponent
  Dim Proj                      As VBProject
  Dim strFind                   As String
  Dim strReplace                As String
  Dim CurProc                   As String
  Dim curModule                 As String
  Dim ProcName                  As String
  Dim StartLine                 As Long
  Dim EndLine                   As Long
  Dim StartCol                  As Long
  Dim EndCol                    As Long
  Dim SelStartLine              As Long
  Dim SelEndLine                As Long
  Dim SelStartCol               As Long
  Dim SelEndCol                 As Long
  Dim AutoRangeRevert           As Boolean
  Dim PrevCurCodePane           As Long

  If bFilterWarning Then
    MSg = IIf(bWholeWordonly, "ON ", "OFF") & " Whole Word"
    MSg = MSg & vbNewLine & IIf(bCaseSensitive, "ON ", "OFF") & " Case Sensitive"
    MSg = MSg & vbNewLine & IIf(bNoComments, "ON ", "OFF") & " No Comments"
    MSg = MSg & vbNewLine & IIf(bCommentsOnly, "ON ", "OFF") & " Comments Only"
    MSg = MSg & vbNewLine & IIf(bNoStrings, "ON ", "OFF") & " No Strings"
    MSg = MSg & vbNewLine & IIf(bStringsOnly, "ON ", "OFF") & " Strings Only"
    If vbCancel = MsgBox(MSg & vbNewLine & vbNewLine & "Proceed with Replace anyway?", vbExclamation + vbOKCancel, "Filter Warning " & AppDetails) Then
      Exit Sub
    End If
  End If
  strFind = cmbF.Text
  strReplace = cmbR.Text
  If LenB(strFind) = 0 Then
    Exit Sub
  End If
  If bBlankWarning Then
    If LenB(strReplace) = 0 Then
      If vbCancel = MsgBox("Replace '" & strFind & "' with blank?", vbExclamation + vbOKCancel, "Blank Warning " & AppDetails) Then
        Exit Sub
      End If
    End If
  End If
  On Error Resume Next
  ComboBoxSavePreviousSearch cmbF, , HistDeep
  ComboBoxSavePreviousSearch cmbR, , HistDeep
  If InStr(strReplace, "^N^") Then
    strReplace = Replace$(strReplace, "^N^", vbNewLine)
  End If
  If InStr(strReplace, "^T^") Then
    strReplace = Replace$(strReplace, "^T^", vbTab)
  End If
  DoEvents
  bCancel = False
  GetCounts
  CurProc = GetCurrentProcedure
  curModule = GetCurrentModule
  AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
  If iCurCodePane = 3 Then
    '    HiLitSelection$ = GetSelectedText(VBInstance)
    VBInstance.ActiveCodePane.GetSelection SelStartLine, SelStartCol, SelEndLine, SelEndCol
  End If
  ReplaceCount = 0
  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If SafeCompToProcess(Comp) Then
        If iCurCodePane > 0 Then
          If Comp.Name <> curModule Then
            GoTo SkipComp
          End If
        End If
        Set CompMod = Comp.CodeModule
        StartLine = 1
        If CompMod.Find(strFind, StartLine, 1, CompMod.CountOfLines, -1, bWholeWordonly, bCaseSensitive, False) Then
          Do
            EndLine = -1
            StartCol = 1
            EndCol = -1
            If CompMod.Find(strFind, StartLine, StartCol, EndLine, EndCol, bWholeWordonly, bCaseSensitive, False) Then
              ProcName = CompMod.ProcOfLine(StartLine, vbext_pk_Proc)
              If LenB(ProcName) = 0 Then
                ProcName = "(Declarations)"
              End If
              If iCurCodePane = 2 Then
                If ProcName <> CurProc Then
                  GoTo SkipProc
                End If
              End If
              code$ = CompMod.Lines(StartLine, 1)
              ApplyStringCommentFilters code$, strFind
              ApplySelectedTextRestriction code$, StartLine, StartCol, EndLine, EndCol, SelStartLine, SelStartCol, SelEndLine, SelEndCol
              If Len(code$) Then
                code$ = Replace$(code$, strFind, strReplace, , , IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
                CompMod.ReplaceLine StartLine, code$
                ReplaceCount = ReplaceCount + 1
              End If
SkipProc:
            End If
            code$ = vbNullString
            StartLine = StartLine + 1
            If CancelSearch(Fgrd.Rows) Then
              Exit Do
            End If
            If StartLine > CompMod.CountOfLines Then
              Exit Do
            End If
          Loop While CompMod.Find(strFind, StartLine, 1, -1, -1, bWholeWordonly, bCaseSensitive, False) 'StartLine > 0 And StartLine <= CompMod.CountOfLines
        End If
      End If
SkipComp:
      Set Comp = Nothing
      If CancelSearch(Fgrd.Rows) Then
        Exit For
      End If
    Next Comp
    If CancelSearch(Fgrd.Rows) Then
      Exit For
    End If
  Next Proj
  'this turns off auto Selected text only
  If AutoRangeRevert Then
    iCurCodePane = PrevCurCodePane
    mobjDoc.ToggleButtonFaces
  End If
  Set Proj = Nothing
  Set CompMod = Nothing
  ReportAction Fgrd, replaced
  If bReplace2Search Then
    If Len(strReplace) Then
      ComboBoxSavePreviousSearch cmbR, strFind, HistDeep
      ComboBoxSavePreviousSearch cmbF, strReplace, HistDeep
    End If
  End If
  On Error GoTo 0

End Sub

Public Sub DoSearch(Fgrd As MSFlexGrid, _
                    cmb As ComboBox)

  'ver1.1.02 major rewrite to allow simple pattern searching
  
  Dim code                      As String
  Dim ProcName                  As String
  Dim ProcLineNo                As Long
  Dim CompMod                   As CodeModule
  Dim Comp                      As VBComponent
  Dim Proj                      As VBProject
  Dim strFind                   As String
  Dim LongestPrj                As String
  Dim longestPrc                As String
  Dim LongestCmp                As String
  Dim LongestCNo                As String
  Dim LongestPNo                As String
  Dim ResizeNeeded              As Boolean
  Dim strStrComTest             As String
  Dim StartLine                 As Long
  Dim StartCol                  As Long
  Dim EndLine                   As Long
  Dim EndCol                    As Long
  Dim SecondRun                 As Boolean
  Dim CurProc                   As String
  Dim curModule                 As String
  Dim StartProcRow              As Long
  Dim PrevCurCodePane           As Long
  Dim AutoRangeRevert           As Boolean
  Dim HiLitSelection            As String
  Dim SelStartLine              As Long
  Dim SeEndLine                 As Long
  Dim SelStartCol               As Long
  Dim SelEndCol                 As Long

  On Error Resume Next
  strFind = cmb.Text
  'if strFind doesn't have any triggers and Pattern Search is on then turn it off
  AutoPatternOff strFind
  If BPatternSearch Then
    If InStr(strFind, "\ASC(") Then
      strFind = ConvertAsciiSearch(strFind)
      cmb.Text = strFind
    End If
  End If
  If LenB(strFind) = 0 Then
    Exit Sub
  End If
  If strFind = " " Then
    MsgBox "Search for single spaces is cancelled, it overloads the system", vbInformation
    SetFocus_Safe cmb
    Exit Sub
  End If
  StartProcRow = 1
ReTry:
  Fgrd.BackColorFixed = ColourHeadWork
  If LenB(strFind) > 0 Then
    LongestPrj$ = vbNullString
    bComplete = False
    ClearFGrid Fgrd
    ComboBoxSavePreviousSearch cmb, , HistDeep
    mobjDoc.CancelButton True
    DoEvents
    bCancel = False
    GetCounts
    CurProc = GetCurrentProcedure
    curModule = GetCurrentModule
    'this does the auto switching if multiple code lines are selected
    AutoSelectInitialize PrevCurCodePane, AutoRangeRevert
    ' this sets search limits if there is selected code.
    If iCurCodePane = 3 Then
      HiLitSelection$ = GetSelectedText(VBInstance)
      VBInstance.ActiveCodePane.GetSelection SelStartLine, SelStartCol, SeEndLine, SelEndCol
    End If
    For Each Proj In VBInstance.VBProjects
      If Len(Proj.Name) > Len(LongestPrj$) Then
        LongestPrj$ = Proj.Name
        ResizeNeeded = True
      End If
      For Each Comp In Proj.VBComponents
        If Fgrd.Rows < LongLimit Then
          If SafeCompToProcess(Comp) Then
            If iCurCodePane > 0 Then
              If Comp.Name <> curModule Then
                GoTo SkipComp
              End If
            End If
            If Len(Comp.Name) > Len(LongestCmp$) Then
              LongestCmp$ = Comp.Name
              ResizeNeeded = True
            End If
            With Comp
              Set CompMod = .CodeModule
              '5hould I quit?
              ReportAction Fgrd, Search
              If LenB(.Name) = 0 Then
                bCancel = True
                bComplete = True
              End If
            End With
            If CancelSearch(Fgrd.Rows) Then
              Exit For
            End If
            'Safety turns off filters if comment/double quote is actually in the search phrase
            If bNoComments Then
              If InStr(strFind, Apostrophe) > 0 Then
                bNoComments = True
                mobjDoc.SetFilterButtons
              End If
            End If
            If bNoStrings Then
              If InStr(strFind, Chr$(34)) > 0 Then
                bNoStrings = False
                mobjDoc.SetFilterButtons
              End If
            End If
            StartLine = 1 'initialize search range
            Do
              StartCol = 1
              EndLine = -1
              EndCol = -1
              If CompMod.Find(strFind, StartLine, StartCol, EndLine, EndCol, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch) Then
                DoEvents
                If CancelSearch(Fgrd.Rows) Then
                  Exit For
                End If
                ProcName = GetProcName(CompMod, StartLine)
                If iCurCodePane = 2 Then
                  If ProcName <> CurProc Then
                    GoTo SkipProc
                  End If
                End If
                ProcLineNo = GetProcLineNumber(CompMod, StartLine)
                If Len(ProcName) > Len(longestPrc$) Then
                  longestPrc$ = ProcName
                  ResizeNeeded = True
                End If
                If Len(CStr(StartLine)) > Len(LongestCNo$) Then
                  LongestCNo$ = CStr(StartLine)
                  ResizeNeeded = True
                End If
                If Len(CStr(StartLine)) > Len(LongestCNo$) Then
                  LongestCNo$ = CStr(StartLine)
                  ResizeNeeded = True
                End If
                If Len(CStr(ProcLineNo)) > Len(LongestPNo$) Then
                  LongestPNo$ = CStr(ProcLineNo)
                  ResizeNeeded = True
                End If
                code$ = CompMod.Lines(StartLine, 1)
                'apply nostring no comment filters
                If BPatternSearch Then
                  ' the string/comment filters cannot work on a PatternSearch StrFind
                  'but you can get the actual string found and test that
                  CompMod.CodePane.SetSelection StartLine, StartCol, EndLine, EndCol
                  strStrComTest = Mid$(code$, StartCol, EndCol - StartCol)
                 Else
                  strStrComTest = strFind
                End If
                ApplyStringCommentFilters code$, strStrComTest
                ApplySelectedTextRestriction code$, StartLine, StartCol, EndLine, EndCol, SelStartLine, SelStartCol, SeEndLine, SelEndCol
                If LenB(code$) Then
                  If ResizeNeeded Then
                    'slight speed advantage of not doing this unless called for
                    mobjDoc.GridReSize LongestPrj$, LongestCmp$, LongestCNo$, longestPrc$, LongestPNo$
                    ResizeNeeded = False
                  End If
                  With Fgrd
                    .Rows = .Rows + 1
                    .TextMatrix(.Row, 0) = Proj.Name
                    .TextMatrix(.Row, 1) = Comp.Name
                    .TextMatrix(.Row, 2) = StartLine
                    .TextMatrix(.Row, 3) = ProcName
                    .TextMatrix(.Row, 4) = ProcLineNo
                    .TextMatrix(.Row, 5) = code$
                    .Row = .Row + 1
                    code$ = vbNullString
                  End With 'Fgrd
                End If
SkipProc:
              End If
              StartLine = StartLine + 1
              If CancelSearch(Fgrd.Rows) Then
                Exit Do
              End If
              If StartLine > CompMod.CountOfLines Then
                Exit Do
              End If
            Loop While CompMod.Find(strFind, StartLine, 1, -1, -1, bWholeWordonly, IIf(BPatternSearch, False, bCaseSensitive), BPatternSearch)
            ' End If
          End If
        End If
SkipComp:
        Set Comp = Nothing
        If CancelSearch(Fgrd.Rows) Then
          Exit For
        End If
      Next Comp
      If CancelSearch(Fgrd.Rows) Then
        Exit For
      End If
    Next Proj
    Set Proj = Nothing
    Set CompMod = Nothing
    mobjDoc.CancelButton False
  End If
  With Fgrd
    .Rows = .Rows - 1
    If .Row = 0 Then
      'nothing found
      .BackColorSel = ColourHeadNoFind
      .ForeColorSel = ColourHeadFore
     Else
      .BackColorSel = ColourFindSelectBack
      .ForeColorSel = ColourFindSelectFore
    End If
  End With 'Fgrd
  'automatically switch to pattern search if ordinary fails
  If Fgrd.Row = 0 Then
    If Not BPatternSearch Then
      If instrAny(strFind, "*", "!", "[", "]", "\") Then
        BPatternSearch = Not BPatternSearch
        mobjDoc.ClearForPattern
        SecondRun = True
        GoTo ReTry
      End If
    End If
    If SecondRun Then
      'autoPattern Search is on
      If Fgrd.Row = 0 Then
        'turn it off it still no finds
        BPatternSearch = Not BPatternSearch
        mobjDoc.ClearForPattern
      End If
    End If
  End If
  If CancelSearch(Fgrd.Rows) Then
    ReportAction Fgrd, IIf(bComplete, Found, inComplete)
  End If
  Fgrd.Refresh
  If Fgrd.Rows > 1 Then
    ReportAction Fgrd, IIf(bComplete, Found, Complete)
    SetFocus_Safe Fgrd
   Else
    ReportAction Fgrd, IIf(bComplete, Found, Missing)
  End If
  'this turns off auto Selected text only
  If AutoRangeRevert Then
    iCurCodePane = PrevCurCodePane
    mobjDoc.ToggleButtonFaces
  End If
  'this turns auto switch to pattern search off if it was used
  If Fgrd.Rows = LongLimit Then
    ' as this is 2147483647 rows it is unlikely that this will ever hit but just in case :)
    MsgBox "Search halted because number of finds reached limit of Find ComboBox", vbCritical
  End If
  On Error GoTo 0

End Sub

Private Function FilteredInStr(strSearch As String, _
                               strFind As String) As Long

  'ver 1.1.02
  'improved word detection makes sure that the DoFind routine
  'highlights the correct (or at least first instance)
  'in string that matches all filters
  
  Dim Lbit As String
  Dim Rbit As String

  Do
    FilteredInStr = InStr(FilteredInStr + 1, strSearch, strFind, IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
    If strSearch = strFind Then
      Exit Do
    End If
    If bWholeWordonly Then
      Select Case FilteredInStr
       Case 1
        Rbit$ = Mid$(strSearch, Len(strFind) + 1, 1)
        If IsPunct(Rbit) Then
          Exit Do
        End If
       Case Len(strSearch) - Len(strFind) + 1
        Lbit$ = Mid$(strSearch, Len(strSearch) - Len(strFind), 1)
        If IsPunct(Lbit) Then
          Exit Function
        End If
       Case 0
        Exit Do 'not found do nothing
       Case Else
        Lbit$ = Mid$(strSearch, FilteredInStr - 1, 1)
        Rbit$ = Mid$(strSearch, FilteredInStr + Len(strFind), 1)
        If IsPunct(Lbit) And IsPunct(Rbit) Then
          Exit Do
        End If
      End Select
     Else
      Exit Do
    End If
  Loop

End Function

Public Function GetCurrentModule() As String

  GetCurrentModule = VBInstance.ActiveCodePane.CodeModule.Name

End Function

Public Function GetCurrentProcedure() As String

  Dim sl As Long
  Dim sc As Long
  Dim el As Long
  Dim ec As Long

  VBInstance.ActiveCodePane.GetSelection sl, sc, el, ec
  GetCurrentProcedure = VBInstance.ActiveCodePane.CodeModule.ProcOfLine(sl, vbext_pk_Proc)
  If Len(GetCurrentProcedure) = 0 Then
    GetCurrentProcedure = "(Declarations)"
  End If

End Function

Public Function GetProcLineNumber(CmpMod As CodeModule, _
                                  CodeLineNo As Long) As String

  Dim LProcName As String

  LProcName = GetProcName(CmpMod, CodeLineNo)
  If LProcName = "(Declarations)" Then
    GetProcLineNumber = CodeLineNo
   Else
    'The + 1 is because ProcBodyLine returns a 0 based count but most people like 1 based counts
    'Oddly CodeLineNo which is generated by VB's Find is 1 based
    GetProcLineNumber = CodeLineNo - CmpMod.ProcBodyLine(LProcName, vbext_pk_Proc) + 1
  End If

End Function

Public Function GetProcName(CmpMod As CodeModule, _
                            CodeLineNo As Long) As String

  GetProcName = CmpMod.ProcOfLine(CodeLineNo, vbext_pk_Proc)
  If LenB(GetProcName) = 0 Then
    GetProcName = "(Declarations)"
  End If

End Function

Public Function InComment(ByVal code As String, _
                          ByVal Codepos As Long) As Boolean

  Dim SQuotePos As Long

  On Error Resume Next
  SQuotePos = InStr(code$, Apostrophe)
  If SQuotePos = 1 Then
    InComment = True
  End If
  If SQuotePos > 1 Then
    If Codepos > SQuotePos Then
      InComment = True
    End If
  End If
  On Error GoTo 0

End Function

Public Function InQuotes(ByVal code As String, _
                         ByVal Codepos As Long) As Boolean

  Dim LQ As Long
  Dim FQ As Long

  On Error Resume Next
  LQ = InStr(StrReverse(code$), Chr$(34))
  If LQ > 0 Then
    LQ = Len(code$) - LQ + 1
  End If
  FQ = InStr(code$, Chr$(34))
  If LQ = 0 Then
    If FQ = 0 Then
      Exit Function
    End If
  End If
  If LQ = FQ Then
    Exit Function
  End If
  If FQ < Codepos Then
    If Codepos < LQ Then
      InQuotes = True
    End If
  End If
  On Error GoTo 0

End Function

Public Function IsAlphaIntl(ByVal sChar As String) As Boolean

  IsAlphaIntl = Not (UCase$(sChar) = LCase$(sChar))

End Function

Public Function IsNumeral(ByVal strTest As String) As Boolean

  IsNumeral = InStr("1234567890", strTest) > 0

End Function

Public Function IsPunct(ByVal strTest As String) As Boolean

  'Detect punctuation

  If IsNumeral(strTest) Then
    IsPunct = False
   Else
    IsPunct = Not IsAlphaIntl(strTest)
  End If

End Function

Public Sub ReportAction(Fgrd As MSFlexGrid, _
                        ByVal Act As EnumMsg, _
                        Optional ByVal AppendStr As String)

  Dim MSg                As String
  Dim StrItems           As String
  Dim StrFilterWarning   As String
  Dim strSearchEndStatus As String
  Dim StrPatternWarning  As String

  StrItems = "(" & Fgrd.Rows - 1 & ") Item" & IIf(Fgrd.Rows - 1 <> 1, "s", vbNullString)
  StrFilterWarning = IIf(mobjDoc.AnyFilterOn, " <Filter>", vbNullString)
  StrPatternWarning = IIf(BPatternSearch, " <Pattern>", vbNullString)
  Select Case Act
   Case Search
    strSearchEndStatus = " Searching " & IIf(Len(AppendStr), " in " & AppendStr, "...")
    Fgrd.BackColorFixed = ColourHeadWork
   Case Complete
    strSearchEndStatus = " Search Complete."
    Fgrd.BackColorFixed = IIf(BPatternSearch, ColourHeadPattern, ColourHeadDefault)
   Case inComplete
    strSearchEndStatus = " Search Cancelled."
    Fgrd.BackColorFixed = IIf(BPatternSearch, ColourHeadPattern, ColourHeadDefault)
  End Select
  Select Case Act
   Case replaced
    MSg = "Code: " & "(" & ReplaceCount & ") Item" & IIf(ReplaceCount <> 1, "s", vbNullString) & " replaced"
    Fgrd.BackColorFixed = IIf(BPatternSearch, ColourHeadPattern, ColourHeadDefault)
   Case Missing
    MSg = "Code: " & StrItems & " Found" & StrFilterWarning & strSearchEndStatus & StrPatternWarning
    Fgrd.BackColorFixed = ColourHeadNoFind
   Case Else
    MSg = "Code: " & StrItems & " Found" & StrFilterWarning & strSearchEndStatus & StrPatternWarning
  End Select
  Fgrd.TextMatrix(0, 5) = MSg
  Fgrd.Refresh
  DoEvents

End Sub

Public Sub SelectedText(cmb As ComboBox, _
                        Cmd As CommandButton)

  Dim HiLitSelection As String

  HiLitSelection$ = GetSelectedText(VBInstance)
  If LenB(HiLitSelection$) Then
    If InStr(HiLitSelection$, vbNewLine) Then
      If (HiLitSelection$ <> vbNewLine) Then
        HiLitSelection$ = Left$(HiLitSelection$, InStr(HiLitSelection$, vbNewLine))
      End If
    End If
    If LenB(HiLitSelection$) Then
      cmb.SetFocus
      cmb.Text = HiLitSelection$
      Cmd = True
    End If
  End If

End Sub

':) Roja's VB Code Fixer V1.1.2 (7/07/2003 2:03:13 AM) 49 + 944 = 993 Lines Thanks Ulli for inspiration and lots of code.

