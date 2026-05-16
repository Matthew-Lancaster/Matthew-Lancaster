Attribute VB_Name = "support"
Option Explicit
Public mobjDoc                   As docfind
Public VBInstance                As VBE
Public ProjCount                 As Long
Public CompCount                 As Long
Public bLaunchOnStart            As Boolean
Public bSaveHistory              As Boolean
Public bBlankWarning             As Boolean
Public bFilterWarning            As Boolean
Public bReplace2Search           As Boolean
Public bRemFilters               As Boolean
Public HistDeep                  As Long
Public bLoadingSettings          As Boolean
Public bAutoSelectText           As Boolean
Public ColourTextFore            As Long
Public ColourTextBack            As Long
Public ColourFindSelectBack      As Long
Public ColourFindSelectFore      As Long
Public ColourHeadFore            As Long
Public ColourHeadDefault         As Long
Public ColourHeadWork            As Long
Public ColourHeadPattern         As Long
Public ColourHeadNoFind          As Long

Public Function AppDetails() As String

  With App
    AppDetails = .ProductName & " V" & .Major & "." & .Minor & "." & .Revision
  End With

End Function

Public Function Bool2Int(bValue As Boolean) As Integer

  'used to simplify settings

  Bool2Int = IIf(bValue, 1, 0)

End Function

Public Function Bool2Str(bValue As Boolean) As String

  'used to simplify settings

  Bool2Str = IIf(bValue, "1", "0")

End Function

Public Sub GetCounts()

  'the counts are used to control column visiblity
  
  Dim Proj      As VBProject

  On Error Resume Next
  ProjCount = VBInstance.VBProjects.Count
  'CompCount includes ProjCount just in case a group includes projects with only one component
  For Each Proj In VBInstance.VBProjects
    CompCount = CompCount + Proj.VBComponents.Count + IIf(ProjCount > 1, 1, 0)
  Next Proj
  On Error GoTo 0

End Sub

Public Function GetSelectedText(VBInstance As vbide.VBE) As String

  Dim StartLine As Long
  Dim cmo       As vbide.CodeModule
  Dim codeText  As String
  Dim cpa       As vbide.CodePane
  Dim EndCol    As Long
  Dim EndLine   As Long
  Dim StartCol  As Long

  'Date: 4/27/1999
  'Versions: VB5 VB6 Level: Intermediate
  'Author: The VB2TheMax Team
  ' Return the string of code the is selected in the code window
  ' that is currently active.
  ' This function can only be used inside an add-in.
  On Error Resume Next
  ' get a reference to the active code window and the underlying module
  ' exit if no one is available
  Set cpa = VBInstance.ActiveCodePane
  Set cmo = cpa.CodeModule
  If Err.Number Then
    Exit Function
  End If
  ' get the current selection coordinates
  cpa.GetSelection StartLine, StartCol, EndLine, EndCol
  ' exit if no text is highlighted
  If StartLine = EndLine Then
    If StartCol = EndCol Then
      Exit Function
    End If
  End If
  ' get the code text
  If StartLine = EndLine Then
    ' only one line is partially or fully highlighted
    codeText = Mid$(cmo.Lines(StartLine, 1), StartCol, EndCol - StartCol)
   Else
    ' the selection spans multiple lines of code
    ' first, get the selection of the first line
    codeText = Mid$(cmo.Lines(StartLine, 1), StartCol) & vbNewLine
    ' then get the lines in the middle, that are fully highlighted
    If StartLine + 1 < EndLine Then
      codeText = codeText & cmo.Lines(StartLine + 1, EndLine - StartLine - 1)
    End If
    ' finally, get the highlighted portion of the last line
    codeText = codeText & Left$(cmo.Lines(EndLine, 1), EndCol - 1)
  End If
  GetSelectedText = codeText
  On Error GoTo 0

End Function

Public Function instrAny(ByVal strSearch As String, _
                         ParamArray finds() As Variant) As Long

  Dim fnd    As Variant
  Dim TMpAny As Long

  instrAny = LongLimit
  For Each fnd In finds
    TMpAny = InStr(strSearch, fnd)
    If TMpAny < instrAny Then
      If TMpAny > 0 Then
        instrAny = TMpAny
      End If
    End If
  Next fnd
  If instrAny = LongLimit Then
    instrAny = 0
  End If

End Function

Public Function MultiLeft(ByVal varSearch As Variant, _
                          ByVal CaseSensitive As Boolean, _
                          ParamArray Afind() As Variant) As Boolean

  Dim FindIt As Variant

  'This routine was originally designed to test multiple possible left strings
  'BUT I also use it as a simple way of testing even a single left string
  'without having to separately code the length at every instance
  If Not CaseSensitive Then
    varSearch = LCase$(varSearch)
  End If
  For Each FindIt In Afind
    If CaseSensitive Then
      If Left$(varSearch, Len(FindIt)) = FindIt Then
        MultiLeft = True
        Exit Function
      End If
     Else
      If Left$(varSearch, Len(FindIt)) = LCase$(FindIt) Then
        MultiLeft = True
        Exit Function
      End If
    End If
  Next FindIt

End Function

Public Function MultiRight(ByVal varSearch As Variant, _
                           ByVal CaseSensitive As Boolean, _
                           ParamArray Afind() As Variant) As Boolean

  Dim FindIt As Variant

  'This routine was originally designed to test multiple possible left strings
  'BUT I also use it as a simple way of testing even a single left string
  'without having to separately code the length at every instance
  'CaseSensitive was added to solve a problem with hand coding of standard VB routines with wrong case
  If Not CaseSensitive Then
    varSearch = LCase$(varSearch)
  End If
  For Each FindIt In Afind
    If CaseSensitive Then
      If Right$(varSearch, Len(FindIt)) = FindIt Then
        MultiRight = True
        Exit Function
      End If
     Else
      If Right$(varSearch, Len(FindIt)) = LCase$(FindIt) Then
        MultiRight = True
        Exit Function
      End If
    End If
  Next FindIt

End Function

Public Function SafeCompToProcess(ByVal cmp As VBComponent) As Boolean

  'returns True if the component is anything that can be processed by program

  SafeCompToProcess = cmp.Type <> vbext_ct_ResFile And cmp.Type <> vbext_ct_RelatedDocument
  ' this routine is called at start of all rewrite code so
  ' this is a good spot to make sure that suspend is off

End Function

Public Sub SetFocus_Safe(CTL As Control)

  '*PURPOSE: protect SetFocus from any of the many conditions which can stuff it

  On Error Resume Next
  With CTL
    If .Visible Then
      If .Enabled Then
        .SetFocus
      End If
    End If
  End With
  On Error GoTo 0

End Sub

':) Roja's VB Code Fixer V1.1.2 (7/07/2003 2:03:00 AM) 23 + 194 = 217 Lines Thanks Ulli for inspiration and lots of code.

