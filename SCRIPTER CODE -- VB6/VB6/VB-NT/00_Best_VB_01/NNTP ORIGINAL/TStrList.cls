VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim Cnt As Long 'number of stored items
Dim Max As Long 'upper limit of current dimension
Dim Dat() As String 'place for items (strings)
Dim Idx() As Long 'sort-index
Dim Tag1 As Variant 'any data
Dim Sort1 As TSort 'sort flag

'types of sort
Public Enum TSort
    None = 0
    Ascending = 1
    Descending = 2
End Enum

Public Sub Add(Data As String, Optional Index As Long = 0)
Attribute Add.VB_Description = "Adds an item to the list and increment the count property."
    On Error Resume Next
    Dim i As Integer, j As Integer
    
    'add an item
    Cnt = Cnt + 1
    
    'widen array if required
    If Cnt > Max Then
        If Max = 0 Then
            Max = 512
        Else
            Max = Max + Max
        End If
        ReDim Preserve Dat(1 To Max)
        ReDim Preserve Idx(1 To Max)
    End If
    
    'add data
    Dat(Cnt) = Data
    Idx(Cnt) = Cnt
    
    'insert at given position
    If Index > 0 Then
        For i = Cnt To Index + 1 Step -1
            Idx(i) = Idx(i - 1)
        Next i
        Idx(Index) = Cnt
    End If
    
    'repair index if sorted
    RepairIndex Cnt
End Sub

Public Property Get Item(ByVal Index As Long) As String
Attribute Item.VB_Description = "Set or return the value of an item."
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    'out of range
    If Index < 1 Or Index > Cnt Then
        Err.Raise 9
    'returns an item
    Else
        Item = Dat(Idx(Index))
    End If
End Property

Public Property Let Item(ByVal Index As Long, vNewValue As String)
    On Error Resume Next
    'out of range
    If Index < 1 Or Index > Cnt Then
        Err.Raise 9
    'sets an item
    Else
        Dat(Idx(Index)) = vNewValue
        'repair index if sorted
        RepairIndex Index
    End If
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the complete itemlist and sets the count property to zero."
    On Error Resume Next
    'clear data
    Cnt = 0
    Max = 0
    ReDim Dat(1 To 1)
    ReDim Idx(1 To 1)
End Sub

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the list."
    On Error Resume Next
    'returns the number of items currently stored
    Count = Cnt
End Property

Public Sub Remove(Index As Long)
Attribute Remove.VB_Description = "Removes an item from the list."
    On Error Resume Next
    Dim Old As Long, i As Long
    
    'out of range
    If Index < 1 Or Index > Cnt Then
        Err.Raise 9
    
    'remove an item
    Else
        Old = Idx(Index)
        
        'remove it
        For i = Index To Cnt - 1
            Dat(i) = Dat(i + 1)
            Idx(i) = Idx(i + 1)
        Next i
        
        'repair index
        For i = 1 To Cnt
            If Idx(i) > Old Then
                Idx(i) = Idx(i) - 1
            End If
        Next i
        
        'decrease count
        Cnt = Cnt - 1
    End If
End Sub

Public Property Get Tag() As Variant
    On Error Resume Next
    'returns the tag
    Tag = Tag1
End Property

Public Property Let Tag(ByVal vNewValue As Variant)
    On Error Resume Next
    'sets the tag
    Tag1 = vNewValue
End Property

Public Function Find(Text As String) As Long
    On Error Resume Next
    Dim i As Long, j As Long, Lb As Long, Ub As Long
    
    'find an item
    If Cnt > 0 Then
        
        'unsorted - search from start to end of array
        If Sort1 = None Then
            For i = 1 To Cnt
                If Dat(Idx(i)) = Text Then
                    Find = i
                    Exit Function
                End If
            Next i
            
        'sorted - quickfind
        Else
            Lb = 1
            Ub = Cnt
            Do
                i = Lb + (Ub - Lb) / 2
                'found
                If Dat(Idx(i)) = Text Then
                    Find = i
                    Exit Function
                ElseIf Lb >= Ub Then
                    Exit Function
                'continue search
                ElseIf (Sort1 = Ascending And Dat(Idx(i)) < Text) Or (Sort1 = Descending And Dat(Idx(i)) > Text) Then
                    Lb = i + 1
                Else
                    Ub = i - 1
                End If
            Loop
                        
        End If
    End If
End Function

Public Property Get Sort() As TSort
    On Error Resume Next
    'returns the sort setup
    Sort = Sort1
End Property

Public Property Let Sort(ByVal vNewValue As TSort)
    On Error Resume Next
    Dim i As Long, j As Long, k As Long, l As Long, X As Long
    Dim Si() As Double
    
    'sets the sort value
    If vNewValue <> Sort1 Then
        Sort1 = vNewValue
        
        'sorting required?
        If Cnt > 1 Then
        
            'completely resort
            If (Cnt > 1) And (vNewValue = Ascending Or vNewValue = Descending) Then
                
                'count chars which are equal from the beginning
                X = Len(Dat(1))
                For i = 2 To Cnt
                    l = Len(Dat(i))
                    If l < X Then X = l
                    For j = 1 To X
                        If Mid$(Dat(i), j, 1) <> Mid$(Dat(1), j, 1) Then
                            X = j - 1
                            Exit For
                        End If
                    Next j
                Next i
                
                'create index on the first 8 chars
                ReDim Si(1 To Cnt)
                For i = 1 To Cnt
                    For j = 1 To IIf(8 < Len(Dat(i)) - X, 8, Len(Dat(i)) - X)
                        Si(i) = Si(i) + Asc(UCase(Mid$(Dat(i), X + j, 1))) * 256 ^ (8 - j)
                    Next j
                Next i
                
                'sort ascending
                If Sort1 = Ascending Then
                    For i = 1 To Cnt - 1
                        For j = i + 1 To Cnt
                            If Si(Idx(i)) > Si(Idx(j)) Then
                                k = Idx(i)
                                Idx(i) = Idx(j)
                                Idx(j) = k
                            ElseIf Si(Idx(i)) = Si(Idx(j)) Then
                                If Dat(Idx(i)) > Dat(Idx(j)) Then
                                    k = Idx(i)
                                    Idx(i) = Idx(j)
                                    Idx(j) = k
                                End If
                            End If
                        Next j
                    Next i
            
                'sort descending
                Else
                    For i = 1 To Cnt - 1
                        For j = i + 1 To Cnt
                            If Si(Idx(i)) < Si(Idx(j)) Then
                                k = Idx(i)
                                Idx(i) = Idx(j)
                                Idx(j) = k
                            ElseIf Si(Idx(i)) = Si(Idx(j)) Then
                                If Dat(Idx(i)) < Dat(Idx(j)) Then
                                    k = Idx(i)
                                    Idx(i) = Idx(j)
                                    Idx(j) = k
                                End If
                            End If
                        Next j
                    Next i
                    
                End If
            End If
        End If
    End If
End Property

Private Sub RepairIndex(Item As Long)
    On Error Resume Next
    Dim i As Long, j As Long
    
    '***this can be optimized -> see http://vb-tec.de/arrmove.htm
    
    'repair index
    If Cnt > 1 Then
        
        'sort ascending
        If Sort1 = Ascending Then
            For i = 1 To Cnt
                If Dat(Idx(i)) > Dat(Idx(Item)) Then
                    j = Idx(Item)
                    Idx(Item) = Idx(i)
                    Idx(i) = j
                End If
            Next i
        
        'sort descending
        ElseIf Sort1 = Descending Then
            For i = 1 To Cnt
                If Dat(Idx(i)) < Dat(Idx(Item)) Then
                    j = Idx(Item)
                    Idx(Item) = Idx(i)
                    Idx(i) = j
                End If
            Next i
        
        End If
    End If
End Sub

