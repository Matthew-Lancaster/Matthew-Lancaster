Attribute VB_Name = "modSort"
Option Explicit
Option Base 0

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' What is a merge sort?  That's a difficult question to demonstrate, but not one to answer
'
' basically, the MergeSort function takes a group of elements, and you split the group in
' half or as near to half as possible.
'
' Then it runs MergeSort on each half.
' If there is two or fewer elements in the array when called, however,
' mergesort compares the elements, sorts them and returns them
' to the previous calling MergeSort.  We are said, at that point,
' to have descended the tree as far as we can.
'
' then, the elements are merged together in order by the Merge function.
' and on back up the chain until the whole thing is in order
'
'
' It's very quick.  It's much quicker than the simple BubbleSort
' routine at the bottom that sorts the two element arrays for the Merge function
'
' Bubble sort's maximum iterations are x^x, Merge sort's maximum iterations are x^log x

Public Function MergeSort(Strings() As String) As String()
    MergeSort = MergeSortString(Strings)
End Function

Public Function MergeSortString(Strings() As String, Optional CompareMethod As VbCompareMethod = vbTextCompare) As String()
    
    Dim s As Long, i As Long, n() As String, m() As String
    Dim u As Long, l As Long, j As Long, x() As String
    
    ' we don't sort 1 element!
    On Error Resume Next
    i = -1&
    i = UBound(Strings)
    If i = -1& Then Exit Function
    
    
    If i = 0 Then
        ReDim n(0)
        n(0) = Strings(0)
        MergeSortString = n
        
        Exit Function
    ElseIf i = 1 Then
        MergeSortString = BubbleSort(Strings, CompareMethod)
        Exit Function
    Else
        
        l = SplitArray(Strings, m, n)
        m = MergeSortString(m, CompareMethod)
        n = MergeSortString(n, CompareMethod)
        
        MergeSortString = MergeArray(m, n, CompareMethod)
    
    End If
       
End Function

Private Function SplitArray(Strings() As String, StringOut1() As String, StringOut2() As String) As Long
' Splits it 50/50 or as close as possible

' if it's just one string, we return StringOut1(0)
' The return number is the total elements of the largest of the arrays
    On Error Resume Next
    
    Dim i As Long, j As Long, s1() As String, s2() As String
    Dim z As Long, x As Long
    
    x = -1&
    x = UBound(Strings)
    If (x = -1&) Then Exit Function
    
    If (x = 0) Then
        ReDim s1(0)
        s1(0) = Strings(0)
        
        StringOut1 = s1
        s1 = Empty
        
        Exit Function
    End If
    
    i = (x + 1) / 2
    
    z = 0
    ReDim s1(0 To (i - 1))
    
    For j = 0 To (i - 1)
        s1(j) = Strings(z)
        z = z + 1
    Next j
    
    i = (x + 1) - i
    ReDim s2(0 To (i - 1))
    
    For j = 0 To (i - 1)
        s2(j) = Strings(z)
        z = z + 1
    Next j

    StringOut1 = s1
    s1 = Empty
    
    StringOut2 = s2
    s2 = Empty

    If ((x + 1) - i) > i Then
        SplitArray = (x + 1) - i
    Else
        SplitArray = i
    End If
    
End Function

Private Function MergeArray(String1() As String, String2() As String, ByVal CompareMethod As VbCompareMethod) As String()

    Dim i As Long, j As Long
    Dim n() As String, x As Long, y As Long, c As Integer
    On Error Resume Next
    
    i = -1&
    j = -1&
    
    i = UBound(String1)
    j = UBound(String2)
    
    If (i < 0) And (j < 0) Then Exit Function
    
    If (i > -1&) Then
        i = i + 1
    ElseIf (i = -1&) Or ((i = 0) And (String1(0) = "")) Then
        MergeArray = String2
        Exit Function
    End If
    
    If j > -1& Then
        i = i + ((UBound(String2) - LBound(String2)) + 1)
    ElseIf (j = -1) Or ((j = 0) And (String2(0) = "")) Then
        MergeArray = String1
        Exit Function
    End If
    
    x = LBound(String1)
    y = LBound(String2)
    
    ReDim n(0 To i - 1)
    
    For j = 0 To (i - 1) Step 0
        If (x > UBound(String1)) And (y > UBound(String2)) Then
            MergeArray = n
            Exit Function
        End If
        
        c = StrComp(String1(x), String2(y), CompareMethod)
        If (c = 0) And (x <= UBound(String1) And (y <= UBound(String2))) Then
            n(j) = String1(x)
            n(j + 1) = String2(y)
            j = j + 2
            y = y + 1
            x = x + 1
        ElseIf ((c < 0) Or (y > UBound(String2))) And (x <= UBound(String1)) Then
            n(j) = String1(x)
            x = x + 1
            j = j + 1
        ElseIf ((c > 0) Or (x > UBound(String1))) And (y <= UBound(String2)) Then
            n(j) = String2(y)
            y = y + 1
            j = j + 1
        End If
    Next
    MergeArray = n
    
End Function

Public Function BubbleSort(Strings() As String, Optional CompareMethod As VbCompareMethod = vbTextCompare) As String()

    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Integer, f As Integer, g As Integer

    Dim i As String, j As String
    Dim m() As String, n() As String
    
    e = 1
    ReDim n(LBound(Strings) To UBound(Strings))
    
    For d = LBound(n) To UBound(n)
        n(d) = Strings(d)
    Next d
    
    Do While e <> -1
    
        For a = 0 To UBound(Strings) - 1
            i = n(a)
            j = n(a + 1)
            f = StrComp(i, j, CompareMethod)
            If f <= -1 Then
                n(a) = i
                n(a + 1) = j
            Else
                n(a) = j
                n(a + 1) = i
                g = 1
            End If
        Next a
        If g = 1 Then
            e = 1
        Else
            e = -1
        End If
        
        g = 0
    Loop
    BubbleSort = n
End Function
