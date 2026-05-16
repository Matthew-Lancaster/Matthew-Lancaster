VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDHQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Thanks to ShareVB

Option Explicit

Private Const ERROR_SUCCESS As Long = 0&

' counter type
'==========================================
'text
Private Const PERF_TYPE_TEXT As Long = &H800
'number
Private Const PERF_TYPE_NUMBER As Long = &H0
'ASCII text
Private Const PERF_TEXT_ASCII As Long = &H10000
'UNICODE text
Private Const PERF_TEXT_UNICODE As Long = &H0
'hex
Private Const PERF_NUMBER_HEX As Long = &H0
'decimal
Private Const PERF_NUMBER_DECIMAL As Long = &H10000

Private Const PERF_NUMBER_DEC_1000 As Long = &H20000
'no suffix
Private Const PERF_DISPLAY_NO_SUFFIX As Long = &H0
'no show
Private Const PERF_DISPLAY_NOSHOW As Long = &H40000000
' per second
Private Const PERF_DISPLAY_PER_SEC As Long = &H10000000
'percent
Private Const PERF_DISPLAY_PERCENT As Long = &H20000000
'seconds
Private Const PERF_DISPLAY_SECONDS As Long = &H30000000

'format double
Private Const PDH_FMT_DOUBLE As Long = (&H200)

'type winow
Private Const BDC_IncludeInstanceIndex As Long = &H1
Private Const BDC_SingleCounterPerAdd As Long = &H2
Private Const BDC_SingleCounterPerDialog As Long = &H4
Private Const BDC_LocalCountersOnly As Long = &H8
Private Const BDC_WildCardInstances As Long = &H10
Private Const BDC_HideDetailBox As Long = &H20
Private Const BDC_InitializePath As Long = &H40
Private Const BDC_DisableMachineSelection As Long = &H80
Private Const BDC_IncludeCostlyObjects As Long = &H100
Private Const BDC_ShowObjectBrowser As Long = &H200

'the user pressed cancel
Private Const PDH_DIALOG_CANCELLED As Long = &H800007D9

' counter description
Private Type TCounterDesc
    hCounter As Long  'counter handle
    szCounterName As String   'counter name
End Type

'contain a counter value
Private Type PDH_FMT_COUNTERVALUE
    CStatus As Long
    '  case Cardinal of
    '    0: (longValue: Integer);
    '    1: (doubleValue: Double);
    '    2: (largeValue: Int64);
    '    3: (AnsiStringValue: PChar);
    '    4: (WideStringValue: PWideChar);
    padding As Long
    ulValueLow As Long
    ulValueHigh As Long
End Type

'for the window that permits to choose the counter
Private Type PDH_BROWSE_DLG_CONFIG
    ' Configuration flags
    '    {DWORD   bIncludeInstanceIndex:1,
    '            bSingleCounterPerAdd:1,
    '            bSingleCounterPerDialog:1,
    '            bLocalCountersOnly:1,
    '            bWildCardInstances:1,
    '            bHideDetailBox:1,
    '            bInitializePath:1,
    '            bDisableMachineSelection:1,
    '            bIncludeCostlyObjects:1,
    '            bShowObjectBrowser:1,
    '            bReserved:22;}
    dwConfigFlags As Long
    
    hWndOwner As Long
    szDataSource As Long
    szReturnPathBuffer As Long
    cchReturnPathLength As Long
    pCallBack As Long
    dwCallBackArg As Long
    CallBackStatus As Long
    dwDefaultDetailLevel As Long
    szDialogBoxCaption As Long
End Type

'open a query
Private Declare Sub PdhOpenQuery Lib "pdh.dll" (ByVal szDataSource As Long, ByVal dwUserData As Long, ByRef phQuery As Long)
'collect the counter values from a query
Private Declare Function PdhCollectQueryData Lib "pdh.dll" (ByVal hQuery As Long) As Long
'close a query
Private Declare Sub PdhCloseQuery Lib "pdh.dll" (ByVal hQuery As Long)

'show a window "Add a counter"
Private Declare Function PdhBrowseCounters Lib "pdh.dll" Alias "PdhBrowseCountersW" (ByRef pBrowseDlgData As PDH_BROWSE_DLG_CONFIG) As Long
'get the counter index with his name
Private Declare Sub PdhLookupPerfIndexByName Lib "pdh.dll" Alias "PdhLookupPerfIndexByNameA" (ByVal szMachineName As String, ByVal szNameBuffer As String, ByRef pdwIndex As Long)
'get the counter name with his index
Private Declare Sub PdhLookupPerfNameByIndex Lib "pdh.dll" Alias "PdhLookupPerfNameByIndexA" (ByVal szMachineName As String, ByVal dwNameIndex As Long, ByVal szNameBuffer As String, ByRef pcchNameBufferSize As Long)
'TRUE if the counter path is ok
Private Declare Function PdhValidatePath Lib "pdh.dll" Alias "PdhValidatePathA" (ByVal szFullPathBuffer As String) As Long

'add a counter to a query
Private Declare Function PdhAddCounter Lib "pdh.dll" Alias "PdhAddCounterA" (ByVal hQuery As Long, ByVal szFullCounterPath As String, ByVal dwUserData As Long, ByRef phCounter As Long) As Long
Private Declare Function PdhGetFormattedCounterValue Lib "pdh.dll" (ByVal hCounter As Long, ByVal dwFormat As Long, ByRef lpdwType As Long, ByRef pValue As PDH_FMT_COUNTERVALUE) As Long
'delete a counter of a query
Private Declare Function PdhRemoveCounter Lib "pdh.dll" (ByVal hCounter As Long) As Long

Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' query handle
Private hQuery As Long
' counters of the query
Private hCounters() As TCounterDesc
' counters number
Private nCounters As Long


Public Sub Class_Initialize()
    PdhOpenQuery ByVal 0&, ByVal 0&, hQuery
End Sub

Public Sub Class_Terminate()
    PdhCloseQuery hQuery
End Sub

Property Get Count() As Long
    Count = nCounters
End Property

' TRUE if the counterpath is ok
Public Function IsCounterPath(CounterPath As String) As Boolean
    IsCounterPath = (PdhValidatePath(CounterPath) = ERROR_SUCCESS)
End Function

' get the counter index from his counter path
Private Function GetIndexFromCounterPath(CounterPath As String) As Long
    Dim i As Long
    GetIndexFromCounterPath = -1
    For i = 0 To nCounters - 1
        If (hCounters(i).szCounterName = CounterPath) Then
            GetIndexFromCounterPath = i
            Exit For
        End If
    Next
End Function

' Format the counter value
' lpValue : counter value
' dwType : counter type
Private Function GetDataToShow(lpValue As PDH_FMT_COUNTERVALUE, dwType As Long) As String
    Dim doubleValue As Double
    ' TEXT
    If (dwType And &HF00&) = PERF_TYPE_TEXT Then
        ' UNICODE
        If (dwType And &HF0000) = PERF_TEXT_UNICODE Then
            ' convert the string
            GetDataToShow = Space(lstrlenW(lpValue.ulValueLow))
            CopyMemory ByVal StrPtr(GetDataToShow), ByVal lpValue.ulValueLow, Len(GetDataToShow)
        ' ASCII
        ElseIf (dwType And &HF0000) = PERF_TEXT_ASCII Then
            ' copy the string
            GetDataToShow = Space(lstrlenA(lpValue.ulValueLow))
            CopyMemory ByVal StrPtr(GetDataToShow), ByVal lpValue.ulValueLow, Len(GetDataToShow)
            GetDataToShow = StrConv(GetDataToShow, vbFromUnicode)
        End If
    'Numeric value
    Else
        CopyMemory ByVal VarPtr(doubleValue), lpValue.ulValueLow, 8&
        ' hexa
        If (dwType And &HF0000) = PERF_NUMBER_HEX Then
            GetDataToShow = Hex$(Round(doubleValue, 8))
        Else
            GetDataToShow = CStr(doubleValue)
        End If
    End If
    ' if the value needs a suffix
    If (dwType And &HF0000000) = PERF_DISPLAY_PER_SEC Then
        GetDataToShow = GetDataToShow + "/sec"
    ElseIf (dwType And &HF0000000) = PERF_DISPLAY_PERCENT Then
        GetDataToShow = GetDataToShow '+ "%"
    ElseIf (dwType And &HF0000000) = PERF_DISPLAY_SECONDS Then
        GetDataToShow = GetDataToShow + "secs"
    End If
End Function

' Add a counter with his path
' example : "\Process(explorer)\% Processor Time"
Public Function AddCounter(CounterPath As String) As Long
    If IsCounterPath(CounterPath) = False Then
        ' error
        AddCounter = -1
        Exit Function
    End If
    ' + 1 counter
    nCounters = nCounters + 1
    ReDim Preserve hCounters(nCounters - 1)
    ' add to query
    AddCounter = PdhAddCounter(hQuery, CounterPath, ByVal 0&, hCounters(nCounters - 1).hCounter)
    ' save his name
    hCounters(nCounters - 1).szCounterName = CounterPath
    ' if error
    If AddCounter <> ERROR_SUCCESS Then
        ' delete from the list
        nCounters = nCounters - 1
        If nCounters Then
            ReDim Preserve hCounters(nCounters - 1)
        Else
            Erase hCounters
        End If
    End If
End Function

' Delete a counter by his index
Public Function RemoveCounter(CounterIndex As Long) As Boolean
    Dim i As Long
    If CounterIndex < 0 Then Exit Function
    If (CounterIndex < nCounters) Then
        RemoveCounter = (PdhRemoveCounter(hCounters(CounterIndex).hCounter) = ERROR_SUCCESS)
        For i = CounterIndex To nCounters - 2
            hCounters(i).hCounter = hCounters(i + 1).hCounter
            hCounters(i).szCounterName = hCounters(i + 1).szCounterName
        Next
        nCounters = nCounters - 1
        If nCounters Then
            ReDim Preserve hCounters(nCounters - 1)
        Else
            Erase hCounters
        End If
    Else
        RemoveCounter = False
    End If
End Function

' Collects the counter values
Public Function Collect() As Long
    Collect = PdhCollectQueryData(hQuery)
End Function

' Get the counter value
Public Function GetCounterData(CounterIndex As Long) As String
    Dim dwType As Long    ' counter type
    Dim lpValue As PDH_FMT_COUNTERVALUE    ' counter value
    If (CounterIndex < nCounters) Then
        If (PdhGetFormattedCounterValue(hCounters(CounterIndex).hCounter, PDH_FMT_DOUBLE, dwType, lpValue) = ERROR_SUCCESS) Then
            ' format the value
            GetCounterData = GetDataToShow(lpValue, dwType)
        Else
            GetCounterData = ""
        End If
    Else
        GetCounterData = "End"
    End If
End Function


' Format a counter
' Object : the object to watch (processor, process etc...)
' Instance : the instance to watch (process name etc...)
' Counter : the counter type (% Processor Time etc...)
Public Function GetFormatedCounter(lObject As Long, strInstance As String, lCounter As Long) As String
    Dim ObjectName  As String
    Dim NameLen     As Long
    Dim CounterName As String
    ' Get the string for the object
    NameLen = 260
    ObjectName = Space$(NameLen)
    PdhLookupPerfNameByIndex ByVal vbNullString, lObject, ObjectName, NameLen
    ObjectName = Left$(ObjectName, NameLen - 1)
    ' Get the string for the counter
    NameLen = 260
    CounterName = Space$(NameLen)
    PdhLookupPerfNameByIndex ByVal vbNullString, lCounter, CounterName, NameLen
    CounterName = Left$(CounterName, NameLen - 1)
    ' Format
    GetFormatedCounter = "\" + ObjectName + "(" + strInstance + ")\" + CounterName
End Function

