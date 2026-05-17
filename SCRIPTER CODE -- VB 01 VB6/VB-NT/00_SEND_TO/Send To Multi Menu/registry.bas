Attribute VB_Name = "Registry"
'----
'VB - Registry - General - Module - Functions to add/delete/modify Registry Keys-VBForums
'http://www.vbforums.com/showthread.php?231560-VB-Registry-General-Module-Functions-to-add-delete-modify-Registry-Keys
'----
'Thu 07-Jun-2018 18:25:25
'----


Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'--------------------------------------------------
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'--------------------------------------------------
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub CreateKey(hKey As Long, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

If lRegResult <> ERROR_SUCCESS Then
  ' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Sub

Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
Dim lRegResult As Long

lRegResult = RegDeleteKey(hKey, strPath)

End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

lRegResult = RegDeleteValue(hCurKey, strValue)

lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

' Set up default value
If Not IsEmpty(Default) Then
  GetSettingString = Default
Else
  GetSettingString = ""
End If

' Open the key and get length of string
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_SZ Then
    ' initialise string buffer and retrieve string
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
    
    ' format string
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      GetSettingString = Left$(strBuffer, intZeroPos - 1)
    Else
      GetSettingString = strBuffer
    End If

  End If

Else
  ' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long

' Set up default value
If Not IsEmpty(Default) Then
  GetSettingLong = Default
Else
  GetSettingLong = 0
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lDataBufferSize = 4       ' 4 bytes = 32 bits = long

lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_DWORD Then
    GetSettingLong = lBuffer
  End If

Else
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)

If lRegResult <> ERROR_SUCCESS Then
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
Dim lValueType As Long
Dim byBuffer() As Byte
Dim lDataBufferSize As Long
Dim lRegResult As Long
Dim hCurKey As Long

' setup default value
If Not IsEmpty(Default) Then
  If VarType(Default) = vbArray + vbByte Then
    GetSettingByte = Default
  Else
    GetSettingByte = 0
  End If

Else
  GetSettingByte = 0
End If

' Open the key and get number of bytes
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

  If lValueType = REG_BINARY Then
  
    ' initialise buffers and retrieve value
    ReDim byBuffer(lDataBufferSize - 1) As Byte
    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
    
    GetSettingByte = byBuffer

  End If

Else
  'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)

End Function

Public Sub SaveSettingByte(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, byData() As Byte)
' Make sure that the array starts with element 0 before passing it!
' (otherwise it will not be saved!)

Dim lRegResult As Long
Dim hCurKey As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

' Pass the first array element and length of array
lRegResult = RegSetValueEx(hCurKey, strValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)

lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function GetAllKeys(hKey As Long, strPath As String) As Variant
' Returns: an array in a variant of strings

Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim strNames() As String
Dim intZeroPos As Integer

lCounter = 0

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do

  'initialise buffers (longest possible length=255)
  lDataBufferSize = 255
  strBuffer = String(lDataBufferSize, " ")
  lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)

  If lRegResult = ERROR_SUCCESS Then
  
    'tidy up string and save it
    ReDim Preserve strNames(lCounter) As String
    
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      strNames(UBound(strNames)) = Left$(strBuffer, intZeroPos - 1)
    Else
      strNames(UBound(strNames)) = strBuffer
    End If

    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop

GetAllKeys = strNames
End Function

Public Function GetAllValues(hKey As Long, strPath As String) As Variant
' Returns: a 2D array.
' (x,0) is value name
' (x,1) is value type (see constants)

Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim strValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim strNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

Do
  ' Initialise bufffers
  lValueNameSize = 255
  strValueName = String$(lValueNameSize, " ")
  lDataBufferSize = 4000
  
  lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
  
  If lRegResult = ERROR_SUCCESS Then
    
    ' Save the type
    ReDim Preserve strNames(lCounter) As String
    ReDim Preserve lTypes(lCounter) As Long
    lTypes(UBound(lTypes)) = lValueType
    
    'Tidy up string and save it
    intZeroPos = InStr(strValueName, Chr$(0))
    If intZeroPos > 0 Then
      strNames(UBound(strNames)) = Left$(strValueName, intZeroPos - 1)
    Else
      strNames(UBound(strNames)) = strValueName
    End If

    lCounter = lCounter + 1

  Else
    Exit Do
  End If
Loop

'Move data into array
Dim Finisheddata() As Variant
ReDim Finisheddata(UBound(strNames), 0 To 1) As Variant

For lCounter = 0 To UBound(strNames)
  Finisheddata(lCounter, 0) = strNames(lCounter)
  Finisheddata(lCounter, 1) = lTypes(lCounter)
Next

GetAllValues = Finisheddata

End Function
