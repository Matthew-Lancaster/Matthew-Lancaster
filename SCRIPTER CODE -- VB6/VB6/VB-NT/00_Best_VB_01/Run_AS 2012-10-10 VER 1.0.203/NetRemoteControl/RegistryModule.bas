Attribute VB_Name = "RegistryModule"
Option Explicit

Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_SZ = 1                         ' Unicode nul terminated string

Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CLASSES_ROOT = &H80000000

Public Const INVALID_HANDLE_VALUE = -1
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.


Function AddValue(txtData As String) As String
Dim Position As Integer
Dim RootKey As Long, SubKey As String
Dim Value As String, ValueData As String
Dim ValueType As Long, retVal As Long
Dim Handle As Long, Msg As String
    
    Position = InStr(1, txtData, "#")
    RootKey& = CRootKey(Left(txtData, Position - 1))
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    SubKey = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    Value = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    ValueType& = CValueType(Left(txtData, Position - 1))
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    ValueData = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    retVal = RegOpenKey(RootKey, SubKey, Handle)
    If retVal = ERROR_SUCCESS Then
        If (ValueType = REG_BINARY) Or (ValueType = REG_DWORD) Then
            retVal = RegSetValueEx(Handle, Value, &O0, ValueType, CLng(ValueData), &O4)
        Else
            retVal = RegSetValueEx(Handle, Value, &O0, ValueType, ByVal ValueData, Len(ValueData))
        End If
        If retVal = ERROR_SUCCESS Then
            Msg = "SUCCESS#Value Successfuly Create."
        Else
            Msg = "ERROR#Value not create."
        End If
    Else
        Msg = "ERROR#Value Handle Not Create"
    End If
    RegCloseKey (Handle)
    AddValue = Msg
End Function

Function DeleteValue(txtData As String) As String
Dim Position As Integer
Dim RootKey As Long, SubKey As String
Dim Value As String, Msg As String
Dim Handle As Long, retVal As Long
    
    Position = InStr(1, txtData, "#")
    RootKey = CRootKey(Left(txtData, Position - 1))
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    SubKey = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    Value = Left(txtData, Position - 1)
    
    retVal = RegOpenKey(RootKey, SubKey, Handle)
    If retVal = ERROR_SUCCESS Then
        retVal = RegDeleteValue(Handle, Value)
        If retVal = ERROR_SUCCESS Then
            Msg = "SUCCESS#Value successfuly delete."
        Else
            Msg = "ERROR#Value not delete."
        End If
    Else
        Msg = "ERROR#Value handle not create."
    End If
    RegCloseKey (Handle)
    DeleteValue = Msg
End Function

Function ShowValue(txtData As String) As String
Dim Position As Integer
Dim RootKey As Long, SubKey As String
Dim Value As String, strValueData As String, lngValueData As Long
Dim ValueType As Long, ValueLength As Long
Dim Handle As Long, retVal As Long
Dim Msg As String

    Position = InStr(1, txtData, "#")
    RootKey = CRootKey(Left(txtData, Position - 1))
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    SubKey = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    Value = Left(txtData, Position - 1)
    
    retVal = RegOpenKey(RootKey, SubKey, Handle)
    If retVal = ERROR_SUCCESS Then
        retVal = RegQueryValueEx(Handle, Value, 0&, ValueType, ByVal 0, ValueLength)
        Select Case ValueType
            Case REG_SZ:
                strValueData = String(ValueLength, " ")
                retVal = RegQueryValueEx(Handle, Value, 0&, ValueType, ByVal strValueData, Len(strValueData))
                If retVal = ERROR_SUCCESS Then
                    Msg = "ShowValue#" & strValueData & "#" & "String Value"
                Else
                    Msg = "ERROR#Value not show."
                End If
            Case REG_BINARY:
                retVal = RegQueryValueEx(Handle, Value, 0&, ValueType, ByVal lngValueData, ValueLength)
                If retVal = ERROR_SUCCESS Then
                    Msg = "ShowValue#" & lngValueData & "#" & "Binary Value"
                Else
                    Msg = "ERROR#Value not show."
                End If
            Case REG_DWORD:
                retVal = RegQueryValueEx(Handle, Value, 0&, ValueType, lngValueData, 4&)
                If retVal = ERROR_SUCCESS Then
                    Msg = "ShowValue#" & lngValueData & "#" & "DWORD Value"
                Else
                    Msg = "ERROR#Value not show."
                End If
        End Select
    Else
        Msg = "ERROR#Value handle not create."
    End If
    RegCloseKey (Handle)
    ShowValue = Msg
End Function

Function AddKey(txtData As String) As String
Dim Position As Integer
Dim RootKey As Long, SubKey As String
Dim Key As String
Dim Handle As Long, retVal As Long
Dim Msg As String
    
    Position = InStr(1, txtData, "#")
    RootKey = CRootKey(Left(txtData, Position - 1))
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    SubKey = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    Key = Left(txtData, Position - 1)
    
    retVal = RegOpenKey(RootKey, SubKey, Handle)
    If retVal = ERROR_SUCCESS Then
        retVal = RegCreateKey(Handle, Key, Handle)
        If retVal = ERROR_SUCCESS Then
            Msg = "SUCCESS#Key successfuly create."
        Else
            Msg = "ERROR#Key not create."
        End If
    Else
        Msg = "ERROR#Key handle not create."
    End If
    RegCloseKey (Handle)
    AddKey = Msg
End Function

Function DeleteKey(txtData As String) As String
Dim Position As Integer
Dim RootKey As Long, SubKey As String
Dim Key As String
Dim Handle As Long, retVal As Long
Dim Msg As String
    
    Position = InStr(1, txtData, "#")
    RootKey = CRootKey(Left(txtData, Position - 1))
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    SubKey = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Position = InStr(1, txtData, "#")
    Key = Left(txtData, Position - 1)
    
    retVal = RegOpenKey(RootKey, SubKey, Handle)
    If retVal = ERROR_SUCCESS Then
        retVal = RegDeleteKey(Handle, Key)
        If retVal = ERROR_SUCCESS Then
            Msg = "SUCCESS#Key successfuly delete."
        Else
            Msg = "ERROR#Key not delete."
        End If
    Else
        Msg = "ERROR#Key handle not create."
    End If
    RegCloseKey (Handle)
    DeleteKey = Msg
End Function
Function CRootKey(RootKey As String) As Long
    Select Case RootKey
        Case "HKEY_CURRENT_CONFIG"
            CRootKey = HKEY_CURRENT_CONFIG
        Case "HKEY_CURRENT_USER"
            CRootKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            CRootKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            CRootKey = HKEY_USERS
        Case "HKEY_CLASSES_ROOT"
    End Select
End Function

Function CValueType(ValueType) As Long
    Select Case ValueType
        Case "Binary Value"
            CValueType = REG_BINARY
        Case "DWORD Value"
            CValueType = REG_DWORD
        Case "String Value"
            CValueType = REG_SZ
    End Select
End Function
