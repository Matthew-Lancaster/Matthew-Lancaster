Attribute VB_Name = "modRegistry"
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0&
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = (1)

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Function bSetRegValue(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
    
    On Error Resume Next
    Dim phkResult As Long
    Dim lResult As Long
    Dim SA As SECURITY_ATTRIBUTES
    Dim lCreate As Long
    RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, lCreate
    lResult = RegSetValueEx(phkResult, sSetValue, 0, REG_SZ, sValue, CLng(Len(sValue) + 1))
    RegCloseKey phkResult
    bSetRegValue = (lResult = ERROR_SUCCESS)
    
End Function


Public Function bGetRegValue(ByVal hKey As Long, ByVal sKey As String, ByVal sSubKey As String) As String
    
    Dim lResult As Long
    Dim phkResult As Long
    Dim dWReserved As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    Dim szBuffer2 As String
    Dim lBuffSize2 As Long
    Dim lIndex As Long
    Dim lType As Long
    Dim sCompKey As String
    Dim bFound As Boolean
    
    lIndex = 0
    lResult = RegOpenKeyEx(hKey, sKey, 0, 1, phkResult)


    Do While lResult = ERROR_SUCCESS And Not (bFound)
        szBuffer = Space(255)
        lBuffSize = Len(szBuffer)
        szBuffer2 = Space(255)
        lBuffSize2 = Len(szBuffer2)
        lResult = RegEnumValue(phkResult, lIndex, szBuffer, lBuffSize, dWReserved, lType, szBuffer2, lBuffSize2)


        If (lResult = ERROR_SUCCESS) Then
            sCompKey = Left(szBuffer, lBuffSize)


            If (sCompKey = sSubKey) Then
                bGetRegValue = Left(szBuffer2, lBuffSize2 - 1)
            End If
        End If
        lIndex = lIndex + 1
        
    Loop
    RegCloseKey phkResult
End Function

Public Sub RegisterFileExtension(sExt As String)

bSetRegValue HKEY_CLASSES_ROOT, "." & sExt, "", "FStudio.paramfile"
bSetRegValue HKEY_CLASSES_ROOT, "FStudio.paramfile", "", "FStudio Param File"
bSetRegValue HKEY_CLASSES_ROOT, "FStudio.paramfile\DefaultIcon", "", aData.szAppPath & "Fstudio.exe, 1"
bSetRegValue HKEY_CLASSES_ROOT, "FStudio.paramfile\shell", "", ""
bSetRegValue HKEY_CLASSES_ROOT, "FStudio.paramfile\shell\open\command", "", aData.szAppPath & "FStudio.exe %1"

End Sub
