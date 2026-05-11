Attribute VB_Name = "modkWab"
Option Explicit
'Public Halo

Public Type tWab
   AddrBook As Long
   WabObject As Long
   InstWab As Long
End Type

Public Type SBinary
   cb As Long
   lpb As Long
End Type

Public Type SysTime
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type LONGLONG
   Word1 As Long
   Word2 As Long
End Type

Public Type CLSID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4_0 As Byte
   Data4_1 As Byte
   Data4_2 As Byte
   Data4_3 As Byte
   Data4_4 As Byte
   Data4_5 As Byte
   Data4_6 As Byte
   Data4_7 As Byte
End Type

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Public Declare Function kWabOpen Lib "kwab.dll" (ByRef kw As tWab) As Boolean
Public Declare Sub kWabClose Lib "kwab.dll" (ByRef kw As tWab)
Public Declare Function kWabGetNumEntries Lib "kwab.dll" (ByRef kw As tWab) As Long
Public Declare Function kWabGetEntries Lib "kwab.dll" (ByRef kw As tWab, ByRef FirstElementOfEntriesArray As SBinary, ByVal nBuf As Long) As Long
Public Declare Sub kWabFreeEntries Lib "kwab.dll" (ByRef kw As tWab, ByRef FirstElementOfEntriesArray As SBinary, ByVal nBuf As Long)
Public Declare Function kWabDetails Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal hWnd As Long) As Boolean

Public Declare Function kWabGetNumProps Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary) As Long
Public Declare Function kWabGetPropTag Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropertyIndex As Long, ByVal UseUnicode As Boolean) As Long
Public Declare Function kWabGetPropID Lib "kwab.dll" (ByVal PropTag As Long) As Long
Public Declare Function kWabGetPropType Lib "kwab.dll" (ByVal PropTag As Long) As Long

Public Declare Function kWabGetProp_String Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal Str As String, ByVal LnStr As Long, ByVal UseUnicode As Boolean) As Boolean
Public Declare Function kWabGetProp_LongLong Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As LONGLONG) As Boolean
Public Declare Function kWabGetProp_Long Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Long) As Boolean
Public Declare Function kWabGetProp_Short Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Integer) As Boolean
Public Declare Function kWabGetProp_SysTime Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef ft As SysTime) As Boolean
Public Declare Function kWabGetProp_Binary_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef NumBytes As Long) As Boolean
Public Declare Function kWabGetProp_Binary_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef FirstElementOfArray As Byte, ByRef NumBytes As Long) As Boolean
Public Declare Function kWabGetProp_Boolean Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Boolean) As Boolean
Public Declare Function kWabGetProp_Float Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Single) As Boolean
Public Declare Function kWabGetProp_Double Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Double) As Boolean
Public Declare Function kWabGetProp_Currency Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Currency) As Boolean
Public Declare Function kWabGetProp_AppTime Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Date) As Boolean
Public Declare Function kWabGetProp_CLSID Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As CLSID) As Boolean

Public Declare Function kWabGetProp_MV_String_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_String_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal Str As String, ByVal LnStr As Long, ByVal UseUnicode As Boolean, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_LongLong_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_LongLong_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As LONGLONG, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_Long_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_Long_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Long, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_Short_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_Short_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Integer, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_SysTime_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_SysTime_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef ft As SysTime, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_Binary_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_Binary_Data_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef NumBytes As Long, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_Binary_Data_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef FirstElementOfArray As Byte, ByRef NumBytes As Long, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_Float_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_Float_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Single, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_Double_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_Double_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Double, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_Currency_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_Currency_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Currency, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_AppTime_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_AppTime_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Date, ByVal idx As Long) As Boolean
Public Declare Function kWabGetProp_MV_CLSID_Count Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Count As Long) As Boolean
Public Declare Function kWabGetProp_MV_CLSID_Data Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As CLSID, ByVal idx As Long) As Boolean

Public Declare Function kWabDeleteEntry Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary) As Boolean
Public Declare Function kWabDeleteProperty Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long) As Boolean
Public Declare Function kWabAddNewGUI Lib "kwab.dll" (ByRef kw As tWab, ByVal hWnd As Long) As Boolean
Public Declare Function kWabAddNew Lib "kwab.dll" (ByRef kw As tWab, ByVal DisplayName As String, ByVal UseUnicode As Boolean) As Long


Public Declare Function kWabSetProp_Short Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Integer) As Boolean
Public Declare Function kWabSetProp_Long Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Long) As Boolean
Public Declare Function kWabSetProp_LongLong Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As LONGLONG) As Boolean
Public Declare Function kWabSetProp_SysTime Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As SysTime) As Boolean
Public Declare Function kWabSetProp_Boolean Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Boolean) As Boolean
Public Declare Function kWabSetProp_Float Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Single) As Boolean
Public Declare Function kWabSetProp_Double Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Double) As Boolean
Public Declare Function kWabSetProp_Currency Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Currency) As Boolean
Public Declare Function kWabSetProp_AppTime Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Date) As Boolean
Public Declare Function kWabSetProp_String Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal Value As String) As Boolean
Public Declare Function kWabSetProp_CLSID Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As CLSID) As Boolean
Public Declare Function kWabSetProp_Binary Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef FirstElementOfArray As Byte, ByVal NumBytes As Long) As Boolean

Public Declare Function kWabSetProp_MV_DelValue Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long) As Boolean

Public Declare Function kWabSetProp_MV_AddValue_Short Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Integer) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_Long Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Long) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_Float Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Single) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_Double Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Double) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_Currency Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Currency) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_AppTime Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As Date) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_SysTime Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As SysTime) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_String Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal Value As String) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_Binary Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef FirstElementOfArrayValue As Byte, ByVal NumberOfBytes As Long) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_CLSID Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As CLSID) As Boolean
Public Declare Function kWabSetProp_MV_AddValue_LongLong Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByRef Value As LONGLONG) As Boolean

Public Declare Function kWabSetProp_MV_EditValue_Short Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As Integer) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_Long Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As Long) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_Float Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As Single) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_Double Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As Double) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_Currency Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As Currency) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_AppTime Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As Date) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_SysTime Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As SysTime) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_String Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByVal Value As String) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_Binary Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef FirstElementOfArrayValue As Byte, ByVal NumberOfBytes As Long) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_CLSID Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As CLSID) As Boolean
Public Declare Function kWabSetProp_MV_EditValue_LongLong Lib "kwab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal ValueNumber As Long, ByRef Value As LONGLONG) As Boolean

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As SysTime, lpSystemTime As SYSTEMTIME) As Long
Public Function ExpandSysTime(inFT As SysTime, ByRef outYear As Integer, ByRef outMonth As Integer, ByRef outDayOfWeek As Integer, ByRef outDay As Integer, ByRef outHour As Integer, ByRef outMinute As Integer, ByRef outSecond As Integer, ByRef outMilliseconds As Integer) As Boolean
   Dim st As SYSTEMTIME

   If FileTimeToSystemTime(inFT, st) <> 0 Then
      With st
         outYear = .wYear
         outMonth = .wMonth
         outDayOfWeek = .wDayOfWeek
         outDay = .wDay
         outHour = .wHour
         outMinute = .wMinute
         outSecond = .wSecond
         outMilliseconds = .wMilliseconds
      End With
      ExpandSysTime = True
   Else
      ExpandSysTime = False
   End If
End Function
Public Function SysTime2Date(inFT As SysTime, ByRef outDate As Date) As Boolean
   Dim y As Integer, m As Integer, dow As Integer, d As Integer, h As Integer, mn As Integer, s As Integer, ms As Integer
   
   SysTime2Date = False
   If ExpandSysTime(inFT, y, m, dow, d, h, mn, s, ms) Then
      outDate = DateSerial(y, m, d) + TimeSerial(h, mn, s)
      SysTime2Date = True
   End If
End Function
Public Function GetDayOfWeek(ByVal dow As Integer) As String
   Select Case dow
      Case 0
         GetDayOfWeek = "sunday"
      Case 1
         GetDayOfWeek = "monday"
      Case 2
         GetDayOfWeek = "tuesday"
      Case 3
         GetDayOfWeek = "wednesday"
      Case 4
         GetDayOfWeek = "thursday"
      Case 5
         GetDayOfWeek = "friday"
      Case 6
         GetDayOfWeek = "saturday"
      Case Else
         GetDayOfWeek = "???"
   End Select
End Function

Public Function CLSID2String(ci As CLSID) As String
   CLSID2String = Right("00000000" & Hex(ci.Data1), 8) & "-" & Right("0000" & Hex(ci.Data2), 4) & "-" & Right("0000" & Hex(ci.Data3), 4) & "-" & Right("00" & Hex(ci.Data4_0), 2) & Right("00" & Hex(ci.Data4_1), 2) & "-" & Right("00" & Hex(ci.Data4_2), 2) & Right("00" & Hex(ci.Data4_3), 2) & Right("00" & Hex(ci.Data4_4), 2) & Right("00" & Hex(ci.Data4_5), 2) & Right("00" & Hex(ci.Data4_6), 2) & Right("00" & Hex(ci.Data4_7), 2)
End Function

Public Function PropIsMultiple(ByVal PropTag As Long)
   PropIsMultiple = ((kWabGetPropType(PropTag) And &H1000&) <> 0)
End Function

