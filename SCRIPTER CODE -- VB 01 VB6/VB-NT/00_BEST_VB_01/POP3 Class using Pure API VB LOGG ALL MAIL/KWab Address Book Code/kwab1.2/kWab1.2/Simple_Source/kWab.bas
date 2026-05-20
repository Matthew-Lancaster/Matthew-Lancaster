Attribute VB_Name = "modkWab"
Option Explicit

Public Type tWab
   AddrBook As Long
   WabObject As Long
   InstWab As Long
End Type

Public Type SBinary
   cb As Long
   lpb As Long
End Type

Public Declare Function kWabOpen Lib "kWab.dll" (ByRef kw As tWab) As Long
Public Declare Sub kWabClose Lib "kWab.dll" (ByRef kw As tWab)
Public Declare Function kWabGetNumEntries Lib "kWab.dll" (ByRef kw As tWab) As Long
Public Declare Function kWabGetEntries Lib "kWab.dll" (ByRef kw As tWab, ByRef FirstElementOfEntriesArray As SBinary, ByVal nBuf As Long) As Long
Public Declare Sub kWabFreeEntries Lib "kWab.dll" (ByRef kw As tWab, ByRef FirstElementOfEntriesArray As SBinary, ByVal nBuf As Long)

Public Declare Function kWabGetProp_String Lib "kWab.dll" (ByRef kw As tWab, ByRef Entry As SBinary, ByVal PropTag As Long, ByVal Str As String, ByVal LnStr As Long, ByVal UseUnicode As Boolean) As Boolean
