Attribute VB_Name = "KWABCOMM"
Public kw As tWab
Public WabIsOpen As Boolean

Public lContacts() As SBinary
Public lContacts_ok As Boolean

Public Const UseUnicode As Boolean = False  'Should be set to False

Public idxEntry As Long, nProps As Long, k As Long, PropTag As Long, propTagName As String
