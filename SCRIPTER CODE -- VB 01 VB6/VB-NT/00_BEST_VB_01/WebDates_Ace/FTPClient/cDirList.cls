VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDirList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000

Private mCol As New Collection

Public Function Add(lAttrib As Long, dtCreationTime As Date, dtLastAccessTime As Date, dtLastWriteTime As Date, lFileSize As Long, sFilename As String) As cDirItem
   Dim newItem As New cDirItem
   
   With newItem
      .Archive = (lAttrib And FILE_ATTRIBUTE_ARCHIVE)
      .Compressed = (lAttrib And FILE_ATTRIBUTE_COMPRESSED)
      .Directory = (lAttrib And FILE_ATTRIBUTE_DIRECTORY)
      .Hidden = (lAttrib And FILE_ATTRIBUTE_HIDDEN)
      .Normal = (lAttrib And FILE_ATTRIBUTE_NORMAL)
      .Offline = (lAttrib And FILE_ATTRIBUTE_OFFLINE)
      .ReadOnly = (lAttrib And FILE_ATTRIBUTE_READONLY)
      .System = (lAttrib And FILE_ATTRIBUTE_SYSTEM)
      .Temporary = (lAttrib And FILE_ATTRIBUTE_TEMPORARY)
      .CreationTime = dtCreationTime
      .LastAccessTime = dtLastAccessTime
      .LastWriteTime = dtLastWriteTime
      .FileSize = lFileSize
      .Filename = sFilename
   End With
   mCol.Add newItem, sFilename
End Function

Public Function Clear()
   Dim lIndex As Long
   If mCol.Count > 0 Then
      For lIndex = 1 To mCol.Count - 1
         mCol.Remove lIndex
      Next
   End If
End Function

Public Function Item(Index As Variant) As cDirItem
Attribute Item.VB_UserMemId = 0
   Set Item = mCol(Index)
End Function

Public Function Count() As Long
   Count = mCol.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mCol.[_NewEnum]
End Function

