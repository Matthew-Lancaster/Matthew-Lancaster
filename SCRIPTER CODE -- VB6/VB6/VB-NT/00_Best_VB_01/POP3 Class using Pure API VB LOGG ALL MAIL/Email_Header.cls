VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Email_Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' This Class Object holds an Email Header and provides a easy method to
' Retreive the Elements

Private Elements        As Collection ' Holds the Element Names
Private ElementsValues  As Collection ' Holds the Element Values
Private mvarRawHeader   As String     ' Holds the RAW header data
Private mvarUID         As String     ' Used so we can store a Unique ID for the Header

Public Property Let UID(ByVal vData As String)
Attribute UID.VB_Description = "Unique ID of this Header"
      mvarUID = vData
End Property
Public Property Get UID() As String
      UID = mvarUID
End Property

Public Property Get RawHeader() As String
      RawHeader = mvarRawHeader
End Property

Public Sub PopulateHeader(ByVal sRawHeader As String)
Dim sName As String
Dim sValue As String
Dim asElements() As String
Dim sTemp As String
Dim sPrevName As String
Dim i As Long

100   On Error Resume Next

      ' Re - Initialize the collections - Clears out any Previous data
110   Set Elements = New Collection
120   Set ElementsValues = New Collection
      
130   If sRawHeader <> "" Then
      
140      asElements = Split(sRawHeader, vbCrLf)
         mvarRawHeader = sRawHeader
150      For i = 0 To UBound(asElements)

200         If InStr(1, asElements(i), ": ", vbTextCompare) = 0 Then
210            sName = sPrevName
220            sValue = asElements(i)
            Else
230            sName = Split(asElements(i), ": ")(0)
240            sValue = Split(asElements(i), ": ", 2)(1)
            End If

250         sPrevName = sName

            ' See if Element is already in the collection
260         Err.Clear
270         If Elements.Count > 0 Then
280            If Elements.item(UCase$(sName)) <> "" Then
290               If Err = 0 Then
                  
                     ' Remove the Element and add the values together
                  
300                  sTemp = ElementsValues.item(UCase$(sName))
310                  sValue = sTemp & vbCrLf & sValue
                     
320                  Elements.Remove UCase$(sName)
330                  ElementsValues.Remove UCase$(sName)
                  End If
               End If
            End If
            
340         Elements.Add sName, UCase$(sName)
350         ElementsValues.Add sValue, UCase$(sName)
         Next
      End If
      
360   Exit Sub

Err_PopulateHeader:
      Debug.Print Err.Description, Erl
      
End Sub

' Returns the Number of Elements in the Header
Public Property Get ElementCount() As Long
      On Error Resume Next
      ElementCount = Elements.Count
End Property

' Will return a Value for a Header Element using its Index Number
Public Property Get ElementValueFromIndex(ByVal IndexID As Long) As String
      On Error Resume Next
      ElementValueFromIndex = ElementsValues.item(IndexID)
End Property

' Will return a Name for a Header Element using its Index Number
Public Property Get ElementNameFromIndex(ByVal IndexID As Long) As String
      On Error Resume Next
      ElementNameFromIndex = Elements.item(IndexID)
End Property

' Will return a Value for a Header Element using its Element Name
Public Property Get ElementValue(ByVal ElementName As String) As String
      On Error Resume Next
      ElementValue = ElementsValues.item(UCase$(ElementName))
End Property

Private Sub Class_Initialize()
      Set Elements = New Collection
      Set ElementsValues = New Collection
End Sub
Private Sub Class_Terminate()
      Set Elements = Nothing
      Set ElementsValues = Nothing
End Sub
