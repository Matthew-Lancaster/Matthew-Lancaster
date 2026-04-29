VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'   THE RESEIZE CLASS
'   ~~~~~~~~~~~~~~~~~

'You are free to use this class in your own projects n give
'me some credits when you do. Dont forget to visit my web
'site k?

'                               --- Author, Muhammad Abubakar
'                                       <joehacker@yahoo.com>
'                                       http://go.to/abubakar

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Key codes:
'1  ->  top only
'2  ->  left only
'3  ->  top and left
'4  ->  height only
'5  ->  width only
'6  ->  height and width
'-----------------------
Option Explicit

Enum eParams
    RS_TopOnly = 1
    RS_LeftOnly = 2
    RS_Top_Left = 3
    RS_HeightOnly = 4
    RS_WidthOnly = 5
    RS_Height_Width = 6
End Enum
Private Type cInfo
    cControl As Control
    cHeight As Integer
    cWidth As Integer
    cTop As Integer
    cLeft As Integer
    cInfo As Integer
End Type

Private cArray() As cInfo
Private Count As Integer

Private FormHeight As Integer
Private FormWidth As Integer

Public Property Let hParam(ByVal fh As Integer)
    FormHeight = fh
        
End Property
Public Property Let wParam(ByVal fw As Integer)
    FormWidth = fw
End Property
Public Sub Map(rCont As Control, SizeInfo As eParams)
    Count = Count + 1
    ReDim Preserve cArray(Count)
    Set cArray(Count).cControl = rCont
    cArray(Count).cInfo = SizeInfo
    
    Select Case SizeInfo
        Case 1:
            cArray(Count).cTop = FormHeight - rCont.Top
        Case 2:
            cArray(Count).cLeft = FormWidth - rCont.Left
        Case 3:
            cArray(Count).cTop = FormHeight - rCont.Top
            cArray(Count).cLeft = FormWidth - rCont.Left
        Case 4:
            cArray(Count).cHeight = FormHeight - rCont.Height
        Case 5:
            cArray(Count).cWidth = FormWidth - rCont.Width
        Case 6:
            cArray(Count).cHeight = FormHeight - rCont.Height
            cArray(Count).cWidth = FormWidth - rCont.Width
        Case Else:
            Exit Sub
    End Select
    
End Sub

Public Sub rSize(cForm As Form)
    
    On Error Resume Next
    Dim i As Integer, a As Integer, b As Integer
    For i = 1 To Count
        Select Case cArray(i).cInfo
            Case 1:
                cArray(i).cControl.Top = cForm.Height - cArray(i).cTop
            Case 2:
                cArray(i).cControl.Left = cForm.Width - cArray(i).cLeft
            Case 3:
                cArray(i).cControl.Top = cForm.Height - cArray(i).cTop
                cArray(i).cControl.Left = cForm.Width - cArray(i).cLeft
            Case 4:
                b = cForm.Height - cArray(i).cHeight
                If b < 0 Then b = 0
                cArray(i).cControl.Height = b 'cForm.Height - cArray(i).cHeight
            Case 5:
                a = cForm.Width - cArray(i).cWidth
                If a < 0 Then a = 0
                cArray(i).cControl.Width = a 'cForm.Width - cArray(i).cWidth
            Case 6:
                a = cForm.Width - cArray(i).cWidth
                b = cForm.Height - cArray(i).cHeight
                If a < 0 Then a = 0
                If b < 0 Then b = 0
                cArray(i).cControl.Height = b 'cForm.Height - cArray(i).cHeight
                cArray(i).cControl.Width = a 'cForm.Width - cArray(i).cWidth
            
                
        End Select
    Next
    Exit Sub
End Sub

