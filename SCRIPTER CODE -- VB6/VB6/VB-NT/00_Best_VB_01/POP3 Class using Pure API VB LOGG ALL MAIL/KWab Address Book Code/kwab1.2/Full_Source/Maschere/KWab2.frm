VERSION 5.00
Begin VB.Form KWab2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "kWab"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCloseWab 
      Caption         =   "Close WAB"
      Height          =   375
      Left            =   1815
      TabIndex        =   1
      Top             =   120
      Width           =   1785
   End
   Begin VB.CommandButton cmdHowManyContacts 
      Caption         =   "How many contacts?"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1785
   End
   Begin VB.CommandButton cmdListContacts 
      Caption         =   "(Refresh) Contacts list"
      Height          =   375
      Left            =   5385
      TabIndex        =   3
      Top             =   120
      Width           =   1785
   End
   Begin VB.CommandButton cmdAddContactGUI 
      Caption         =   "Add contact (OS GUI)"
      Height          =   375
      Left            =   7170
      TabIndex        =   4
      Top             =   120
      Width           =   1785
   End
   Begin VB.CommandButton cmdOpenWab 
      Caption         =   "Open WAB"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   1785
   End
   Begin VB.CommandButton cmdAddContact 
      Caption         =   "Add contact"
      Height          =   375
      Left            =   8955
      TabIndex        =   5
      Top             =   120
      Width           =   1785
   End
   Begin VB.CommandButton cmdAddProp 
      Caption         =   "Add property"
      Height          =   285
      Left            =   2355
      TabIndex        =   26
      Top             =   6240
      Width           =   2145
   End
   Begin VB.Frame fraMultipleProperty 
      Caption         =   "Multiple-value property"
      Height          =   1095
      Left            =   4560
      TabIndex        =   19
      Top             =   5430
      Width           =   3375
      Begin VB.CommandButton cmdMultiEditProp 
         Caption         =   "Edit"
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   660
         Width           =   855
      End
      Begin VB.CommandButton cmdMultiDelProp 
         Caption         =   "Remove"
         Height          =   285
         Left            =   990
         TabIndex        =   24
         Top             =   660
         Width           =   855
      End
      Begin VB.ComboBox cmbWichMultiple 
         Height          =   315
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   630
         Width           =   705
      End
      Begin VB.CommandButton cmdMultiAddProp 
         Caption         =   "Add a new value"
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   3135
      End
      Begin VB.Label lblWichMulti 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Value #"
         Height          =   195
         Left            =   1920
         TabIndex        =   21
         Top             =   690
         Width           =   555
      End
   End
   Begin VB.Frame fraWholeProperty 
      Caption         =   "Whole property"
      Height          =   675
      Left            =   4560
      TabIndex        =   14
      Top             =   4695
      Width           =   3375
      Begin VB.CommandButton cmdEditProp 
         Caption         =   "Edit value"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   1515
      End
      Begin VB.CommandButton cmdDelProp 
         Caption         =   "Remove property"
         Height          =   285
         Left            =   1740
         TabIndex        =   16
         Top             =   270
         Width           =   1515
      End
   End
   Begin VB.TextBox txtBinData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Index           =   0
      Left            =   7980
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "KWab2.frx":0000
      Top             =   1200
      Width           =   2625
   End
   Begin VB.ComboBox cmbProbBinData 
      Height          =   315
      Left            =   7980
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   810
      Width           =   2625
   End
   Begin VB.ListBox lstPropDett 
      Height          =   3810
      IntegralHeight  =   0   'False
      Left            =   4530
      TabIndex        =   13
      Top             =   825
      Width           =   3405
   End
   Begin VB.CommandButton cmdDelContact 
      Caption         =   "Delete contact"
      Height          =   285
      Left            =   180
      TabIndex        =   25
      Top             =   6240
      Width           =   2145
   End
   Begin VB.CommandButton cmdEditContactGUI 
      Caption         =   "Edit contact (OS GUI)"
      Height          =   285
      Left            =   180
      TabIndex        =   17
      Top             =   5910
      Width           =   2145
   End
   Begin VB.ListBox lstContacts 
      Height          =   5040
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   11
      Top             =   825
      Width           =   2145
   End
   Begin VB.ListBox lstProps 
      Height          =   5370
      IntegralHeight  =   0   'False
      Left            =   2355
      TabIndex        =   12
      Top             =   825
      Width           =   2145
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Property binary data"
      Height          =   195
      Left            =   7950
      TabIndex        =   6
      Top             =   570
      Width           =   2625
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Property details"
      Height          =   195
      Left            =   4530
      TabIndex        =   9
      Top             =   600
      Width           =   3405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Contact properties"
      Height          =   195
      Left            =   2355
      TabIndex        =   8
      Top             =   600
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Contacts list"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   600
      Width           =   2145
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Result"
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   6600
      Width           =   450
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "lblStato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   660
      TabIndex        =   28
      Top             =   6600
      Width           =   660
   End
End
Attribute VB_Name = "KWab2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private WabIsOpen As Boolean


Private Const UseUnicode As Boolean = False  'Should be set to False
Private Sub cmdAddContact_Click()
   Dim NewDisplayName As String
   If Not WabIsOpen Then Exit Sub
   
   NewDisplayName = Trim(InputBox("Enter the new contact's display name"))
   If Len(NewDisplayName) = 0 Then Exit Sub
   Me.lblResult = kWabAddNew(kw, NewDisplayName, UseUnicode)
   cmdListContacts_Click
End Sub

Public Sub cmdOpenWab_Click()
   Dim R As Long
   If WabIsOpen Then Exit Sub
   WabIsOpen = kWabOpen(kw)
   WabIsOpen = True
   Me.lblResult = WabIsOpen
   ChangeStates
End Sub
Public Sub cmdCloseWab_Click()
   If WabIsOpen = False Then
      If lContacts_ok Then
         kWabFreeEntries kw, lContacts(1), UBound(lContacts)
         ReDim lContacts(0 To 0)
         lContacts_ok = False
      End If
      kWabClose kw
      WabIsOpen = False
      ChangeStates
   End If
End Sub
Private Sub cmdAddContactGUI_Click()
   If Not WabIsOpen Then Exit Sub
   Me.lblResult = kWabAddNewGUI(kw, Me.hWnd)
End Sub
Private Sub cmdHowManyContacts_Click()
   If Not WabIsOpen Then Exit Sub
   Me.lblResult = kWabGetNumEntries(kw)
End Sub

Public Sub cmdListContacts_Click()
    'Exit Sub
   Dim N As Long, R As Long
   Dim K As Long
   
   If Not WabIsOpen Then
   'Call cmdOpenWab_Click
   Exit Sub
   End If
   If lContacts_ok Then
      kWabFreeEntries kw, lContacts(1), UBound(lContacts)
      ReDim lContacts(0 To 0)
      lContacts_ok = False
   End If
   N = kWabGetNumEntries(kw)
   If N < 1 Then Exit Sub
   ReDim lContacts(1 To N)
   lContacts_ok = True
   R = kWabGetEntries(kw, lContacts(1), N)
   With Me.lstContacts
      .Clear
      For K = 1 To R
         .AddItem GetVisName(lContacts(K))
         .ItemData(.NewIndex) = K
        'If K Mod 10 = 0 Then DoEvents
      Next
   End With
   ChangeStates
End Sub

Private Sub Form_Load()
   WabIsOpen = False
   Me.lblResult = ""
   lContacts_ok = False
   ChangeStates
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      If WabIsOpen Then
         cmdCloseWab_Click
      Else
         End
      End If
   End If
End Sub
Sub ChangeStates()
   Dim EntryIsSelected As Boolean, PropIsSelected As Boolean, MultipleValues As Boolean
   Dim SelectedPropTag As Long

   EntryIsSelected = False
   PropIsSelected = False
   MultipleValues = False

   If WabIsOpen Then
      With Me.lstContacts
         If .ListCount > 0 Then
            If .ListIndex >= 0 Then
               If .ItemData(.ListIndex) > 0 Then
                  EntryIsSelected = True
               End If
            End If
         End If
      End With
      If EntryIsSelected Then
         With Me.lstProps
            If .ListCount > 0 Then
               If .ListIndex >= 0 Then
                  If .ItemData(.ListIndex) <> 0 Then
                     PropIsSelected = True
                     SelectedPropTag = .ItemData(.ListIndex)
                     MultipleValues = PropIsMultiple(SelectedPropTag)
                  End If
               End If
            End If
         End With
      End If
   End If

   If Not PropIsSelected Then
      SetNumBinData 0
      SetNumMultiple 0
   End If

   Me.cmdOpenWab.Enabled = Not WabIsOpen
   Me.cmdCloseWab.Enabled = WabIsOpen

   Me.cmdHowManyContacts.Enabled = WabIsOpen
   Me.cmdListContacts.Enabled = WabIsOpen
   Me.cmdEditContactGUI.Enabled = EntryIsSelected
   Me.cmdDelContact.Enabled = EntryIsSelected
   Me.lstContacts.Enabled = WabIsOpen
   Me.lstContacts.BackColor = IIf(WabIsOpen, &H80000005, Me.BackColor)
   Me.lstProps.Enabled = EntryIsSelected
   Me.lstProps.BackColor = IIf(EntryIsSelected, &H80000005, Me.BackColor)
   Me.lstPropDett.Enabled = PropIsSelected
   Me.lstPropDett.BackColor = IIf(PropIsSelected, &H80000005, Me.BackColor)
   Me.cmdAddContactGUI.Enabled = WabIsOpen
   Me.cmdAddContact.Enabled = WabIsOpen

   Me.fraWholeProperty.Enabled = EntryIsSelected
   Me.cmdEditProp.Enabled = PropIsSelected And (Not MultipleValues)
   Me.cmdDelProp.Enabled = PropIsSelected
   Me.cmdAddProp.Enabled = EntryIsSelected

   Me.fraMultipleProperty.Enabled = MultipleValues
   Me.cmdMultiEditProp.Enabled = MultipleValues And (Me.cmbWichMultiple.ListCount > 0)
   Me.cmdMultiAddProp.Enabled = MultipleValues
   Me.cmdMultiDelProp.Enabled = MultipleValues And (Me.cmbWichMultiple.ListCount > 0)

   If Not PropIsSelected Then
      Me.lstPropDett.Clear
      If Not EntryIsSelected Then
         Me.lstProps.Clear
         If Not WabIsOpen Then
            Me.lstContacts.Clear
         End If
      End If
   End If
End Sub
Private Sub cmdEditContactGUI_Click()
   Dim ID As Long
   
   If Not WabIsOpen Then Exit Sub
   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      ID = .ItemData(.ListIndex)
      
      If ID < 1 Then Exit Sub
      Me.lblResult = kWabDetails(kw, lContacts(ID), Me.hWnd)
   End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   cmdCloseWab_Click
End Sub

Public Sub lstContactsOld_Click()
   Dim idxEntry As Long, nProps As Long, K As Long, PropTag As Long, propTagName As String

   If Not WabIsOpen Then Exit Sub
   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      idxEntry = .ItemData(.ListIndex)
   End With
   
   If idxEntry < 1 Then Exit Sub
   nProps = kWabGetNumProps(kw, lContacts(idxEntry))
   With Me.lstProps
      .Clear
      For K = 0 To nProps - 1
         PropTag = kWabGetPropTag(kw, lContacts(idxEntry), K, UseUnicode)
         .AddItem kWabPropTagName(PropTag)
         .ItemData(.NewIndex) = PropTag
      Next
   End With
   ChangeStates
End Sub

Public Sub lstContacts_Click()
   Dim idxEntry As Long, nProps As Long, K As Long, PropTag As Long, propTagName As String, B1$

   'If Not WabIsOpen Then Exit Sub
   'With Me.lstContacts
   '   If .ListCount < 1 Then Exit Sub
   '   If .ListIndex < 0 Then Exit Sub
   '   idxEntry = .ItemData(.ListIndex)
   'End With
   
   
   'HALO3 = HALO1Email
   
   If idxEntry < 1 Then MsgBox "Entry Not Found in List": Exit Sub
   nProps = kWabGetNumProps(kw, lContacts(idxEntry))
   With Me.lstProps
      .Clear
      For K = 0 To nProps - 1
         PropTag = kWabGetPropTag(kw, lContacts(idxEntry), K, UseUnicode)
         
         .AddItem kWabPropTagName(PropTag)
         .ItemData(.NewIndex) = PropTag
      Next
   End With
   ChangeStates
End Sub

Private Sub lstProps_Click()
   Dim idxEntry As Long, PropTag As Long, propID As Long, PropType As Long
   
   'For string types
   Dim sBuf As String, lBuf As Long
   
   'For LONGLONG types
   Dim lnglng As LONGLONG
   
   'For LONG types
   Dim lng As Long

   'For SHORT types
   Dim srt As Integer

   'For FILETIME types
   Dim st As SysTime, yr As Integer, mh As Integer, dow As Integer, dy As Integer, hr As Integer, mt As Integer, ss As Integer, ms As Integer, dat As Date

   'For BINARY types
   Dim sBin() As Byte, nBinPRE As Long, nBinINOUT As Long

   'For FLOAT types
   Dim sng As Single
   
   'For DOUBLE types
   Dim dbl As Double
   
   'For CURRENCY types
   Dim cy As Currency
   
   'For APPTIME types
   Dim at As Date
   
   'For CLSID types
   Dim cid As CLSID
   
   'For BOOLEAN types
   Dim boo As Boolean
   
   'For the MV (MultipleValues) types
   Dim N As Long, Count As Long
   
   SetNumBinData 0
   SetNumMultiple 0
   
   Me.lstPropDett.Clear
   If Not WabIsOpen Then Exit Sub
   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      idxEntry = .ItemData(.ListIndex)
      If idxEntry = 0 Then Exit Sub
   End With

   With Me.lstProps
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      PropTag = .ItemData(.ListIndex)
      If PropTag = 0 Then Exit Sub
   End With

   propID = kWabGetPropID(PropTag)
   PropType = kWabGetPropType(PropTag)
   With Me.lstPropDett
      .Clear
      .AddItem "tag: " & Right("00000000" & Hex(PropTag), 8)
      .AddItem "type: " & kWabPropTypeName(PropType)
      .AddItem "group: " & kWabPropIDGroup(propID)
      .AddItem "subgroup: " & kWabPropIDSubGroup(propID)
      .AddItem "tag name: " & kWabPropTagName(PropTag)
      .AddItem "trasmitteable: " & kWabPropIsTransmittable(propID)
      Select Case PropType
            
         Case PT_STRING8, PT_UNICODE
            lBuf = 500
            sBuf = String(lBuf, vbNullChar)
            If kWabGetProp_String(kw, lContacts(idxEntry), PropTag, sBuf, lBuf, UseUnicode) Then
               sBuf = StripNullChars(sBuf)
               .AddItem "value: """ & sBuf & """"
            Else
               .AddItem "value: <error>"
            End If
         
         Case PT_LONGLONG
            If kWabGetProp_LongLong(kw, lContacts(idxEntry), PropTag, lnglng) Then
               .AddItem "value: hex " & Right("00000000" & Hex(lnglng.Word1), 8) & Right("00000000" & Hex(lnglng.Word2), 8) & ")"
            Else
               .AddItem "value: <error>"
            End If
         
         Case PT_LONG
            If kWabGetProp_Long(kw, lContacts(idxEntry), PropTag, lng) Then
               .AddItem "value: " & lng & " (hex: " & Right("00000000" & Hex(lng), 8) & ")"
            Else
               .AddItem "value: <error>"
            End If

         Case PT_SHORT
            If kWabGetProp_Short(kw, lContacts(idxEntry), PropTag, srt) Then
               .AddItem "value: " & srt & " (hex: " & Right("0000" & Hex(srt), 4) & ")"
            Else
               .AddItem "value: <error>"
            End If

         Case PT_SYSTIME
            If kWabGetProp_SysTime(kw, lContacts(idxEntry), PropTag, st) Then
               If ExpandSysTime(st, yr, mh, dow, dy, hr, mt, ss, ms) Then
                  If SysTime2Date(st, dat) Then
                     .AddItem "value: " & GetDayOfWeek(dow) & " " & Format(dy, "00") & "/" & Format(mh, "00") & "/" & Format(yr, "0000") & " " & Format(hr, "00") & ":" & Format(mt, "00") & ":" & Format(ss, "00") & ":" & Format(ms, "000") & " (" & dat & ")"
                  Else
                     .AddItem "value: <error expanding time 2nd time>"
                  End If
               Else
                  .AddItem "value: <error expanding time>"
               End If
            Else
               .AddItem "value: <error>"
            End If

         Case PT_BINARY
            If kWabGetProp_Binary_Count(kw, lContacts(idxEntry), PropTag, nBinPRE) Then
               If nBinPRE = 0 Then
                  .AddItem "value: 0 bytes"
               Else
                  nBinINOUT = nBinPRE + 50   '+50 just for test: kWabGetProp_Binary_Data should re-set it to nBinPRE
                  ReDim sBin(1 To nBinINOUT)
                  If kWabGetProp_Binary_Data(kw, lContacts(idxEntry), PropTag, sBin(1), nBinINOUT) Then
                     SetNumBinData 1
                     If nBinINOUT > 0 Then
                        ReDim Preserve sBin(1 To nBinINOUT)
                        SetBinData 0, sBin
                     End If
                     .AddItem "value: " & nBinPRE & " bytes - read ok (" & nBinINOUT & ")"
                  Else
                     .AddItem "value: " & nBinPRE & " bytes - NOT READ"
                  End If
                  ReDim sBin(0 To 0)
               End If
            Else
               .AddItem "value: <error>"
            End If

         Case PT_FLOAT
            If kWabGetProp_Float(kw, lContacts(idxEntry), PropTag, sng) Then
               .AddItem "value: " & sng
            Else
               .AddItem "value: <error>"
            End If

         Case PT_DOUBLE
            If kWabGetProp_Double(kw, lContacts(idxEntry), PropTag, dbl) Then
               .AddItem "value: " & dbl
            Else
               .AddItem "value: <error>"
            End If

         Case PT_CURRENCY
            If kWabGetProp_Currency(kw, lContacts(idxEntry), PropTag, cy) Then
               .AddItem "value: " & cy
            Else
               .AddItem "value: <error>"
            End If

         Case PT_APPTIME
            If kWabGetProp_AppTime(kw, lContacts(idxEntry), PropTag, at) Then
               .AddItem "value: " & at
            Else
               .AddItem "value: <error>"
            End If

         Case PT_CLSID
            If kWabGetProp_CLSID(kw, lContacts(idxEntry), PropTag, cid) Then
               .AddItem "value: " & CLSID2String(cid)
            Else
               .AddItem "value: <error>"
            End If

         Case PT_BOOLEAN
            If kWabGetProp_Boolean(kw, lContacts(idxEntry), PropTag, boo) Then
               .AddItem "value: " & boo
            Else
               .AddItem "value: <error>"
            End If

      'Multiple values
         Case PT_MV_STRING8, PT_MV_UNICODE
            If kWabGetProp_MV_String_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  lBuf = 500
                  sBuf = String(lBuf, vbNullChar)
                  If kWabGetProp_MV_String_Data(kw, lContacts(idxEntry), PropTag, sBuf, lBuf, UseUnicode, Count - 1) Then
                     sBuf = StripNullChars(sBuf)
                     .AddItem "  " & Count & ") """ & sBuf & """"
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_LONGLONG
            If kWabGetProp_MV_LongLong_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_LongLong_Data(kw, lContacts(idxEntry), PropTag, lnglng, Count - 1) Then
                     .AddItem "  " & Count & ") " & Right("00000000" & Hex(lnglng.Word1), 8) & Right("00000000" & Hex(lnglng.Word2), 8) & ")"
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_LONG
            If kWabGetProp_MV_Long_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_Long_Data(kw, lContacts(idxEntry), PropTag, lng, Count - 1) Then
                     .AddItem "  " & Count & ") " & Right("00000000" & Hex(lng), 8)
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_SHORT
            If kWabGetProp_MV_Short_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_Short_Data(kw, lContacts(idxEntry), PropTag, srt, Count - 1) Then
                     .AddItem "  " & Count & ") " & srt & " (hex: " & Right("0000" & Hex(srt), 4) & ")"
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_SYSTIME
            If kWabGetProp_MV_SysTime_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_SysTime_Data(kw, lContacts(idxEntry), PropTag, st, Count - 1) Then
                     If ExpandSysTime(st, yr, mh, dow, dy, hr, mt, ss, ms) Then
                        If SysTime2Date(st, dat) Then
                           .AddItem "  " & Count & ") " & GetDayOfWeek(dow) & " " & Format(dy, "00") & "/" & Format(mh, "00") & "/" & Format(yr, "0000") & " " & Format(hr, "00") & ":" & Format(mt, "00") & ":" & Format(ss, "00") & ":" & Format(ms, "000") & " (" & dat & ")"
                        Else
                           .AddItem "  " & Count & ") <error expanding time 2nd time>"
                        End If
                     Else
                        .AddItem "  " & Count & ") <error expanding time>"
                     End If
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If


         Case PT_MV_BINARY
            If kWabGetProp_MV_Binary_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               SetNumBinData N
               For Count = 1 To N
                  If kWabGetProp_MV_Binary_Data_Count(kw, lContacts(idxEntry), PropTag, nBinPRE, Count - 1) Then
                     If nBinPRE = 0 Then
                        .AddItem "  " & Count & ") 0 bytes"
                     Else
                        nBinINOUT = nBinPRE + 50   '+50 just for test
                        ReDim sBin(1 To nBinINOUT)
                        If kWabGetProp_MV_Binary_Data_Data(kw, lContacts(idxEntry), PropTag, sBin(1), nBinINOUT, Count - 1) Then
                           If nBinINOUT > 0 Then
                              ReDim Preserve sBin(1 To nBinINOUT)
                              SetBinData Count - 1, sBin
                           End If
                           .AddItem "  " & Count & ") " & nBinPRE & " bytes - read ok (" & nBinINOUT & ")"
                        Else
                           .AddItem "  " & Count & ") " & nBinPRE & " bytes - NOT READ"
                        End If
                        ReDim sBin(0 To 0)
                     End If
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_FLOAT
            If kWabGetProp_MV_Float_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_Float_Data(kw, lContacts(idxEntry), PropTag, sng, Count - 1) Then
                     .AddItem "  " & Count & ") " & sng
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_DOUBLE
            If kWabGetProp_MV_Double_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_Double_Data(kw, lContacts(idxEntry), PropTag, dbl, Count - 1) Then
                     .AddItem "  " & Count & ") " & dbl
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_CURRENCY
            If kWabGetProp_MV_Currency_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_Currency_Data(kw, lContacts(idxEntry), PropTag, cy, Count - 1) Then
                     .AddItem "  " & Count & ") " & cy
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_APPTIME
            If kWabGetProp_MV_AppTime_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_AppTime_Data(kw, lContacts(idxEntry), PropTag, at, Count - 1) Then
                     .AddItem "  " & Count & ") " & at
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case PT_MV_CLSID
            If kWabGetProp_MV_CLSID_Count(kw, lContacts(idxEntry), PropTag, N) Then
               SetNumMultiple N
               .AddItem "" & N & " values:"
               For Count = 1 To N
                  If kWabGetProp_MV_CLSID_Data(kw, lContacts(idxEntry), PropTag, cid, Count - 1) Then
                     .AddItem "  " & Count & ") reak ok"
                  Else
                     .AddItem "  " & Count & ") <error>"
                  End If
               Next
            Else
               .AddItem "unknown # of values"
            End If

         Case Else
            .AddItem "value: <unknown/unsupported type>"
      
      End Select
   End With
   
   ChangeStates
End Sub
Private Function StripNullChars(ByVal s As String) As String
   Do While Len(s) > 0
      If Asc(Right(s, 1)) <> 0 Then Exit Do
      s = Left(s, Len(s) - 1)
   Loop
   StripNullChars = s
End Function
Private Sub cmdDelContact_Click()
   Dim idx As Long, li As Long, K As Long

   If Not WabIsOpen Then Exit Sub
   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      li = .ListIndex
      If li < 0 Then Exit Sub
      idx = .ItemData(.ListIndex)
   End With
   If idx < 1 Then Exit Sub
   
   If MsgBox("Are you sure to delete the selected item?", vbYesNo Or vbQuestion Or vbDefaultButton2) <> vbYes Then Exit Sub
   If kWabDeleteEntry(kw, lContacts(idx)) Then
      Me.lstContacts.RemoveItem li
      kWabFreeEntries kw, lContacts(idx), 1
      If UBound(lContacts) = 1 Then
         ReDim lContacts(0 To 0)
         lContacts_ok = False
      Else
         For K = idx To UBound(lContacts) - 1
            lContacts(K).cb = lContacts(K + 1).cb
            lContacts(K).lpb = lContacts(K + 1).lpb
         Next
         ReDim Preserve lContacts(1 To UBound(lContacts) - 1)
      End If
      Me.lblResult = "Done"
   Else
      Me.lblResult = "Error"
   End If

End Sub
Private Function GetVisName(e As SBinary) As String
   Dim N As String, s As String

   GetVisName = ""
   N = 255
   s = String(N, vbNullChar)
   If kWabGetProp_String(kw, e, PR_DISPLAY_NAME_A, s, N, False) Then
      s = StripNullChars(s)
      GetVisName = s
   End If
End Function
Private Sub cmdEditProp_Click()
   Dim idxEntry As Long, PropTag As Long
   Dim vLONGLONG As LONGLONG
   Dim vLONG As Long
   Dim vSHORT As Integer
   Dim vSTRING As String, vSTRINGlen As Long
   Dim vSYSTIME As SysTime
   Dim vAPPTIME As Date
   Dim vFLOAT As Single
   Dim vDOUBLE As Double
   Dim vBOOLEAN As Boolean
   Dim vCURRENCY As Currency
   Dim vCLSID As CLSID
   Dim vBINARY() As Byte, vBINARYlen As Long

   If Not WabIsOpen Then Exit Sub
   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      idxEntry = .ItemData(.ListIndex)
      If idxEntry = 0 Then Exit Sub
   End With

   With Me.lstProps
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      PropTag = .ItemData(.ListIndex)
      If PropTag = 0 Then Exit Sub
   End With
   
   Select Case kWabGetPropType(PropTag)
      
      Case PT_LONGLONG
         kWabGetProp_LongLong kw, lContacts(idxEntry), PropTag, vLONGLONG
         If Not AskLongLong(vLONGLONG, vLONGLONG) Then Exit Sub
         Me.lblResult = kWabSetProp_LongLong(kw, lContacts(idxEntry), PropTag, vLONGLONG)

      Case PT_LONG
         kWabGetProp_Long kw, lContacts(idxEntry), PropTag, vLONG
         If Not AskLong(vLONG, vLONG) Then Exit Sub
         Me.lblResult = kWabSetProp_Long(kw, lContacts(idxEntry), PropTag, vLONG)

      Case PT_SHORT
         kWabGetProp_Short kw, lContacts(idxEntry), PropTag, vSHORT
         If Not AskShort(vSHORT, vSHORT) Then Exit Sub
         Me.lblResult = kWabSetProp_Short(kw, lContacts(idxEntry), PropTag, vSHORT)

      Case PT_STRING8, PT_UNICODE
         vSTRINGlen = 500
         vSTRING = String(vSTRINGlen, vbNullChar)
         kWabGetProp_String kw, lContacts(idxEntry), PropTag, vSTRING, vSTRINGlen, UseUnicode
         vSTRING = StripNullChars(vSTRING)
         If Not AskString(vSTRING, vSTRING) Then Exit Sub
         Me.lblResult = kWabSetProp_String(kw, lContacts(idxEntry), PropTag, vSTRING)

      Case PT_SYSTIME
         kWabGetProp_SysTime kw, lContacts(idxEntry), PropTag, vSYSTIME
         If Not AskSysTime(vSYSTIME, vSYSTIME) Then Exit Sub
         Me.lblResult = kWabSetProp_SysTime(kw, lContacts(idxEntry), PropTag, vSYSTIME)

      Case PT_APPTIME
         kWabGetProp_AppTime kw, lContacts(idxEntry), PropTag, vAPPTIME
         If Not AskAppTime(vAPPTIME, vAPPTIME) Then Exit Sub
         Me.lblResult = kWabSetProp_AppTime(kw, lContacts(idxEntry), PropTag, vAPPTIME)

      Case PT_FLOAT
         kWabGetProp_Float kw, lContacts(idxEntry), PropTag, vFLOAT
         If Not AskFloat(vFLOAT, vFLOAT) Then Exit Sub
         Me.lblResult = kWabSetProp_Float(kw, lContacts(idxEntry), PropTag, vFLOAT)

      Case PT_DOUBLE
         kWabGetProp_Double kw, lContacts(idxEntry), PropTag, vDOUBLE
         If Not AskDouble(vDOUBLE, vDOUBLE) Then Exit Sub
         Me.lblResult = kWabSetProp_Double(kw, lContacts(idxEntry), PropTag, vDOUBLE)

      Case PT_BOOLEAN
         kWabGetProp_Boolean kw, lContacts(idxEntry), PropTag, vBOOLEAN
         If Not AskBoolean(vBOOLEAN, vBOOLEAN) Then Exit Sub
         Me.lblResult = kWabSetProp_Boolean(kw, lContacts(idxEntry), PropTag, vBOOLEAN)

      Case PT_CURRENCY
         kWabGetProp_Currency kw, lContacts(idxEntry), PropTag, vCURRENCY
         If Not AskCurrency(vCURRENCY, vCURRENCY) Then Exit Sub
         Me.lblResult = kWabSetProp_Currency(kw, lContacts(idxEntry), PropTag, vCURRENCY)

      Case PT_CLSID
         kWabGetProp_CLSID kw, lContacts(idxEntry), PropTag, vCLSID
         If Not AskCLSID(vCLSID, vCLSID) Then Exit Sub
         Me.lblResult = kWabSetProp_CLSID(kw, lContacts(idxEntry), PropTag, vCLSID)

      Case PT_BINARY
         If kWabGetProp_Binary_Count(kw, lContacts(idxEntry), PropTag, vBINARYlen) Then
            If vBINARYlen > 0 Then
               ReDim Preserve vBINARY(1 To vBINARYlen)
               kWabGetProp_Binary_Data kw, lContacts(idxEntry), PropTag, vBINARY(1), vBINARYlen
            End If
         End If
         If Not AskBinary(vBINARY, vBINARY) Then Exit Sub
         Me.lblResult = kWabSetProp_Binary(kw, lContacts(idxEntry), PropTag, vBINARY(LBound(vBINARY)), UBound(vBINARY) - LBound(vBINARY) + 1)

      Case Else
         MsgBox "Unknow property type", vbExclamation Or vbOKOnly
   End Select
   lstProps_Click
End Sub
Private Sub cmdDelProp_Click()
   Dim idxEntry As Long, PropTag As Long

   If Not WabIsOpen Then Exit Sub
   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      idxEntry = .ItemData(.ListIndex)
      If idxEntry = 0 Then Exit Sub
   End With

   With Me.lstProps
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      PropTag = .ItemData(.ListIndex)
      If PropTag = 0 Then Exit Sub
   End With
   
   If MsgBox("Are you sure to remove the property?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then Exit Sub
   Me.lblResult = kWabDeleteProperty(kw, lContacts(idxEntry), PropTag)
   lstContacts_Click
End Sub

Private Sub SetNumBinData(ByVal N As Integer)
   Dim i As Integer

   If Me.txtBinData.UBound > 0 Then
      For i = Me.txtBinData.UBound To IIf(N > 1, N - 1, 1)
         Unload Me.txtBinData(i)
      Next
   End If
   For i = 0 To N - 1
      If i > 0 Then
         If i > Me.txtBinData.UBound Then Load Me.txtBinData(i)
         Me.txtBinData(i).Visible = True
      End If
      Me.txtBinData(i) = ""
   Next
   With Me.cmbProbBinData
      .Clear
      For i = 1 To N
         .AddItem "Data " & i
      Next
   End With
   Select Case N
      Case 0
         Me.cmbProbBinData.Enabled = False
         Me.cmbProbBinData.BackColor = Me.BackColor
         Me.txtBinData(0) = ""
         Me.txtBinData(0).Enabled = False
         Me.txtBinData(0).BackColor = Me.BackColor
      Case 1
         Me.cmbProbBinData.Enabled = False
         Me.cmbProbBinData.BackColor = Me.BackColor
         Me.txtBinData(0).Enabled = True
         Me.txtBinData(0).BackColor = &H80000005
         Me.cmbProbBinData.ListIndex = 0
      Case Else
         Me.cmbProbBinData.Enabled = True
         Me.cmbProbBinData.BackColor = &H80000005
         For i = 0 To N - 1
            Me.txtBinData(i).Enabled = True
            Me.txtBinData(i).BackColor = &H80000005
            Me.txtBinData(i).Visible = (i = 0)
         Next
         Me.cmbProbBinData.ListIndex = 0
   End Select
End Sub

Private Sub SetBinData(N As Integer, ba() As Byte)
   Dim ris As String, p As Long, p2 As Long, lb As Long, ub As Long, LineB As String, LineA As String
   Const BytesPerLine As Long = 4

   ris = ""
   lb = LBound(ba)
   ub = UBound(ba)
   For p = lb To ub
      LineB = ""
      LineA = ""
      For p2 = p To p + BytesPerLine - 1
         If p2 <= ub Then
            If p2 > p Then
               LineB = LineB & " "
               'LineA = LineA & " "
            End If
            LineB = LineB & Right("00" & Hex(ba(p2)), 2)
            If ba(p2) >= 32 Then
               LineA = LineA & Chr(ba(p2))
            Else
               LineA = LineA & "?"
            End If
         Else
            LineB = LineB & "  "
            LineA = LineA & " "
         End If
      Next
      LineB = Left(LineB & String(3 * BytesPerLine - 1, " "), 3 * BytesPerLine - 1)
      LineA = Left(LineA & String(BytesPerLine, " "), BytesPerLine)
      If p = lb Then
         ris = LineB & "  |  " & LineA
      Else
         ris = ris & vbNewLine & LineB & "  |  " & LineA
      End If
      p = p + BytesPerLine - 1
   Next
   Me.txtBinData(N) = ris
End Sub
Private Sub cmbProbBinData_Click()
   Dim i As Integer
   For i = Me.txtBinData.LBound To Me.txtBinData.UBound
      Me.txtBinData(i).Visible = (i = Me.cmbProbBinData.ListIndex)
   Next
End Sub
Public Sub cmdAddProp_Click()
    Exit Sub
   Dim idxEntry As Long, PropTag As Long
   Dim vLONGLONG As LONGLONG
   Dim vLONG As Long
   Dim vSHORT As Integer
   Dim vSTRING As String, vSTRINGlen As Long
   Dim vSYSTIME As SysTime
   Dim vAPPTIME As Date
   Dim vFLOAT As Single
   Dim vDOUBLE As Double
   Dim vBOOLEAN As Boolean
   Dim vCURRENCY As Currency
   Dim vCLSID As CLSID
   Dim vBINARY() As Byte, vBINARYlen As Long

   If Not WabIsOpen Then Exit Sub
   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      idxEntry = .ItemData(.ListIndex)
      If idxEntry = 0 Then Exit Sub
   End With

    
   'If Not modAsk.AskPropTag(PropTag) Then Exit Sub

   Select Case kWabGetPropType(PropTag)
      Case PT_LONGLONG
         kWabGetProp_LongLong kw, lContacts(idxEntry), PropTag, vLONGLONG
         If Not AskLongLong(vLONGLONG, vLONGLONG) Then Exit Sub
         Me.lblResult = kWabSetProp_LongLong(kw, lContacts(idxEntry), PropTag, vLONGLONG)

      Case PT_LONG
         kWabGetProp_Long kw, lContacts(idxEntry), PropTag, vLONG
         If Not AskLong(vLONG, vLONG) Then Exit Sub
         Me.lblResult = kWabSetProp_Long(kw, lContacts(idxEntry), PropTag, vLONG)

      Case PT_SHORT
         kWabGetProp_Short kw, lContacts(idxEntry), PropTag, vSHORT
         If Not AskShort(vSHORT, vSHORT) Then Exit Sub
         Me.lblResult = kWabSetProp_Short(kw, lContacts(idxEntry), PropTag, vSHORT)

      '-----------HERE
      Case PT_STRING8, PT_UNICODE
         vSTRINGlen = 500
         vSTRING = String(vSTRINGlen, vbNullChar)
         kWabGetProp_String kw, lContacts(idxEntry), PropTag, vSTRING, vSTRINGlen, UseUnicode
         vSTRING = StripNullChars(vSTRING)
         If Not AskString(vSTRING, vSTRING) Then Exit Sub
         Me.lblResult = kWabSetProp_String(kw, lContacts(idxEntry), PropTag, vSTRING)

      Case PT_SYSTIME
         kWabGetProp_SysTime kw, lContacts(idxEntry), PropTag, vSYSTIME
         If Not AskSysTime(vSYSTIME, vSYSTIME) Then Exit Sub
         Me.lblResult = kWabSetProp_SysTime(kw, lContacts(idxEntry), PropTag, vSYSTIME)

      Case PT_APPTIME
         kWabGetProp_AppTime kw, lContacts(idxEntry), PropTag, vAPPTIME
         If Not AskAppTime(vAPPTIME, vAPPTIME) Then Exit Sub
         Me.lblResult = kWabSetProp_AppTime(kw, lContacts(idxEntry), PropTag, vAPPTIME)

      Case PT_FLOAT
         kWabGetProp_Float kw, lContacts(idxEntry), PropTag, vFLOAT
         If Not AskFloat(vFLOAT, vFLOAT) Then Exit Sub
         Me.lblResult = kWabSetProp_Float(kw, lContacts(idxEntry), PropTag, vFLOAT)

      Case PT_DOUBLE
         kWabGetProp_Double kw, lContacts(idxEntry), PropTag, vDOUBLE
         If Not AskDouble(vDOUBLE, vDOUBLE) Then Exit Sub
         Me.lblResult = kWabSetProp_Double(kw, lContacts(idxEntry), PropTag, vDOUBLE)

      Case PT_BOOLEAN
         kWabGetProp_Boolean kw, lContacts(idxEntry), PropTag, vBOOLEAN
         If Not AskBoolean(vBOOLEAN, vBOOLEAN) Then Exit Sub
         Me.lblResult = kWabSetProp_Boolean(kw, lContacts(idxEntry), PropTag, vBOOLEAN)

      Case PT_CURRENCY
         kWabGetProp_Currency kw, lContacts(idxEntry), PropTag, vCURRENCY
         If Not AskCurrency(vCURRENCY, vCURRENCY) Then Exit Sub
         Me.lblResult = kWabSetProp_Currency(kw, lContacts(idxEntry), PropTag, vCURRENCY)

      Case PT_CLSID
         kWabGetProp_CLSID kw, lContacts(idxEntry), PropTag, vCLSID
         If Not AskCLSID(vCLSID, vCLSID) Then Exit Sub
         Me.lblResult = kWabSetProp_CLSID(kw, lContacts(idxEntry), PropTag, vCLSID)

      Case PT_BINARY
         If kWabGetProp_Binary_Count(kw, lContacts(idxEntry), PropTag, vBINARYlen) Then
            If vBINARYlen > 0 Then
               ReDim Preserve vBINARY(1 To vBINARYlen)
               kWabGetProp_Binary_Data kw, lContacts(idxEntry), PropTag, vBINARY(1), vBINARYlen
            End If
         End If
         If Not AskBinary(vBINARY, vBINARY) Then Exit Sub
         Me.lblResult = kWabSetProp_Binary(kw, lContacts(idxEntry), PropTag, vBINARY(LBound(vBINARY)), UBound(vBINARY) - LBound(vBINARY) + 1)

      Case PT_MV_LONGLONG
         If Not AskLongLong(vLONGLONG, vLONGLONG) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_LongLong(kw, lContacts(idxEntry), PropTag, vLONGLONG)

      Case PT_MV_LONG
         If Not AskLong(vLONG, vLONG) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_Long(kw, lContacts(idxEntry), PropTag, vLONG)
      
      Case PT_MV_SHORT
         If Not AskShort(vSHORT, vSHORT) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_Short(kw, lContacts(idxEntry), PropTag, vSHORT)

      Case PT_MV_STRING8, PT_MV_UNICODE
         If Not AskString(vSTRING, vSTRING) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_String(kw, lContacts(idxEntry), PropTag, vSTRING)
     
      Case PT_MV_SYSTIME
         If Not AskSysTime(vSYSTIME, vSYSTIME) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_SysTime(kw, lContacts(idxEntry), PropTag, vSYSTIME)

      Case PT_MV_APPTIME
         If Not AskAppTime(vAPPTIME, vAPPTIME) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_AppTime(kw, lContacts(idxEntry), PropTag, vAPPTIME)

      Case PT_MV_FLOAT
         If Not AskFloat(vFLOAT, vFLOAT) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_Float(kw, lContacts(idxEntry), PropTag, vFLOAT)

      Case PT_MV_DOUBLE
         If Not AskDouble(vDOUBLE, vDOUBLE) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_Double(kw, lContacts(idxEntry), PropTag, vDOUBLE)

      Case PT_MV_CURRENCY
         If Not AskCurrency(vCURRENCY, vCURRENCY) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_Currency(kw, lContacts(idxEntry), PropTag, vCURRENCY)

      Case PT_MV_CLSID
         If Not AskCLSID(vCLSID, vCLSID) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_CLSID(kw, lContacts(idxEntry), PropTag, vCLSID)

      Case PT_MV_BINARY
         If Not AskBinary(vBINARY, vBINARY) Then Exit Sub
         Me.lblResult = kWabSetProp_MV_AddValue_Binary(kw, lContacts(idxEntry), PropTag, vBINARY(LBound(vBINARY)), UBound(vBINARY) - LBound(vBINARY) + 1)

      Case Else
         MsgBox "Unknown/unsupported property type", vbExclamation Or vbOKOnly
   
   End Select

   lstContacts_Click
End Sub

Private Sub SetNumMultiple(ByVal HowManyValues As Long)
   Dim K As Long

   With Me.cmbWichMultiple
      .Clear
      For K = 1 To HowManyValues
         .AddItem K
      Next
      If HowManyValues > 0 Then .ListIndex = 0
      .Enabled = (HowManyValues > 0)
      .BackColor = IIf(HowManyValues > 0, &H80000005, Me.BackColor)
   End With
   Me.lblWichMulti.Enabled = (HowManyValues > 0)
End Sub
Private Sub cmdMultiAddProp_Click()
   Dim idxEntry As Long, PropTag As Long, PropType As Long
   Dim vLONGLONG As LONGLONG
   Dim vLONG As Long
   Dim vSHORT As Integer
   Dim vSTRING As String, vSTRINGlen As Long
   Dim vSYSTIME As SysTime
   Dim vAPPTIME As Date
   Dim vFLOAT As Single
   Dim vDOUBLE As Double
   Dim vCURRENCY As Currency
   Dim vCLSID As CLSID
   Dim vBINARY() As Byte

   If Not WabIsOpen Then Exit Sub
   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      idxEntry = .ItemData(.ListIndex)
      If idxEntry = 0 Then Exit Sub
   End With

   With Me.lstProps
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      PropTag = .ItemData(.ListIndex)
      If PropTag = 0 Then Exit Sub
   End With
   
   If Not PropIsMultiple(PropTag) Then Exit Sub

   Select Case kWabGetPropType(PropTag)
      Case PT_MV_LONGLONG
         If Not AskLongLong(vLONGLONG, vLONGLONG) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_LongLong(kw, lContacts(idxEntry), PropTag, vLONGLONG)

      Case PT_MV_LONG
         If Not AskLong(vLONG, vLONG) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_Long(kw, lContacts(idxEntry), PropTag, vLONG)

      Case PT_MV_SHORT
         If Not AskShort(vSHORT, vSHORT) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_Short(kw, lContacts(idxEntry), PropTag, vSHORT)

      Case PT_MV_STRING8, PT_MV_UNICODE
         If Not AskString(vSTRING, vSTRING) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_String(kw, lContacts(idxEntry), PropTag, vSTRING)

      Case PT_MV_SYSTIME
         If Not AskSysTime(vSYSTIME, vSYSTIME) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_SysTime(kw, lContacts(idxEntry), PropTag, vSYSTIME)

      Case PT_MV_APPTIME
         If Not AskAppTime(vAPPTIME, vAPPTIME) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_AppTime(kw, lContacts(idxEntry), PropTag, vAPPTIME)

      Case PT_MV_FLOAT
         If Not AskFloat(vFLOAT, vFLOAT) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_Float(kw, lContacts(idxEntry), PropTag, vFLOAT)

      Case PT_MV_DOUBLE
         If Not AskDouble(vDOUBLE, vDOUBLE) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_Double(kw, lContacts(idxEntry), PropTag, vDOUBLE)

      Case PT_MV_CURRENCY
         If Not AskCurrency(vCURRENCY, vCURRENCY) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_Currency(kw, lContacts(idxEntry), PropTag, vCURRENCY)

      Case PT_MV_CLSID
         If Not AskCLSID(vCLSID, vCLSID) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_CLSID(kw, lContacts(idxEntry), PropTag, vCLSID)

      Case PT_MV_BINARY
         If Not AskBinary(vBINARY, vBINARY) Then Exit Sub
         Me.lblResult = modkWab.kWabSetProp_MV_AddValue_Binary(kw, lContacts(idxEntry), PropTag, vBINARY(LBound(vBINARY)), UBound(vBINARY) - LBound(vBINARY) + 1)

      Case Else
         MsgBox "Unknown/unsupported multiple value type", vbExclamation Or vbOKOnly
   End Select
   
   lstProps_Click
End Sub
Private Sub cmdMultiEditProp_Click()
   Dim idxEntry As Long, PropTag As Long, nToEdit As Long
   Dim N As Long
   Dim vLONGLONG As LONGLONG
   Dim vLONG As Long
   Dim vSHORT As Integer
   Dim vSTRING As String, vSTRINGlen As Long
   Dim vSYSTIME As SysTime
   Dim vAPPTIME As Date
   Dim vFLOAT As Single
   Dim vDOUBLE As Double
   Dim vCURRENCY As Currency
   Dim vCLSID As CLSID
   Dim vBINARY() As Byte, vBINARYlen As Long


   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      idxEntry = .ItemData(.ListIndex)
      If idxEntry = 0 Then Exit Sub
   End With

   With Me.lstProps
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      PropTag = .ItemData(.ListIndex)
      If PropTag = 0 Then Exit Sub
      If Not PropIsMultiple(PropTag) Then Exit Sub
   End With

   With Me.cmbWichMultiple
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      nToEdit = .ListIndex
   End With
   
   If Not PropIsMultiple(PropTag) Then Exit Sub

   Select Case kWabGetPropType(PropTag)
      Case PT_MV_LONGLONG
         If kWabGetProp_MV_LongLong_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_LongLong_Data(kw, lContacts(idxEntry), PropTag, vLONGLONG, nToEdit) Then
                  If AskLongLong(vLONGLONG, vLONGLONG) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_LongLong(kw, lContacts(idxEntry), PropTag, nToEdit, vLONGLONG)
                  End If
               End If
            End If
         End If

      Case PT_MV_LONG
         If kWabGetProp_MV_Long_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_Long_Data(kw, lContacts(idxEntry), PropTag, vLONG, nToEdit) Then
                  If AskLong(vLONG, vLONG) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_Long(kw, lContacts(idxEntry), PropTag, nToEdit, vLONG)
                  End If
               End If
            End If
         End If
      
      Case PT_MV_SHORT
         If kWabGetProp_MV_Short_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_Short_Data(kw, lContacts(idxEntry), PropTag, vSHORT, nToEdit) Then
                  If AskShort(vSHORT, vSHORT) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_Short(kw, lContacts(idxEntry), PropTag, nToEdit, vSHORT)
                  End If
               End If
            End If
         End If

      Case PT_MV_STRING8, PT_MV_UNICODE
         If kWabGetProp_MV_String_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               vSTRINGlen = 500
               vSTRING = String(vSTRINGlen, vbNullChar)
               If kWabGetProp_MV_String_Data(kw, lContacts(idxEntry), PropTag, vSTRING, vSTRINGlen, UseUnicode, nToEdit) Then
                  vSTRING = StripNullChars(vSTRING)
                  If AskString(vSTRING, vSTRING) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_String(kw, lContacts(idxEntry), PropTag, nToEdit, vSTRING)
                  End If
               End If
            End If
         End If
     
      Case PT_MV_SYSTIME
         If kWabGetProp_MV_SysTime_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_SysTime_Data(kw, lContacts(idxEntry), PropTag, vSYSTIME, nToEdit) Then
                  If AskSysTime(vSYSTIME, vSYSTIME) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_SysTime(kw, lContacts(idxEntry), PropTag, nToEdit, vSYSTIME)
                  End If
               End If
            End If
         End If

      Case PT_MV_APPTIME
         If kWabGetProp_MV_AppTime_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_AppTime_Data(kw, lContacts(idxEntry), PropTag, vAPPTIME, nToEdit) Then
                  If AskAppTime(vAPPTIME, vAPPTIME) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_AppTime(kw, lContacts(idxEntry), PropTag, nToEdit, vAPPTIME)
                  End If
               End If
            End If
         End If

      Case PT_MV_FLOAT
         If kWabGetProp_MV_Float_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_Float_Data(kw, lContacts(idxEntry), PropTag, vFLOAT, nToEdit) Then
                  If AskFloat(vFLOAT, vFLOAT) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_Float(kw, lContacts(idxEntry), PropTag, nToEdit, vFLOAT)
                  End If
               End If
            End If
         End If

      Case PT_MV_DOUBLE
         If kWabGetProp_MV_Double_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_Double_Data(kw, lContacts(idxEntry), PropTag, vDOUBLE, nToEdit) Then
                  If AskDouble(vDOUBLE, vDOUBLE) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_Double(kw, lContacts(idxEntry), PropTag, nToEdit, vDOUBLE)
                  End If
               End If
            End If
         End If

      Case PT_MV_CURRENCY
         If kWabGetProp_MV_Currency_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_Currency_Data(kw, lContacts(idxEntry), PropTag, vCURRENCY, nToEdit) Then
                  If AskCurrency(vCURRENCY, vCURRENCY) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_Currency(kw, lContacts(idxEntry), PropTag, nToEdit, vCURRENCY)
                  End If
               End If
            End If
         End If

      Case PT_MV_CLSID
         If kWabGetProp_MV_CLSID_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               If kWabGetProp_MV_CLSID_Data(kw, lContacts(idxEntry), PropTag, vCLSID, nToEdit) Then
                  If AskCLSID(vCLSID, vCLSID) Then
                     Me.lblResult = kWabSetProp_MV_EditValue_CLSID(kw, lContacts(idxEntry), PropTag, nToEdit, vCLSID)
                  End If
               End If
            End If
         End If

      Case PT_MV_BINARY
         If kWabGetProp_MV_Binary_Count(kw, lContacts(idxEntry), PropTag, N) Then
            If nToEdit < N Then
               vBINARYlen = 0
               ReDim vBINARY(0 To 0)
               If kWabGetProp_MV_Binary_Data_Count(kw, lContacts(idxEntry), PropTag, vBINARYlen, nToEdit) Then
                  If vBINARYlen > 0 Then
                     ReDim vBINARY(1 To vBINARYlen)
                     If kWabGetProp_MV_Binary_Data_Data(kw, lContacts(idxEntry), PropTag, vBINARY(1), vBINARYlen, nToEdit) Then
                        If AskBinary(vBINARY, vBINARY) Then
                           Me.lblResult = kWabSetProp_MV_EditValue_Binary(kw, lContacts(idxEntry), PropTag, nToEdit, vBINARY(LBound(vBINARY)), UBound(vBINARY) - LBound(vBINARY) + 1)
                        End If
                     End If
                  End If
               End If
            End If
         End If

      Case Else
         MsgBox "Unknown/unsupported multiple value type", vbExclamation Or vbOKOnly
   End Select
   
   lstProps_Click
End Sub
Private Sub cmdMultiDelProp_Click()
   Dim idxEntry As Long, PropTag As Long, N As Long

   With Me.lstContacts
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      idxEntry = .ItemData(.ListIndex)
      If idxEntry = 0 Then Exit Sub
   End With

   With Me.lstProps
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      PropTag = .ItemData(.ListIndex)
      If PropTag = 0 Then Exit Sub
      If Not PropIsMultiple(PropTag) Then Exit Sub
   End With

   With Me.cmbWichMultiple
      If .ListCount < 1 Then Exit Sub
      If .ListIndex < 0 Then Exit Sub
      N = .ListIndex
   End With
   
   If MsgBox("Are you sure to remove the selected value from the property?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then Exit Sub

   Me.lblResult = kWabSetProp_MV_DelValue(kw, lContacts(idxEntry), PropTag, Me.cmbWichMultiple.ListIndex)
   
   lstProps_Click
End Sub


