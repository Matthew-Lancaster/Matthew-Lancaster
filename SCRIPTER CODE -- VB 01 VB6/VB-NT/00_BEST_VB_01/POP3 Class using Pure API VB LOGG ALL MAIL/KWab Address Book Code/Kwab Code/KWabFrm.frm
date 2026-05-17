VERSION 5.00
Begin VB.Form KWab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "kWab - 2002, Michele Locati (mlocati@tiscalinet.it)"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.ListBox lstContactsEmail 
      Height          =   5040
      IntegralHeight  =   0   'False
      Left            =   7740
      TabIndex        =   3
      Top             =   90
      Width           =   2145
   End
   Begin VB.ListBox lstContactsUser 
      Height          =   5040
      IntegralHeight  =   0   'False
      Left            =   5505
      TabIndex        =   2
      Top             =   90
      Width           =   2145
   End
   Begin VB.CommandButton cmdGo2 
      Caption         =   "Go!"
      Height          =   645
      Left            =   180
      TabIndex        =   1
      Top             =   4545
      Width           =   1575
   End
   Begin VB.TextBox txtRes 
      Height          =   4230
      Left            =   173
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   210
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8685
      TabIndex        =   5
      Top             =   5190
      Width           =   165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6405
      TabIndex        =   4
      Top             =   5205
      Width           =   165
   End
End
Attribute VB_Name = "KWab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Const UseUnicode As Boolean = False    'Should be set to False
Private Const PR_DISPLAY_NAME_A = &H3001001E
Private Const PR_EMAIL_ADDRESS_A = &H3003001E

Public Sub cmdGo2_Click()
   
   Dim kw As tWab, Mag
   Dim nEntries As Long, lEntries() As SBinary
   Dim K As Long, sBuf As String, nBuf As Long
   Dim ContactName As String, ContactEmail As String
    ChDir App.Path
    'Me.Show

   'Let's reset the textbox
   Me.txtRes = ""

   'Opening the system default wab (future versions will accept a wab file name)
   If Not (kWabOpen(kw) <> 0) Then
      'AddText "Unable to open the system wab"
        MsgBox "Unable to open the system wab"
        Stop
      Exit Sub
   End If

   'Determining the number of contacts
   nEntries = kWabGetNumEntries(kw)
   'AddText "Number of entries: " & nEntries

   If nEntries > 0 Then
      'Getting the contacts identifiers (ie used to identify a contact)
      ReDim lEntries(1 To nEntries)
      nEntries = kWabGetEntries(kw, lEntries(1), nEntries)
        
    Me.lstContactsUser.Clear
    Me.lstContactsEmail.Clear
      
      'Let's get the display name and the default email address of every contact
      For K = 1 To nEntries
      
         'We have to pass the function that retrieve a string property an inisialized string
         nBuf = 256  'Length of string
         sBuf = String(nBuf, vbNullChar)  'he string
         
        ContactEmail = ""
         'Let's get the property
         If kWabGetProp_String(kw, lEntries(K), PR_DISPLAY_NAME_A, sBuf, nBuf, UseUnicode) Then
         
            'Returned ok: strip null chars from the end of the string
            ContactName = StripNullChars(sBuf)
            
            'Me.lstContactsUser.ItemData(Me.lstContactsUser.NewIndex) = K
            'Me.lstContactsUser.ItemData = K
                     
                     
            With Me.lstContactsUser
                'GetVisName(contactname)
                .AddItem ContactName 'GetVisName(lContacts(K))
                .ItemData(.NewIndex) = K
            End With
            
            'Let's get the default email address...
            nBuf = 256
            sBuf = String(nBuf, vbNullChar)
            If kWabGetProp_String(kw, lEntries(K), PR_EMAIL_ADDRESS_A, sBuf, nBuf, UseUnicode) Then
               ContactEmail = StripNullChars(sBuf)
                
                
                Mag = 0
                If InStr(ContactEmail, "Noemailaddressfound") > 0 Then Mag = 1
                If Trim(ContactEmail) = "" Then Mag = 1
                If InStr(ContactEmail, "No email address found") > 0 Then Mag = 1
                
                                
                If Mag = 1 Then ContactEmail = ""
                
                AddText ContactEmail
                Me.lstContactsEmail.AddItem ContactEmail
                
            'Print the results
            'AddText ContactName & ": " & ContactEmail
         Else
            AddText ContactEmail
            Me.lstContactsEmail.AddItem ContactEmail
            
            'Get errors here ont know why not true is display name
            'MsgBox "Error Here -- Error retriving the display name (which is mandatory for each contact)"
            'Error retriving the display name (which is mandatory for each contact)
            'AddText "Error getting entry #" & k
         End If
         End If
        If K Mod 10 = 0 Then DoEvents

      Next

      'If we found some contact, let's free the resources we are holding
      If nEntries > 0 Then
         kWabFreeEntries kw, lEntries(1), UBound(lEntries)
         ReDim lEntries(0 To 0)
         nEntries = 0
      End If
   
   End If
   
 Me.txtRes = UCase(Me.txtRes)
   

Label1 = Val(Me.lstContactsUser.ListCount)
Label2 = Val(Me.lstContactsEmail.ListCount)
   
   
'Let's free the resource holded for maintain the wab open
'kWabClose kw
   
End Sub

Private Sub AddText(s As String) 'Append some text to the textbox
   Me.txtRes = Me.txtRes & s & vbNewLine
End Sub

Private Function StripNullChars(ByVal s As String) As String   'Remove trailing null chars from a string
   'Do While Len(s) > 0
   '   If Asc(Right(s, 1)) <> 0 Then Exit Do
   '   s = Left(s, Len(s) - 1)
   'Loop
   If InStr(s, Chr(0)) > 0 Then s = Mid(s, 1, InStr(s, Chr(0)) - 1)
   
   StripNullChars = s
End Function


Public Sub cmdAddContact_Click(ContactS)
   'Dim kw As tWab
   Dim B1$, idxEntry
   Dim Mag
   Dim nEntries As Long, lEntries() As SBinary
   Dim K As Long, sBuf As String, nBuf As Long
   Dim ContactName As String, ContactEmail As String
   
   
   Dim vLONGLONG As LONGLONG
   Dim vLONG As Long
   Dim vSHORT As Integer
   Dim vSTRING As String, vSTRINGlen2 As Long
   Dim vSYSTIME As SysTime
   Dim vAPPTIME As Date
   Dim vFLOAT As Single
   Dim vDOUBLE As Double
   Dim vBOOLEAN As Boolean
   Dim vCURRENCY As Currency
   Dim vCLSID As CLSID
   Dim vBINARY() As Byte, vBINARYlen As Long
   
   
'Exit Sub
   
   If Not (kWabOpen(kw) <> 0) Then
      'AddText "Unable to open the system wab"
        MsgBox "Unable to open the system wab"
        Stop
      Exit Sub
   End If
   
   Dim NewDisplayName As String
   'If Not WabIsOpen Then Exit Sub
   
   NewDisplayName = ContactS 'Trim(InputBox("Enter the new contact's display name"))
   'If Len(NewDisplayName) = 0 Then Exit Sub
   Dim i
   
   i = kWabAddNew(kw, NewDisplayName, UseUnicode)
   
    
    'KWab2.Show
    
    If WabIsOpen = False Then
         KWab2.cmdCloseWab_Click
    End If
    Call KWab2.cmdOpenWab_Click
    Call KWab2.cmdListContacts_Click
    Call Me.cmdGo2_Click
    
    
    For K = 1 To KWab.lstContactsUser.ListCount
        'A1$ = lstContacts.ListItems.item(WE).SubItems(1)
        B1$ = KWab.lstContactsUser.List(K)
        'Me.lvHeader2.ListItems.Add (K), , B1$
        'Me.lvHeader2.ListItems.item(WE).SubItems(1) = A1$
        If B1$ = HALO1User Then
            'idxEntry = Me.lstContactsUser.ItemData(K): Exit For
            'Chk with other Prog Compare My Patterns to know this is +1
            idxEntry = K + 1
            Exit For
        
        End If
    Next
    
    
            
                
    'Call KWab2.cmdListContacts_Click
    
    'KWab2.lstContacts.ListIndex = idxEntry - 1
    'Double Check Correct if like
    'KWab2.lstContacts.List(idxEntry - 1)
    
    'Email Property Hex
    PropTag = Val(&H3003001E)
    
    'Was In Kwab2
    Dim vSTRINDit As Long
    vSTRINDit = 500
    vSTRING = String(vSTRINDit, vbNullChar)
        
'   cb As Long
'   lpb As Long
    'lContacts(idxEntry).cb = 4
    'lContacts(idxEntry).lpb = 22694856
    'idx = 351
    
    'kWabGetProp_String kw, lContacts(idxEntry), PropTag, vSTRING, vSTRINDit, UseUnicode
    
    
    
    kWabGetProp_String kw, lContacts(idxEntry), PropTag, vSTRING, vSTRINDit, UseUnicode
    
    vSTRING = StripNullChars(vSTRING)
    'If Not AskString(vSTRING, vSTRING) Then Exit Sub
    KWab2.lblResult = kWabSetProp_String(kw, lContacts(idxEntry), PropTag, vSTRING)
    
    'Both the Contact Lists
    Call KWab2.cmdListContacts_Click
    Call Me.cmdGo2_Click
    
    
    frmMain.Command6.BackColor = &HC0C0C0
    
    
End Sub

Private Function GetVisName(e As SBinary) As String
   Dim kw As tWab
   Dim N As String, s As String

   GetVisName = ""
   N = 255
   s = String(N, vbNullChar)
   If kWabGetProp_String(kw, e, PR_DISPLAY_NAME_A, s, N, False) Then
      s = StripNullChars(s)
      GetVisName = s
   End If
End Function

