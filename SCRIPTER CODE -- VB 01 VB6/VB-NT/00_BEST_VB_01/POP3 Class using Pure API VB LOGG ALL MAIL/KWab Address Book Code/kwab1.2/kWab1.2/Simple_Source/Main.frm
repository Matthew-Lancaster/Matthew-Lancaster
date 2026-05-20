VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "kWab - 2002, Michele Locati (mlocati@tiscalinet.it)"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   645
      Left            =   4373
      TabIndex        =   1
      Top             =   4500
      Width           =   1575
   End
   Begin VB.TextBox txtRes 
      Height          =   4155
      Left            =   173
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   210
      Width           =   9975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Const UseUnicode As Boolean = False    'Should be set to False
Private Const PR_DISPLAY_NAME_A = &H3001001E
Private Const PR_EMAIL_ADDRESS_A = &H3003001E
Private Sub cmdGo_Click()
   Dim kw As tWab
   Dim nEntries As Long, lEntries() As SBinary
   Dim k As Long, sBuf As String, nBuf As Long
   Dim ContactName As String, ContactEmail As String

   'Let's reset the textbox
   Me.txtRes = ""

   'Opening the system default wab (future versions will accept a wab file name)
   If Not (kWabOpen(kw) <> 0) Then
      AddText "Unable to open the system wab"
      Exit Sub
   End If

   'Determining the number of contacts
   nEntries = kWabGetNumEntries(kw)
   AddText "Number of entries: " & nEntries

   If nEntries > 0 Then
      'Getting the contacts identifiers (ie used to identify a contact)
      ReDim lEntries(1 To nEntries)
      nEntries = kWabGetEntries(kw, lEntries(1), nEntries)
      
      'Let's get the display name and the default email address of every contact
      For k = 1 To nEntries
      
         'We have to pass the function that retrieve a string property an inisialized string
         nBuf = 256  'Length of string
         sBuf = String(nBuf, vbNullChar)  'he string
         
         'Let's get the property
         If kWabGetProp_String(kw, lEntries(k), PR_DISPLAY_NAME_A, sBuf, nBuf, UseUnicode) Then
         
            'Returned ok: strip null chars from the end of the string
            ContactName = StripNullChars(sBuf)
            
            'Let's get the default email address...
            nBuf = 256
            sBuf = String(nBuf, vbNullChar)
            If kWabGetProp_String(kw, lEntries(k), PR_EMAIL_ADDRESS_A, sBuf, nBuf, UseUnicode) Then
               ContactEmail = StripNullChars(sBuf)
            Else
               ContactEmail = " <email address not found>"
            End If

            'Print the results
            AddText ContactName & ": " & ContactEmail
         Else
            'Error retriving the display name (which is mandatory for each contact)
            AddText "Error getting entry #" & k
         End If
      Next

      'If we found some contact, let's free the resources we are holding
      If nEntries > 0 Then
         kWabFreeEntries kw, lEntries(1), UBound(lEntries)
         ReDim lEntries(0 To 0)
         nEntries = 0
      End If
   End If
   
   'Let's free the resource holded for maintain the wab open
   kWabClose kw
End Sub

Private Sub AddText(s As String) 'Append some text to the textbox
   Me.txtRes = Me.txtRes & s & vbNewLine
End Sub

Private Function StripNullChars(ByVal s As String) As String   'Remove trailing null chars from a string
   Do While Len(s) > 0
      If Asc(Right(s, 1)) <> 0 Then Exit Do
      s = Left(s, Len(s) - 1)
   Loop
   StripNullChars = s
End Function

