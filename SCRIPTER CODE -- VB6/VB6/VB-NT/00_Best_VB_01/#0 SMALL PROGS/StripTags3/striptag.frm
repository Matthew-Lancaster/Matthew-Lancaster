VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStripTag 
   Caption         =   "Strip HTML Tags"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1950
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4485
      Left            =   300
      TabIndex        =   3
      Top             =   630
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   7911
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"striptag.frx":0000
   End
   Begin VB.CommandButton cmdStripTags 
      Height          =   345
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Strip HTML tags"
      Top             =   150
      Width           =   345
   End
   Begin VB.CommandButton cmdFileSave 
      Height          =   345
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Save"
      Top             =   150
      Width           =   345
   End
   Begin VB.CommandButton cmdFileOpen 
      Height          =   345
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Open file"
      Top             =   150
      Width           =   345
   End
End
Attribute VB_Name = "frmStripTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Striptag.frm
'
' By Herman Liu
'
' Stripping HTML tags can hardly be perfect, e.g. angled brackets and their contents
' are removed when in fact they are not tags, and on the other hand, unpaired tags are
' not cleared during the process when they should be removed.  Nevertheless, it is a
' way to obtain certain desired output at times when there are no better alternatives
' on hand.
'
' This code let you open an HTML file, strip its tags, view the resulting text and
' save it to a new file if you want to.
'
' Viewing of resulting text is important, e.g. if the original HTML file does not
' have paired angle brackets in all places, some of the tags would remain there after
' the first stripping process.  In this case, you may check the first occurance of an
' angle tag, follow it up to see whether you have to manually remove the odd one, or
' make it paired, on the screen.  You may then click the "Strip Tags" command button
' again for a second pass of the stripping process, or even a third one, until all
' tags are cleansed.

Option Explicit

Dim mFileSpec As String
Dim gCancel As Boolean
Dim gcdg As Object


Private Sub Form_Load()
    Set gcdg = CommonDialog1
End Sub


Public Sub cmdFileOpen_Click()
    
    
    'On Error GoTo errHandler
    Dim mHandle
    
    GoTo Skip1
    
    gcdg.Filter = "(*.html)|*.html|(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.DefaultExt = "htm"
    gcdg.Flags = cdlOFNFileMustExist
    gcdg.Filename = ""
    gcdg.CancelError = True
    gcdg.ShowOpen
    If gcdg.Filename = "" Then
        Exit Sub
    End If
    
    mFileSpec = gcdg.Filename
    
Skip1:
    mFileSpec = XFileSpec
    
    mHandle = FreeFile
    Open mFileSpec For Input As mHandle
    Text1 = StrConv(InputB(LOF(mHandle), 1), vbUnicode)
    Close #mHandle
    'Text1.SetFocus
    Exit Sub
errHandler:
    If Err <> 32755 Then
         ErrMsgProc "cmdFileOpen_Click"
            'err.description
    End If
    Text1.SetFocus
End Sub



Private Sub cmdFileSave_Click()
    On Error GoTo errHandler
    gCancel = False
    Dim mPath As String
    mPath = CurDir
    gcdg.Filename = mFileSpec
    gcdg.Filter = "(*.txt)|*.txt|(*.*)|*.*|"
    gcdg.DefaultExt = "txt"
    gcdg.FilterIndex = 1
    gcdg.Flags = cdlOFNOverwritePrompt
    gcdg.CancelError = True
    gcdg.ShowSave
    mFileSpec = gcdg.Filename
    Text1.SaveFile mFileSpec, 1
    ChDir mPath
    Text1.SetFocus
    Exit Sub
errHandler:
    If Err <> 32755 Then
         ErrMsgProc "cmdFileSave"
    Else
         gCancel = True
    End If
    Text1.SetFocus
End Sub





 
Public Sub cmdStripTags_click()
    On Error Resume Next
    Dim strContent As String, mString As String
    Dim mStartPos As Long, mEndPos As Long
    Dim i, j
    strContent = Text1.Text
       ' Start process
    mStartPos = InStr(strContent, "<")
    mEndPos = InStr(strContent, ">")
    frmMain.ProgressBar1.Max = Len(Text1.Text)
    
    Do While mStartPos <> 0 And mEndPos <> 0 And mEndPos > mStartPos
          mString = Mid(strContent, mStartPos, mEndPos - mStartPos + 1)
          strContent = Replace(strContent, mString, "")
          mStartPos = InStr(strContent, "<")
          mEndPos = InStr(strContent, ">")
            frmMain.ProgressBar1.Value = mStartPos
    Loop
       ' Translate common escape sequence chars
    'strContent = Replace(LCase$(strContent), "<br>", " ")
    
    strContent = Replace(strContent, "document.write(""Click here for Famous Birthdays for this day."");", " ")
    strContent = Replace(strContent, " document.write(""Click here for Music history for this day."");", " ")
    
    
    strContent = Replace(strContent, "&nbsp;", " ")
    strContent = Replace(strContent, "&amp;", "&")
    strContent = Replace(strContent, "&quot;", "'")
    strContent = Replace(strContent, "&#", "#")
    strContent = Replace(strContent, "&lt;", "<")
    strContent = Replace(strContent, "&gt;", ">")
    strContent = Replace(strContent, "%20", " ")
    strContent = LTrim(Trim(strContent))
    Do While Left(strContent, 1) = Chr$(13) Or Left(strContent, 1) = Chr$(10)
          strContent = Mid(strContent, 2)
    Loop
    Text1.Text = strContent
       ' If any angle brackets still exist, highlight the first one
    i = InStr(Text1.Text, "<")
    j = InStr(Text1.Text, ">")
    If j < i And j > 0 Then i = j
    If i > 0 Then
          Text1.SelStart = i - 1
          Text1.SelLength = 1
    ElseIf j > 0 Then
          Text1.SelStart = j - 1
          Text1.SelLength = 1
    End If
    Text1.SetFocus
End Sub




Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub




