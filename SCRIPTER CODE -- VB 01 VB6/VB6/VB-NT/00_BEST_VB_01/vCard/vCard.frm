VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   14445
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   75
      TabIndex        =   1
      Top             =   120
      Width           =   12540
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   12885
      TabIndex        =   0
      Top             =   195
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   105
      TabIndex        =   2
      Top             =   3270
      Width           =   7395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#### THIS IS FOR THE NOKIA
'#### SEEMS IMPORT CONTACTS AND THEN BACKUP MAYBE CHANGE THEM
'MODIFIY STRIP END SPACES
'MAKE FIRST ONE NAME EVEN IF MULTI

Sub Compare_Without_Name()

'File1.Path = "F:\Others\t\test combine"
For i = 0 To File1.ListCount - 1
    'Debug.Print File1.List(i)
    Label1.Caption = Str(i) + " of " + Str(File1.ListCount - 1)
    DoEvents
    Open File1.Path + "\" + File1.List(i) For Input As #1
    Do
        Line Input #1, LINE1
        If Mid(LINE1, 1, 2) = "N:" Or Mid(LINE1, 1, 2) = "N;" Then dropone = True Else dropone = False
        If dropone = False Then
        If InStr(LINE1, "CELL:") > 0 Then
            LINE1 = Replace(LINE1, "TEL;CELL:", "TEL:")
        End If
        
            storedb1 = storedb1 + LINE1 + vbCrLf
        End If
    Loop Until EOF(1)
    Close #1
Next

'KEEP TOP FRONT OF ADDY BOOK PRIORITY
'SWAP BACK AND FORWARD A BIT
For i = File1.ListCount - 1 To 0 Step -1
'For i = 0 To File1.ListCount - 1
    'Debug.Print File1.List(i)
    Label1.Caption = Str(i) + " of " + Str(File1.ListCount - 1)
    DoEvents
    
    Open File1.Path + "\" + File1.List(i) For Input As #1
    storedb2 = ""
    Do
        Line Input #1, LINE1
        If Mid(LINE1, 1, 2) = "N:" Or Mid(LINE1, 1, 2) = "N;" Then dropone = True Else dropone = False
        If dropone = False Then
            If InStr(LINE1, "CELL:") > 0 Then
                LINE1 = Replace(LINE1, "TEL;CELL:", "TEL:")
            End If
            storedb2 = storedb2 + LINE1 + vbCrLf
        End If
    Loop Until EOF(1)
    Close #1
    
    'X1 = 1
    'Do
        A1 = InStr(storedb1, storedb2)
        a2 = InStr(A1 + 1, storedb1, storedb2)
        
        X1 = A1 + 1
        If a2 > 0 Then
'            Kill File1.Path + "\" + File1.List(i)
            xc = xc + 1
            List1.AddItem Str(xc) + "-" + Str(i) + "-" + File1.Path + "\" + File1.List(i)
            Mid(storedb1, A1, Len(storedb2)) = Space(Len(storedb2))
        End If
        
    'Loop Until a2 = 0
Next



End Sub


Sub PAGE_OF_TEXT_TO_VCARD()


Open "F:\My Folder\Documents\My Docs - Dad\00 Blue Tooth Exchange Folder\00 CONTACTS\CallerID 2nd Copy\Dad'S Phone.txt" For Input As #1


xc = 0
Do
    Line Input #1, LINE1
    xc = xc + 1
    Label1.Caption = Str(xc)
    DoEvents
    If Mid(LINE1, 1, 5) = "#MATT" Then
        EXTRADAD = True
        
    End If
    
    If Trim(LINE1) <> "" And Mid(LINE1, 1, 1) <> "#" Then
        '# NUMBER
        NAMEADDY1 = Trim(Mid(LINE1, 1, InStr(LINE1, ";") - 1))
        If InStr(NAMEADDY1, " ") > 0 Then
            NAMEADDY1 = Trim(Mid(NAMEADDY1, 1, InStr(NAMEADDY1, " ")))
        End If
        '# NAME
        NAMEADDY2 = Trim(Mid(LINE1, InStr(LINE1, ";") + 1))
        'nameaddy2 = Replace(nameaddy2, " ", "_")
        
    
        
        If EXTRADAD = False Then
            NAMEADDY2 = "RAY " + NAMEADDY2
        End If
        
        manipulate1 = NAMEADDY2
        
        
        'Phone Strip's Wild Card for Vcard
        manipulate1 = Replace(manipulate1, "*", "")
        manipulate1 = Replace(manipulate1, "?", "")
        manipulate1 = Replace(manipulate1, ">", "")
        manipulate1 = Replace(manipulate1, "<", "")
        manipulate1 = Replace(manipulate1, "\", "")
        manipulate1 = Replace(manipulate1, "/", "")
        manipulate1 = Replace(manipulate1, ":", "")
        manipulate1 = Replace(manipulate1, "|", "")
        manipulate1 = Replace(manipulate1, """", "")
        
        dupecount = 1
        Do
            If Dir(File1.Path + "-Test\" + manipulate1 + ".vcf") <> "" Then
                dupecount = dupecount + 1
                manipulate1 = manipulate1 + "[Dupe" + Format(dupecount, "00") + "]"
            End If
        Loop Until Dir(File1.Path + "-Test\" + manipulate1 + ".vcf") = ""
        
        Open File1.Path + "-Test\" + manipulate1 + ".vcf" For Output As #5
            Print #5, "BEGIN: VCARD"
            Print #5, "VERSION:2.1"
            Print #5, "N:;" + NAMEADDY2 + ";;;"
            Print #5, "TEL;CELL:" + NAMEADDY1
            Print #5, "X-CLASS:private"
            Print #5, "End: VCARD"
        Close #5
    End If
    
Loop Until EOF(1)
Close #1



'BEGIN: VCARD
'VERSION:2.1
'N:;# 0 Matt Home;;;
'TEL;CELL:01273421969
'X-CLASS:private
'End: VCARD
'Stop


End Sub

Private Sub Form_Load()
Me.Show
DoEvents

FILEPATH = "D:\0 00 MOBILE\MOBILE\MOBILE NOKIA - COMMON\Others\Contacts"
File1.Path = FILEPATH

'Call Compare_Without_Name
'Exit Sub



If Dir(File1.Path + "-Test", vbDirectory) = "" Then
    MkDir File1.Path + "-Test"
End If

File1.Path = FILEPATH + "-Test"

For i = 0 To File1.ListCount - 1
    Kill File1.Path + "\" + File1.List(i)
    Label1.Caption = Str(i) + " of " + Str(File1.ListCount - 1)
    DoEvents
Next

'File1.Path = "F:\Others\Contacts"
File1.Path = FILEPATH



Call PAGE_OF_TEXT_TO_VCARD


'----------------------------------------

tempfile1 = "C:\temp\Vcard_Test.vcf"
tempfile2 = "C:\temp\TempVcard.vcf"
Open tempfile1 For Output As #2


'File1.Path = "F:\Others\Contacts"
For i = 0 To File1.ListCount - 1
    'Debug.Print File1.List(i)
    Label1.Caption = Str(i) + " of " + Str(File1.ListCount - 1)
    DoEvents
    Open tempfile2 For Output As #3
    Open File1.Path + "\" + File1.List(i) For Input As #1
'    If File1.List(i) = "BT LadL21Num#Divert All.vcf" Then Stop
'    If File1.List(i) = "# 0 Matt Home.vcf" Then Stop
    
        
    
    Do
        
        
        Line Input #1, LINE1
        
        'If Mid(LINE1, 1, 2) = "N;" Then namevar1 = Mid(File1.List(i), 1, InStrRev(File1.List(i), ".") - 1)
        
        
        If Mid(LINE1, 1, 2) = "N:" Or Mid(LINE1, 1, 2) = "N;" Then
            List1.AddItem "--" + LINE1
            
            If Mid(LINE1, 3, 1) = ";" Then
                If InStr(LINE1, ";;;;") > 0 Then Stop ' Never
                LINE1 = "N:" + Mid(LINE1, 4) + ";"
            End If
            
            Do
                'If InStr(line1, "#9") > 0 Then Stop
                
                'If Mid(line1, 1, 2) = "N;" Then Stop
                
                line2 = LINE1
                A1 = InStr(3, LINE1, ";")
                a2 = InStr(LINE1, ";;")
                If A1 > 0 And A1 <> a2 Then
                    
                    manipulate1 = Mid(LINE1, A1 + 1, InStr(A1 + 1, LINE1, ";") - A1 - 1)
                    
                    manipulate2 = Mid(LINE1, 1, A1)
                    If InStr(manipulate2, manipulate1) > 0 Then
                        List1.AddItem LINE1
                        
                        LINE1 = Replace(LINE1, ";" + manipulate1 + ";", ";") + ";"
                        
                    End If
                End If
            Loop Until line2 = LINE1
            
            Mid(LINE1, 3, 1) = UCase(Mid(LINE1, 3, 1))
            
            List1.AddItem LINE1
            
            
            namevar1 = LINE1
            
            
            A1 = InStr(LINE1, ";;")
            If A1 > 0 Then
                namevar1 = Mid(namevar1, 3, A1 - 3)
                If InStr(namevar1, ";") > 0 Then
                    namevar2 = Mid(namevar1, InStr(namevar1, ";") + 1)
                    namevar1 = Mid(namevar1, 1, InStr(namevar1, ";")) + namevar2
                    namevar1 = Replace(namevar1, ";", " ")
                End If
            Else
            '"N:Argos shop"
            namevar1 = Mid(namevar1, 3)
            End If
 
 
                
            Do
            
                Modtrue = 0
                line2 = LINE1
                If InStr(LINE1, " ;") > 0 Then Modtrue = 1
                LINE1 = Replace(LINE1, " ;", ";")
                
                If Modtrue = 1 Then List1.AddItem LINE1
                namevar1 = LINE1
                A1 = InStr(LINE1, ";;")
                If A1 > 0 Then
                    namevar1 = Mid(namevar1, 3, A1 - 3)
                Else
                    '"N:Argos shop"
                    namevar1 = Mid(namevar1, 3)
                End If
                
                If InStr(namevar1, ";") > 0 Then
                    namevar2 = Mid(namevar1, InStr(namevar1, ";") + 1)
                    namevar1 = Mid(namevar1, 1, InStr(namevar1, ";")) + namevar2
                    namevar1 = Replace(namevar1, ";", " ")
                End If
            Loop Until line2 = LINE1
            
            
            
        
            If InStr(LINE1, "ENCODING=QUOTED-PRINTABLE") > 0 Then
                'namevar1 = Mid(File1.List(i), 1, InStrRev(File1.List(i), ".") - 1)
                X2 = namevar1
                X2 = Replace(X2, "ENCODING=QUOTED-PRINTABLE:", "")
                X2 = Replace(X2, "=40", "@")
                X2 = Replace(X2, "=20", " ")
                namevar1 = Trim(X2)
            End If
            'MADE MISTAKE IN CODE NOT HAVE LEARN - FIRST TEXT BLOCK IS SURNAME FIRST - SO PUT BACK TO FIRST NAME
            'LIKE THIS
            If Mid(LINE1, 1, 2) = "N:" Then
                If InStr(LINE1, ";;;;") > 0 Then
                    LINE1 = "N:;" + Mid(LINE1, 3)
                    LINE1 = Replace(LINE1, ";;;;", ";;;")
                End If
            End If
        End If
        
        
        
        'List1.AddItem manipulate
        
        Print #2, LINE1
        Print #3, LINE1
    Loop Until EOF(1)
    Close #1, #3


    
    'a1 = InStr(namevar1, ";")
    Mid(namevar1, 1, 1) = UCase(Mid(namevar1, 1, 1))
    manipulate1 = namevar1
    
    'Phone Strip's Wild Card for Vcard
    manipulate1 = Replace(manipulate1, "*", "")
    manipulate1 = Replace(manipulate1, "?", "")
    manipulate1 = Replace(manipulate1, ">", "")
    manipulate1 = Replace(manipulate1, "<", "")
    manipulate1 = Replace(manipulate1, "\", "")
    manipulate1 = Replace(manipulate1, "/", "")
    manipulate1 = Replace(manipulate1, ":", "")
    manipulate1 = Replace(manipulate1, "|", "")
    manipulate1 = Replace(manipulate1, """", "")
    TDK1 = "1": TDK2 = "2"
    dupecount = 1
    Do
        If Dir(File1.Path + "-Test\" + manipulate1 + ".vcf") <> "" Then
            Open File1.Path + "-Test\" + manipulate1 + ".vcf" For Input As #8
                TDK1 = Input$(LOF(8), #8)
            Close #8
            Open tempfile2 For Input As #8
                TDK2 = Input$(LOF(8), #8)
            Close #8
            If TDK2 <> TDK1 Then
'                dupecount = dupecount + 1
'                manipulate1 = manipulate1 + "[Dupe" + Format(dupecount, "00") + "]"
            End If
        End If
    Loop Until Dir(File1.Path + "-Test\" + manipulate1 + ".vcf") = "" Or TDK2 = TDK1 Or 1 = 1
    
    If Dir(File1.Path + "-Test\" + manipulate1 + ".vcf") = "" Then
        Open tempfile2 For Input As #4
            TDK1 = Input$(LOF(4), #4)
        Close #4
        Open File1.Path + "-Test\" + manipulate1 + ".vcf" For Output As #4
            Print #4, TDK1
        Close #4
    End If
    Debug.Print Str(i) + " of" + Str(File1.ListCount - 1)


Next
Close #2






'Shell "C:\Program Files\TextView\Textview.exe """ + vFile + """", vbMaximizedFocus
'Shell "Notepad C:\temp\Vcard_Test.vcf", vbNormalFocus
'End
End
End Sub
