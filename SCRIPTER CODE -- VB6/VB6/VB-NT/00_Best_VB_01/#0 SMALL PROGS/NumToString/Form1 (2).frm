VERSION 5.00
Begin VB.Form Frm_StrToNum 
   BackColor       =   &H80000007&
   Caption         =   "StringToNum"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   13815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2745
      Left            =   15
      TabIndex        =   1
      Top             =   1755
      Width           =   13800
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   13800
   End
End
Attribute VB_Name = "Frm_StrToNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i22 As Currency, i3 As Currency, CmdCancle
Dim Ace, BigWord, LenTT2


Private Sub Form_Load()
Me.Show
Me.Refresh
DoEvents
i3 = 0
i22 = 0

'Dim TEST

If IsIDE = True Then
    On Error Resume Next
        Kill App.Path + "\#0 DataText\*.*"
    On Error GoTo 0
End If

On Error Resume Next
Open App.Path + "\#0 DataText\Number To String List-VarCount1.txt" For Input As #1
    i22 = Val(Input(LOF(1), 1))
Close #1
Open App.Path + "\#0 DataText\Number To String List-VarCount2.txt" For Input As #1
    i3 = Val(Input(LOF(1), 1))
Close #1
Open App.Path + "\#0 DataText\Number To String List-VarBigWord.txt" For Input As #1
    Line Input #1, ll1: BigWord = Val(ll1)
    Line Input #1, ll1: LenTT2 = Val(ll1)
    Line Input #1, ll1: Ace = ll1
Close #1
On Error GoTo 0




f1 = FreeFile
Open App.Path + "\#0 DataText\Number To String List -- " + Format(i3, "00000") + ".txt" For Append Lock Write As #f1
Print #f1, Format(Now, "DDD-DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
i3 = i3 + 10


f2 = FreeFile
Open App.Path + "\#0 DataText\Number To String List-VarCount2.txt" For Output As #f2
    Print #f2, Str(i3)
Close #f2


f2 = FreeFile
Open App.Path + "\#0 DataText\Number To String List-Big Word Logg.txt" For Append As #f2
Print #f2, Format(Now, "DDD-DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf

f3 = FreeFile
Open App.Path + "\#0 DataText\Number To String List-Acroynm List.txt" For Append As #f3
Print #f3, Format(Now, "DDD-DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf


f4 = FreeFile
Open App.Path + "\#0 DataText\Number To String List Simple -- " + Format(i3, "00000") + ".txt" For Append Lock Write As #f4
Print #f4, Format(Now, "DDD-DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf


i = 0
Do
    wc = 0
    DoEvents
    i = i + 1
    tt = NumToString(i22)
    
    agwa = ""
    
    outstg = UCase(tt)
    x = 0
    Do
        x = InStr(x + 1, outstg, "-")
        If x > 0 Then
            Mid$(outstg, x, 1) = " "
        End If
    Loop Until x = 0
    
    x = 0
    Do
        x = InStr(x + 1, outstg, " ")
        If x > 0 Then
            agwa = agwa + Trim(Mid$(outstg, x + 1, 1))
        End If
    Loop Until x = 0
    
    
    x = 0
    Do
        x = InStr(x + 1, tt, " ")
        If x > 0 Then Mid$(tt, x + 1, 1) = UCase(Mid$(tt, x + 1, 1))
    Loop Until x = 0
    
    x = 0
    Do
        x = InStr(x + 1, tt, "-")
        If x > 0 Then Mid$(tt, x + 1, 1) = UCase(Mid$(tt, x + 1, 1))
    Loop Until x = 0
    wc = Len(agwa)
    
    tt = Trim(tt)
    Print #f1, Format(i22, "000,000,000,000") + " -- Acronym " + agwa + " -- #-Words " + Format(wc, "000") + " -- LEN " + Format(Len(tt), "0000") + " -- " + tt
    
    Print #f3, Format(i22, "000,000,000,000") + " -- " + agwa
    Print #f4, Format(i22, "000,000,000,000") + " -- " + tt
    
    i22 = i22 + 1
    
    If wc > BigWord And Len(tt) > LenTT2 Then
        BigWord = wc: LenTT2 = Len(tt): Ace = tt
        Print #f2, Format(i22, "000,000,000,000") + " -- #-Words  " + Format(BigWord, "000") + " -- LEN " + Format(Len(tt), "0000") + " -- " + tt
    End If
    
    If i22 Mod 100 = 0 Then
        Label1 = Format(i22, ",###,###,###,###,###,##0") + "-" + Trim(Str(BigWord))
        Label2 = tt
    End If
    
Loop Until i > 1500000 Or CmdCancle = True ' One Million
Close #f1, f2, f3, f4

f1 = FreeFile
Open App.Path + "\#0 DataText\Number To String List-VarCount.txt" For Output As #f1
    Print #f1, Str(i22);
Close #f1
'End

Open App.Path + "\#0 DataText\Number To String List-VarBigWord.txt" For Output As #1
    Print #1, BigWord
    Print #1#, LenTT2
    Print #1, Ace
Close #1

'Reset
'Open App.Path + "\#0 DataText\Number To String List-VarCount.txt" For Output As #1
'    Print #1, Str(i22);
'Close #1
'End

Unload Me
End Sub


'Sub BumToString()





'NumToString

'Description:
' The function NumToString writes out any number, up to about 920 trillion, the limit of the Currency variable type, in English words. It only writes out the integer portion of the number. The DollarToString function does the same, but places the word dollars after the string and also writes out the fractional part of the value as the cents.

'1/8/2001: Added DateToString and NumToStringTh (ie, "23" -> "twenty-third") helper functions.

'7/20/2002: Corrected error in DollarToString that limited the range of values.
  
'Code:
 Public Function DollarToString(ByVal nAmount As Currency) As String
'Example:
'  DollarToString(5.99) = " five dollars and ninety-nine cents"
    Dim nDollar As Currency
    Dim nCent As Currency
    
    nDollar = Int(nAmount)
    nCent = (Abs(nAmount) - Int(Abs(nAmount))) * 100
    
    DollarToString = NumToString(nDollar) & " dollar"
    
    If Abs(nDollar) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
    
    DollarToString = DollarToString & " and" & _
    NumToString(nCent) & " cent"
    
    If Abs(nCent) <> 1 Then
        DollarToString = DollarToString & "s"
    End If
End Function

Public Function NumToString(ByVal nNumber As Currency) As String
'Example: NumToString(123) = " one hundred twenty-three"
    Dim bNegative As Boolean
    Dim bHundred As Boolean

    If nNumber < 0 Then
        bNegative = True
    End If

    nNumber = Abs(Int(nNumber))
    If nNumber < 1000 Then
        If nNumber \ 100 > 0 Then
            NumToString = NumToString & _
                NumToString(nNumber \ 100) & " hundred"
            bHundred = True
        End If
        nNumber = nNumber - ((nNumber \ 100) * 100)
        Dim bNoFirstDigit As Boolean
        bNoFirstDigit = False
        Select Case nNumber \ 10
            Case 0
                Select Case nNumber Mod 10
                    Case 0
                        If Not bHundred Then
                            NumToString = NumToString & " zero"
                        End If
                    Case 1: NumToString = NumToString & " one"
                    Case 2: NumToString = NumToString & " two"
                    Case 3: NumToString = NumToString & " three"
                    Case 4: NumToString = NumToString & " four"
                    Case 5: NumToString = NumToString & " five"
                    Case 6: NumToString = NumToString & " six"
                    Case 7: NumToString = NumToString & " seven"
                    Case 8: NumToString = NumToString & " eight"
                    Case 9: NumToString = NumToString & " nine"
                End Select
                bNoFirstDigit = True
            Case 1
                Select Case nNumber Mod 10
                    Case 0: NumToString = NumToString & " ten"
                    Case 1: NumToString = NumToString & " eleven"
                    Case 2: NumToString = NumToString & " twelve"
                    Case 3: NumToString = NumToString & " thirteen"
                    Case 4: NumToString = NumToString & " fourteen"
                    Case 5: NumToString = NumToString & " fifteen"
                    Case 6: NumToString = NumToString & " sixteen"
                    Case 7: NumToString = NumToString & " seventeen"
                    Case 8: NumToString = NumToString & " eighteen"
                    Case 9: NumToString = NumToString & " nineteen"
                End Select
                bNoFirstDigit = True
            Case 2: NumToString = NumToString & " twenty"
            Case 3: NumToString = NumToString & " thirty"
            Case 4: NumToString = NumToString & " forty"
            Case 5: NumToString = NumToString & " fifty"
            Case 6: NumToString = NumToString & " sixty"
            Case 7: NumToString = NumToString & " seventy"
            Case 8: NumToString = NumToString & " eighty"
            Case 9: NumToString = NumToString & " ninety"
        End Select
        If Not bNoFirstDigit Then
            If nNumber Mod 10 <> 0 Then
                NumToString = NumToString & "-" & _
                    Mid(NumToString(nNumber Mod 10), 2)
            End If
        End If
    Else
        Dim nTemp As Currency
        nTemp = 10 ^ 12 'trillion
        Do While nTemp >= 1
            If nNumber >= nTemp Then
                NumToString = NumToString & _
                    NumToString(Int(nNumber / nTemp))
                Select Case Int(Log(nTemp) / Log(10) + 0.5)
                    Case 12: NumToString = NumToString & " trillion"
                    Case 9: NumToString = NumToString & " billion"
                    Case 6: NumToString = NumToString & " million"
                    Case 3: NumToString = NumToString & " thousand"
                End Select
                nNumber = nNumber - (Int(nNumber / nTemp) * nTemp)
            End If
            nTemp = nTemp / 1000
        Loop
    End If
    If bNegative Then
        NumToString = " negative " & NumToString
    End If
End Function

Public Function NumToStringTh(ByVal nNumber As Currency) As String
'Example: NumToStringTh(123) = " one hundred twenty-third"

    Dim sNum As String
    Dim sExtra As String
    Dim nSpace As String

    'Convert the number
    sNum = NumToString(nNumber)
    'Find the location of the last space or dash
    nSpace = Len(sNum)
    
    Do Until Mid(sNum, nSpace, 1) = " " Or _
        Mid(sNum, nSpace, 1) = "-"
        nSpace = nSpace - 1
    Loop

    sExtra = Mid(sNum, nSpace + 1)
    sNum = Mid(sNum, 1, nSpace)
    
    'Conver the last word ("one" -> "first", etc)
    Select Case sExtra
        Case "hundred":   NumToStringTh = sNum & "hundredth"
        Case "zero" 'no such thing as 'zeroth'
            NumToStringTh = sNum & "zero"
        Case "one":       NumToStringTh = sNum & "first"
        Case "two":       NumToStringTh = sNum & "second"
        Case "three":     NumToStringTh = sNum & "third"
        Case "four":      NumToStringTh = sNum & "fourth"
        Case "five":      NumToStringTh = sNum & "fifth"
        Case "six":       NumToStringTh = sNum & "sixth"
        Case "seven":     NumToStringTh = sNum & "seventh"
        Case "eight":     NumToStringTh = sNum & "eighth"
        Case "nine":      NumToStringTh = sNum & "ninth"
        Case "ten":       NumToStringTh = sNum & "tenth"
        Case "eleven":    NumToStringTh = sNum & "eleventh"
        Case "twelve":    NumToStringTh = sNum & "twelfth"
        Case "thirteen":  NumToStringTh = sNum & "thirteenth"
        Case "fourteen":  NumToStringTh = sNum & "fourteenth"
        Case "fifteen":   NumToStringTh = sNum & "fifteenth"
        Case "sixteen":   NumToStringTh = sNum & "sixteenth"
        Case "seventeen": NumToStringTh = sNum & "seventeenth"
        Case "eighteen":  NumToStringTh = sNum & "eighteenth"
        Case "nineteen":  NumToStringTh = sNum & "nineteenth"
        Case "twenty":    NumToStringTh = sNum & "twentieth"
        Case "thirty":    NumToStringTh = sNum & "thirtieth"
        Case "forty":     NumToStringTh = sNum & "fortieth"
        Case "fifty":     NumToStringTh = sNum & "fiftieth"
        Case "sixty":     NumToStringTh = sNum & "sixtieth"
        Case "seventy":   NumToStringTh = sNum & "seventieth"
        Case "eighty":    NumToStringTh = sNum & "eightieth"
        Case "ninety":    NumToStringTh = sNum & "ninetieth"
        Case "trillion":  NumToStringTh = sNum & "trillionth"
        Case "billion":   NumToStringTh = sNum & "billionth"
        Case "million":   NumToStringTh = sNum & "millionth"
        Case "thousand":  NumToStringTh = sNum & "thousandth"
        Case Else 'This shouldn't happen, but just in case
            NumToStringTh = NumToString(nNumber)
    End Select
    
End Function

Public Function DateToString(FromDate As String) As String
'Example: ' DateToString(#1/1/2001#) = "January the first, two thousand one"

    DateToString = Format(FromDate, "mmmm") & " the" & _
        NumToStringTh(DatePart("d", FromDate)) & "," & _
        NumToString(DatePart("yyyy", FromDate))
        
End Function

 
  
'Sample Usage:
  
     'Debug.Print DollarToString(1234.11)
    'Debug.Print NumToString(-54321)

 
'-----------------------------------------------------------------
'-----------------------------------------------------------------
'-----------------------------------------------------------------



Private Sub Form_Unload(Cancel As Integer)
CmdCancle = True
End Sub
'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************

