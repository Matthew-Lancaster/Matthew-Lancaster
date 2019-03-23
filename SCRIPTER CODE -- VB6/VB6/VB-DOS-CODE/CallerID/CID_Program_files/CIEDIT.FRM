VERSION 5.00
Begin VB.Form CiEDIT 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CiEDIT"
   ClientHeight    =   3270
   ClientLeft      =   10245
   ClientTop       =   1755
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3270
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Add_Next_New_Number 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "&Add Next New Number"
      Height          =   570
      Left            =   6765
      TabIndex        =   33
      Top             =   720
      Width           =   2820
   End
   Begin VB.CommandButton Dont_Treat 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "Dont Treat as Called No."
      Height          =   390
      Left            =   5760
      TabIndex        =   32
      Top             =   2880
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.TextBox New_In 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   4905
      TabIndex        =   31
      Top             =   1290
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5010
      Top             =   2895
   End
   Begin VB.CommandButton let_go_command 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "&Let Rip"
      Height          =   390
      Left            =   1080
      TabIndex        =   29
      Top             =   2865
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton pick_up_Command 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "&Pick up"
      Height          =   390
      Left            =   0
      TabIndex        =   28
      Top             =   2865
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox Speed_search_text 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3240
   End
   Begin VB.TextBox Speedsearchtext 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Speed Search:"
      Top             =   240
      Width           =   1560
   End
   Begin VB.CommandButton findbutton 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "&F3 = Find Again"
      Height          =   240
      Left            =   5640
      TabIndex        =   3
      Top             =   0
      Width           =   2280
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Date & Time :"
      Top             =   0
      Width           =   1560
   End
   Begin VB.TextBox datetime 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "MON 00/JAN/1900 12:00:00 PM"
      Top             =   0
      Width           =   3240
   End
   Begin VB.TextBox BUSY 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   " BUSY.    "
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox recordpositiontext 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Record Position :"
      Top             =   240
      Width           =   1800
   End
   Begin VB.TextBox recpostext 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7440
      TabIndex        =   8
      Top             =   240
      Width           =   2160
   End
   Begin VB.CommandButton Formularbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Edit Menu"
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   9600
   End
   Begin VB.TextBox The_description_Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2160
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   7440
   End
   Begin VB.CommandButton rec_Up_Command 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "&Up."
      Height          =   390
      Left            =   3240
      TabIndex        =   16
      Top             =   1290
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Rec_down_Command 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "&Down."
      Height          =   390
      Left            =   4080
      TabIndex        =   17
      Top             =   1290
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox The_description_Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   5520
      TabIndex        =   22
      Top             =   2070
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.TextBox Number_Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   7965
      TabIndex        =   23
      Text            =   "01273-"
      Top             =   2070
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Date_Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2160
      TabIndex        =   24
      Top             =   2460
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.TextBox Time_Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   6840
      TabIndex        =   25
      Top             =   2460
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.CommandButton Save_command 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "&Save"
      Height          =   390
      Left            =   8145
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton editexitcommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "E&xitbutton"
      Height          =   240
      Left            =   7920
      TabIndex        =   27
      Top             =   0
      Width           =   1680
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5100
      Left            =   0
      TabIndex        =   1
      Top             =   4770
      Width           =   9600
   End
   Begin VB.CommandButton Edit_item 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Edit &item mode."
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   2820
   End
   Begin VB.CommandButton Add_item_command 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "&Add Item to log."
      Height          =   570
      Left            =   3435
      TabIndex        =   18
      Top             =   720
      Width           =   2820
   End
   Begin VB.Label The_description_label 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&The description"
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   0
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label pick_up_label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   2160
      TabIndex        =   30
      Top             =   2865
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Number_label 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2160
      TabIndex        =   13
      Top             =   2070
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.Label RecNo_Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Record No."
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   0
      TabIndex        =   14
      Top             =   1290
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Recno_Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2160
      TabIndex        =   15
      Top             =   1290
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Number_label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Number"
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   0
      TabIndex        =   19
      Top             =   2070
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Date_Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Date"
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   0
      TabIndex        =   20
      Top             =   2460
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Time_Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T&ime"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   4800
      TabIndex        =   21
      Top             =   2460
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "CiEDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This is the main first entry point in this form Here:-
'---------------------------
'Private Sub The_description_Text_Change()

'These Are the 3 Entry Points in the program main menu From
'the Point Above
'------------------------------------
'Private Sub Add_item_command_Click()
'Private Sub Edit_item_Click()
'Private Sub Add_Next_New_Number_click()

Dim Date_An_Time2 As Double
Dim Date_An_Time As Double
Dim Yup As Double
Dim Eat2 As Integer
Dim Drt As Integer
Dim Rtd As Integer
Dim Gdy As Integer
Dim KeyPass As Integer
Dim LineSelected As Integer
Dim NumberToChange As Integer
Dim DateAnTimeToChange As Integer
Dim PickWeek As Integer



Private Sub Add_Item_Command_Click()

DateAnTimeToChange = 0
NumberToChange = 0
CiEDIT.Edit_item.Visible = 0
CiEDIT.Add_item_command.Visible = 0
CiEDIT.Add_Next_New_Number.Visible = 0

CiEDIT.Dont_Treat.Visible = -1
CiEDIT.The_description_label.Visible = -1
CiEDIT.The_description_Text1.Visible = -1
CiEDIT.The_description_Text1.Top = CiEDIT.The_description_Text.Top
CiEDIT.The_description_Text1.Left = CiEDIT.The_description_Text.Left
CiEDIT.The_description_Text1.Width = CiEDIT.The_description_Text.Width

CiEDIT.Number_label2.Visible = -1
CiEDIT.Number_Text1.Visible = -1
CiEDIT.Number_Text1.Top = CiEDIT.Number_label.Top
CiEDIT.Number_Text1.Left = CiEDIT.Number_label.Left
CiEDIT.Number_Text1.Width = CiEDIT.Number_label.Width



CiEDIT.Date_Text1.Visible = -1
CiEDIT.Date_Label1.Visible = -1
CiEDIT.Time_Text2.Visible = -1
CiEDIT.Time_Label1.Visible = -1
CiEDIT.Save_command.Visible = -1
CiEDIT.Save_command.Enabled = 0
CiEDIT.Formularbar.Caption = "Add item to capture log mode."

Number$ = Mid$(Ciform1.List2.List(0), 1, 16)
CiEDIT.Number_Text1.Text = RTrim(Number$)
CiEDIT.Number_Text1.SelStart = Len(Number_Text1.Text)
Code$ = "xx22"
Call Sort_Ci_Find
Date_An_Time = DateSerial(Val(Mid$(Xdate$, 7, 4)), Val(Mid$(Xdate$, 4, 2)), Val(Mid$(Xdate$, 1, 2)))
Date_An_Time = Date_An_Time + TimeSerial(Val(Mid$(Ttime$, 1, 2)), Val(Mid$(Ttime$, 4, 2)), Val(Mid$(Ttime$, 7, 2)))
Date_An_Time2 = Date_An_Time
CiEDIT.Time_Text2.Text = Ttime$
CiEDIT.Date_Text1.Text = Xdate$
CiEDIT.Number_Text1.SetFocus
CiEDIT.Timer1.Enabled = False
CiEDIT.Timer1.Interval = 500
End Sub



Private Sub Edit_Item_Click()

CiEDIT.BUSY.Visible = -1

Edit_item.Visible = 0
Add_item_command.Visible = 0
Add_Next_New_Number.Visible = 0

The_description_label.Visible = -1
The_description_Text.Visible = -1
New_In.Visible = -1
RecNo_Label2.Visible = -1
Recno_Label.Visible = -1
rec_Up_Command.Visible = -1
Rec_down_Command.Visible = -1
Number_label.Visible = -1
Number_label2.Visible = -1
pick_up_Command.Visible = -1
let_go_command.Visible = -1
pick_up_label.Visible = -1
CiEDIT.Formularbar.Caption = "Edit item mode."
dab = -1



If Which = 0 Then dab = Val(Ci.Rec_no_label2.Caption)
If Which = 1 Then dab = Val(Ciform1.recpostext.Text)
If Which = 2 Then dab = Val(Ciform2.recpostext.Text)
If Which = 3 Then dab = Val(Ciform4.recpostext.Text)

If dab > -1 Then
    If Which = 0 Then
        Number_label.Caption = Mid$(Ci.List1.List(Ci.List1.ListIndex), 19, 16)
        Number$ = Number_label.Caption
        Word$ = UCase$(Mid$(Ci.List1.List(Ci.List1.ListIndex), 36, 4))
        Code$ = ""
        If LTrim(Word$) <> "" Then Code$ = "01"
        If Word$ = "RING" Then Code$ = "02"
    End If

    If Which = 1 Then
        Number_label.Caption = Mid$(Ciform1.List2.List(Ciform1.List2.ListIndex), 1, 16)
        Number$ = Number_label.Caption
        Code$ = Mid$(Ciform1.List2.List(Ciform1.List2.ListIndex), 42, 2)
    End If

    If Which = 2 Then
        Number_label.Caption = Mid$(Ciform2.List2.List(Ciform2.List2.ListIndex), 1, 16)
        Number$ = Number_label.Caption
        Code$ = Mid$(Ciform2.List2.List(Ciform2.List2.ListIndex), 42, 2)
    End If

    If Which = 3 Then
        Number_label.Caption = Mid$(Ciform4.List2.List(Ciform4.List2.ListIndex), 1, 16)
        Number$ = Number_label.Caption
        Code$ = Mid$(Ciform4.List2.List(Ciform4.List2.ListIndex), 42, 2)
    End If

    Call Sort_Ci_Find
End If

If dab = 0 And Which = 0 Then
    Number_label.Caption = Mid$(Ciform2.List2.List(0), 1, 16)
    Number$ = Number_label.Caption
    Code$ = Mid$(Ciform2.List2.List(0), 42, 2)
    Call Sort_Ci_Find
End If

Call Add_Strings

If YuBt < We2 Then pick_up_label.Caption = "Pick up." Else pick_up_label.Caption = "Let Rip.."

CiEDIT.BUSY.Visible = 0

awscount2 = 0

For ri1 = 0 To Ci.List1.ListCount - 1
    aws1$ = Mid$(Ci.List1.List(ri1), 46)
    If InStr(UCase$(aws1$), "NEW IN") = 1 Then
        awscount = Val(Mid$(aws1$, 8, InStr(8, aws1$, " ") - 8))
        If awscount > awscount2 Then awscount2 = awscount
    End If

    aws2$ = Mid$(Ci.List1.List(ri1), 1, 8)
    aws3$ = Mid$(Ci.List1.List(ri1), 19, 16)
    aws4$ = Mid$(Ci.List1.List(ri1), 41, 4)

    If aws3$ = Number$ Then
        If Val(aws4$) = 1 Then
            awsdate$ = aws2$
        End If
    End If
Next

New_In.Text = "New In" + Str(awscount2 + 1) + " Number " + awsdate$

If Mid$(UCase$(The_description_Text.Text), 1, 6) = "NEW IN" Then
    New_In.Text = The_description_Text.Text
End If

End Sub


Private Sub Add_Next_New_Number_click()

CiEDIT.BUSY.Visible = -1
Edit_item.Visible = 0
Add_item_command.Visible = 0
Add_Next_New_Number.Visible = 0
The_description_label.Visible = -1
The_description_Text.Visible = -1
New_In.Visible = -1
RecNo_Label2.Visible = -1
Recno_Label.Visible = -1
rec_Up_Command.Visible = -1
Rec_down_Command.Visible = -1
Number_label.Visible = -1
Number_label2.Visible = -1
pick_up_Command.Visible = -1
let_go_command.Visible = -1
pick_up_label.Visible = -1
CiEDIT.Formularbar.Caption = "Add Next New Number."
dab = -1


For ri1 = 0 To Ci.List1.ListCount - 1
    aws1$ = Mid$(Ci.List1.List(ri1), 46)
    If InStr(UCase$(aws1$), "NEW NUMBER") = 1 Then
        xcde = ri1
    End If

    If InStr(UCase$(aws1$), "Recorded Numbers...OLD") > 0 Then
        Exit For
    End If
Next

Ci.List1.ListIndex = xcde
Ci.List1.Refresh
            
Call Ci.Record_Number_text

Which = 0
If Which = 0 Then dab = Val(Ci.Rec_no_label2.Caption)
If Which = 1 Then dab = Val(Ciform1.recpostext.Text)
If Which = 2 Then dab = Val(Ciform2.recpostext.Text)
If Which = 3 Then dab = Val(Ciform4.recpostext.Text)

If dab > -1 Then
    If Which = 0 Then
        Number_label.Caption = Mid$(Ci.List1.List(Ci.List1.ListIndex), 19, 16)
        Number$ = Number_label.Caption
        Word$ = UCase$(Mid$(Ci.List1.List(Ci.List1.ListIndex), 36, 4))
        Code$ = ""
        If LTrim(Word$) <> "" Then Code$ = "01"
        If Word$ = "RING" Then Code$ = "02"
    End If

    If Which = 1 Then
        Number_label.Caption = Mid$(Ciform1.List2.List(Ciform1.List2.ListIndex), 1, 16)
        Number$ = Number_label.Caption
        Code$ = Mid$(Ciform1.List2.List(Ciform1.List2.ListIndex), 42, 2)
    End If

    If Which = 2 Then
        Number_label.Caption = Mid$(Ciform2.List2.List(Ciform2.List2.ListIndex), 1, 16)
        Number$ = Number_label.Caption
        Code$ = Mid$(Ciform2.List2.List(Ciform2.List2.ListIndex), 42, 2)
    End If

    If Which = 3 Then
        Number_label.Caption = Mid$(Ciform4.List2.List(Ciform4.List2.ListIndex), 1, 16)
        Number$ = Number_label.Caption
        Code$ = Mid$(Ciform4.List2.List(Ciform4.List2.ListIndex), 42, 2)
    End If

    Call Sort_Ci_Find

End If

If dab = 0 And Which = 0 Then
    Number_label.Caption = Mid$(Ciform2.List2.List(0), 1, 16)
    Number$ = Number_label.Caption
    Code$ = Mid$(Ciform2.List2.List(0), 42, 2)
    Call Sort_Ci_Find
End If

Call Add_Strings

If YuBt < We2 Then pick_up_label.Caption = "Pick up." Else pick_up_label.Caption = "Let Rip.."

CiEDIT.BUSY.Visible = 0

awscount2 = 0

For ri1 = 0 To Ci.List1.ListCount - 1
    aws1$ = Mid$(Ci.List1.List(ri1), 46)
    If InStr(UCase$(aws1$), "NEW IN") = 1 Then
        awscount = Val(Mid$(aws1$, 8, InStr(8, aws1$, " ") - 8))
        If awscount > awscount2 Then awscount2 = awscount
    End If

    aws2$ = Mid$(Ci.List1.List(ri1), 1, 8)
    aws3$ = Mid$(Ci.List1.List(ri1), 19, 16)
    aws4$ = Mid$(Ci.List1.List(ri1), 41, 4)
   
    If aws3$ = Number$ Then
        If Val(aws4$) = 1 Then
            awsdate$ = aws2$
        End If
    End If
Next

New_In.Text = "New In" + Str(awscount2 + 1) + " Number " + awsdate$

If Mid$(UCase$(The_description_Text.Text), 1, 6) = "NEW IN" Then
    New_In.Text = The_description_Text.Text
End If

End Sub




Private Sub Add_Strings()
If YuBt > -1 Then
    Drt = 1
    CiEDIT.The_description_Text.Text = Find$
    Drt = 0
    CiEDIT.Recno_Label.Caption = Str(Int(YuBt))
End If

CiEDIT.The_description_Text.SetFocus
'''''''''''next step after edit item click
End Sub

Private Sub Change_Ci_list1()

CiEDIT.BUSY.Visible = -1
For r = 0 To Ci.List1.ListCount
    word2$ = Mid$(Ci.List1.List(r), 36, 4)
    code2$ = "00"
    If LTrim(word2$) <> "" Then code2$ = "01"
    If UCase$(word2$) = "RING" Then code2$ = "02"

    If Mid$(Ci.List1.List(r), 19, 16) = Number$ And Code$ = code2$ Then
        a$ = Ci.List1.List(r)
        Ci.List1.List(r) = Mid$(a$, 1, 35) + Word$ + Mid$(a$, 40)
        Ci.List1.List(r) = Mid$(Ci.List1.List(r), 1, 45) + Find$
    End If
Next
CiEDIT.BUSY.Visible = 0

End Sub

Private Sub Date_Text1_Change()

Date_An_Time2 = DateSerial(Val(Mid$(Date_Text1.Text, 7, 4)), Val(Mid$(Date_Text1.Text, 4, 2)), Val(Mid$(Date_Text1.Text, 1, 2)))
Date_An_Time2 = Date_An_Time2 + TimeSerial(Val(Mid$(Time_Text2.Text, 1, 2)), Val(Mid$(Time_Text2.Text, 4, 2)), Val(Mid$(Time_Text2.Text, 7, 2)))

If Date_An_Time2 > Date_An_Time Then
    DateAnTimeToChange = -1
End If

If Date_An_Time2 <= Date_An_Time Then
    DateAnTimeToChange = 0
End If

If Date_An_Time2 >= Now Then
    DateAnTimeToChange = 0
End If

If Date_An_Time2 < Now And Date_An_Time2 > Date_An_Time Then
    DateAnTimeToChange = -1
End If

If NumberToChange = -1 And DateAnTimeToChange = -1 Then Save_command.Enabled = -1 Else Save_command.Enabled = 0

End Sub

Private Sub Dont_Treat_Click()

a$ = String$(16, " ")
LSet a$ = RTrim(Number_Text1.Text)
Slop$ = RTrim(Date_Text1.Text) + "!" + RTrim(Time_Text2.Text) + "!" + a$ + "!" + RTrim(The_description_Text1.Text)
Save_command.Enabled = 0
Mdate$ = Mid$(Slop$, 1, 10)
Tdate$ = Mid$(Slop$, 1, 6) + Mid$(Slop$, 9, 2)
Ttime$ = Mid$(Slop$, 12, 8)
number2$ = Mid$(Slop$, 21, 16)
Number$ = "                "
LSet Number$ = number2$
Word$ = Mid$(Slop$, 38)

If Word$ = "" Then
    Word$ = "    "
End If

Code$ = "00"

''''''''''' in add item mode puts code 01 if you enter a description
rt$ = Word$
'code$ = "00"

If InStr(Find$, RTrim(Number$)) Then Code$ = "01" Else Code$ = "00"

If InStr(UCase$(Word$), "RINGBACK") Then Code$ = "02"
Counter$ = "0000"
Call FastFind
If Code$ = "01" Then Word$ = Mid$(Find$, 1, 4)
If Code$ = "02" Then Word$ = Mid$(Find$, 1, 4)
If InStr(Find$, RTrim(Number$)) Then Word$ = Mid$(Find$, 1, 4)
Counter$ = Format$(Val(Counter$), "0000")
RSet Counter$ = LTrim(Str(Val(Counter$)))

wd77 = InStr(Find$, "NEW NUMBER")
If wd77 > 1 Then
Find$ = Mid$(Find$, 1, wd77 - 2)
End If
If Foundfind = 0 Then
    Ciform2.List2.AddItem Number$ + "00/00/1900 00:00:00 0000 " + Code$ + " " + Find$, 0
    Eat2 = -1
    Else
    MsgBox ("Already Got that Number...")
End If


End Sub

Private Sub EditExitCommand_Click()


If Eat2 = True Then
    Call Change_Ci_list1
    updater = FreeFile
    Open "W:\CALLERID\CI LOGS\update.txt" For Output As #updater
    Close #updater
    Call Ci.BackUp
    FastFind2String$ = "?"
    For rt = 0 To Ciform2.List2.ListCount
        FastFind2String$ = FastFind2String$ + Format$(rt, "00000") + "***v*" + Ciform2.List2.List(rt)
    Next
    Eat2 = False
    Call Ci.New_Number_Counter
End If
    
Unload CiEDIT

End Sub


Private Sub FormularBar_Click()
CiEDIT.Edit_item.SetFocus
End Sub

Private Sub Let_Go_Command_Click()
If pick_up_label.Caption <> "Let Rip.." Then
    pick_up_label.Caption = "Let Rip.."
    PickWeek = 2
    Call PutDown_line
End If
The_description_Text.SetFocus
End Sub

Private Sub New_In_Click()
If UCase$(Trim(The_description_Text.Text)) = "NEW NUMBER" Then
    The_description_Text.Text = New_In.Text
End If
End Sub

Private Sub Number_Text1_Change()
NumberToChange = -1
Code$ = ""
Number_Text1.Text = UCase$(Number_Text1.Text)
Number_Text1.SelStart = Len(Number_Text1.Text)

If NumberToChange = -1 And DateAnTimeToChange = -1 Then Save_command.Enabled = -1
End Sub

Private Sub Pick_Up_Command_Click()
If pick_up_label.Caption <> "Pick up." Then
    pick_up_label.Caption = "Pick up."
    PickWeek = 1
    Call PutDown_line
End If

The_description_Text.SetFocus
End Sub

Private Sub PutDown_line()
Dim pop As Double

If PickWeek = 2 Then
    aq = 0
    For rt = 0 To Ciform2.List2.ListCount
        If Mid$(Ciform2.List2.List(rt), 1, 16) = "****************" Then aq = 1: Exit For
    Next

    If aq = 1 Then Yup = rt
    If aq = 0 Then Yup = 0.0625

End If

LineSelected = Val(CiEDIT.Recno_Label.Caption)

as2$ = Ciform2.List2.List(LineSelected)
Ciform2.List2.RemoveItem LineSelected
backabitmore2:
pop1 = Yup
If Yup < 1 Then Yup = 1

If PickWeek = 1 Then
    Ciform2.List2.AddItem as2$, (0)
End If

If PickWeek = 2 Then
    Ciform2.List2.AddItem as2$, (Yup)
End If

Yup = pop1


aq = 0

For rt = 1 To Ciform2.List2.ListCount
    If Mid$(Ciform2.List2.List(rt), 1, 16) = "****************" Then aq = 1: Exit For
Next

If aq = 1 Then Yup = rt + 1
If aq = 0 Then Yup = 0.0625
We2 = Yup


B = LineSelected
LineSelected = -1

If PickWeek = 2 Then
    CiEDIT.Recno_Label.Caption = Str(Int(We2)): YuBt = Int(We2)
    Ciform2.List2.ListIndex = Int(We2)
End If

If PickWeek = 1 Then
    CiEDIT.Recno_Label.Caption = Str(Int(0)): YuBt = 0.0625
    Ciform2.List2.ListIndex = 0
End If

Eat2 = True
Call Change_Ci_list1

End Sub

Private Sub Rec_Down_Command_Click()

'Call Change_Ci_list1

Do
    YuBt = YuBt - 1
    If YuBt < 0 Then YuBt = YuBt + 1: Exit Do
    Call Sort_CiEdit_Find
    If InStr(Number$, String$(15, "-")) Or Number$ = String$(16, "*") Or RTrim(Number$) = "" Then
        efit = 1
        Else
        efit = 0
    End If
Loop Until efit = 0

Call Sort_CiEdit_Find

Number_label.Caption = Number$

Call Add_Strings

If YuBt < We2 Then pick_up_label.Caption = "Pick up." Else pick_up_label.Caption = "Let Rip.."

End Sub

Private Sub Rec_Up_Command_Click()

'Call Change_Ci_list1

Do
    YuBt = YuBt + 1
    
    Call Sort_CiEdit_Find

    If InStr(Number$, String$(15, "-")) Or Number$ = String$(16, "*") Or RTrim(Number$) = "" Then
        efit = 1
        Else
        efit = 0
    End If
Loop Until efit = 0

Call Sort_CiEdit_Find

Number_label.Caption = Number$
Call Add_Strings
If YuBt < We2 Then pick_up_label.Caption = "Pick up." Else pick_up_label.Caption = "Let Rip.."
End Sub



Private Sub Save_Command_Click()
Tk = 2
a$ = String$(16, " ")
LSet a$ = RTrim(Number_Text1.Text)
Slop$ = RTrim(Date_Text1.Text) + "!" + RTrim(Time_Text2.Text) + "!" + a$ + "!" + RTrim(The_description_Text1.Text)
Save_command.Enabled = 0
If Code$ = "" Then Number$ = Number_Text1.Text: Call Sort_Ci_Find
Call Ci.AddNumberToCIList1
Gdy = -1
Call Ci.BackUp
Unload CiEDIT
End Sub



Private Sub Speed_Search_Text_GotFocus()

If DoIt2 = 1 Then
    DoIt2 = 0
    recpostext.Visible = 0
    recordpositiontext.Visible = 0
    CiEDIT.List2.Visible = 0
    CiEDIT.findbutton.Visible = 0
    CiEDIT.Edit_item.SetFocus
    CiEDIT.Speed_search_text.Visible = 0
    CiEDIT.Speedsearchtext.Visible = 0
End If

If Which > 0 Then
    recpostext.Visible = 0
    recordpositiontext.Visible = 0
    CiEDIT.List2.Visible = 0
    CiEDIT.findbutton.Visible = 0
    CiEDIT.Speed_search_text.Visible = 0
    CiEDIT.Speedsearchtext.Visible = 0
    Call Edit_Item_Click
End If

End Sub

Private Sub The_Description_Text_Change()

If Drt = 0 Then
    If YuBt < Ciform2.List2.ListCount Then
        a$ = Ciform2.List2.List(YuBt)
    End If
    
    If Len(a$) > 45 Then Mid$(a$, 45) = The_description_Text.Text
    B$ = "    "
    
    Ciform2.List2.List(YuBt) = Mid$(a$, 1, 44) + The_description_Text.Text

    Call Sort_Ci_Find
    Eat2 = True
    Call Change_Ci_list1
End If

End Sub

Public Sub Timer1_Timer()



If Gdy = -1 Then
    NumberToChange = 0
    DateAnTimeToChange = 0
    Number$ = Mid$(Ciform1.List2.List(0), 1, 16)
    Number_Text1.Text = RTrim(Number$)
    Number_Text1.SelStart = Len(Number_Text1.Text)
    
    Call Sort_Ci_Find
    
    Date_An_Time = DateSerial(Val(Mid$(Xdate$, 7, 4)), Val(Mid$(Xdate$, 4, 2)), Val(Mid$(Xdate$, 1, 2)))
    Date_An_Time = Date_An_Time + TimeSerial(Val(Mid$(Ttime$, 1, 2)), Val(Mid$(Ttime$, 4, 2)), Val(Mid$(Ttime$, 7, 2)))
    Date_An_Time2 = Date_An_Time
    Number_Text1.SetFocus
    Gdy = 0
    NumberToChange = 0
    DateAnTimeToChange = 0
End If

If Date_An_Time2 < Now And Date_An_Time2 > Date_An_Time Then
    DateAnTimeToChange = -1
End If

If NumberToChange = -1 And DateAnTimeToChange = -1 Then
    If Save_command.Enabled = 0 Then Save_command.Enabled = -1
End If

If NumberToChange = 0 Or DateAnTimeToChange = 0 Then
    If Save_command.Enabled = -1 Then Save_command.Enabled = 0
End If

End Sub

Private Sub Time_Text2_Change()

Date_An_Time2 = DateSerial(Val(Mid$(Date_Text1.Text, 7, 4)), Val(Mid$(Date_Text1.Text, 4, 2)), Val(Mid$(Date_Text1.Text, 1, 2)))
Date_An_Time2 = Date_An_Time2 + TimeSerial(Val(Mid$(Time_Text2.Text, 1, 2)), Val(Mid$(Time_Text2.Text, 4, 2)), Val(Mid$(Time_Text2.Text, 7, 2)))

If Date_An_Time2 > Date_An_Time Then
    DateAnTimeToChange = -1
End If

If Date_An_Time2 <= Date_An_Time Then
    DateAnTimeToChange = 0
End If

If Date_An_Time2 >= Now Then
    DateAnTimeToChange = 0
End If

If Date_An_Time2 < Now And Date_An_Time2 > Date_An_Time Then
    DateAnTimeToChange = -1
End If

If NumberToChange = -1 And DateAnTimeToChange = -1 Then Save_command.Enabled = -1 Else Save_command.Enabled = 0

End Sub

