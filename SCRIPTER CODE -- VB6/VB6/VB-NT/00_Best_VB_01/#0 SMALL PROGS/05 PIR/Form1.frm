VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Pir Rs232 Activity Door Bell The Works.."
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   8955
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   2175
      Top             =   300
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   1800
      Top             =   2400
   End
   Begin MCI.MMControl MMControl1 
      Height          =   540
      Left            =   3450
      TabIndex        =   5
      Top             =   1125
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   953
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1425
      Top             =   225
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   825
      Top             =   150
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   375
      Top             =   3075
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RTSEnable       =   -1  'True
      InputMode       =   1
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   8940
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   0
      TabIndex        =   3
      Top             =   2025
      Width           =   8940
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   0
      TabIndex        =   4
      Top             =   3300
      Width           =   8940
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   7575
      TabIndex        =   8
      Top             =   0
      Width           =   690
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   6825
      TabIndex        =   7
      Top             =   0
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   6075
      TabIndex        =   6
      Top             =   0
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   2925
      TabIndex        =   0
      Top             =   0
      Width           =   3090
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2865
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
   End
   Begin VB.Menu Mnu_LogDIR 
      Caption         =   "Log Dir"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Labbc1 = &HA0F0A0
Const LabFC1 = &H0
Const Labbc2 = &H404040
Const LabFC2 = &HFFFFFF
Public Timer2CNT, Timer3CNT

'MicroSoft Speech Object Libary
'Lookup Referances
Public Voice As SpVoice

Public vc1, vc2, Pp1, PpT1, Pp1Timer, Pp2, Pp1Age, Pp1Old
Public uu$, FormLoading
Public ct1, ct2, ct3
Public fr1, fr2, fr3
Public DayTime, DayCount1, DayCount2


Private Sub form_load()
FormLoading = True




uu$ = "C:\callerid\#00 PIR Rs232 Logging\"


Set Voice = New SpVoice
Set Voice.Voice = Voice.GetVoices().Item(3) 'microsoft mike

'If speak$ = "" Then speak$ = "Warriarr"
Speak$ = "CHaahhrs"


On Local Error Resume Next
Voice.Speak Speak$, SVSFlagsAsync


Form1.WindowState = 0
vc1 = 2
vc2 = 2

MSComm1.CommPort = 14
MSComm1.Settings = "1200,N,8,1"
MSComm1.PortOpen = True
'MSComm1.InputLen = 1
MSComm1.DTREnable = True
MSComm1.Handshaking = comRTS
MSComm1.RTSEnable = True

On Error Resume Next
c1 = FreeFile
Open uu$ + "00 DayTime" For Input As #c1
Line Input #c1, cw1$
Line Input #c1, cw2$
Line Input #c1, cw3$
Close #c1

DayTime = Val(cw1$)
DayCount1 = Val(cw2$)
DayCount2 = Val(cw3$)

If DayTime <> Day(Now) Then
DayTime = 0
c1 = FreeFile
Open uu$ + "00 DayTime" For Input As #c1
Print #c1, cw1$
Print #c1, cw2$
Print #c1, cw3$
Close #c1
End If

c1 = FreeFile
Open uu$ + "31 Bell Count" For Input As #c1
c2 = FreeFile
Open uu$ + "32 Door Count" For Input As #c2
c3 = FreeFile
Open uu$ + "33 Pir Count" For Input As #c3
Line Input #c1, cw1$
Line Input #c2, cw2$
Line Input #c3, cw3$
Close #c1, #c2, #c3

ct1 = Val(cw1$)
ct2 = Val(cw2$)
ct3 = Val(cw3$)
On Error GoTo 0

'MSComm1.
oo = MSComm1.PortOpen
If oo = False Then MsgBox "False RS232"

MSComm1.RTSEnable = True
MSComm1.DTREnable = True


tt1$ = "Bell Trigged - Start Up" + Str(ct1) + " Count": Pp1Age = 1
List1.AddItem Format$(Now, "DDD DD-MM-YYYY HH:MM:SSa/p") + " --- Bell Trig -- StartUp At" + Str(ct1) + " Count"
List1.AddItem String$(50, "-")
tt1$ = "Bell Trigged ---" + Str(ct1) + " Count"
Label1.Caption = tt1$

    tt2$ = "Pir Door" + Str(ct2) + " Count"
    Label2.Caption = tt2$
    
        List2.AddItem Format$(Now, "DDD DD-MM-YYYY HH:MM:SSa/p") + " --- " + "PIr Door Trig -- StartUp" + Str(ct2) + " Count"
        List2.AddItem String$(50, "-")




Pp1 = 0
Call MSComm1_OnComm
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False
'Pp1 = 0
End Sub

Private Sub Mnu_LogDIR_Click()
Shell "Explorer " + Mid$(uu$, 1, Len(uu$) - 1)
End Sub

Private Sub MSComm1_OnComm()
'Call Timer1_Timer
Label3.Caption = MSComm1.PortOpen



Pp1 = MSComm1.DSRHolding ' pin 6 lookin in d connector
Pp2 = MSComm1.CDHolding 'pin '1

Label4 = Pp2
Label5 = Pp1

If Pp1 = False And vc2 = Pp2 Then Exit Sub

'Reset
c1 = FreeFile
Open uu$ + "11 Bell Append" For Output As #c1
c2 = FreeFile
Open uu$ + "12 Door Append" For Output As #c2
c3 = FreeFile
Open uu$ + "13 Pir Append" For Output As #c3
    





Timer2.Enabled = True
Timer2CNT = 0

Form1.WindowState = 0


xx = 0
If Pp1Timer < Now And Pp1 = True Then xx = 1

If xx = 1 Then

Pp1Timer = Now + TimeSerial(0, 0, 4)
ct1 = ct1 + 1
DayCount1 = DayCount1 + 1
Speak$ = "Bell "
Voice.Speak Speak$, SVSFlagsAsync
Speak$ = Str(ct1)
Voice.Speak Speak$, SVSFlagsAsync
Speak$ = Str(DayCount1) + " Today"
Voice.Speak Speak$, SVSFlagsAsync


If Pp1 = True Then
    Label1.BackColor = Labbc1
    Label1.ForeColor = LabFC1
Else
    Label1.BackColor = Labbc2
    Label1.ForeColor = LabFC2
End If



    
 
    
    
    
tt1$ = "Bell Trigged ---" + Str(ct1) + " Count" + Str(DayCount1) + " Today"
Label1.Caption = tt1$
    
    yy$ = Format$(Now, "DDD DD-MM-YYYY HH:MM:SSa/p") + " --- " + tt1$
    List1.AddItem yy$
    Print #c1, yy$

'vc1 = Pp1

End If






If vc2 <> Pp2 Then

If Pp2 = True Then
    ct2 = ct2 + 1
    DayCount2 = DayCount2 + 1
End If

If FormLoading = False Then
If Pp2 = True Then
Speak$ = "Door "
Voice.Speak Speak$, SVSFlagsAsync
Speak$ = Str(ct2)
Voice.Speak Speak$, SVSFlagsAsync
Speak$ = Str(DayCount2) + " Today"
Voice.Speak Speak$, SVSFlagsAsync
Else
'Speak$ = "Ef Off"
Speak$ = "CHaahhrs"
Voice.Speak Speak$, SVSFlagsAsync
End If
End If

    If Pp2 = True Then
        tt2$ = "Pir Door On-" + Str(ct2) + " Count" + Str(DayCount2) + " Today"
        Label2.BackColor = Labbc1
        Label2.ForeColor = LabFC1
    Else
        tt2$ = "Pir Door Off" + Str(ct2) + " Count" + Str(DayCount2) + " Today"
        Label2.BackColor = Labbc2
        Label2.ForeColor = LabFC2
    End If







    Label2.Caption = tt2$

    'Speak$ = Str(ct2)
    'Voice.Speak Speak$, SVSFlagsAsync
    
    
    yy$ = Format$(Now, "DDD DD-MM-YYYY HH:MM:SSa/p") + " --- " + tt2$
    List2.AddItem yy$
    Print #c2, yy$
    vc2 = Pp2
End If

Close #c1, #c2, #c3





c1 = FreeFile
Open uu$ + "11 Bell List " For Output As #c1
c2 = FreeFile
Open uu$ + "12 Door List" For Output As #c2
c3 = FreeFile
Open uu$ + "13 Pir List" For Output As #c3

For r = 1 To List1.ListCount - 1
    
    Print #c1, List1.List(r)
Next
For r = 1 To List2.ListCount - 1
    
    Print #c2, List2.List(r)
Next
For r = 1 To List3.ListCount - 1
    
    Print #c3, List3.List(r)
Next
    
List1.ListIndex = List1.ListCount - 1
List2.ListIndex = List2.ListCount - 1
List3.ListIndex = List3.ListCount - 1
    
    
Close #c1, #c2, #c3



c1 = FreeFile
Open uu$ + "31 Bell Count" For Output As #c1
c2 = FreeFile
Open uu$ + "32 Door Count" For Output As #c2
c3 = FreeFile
Open uu$ + "33 Pir Count" For Output As #c3
Print #c1, Str(ct1)
Print #c2, Str(ct2)
Print #c3, Str(ct3)
Close #c1, #c2, #c3

c1 = FreeFile
Open uu$ + "00 DayTime" For Output As #c1
Print #c1, Str(Day(Now))
Print #c1, Str(DayCount1)
Print #c1, Str(DayCount2)
Close #c1

FormLoading = False
End Sub

Private Sub Timer1_Timer()

If Day(Now) <> DayTime Then
DayTime = Day(Now)
c1 = FreeFile
Open uu$ + "00 DayTime" For Output As #c1
Print #c1, Str(Day(Now))
Print #c1, Str(DayCount1)
Print #c1, Str(DayCount2)
Close #c1
DayCount1 = 0
DayCount2 = 0
End If


If Pp1Timer > Now Then Exit Sub

If Pp1Timer < Now Then
Pp1 = 0
End If

If Pp1 = True Then
    Label1.BackColor = Labbc1
    Label1.ForeColor = LabFC1
Else
    Label1.BackColor = Labbc2
    Label1.ForeColor = LabFC2
End If


End Sub

Private Sub Timer2_Timer()




Timer2CNT = Timer2CNT + 1
If Timer2CNT >= 120 Then
    Timer2CNT = 0
'    Form1.WindowState = 1
    Timer2.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()


Timer3CNT = Timer3CNT + 1
If Timer3CNT >= 60 And 1 = 2 Then
    Timer3CNT = 0
 
MSComm1.PortOpen = False

DoEvents
MSComm1.PortOpen = True
MSComm1.DTREnable = True
MSComm1.Handshaking = comRTS
MSComm1.RTSEnable = True
 
 
 End If

End Sub

Private Sub Timer4_Timer()


Pp1 = MSComm1.DSRHolding ' pin 6 lookin in d connector
Pp2 = MSComm1.CDHolding 'pin '1

PpT3 = "False"
If MSComm1.PortOpen Then PpT3 = "True"

PpT1 = "False"
If Pp1 = True Then PpT1 = "True"
Ppt2 = "False"
If Pp2 = True Then Ppt2 = "True"

Label3 = PpT3 + vbCrLf + "Port Open"
Label4 = PpT1 + vbCrLf + "Door"
Label5 = Ppt2 + vbCrLf + "Bell"


End Sub
