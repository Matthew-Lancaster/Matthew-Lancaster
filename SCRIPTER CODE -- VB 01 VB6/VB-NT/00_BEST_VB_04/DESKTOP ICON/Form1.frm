VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Desktop Icon Tools 3.02"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Hide \UnHide TaskBar"
      Height          =   375
      Left            =   360
      TabIndex        =   32
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Show All           Icons"
      Height          =   495
      Left            =   3120
      TabIndex        =   31
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Resolution Mode"
      Height          =   855
      Left            =   4440
      TabIndex        =   28
      Top             =   3000
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "Use QuickRes dll"
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Use CDS_FORCE"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Hide All Icons"
      Height          =   495
      Left            =   3120
      TabIndex        =   27
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   25
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Change Dynamically"
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   4680
      TabIndex        =   23
      Top             =   4920
      Width           =   2175
   End
   Begin VB.PictureBox picHook 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Hide\UnHide  Current Icon"
      Height          =   615
      Left            =   3120
      TabIndex        =   17
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   3360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   3480
      Width           =   540
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Restore "
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save "
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Swap Icons Randomly"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mark Icon Positions"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   2520
      Width           =   5895
   End
   Begin VB.Line Line21 
      X1              =   24
      X2              =   176
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line Line20 
      X1              =   216
      X2              =   272
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00E0E0E0&
      X1              =   472
      X2              =   472
      Y1              =   192
      Y2              =   16
   End
   Begin VB.Line Line18 
      X1              =   8
      X2              =   472
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line17 
      X1              =   8
      X2              =   8
      Y1              =   16
      Y2              =   192
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   8
      X2              =   472
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Label Label13 
      Caption         =   "Current Resolution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Debug and FunStuff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   22
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00E0E0E0&
      X1              =   288
      X2              =   288
      Y1              =   200
      Y2              =   416
   End
   Begin VB.Line Line15 
      X1              =   200
      X2              =   200
      Y1              =   416
      Y2              =   200
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00E0E0E0&
      X1              =   200
      X2              =   288
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line13 
      X1              =   200
      X2              =   288
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00E0E0E0&
      X1              =   8
      X2              =   192
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line11 
      X1              =   8
      X2              =   8
      Y1              =   200
      Y2              =   416
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00E0E0E0&
      X1              =   192
      X2              =   192
      Y1              =   200
      Y2              =   416
   End
   Begin VB.Line Line9 
      X1              =   8
      X2              =   192
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00E0E0E0&
      X1              =   472
      X2              =   472
      Y1              =   336
      Y2              =   416
   End
   Begin VB.Line Line7 
      X1              =   296
      X2              =   296
      Y1              =   336
      Y2              =   416
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00E0E0E0&
      X1              =   296
      X2              =   472
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line5 
      X1              =   24
      X2              =   176
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   472
      X2              =   472
      Y1              =   336
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   296
      X2              =   296
      Y1              =   336
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   296
      X2              =   472
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Label Label11 
      Caption         =   "Change Screen Resolution   And Keep Icon Pattern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   20
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Current Icon Positions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Hidden?"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Programmers Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Current Icon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   " ImageList   Icon No. "
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Location"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Icon Text"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "  Icon Number"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopupMenu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "DeskTop Icon Tools"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please read this source code before running in the IDE

'DESKTOP ICON TOOLS minimizes to the tray after 1 second

'It is designed to be run from the Startup Folder
'On Startup ,
'The program stores all of the current icon positions.

'///////////////////////////////////////////////
'If the program detects a WM_WININICHANGE event.
'Which is triggered by the taskbar being dragged
'to the bottom or being reinstated.
'Windows would detect our OFF-SCREEN ICONS.
'And then put them back on the Desktop
'Bugger that.
'We then replace the Icons to the position
'we detected at StartUp.
'THIS BEHAVIOUR IS NOT PRESENT WHEN RUNNING IN THE VB IDE
'THIS IS BY DESIGN

'Please read the warnings below
'///////////////////////////////////////////////
'This code has been written to detect the IDE
'and then NOT do the SUBCLASSING.
'*****************************************
'IF YOU CHANGE THE SUBCLASSING BEHAVIOUR
'THIS CODE WILL CRASH THE VISUAL BASIC IDE
'WHEN YOU EXIT THE PROGRAM.
'PLEASE READ THE APP NOTES CONTAINED IN THE SOURCE CODE.
'*****************************************

'    *** Warning - if you change the subclassing code ***
'Subclassing is dangerous. The debugger does not work well when a new
'WindowProc is installed. If you halt the program instead of unloading its
'forms, it will crash and so will the Visual Basic IDE. Save your work often!
'Don't say you weren't warned.

'///////////////////////////////////////////////


'    Disclaimer
'This example program is provided "as is" with no warranty of any kind. It is
'intended for demonstration purposes only.

'Window subclassing code courtesy of www.vb-helper.com
'added this line in the Unload event
'Call SetWindowLong(Me.hwnd, GWL_WNDPROC, OldWindowProc)

'|||||||||||||||||||||||||||||||||||||||||||
'App notes spread through code and in modules
'use as you see fit
'tested on Win 95
'will almost certainly break on NT
'testing on NT soon.
'Paul Pavlic 3 August 1999
'I assert that I wrote this rubbish
':::::::::::::::::::::::::::

Private Sub Command1_Click()
WroteToDeskTop& = 1
MarkIcons

End Sub

Private Sub Command10_Click() ' hide\unhide current icon

'all we do is move the icon off-screen.
'this however does make us vulnerable to windows
'detecting this and replacing our off-screen icons
'back onto the Desktop
'Which is why we subclass our window to get the
'WM_WININICHANGE message
'This is posted by the user pulling down or up
'on the taskbar.
'Further testing required to look for other holes.
'We could perhaps refresh whenever the Desktop is refreshed?
'TODO:
'Is possible to detect desktop events
'Subclass and catch namespace changes

'HIDING ICONS this way allows us to easily hide icons like
'My Computer, Inbox etc.

'TRADEOFFS:
'Currently , when Windows detects the Off-Screen Icons
'It tries to put them on the desktop
'We pick up on this from the above event
'And immediately move them back off.
'On my P150
'The icons are painted black on the screen and then
'immediately disappear.
'I think we can live with this.
'Possibly we may find a way to find all the ways that
'Windows will do its Desktop Icon refresh and
'catch it first <difficult>
'Then hiding the desktop window <easy>
'Then redisplaying
'Hmmmm.....
'Back to the acting...
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
If List1.ListIndex > -1 Then

    If Hidden(List1.ListIndex) = "Y" Then
    'check if at startup this icon was visible on the screen
        If IconPosition(List1.ListIndex).x > StartWidth Or IconPosition(List1.ListIndex).y > StartHeight Then
            'nup put icon back on screen at 30,30
            Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, List1.ListIndex, _
            CLng(30 + 30 * &H10000))
    
    
    
        Else 'yup restore
            Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, List1.ListIndex, _
            CLng(IconPosition(List1.ListIndex).x + IconPosition(List1.ListIndex).y * &H10000))
        
        End If
    
        Text1.Text = "Icon Number " + Str(List1.ListIndex + 1) + "  is now Un_hidden"
        Hidden(List1.ListIndex) = "N"
    Else
        'move the icon off the screen
        Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, List1.ListIndex, _
        CLng((StartWidth + 50) + (StartHeight + 50) * &H10000))
        '                   ^
        '                   50 more than width of screen
        'There is a limit to how far Win95 will
        'allow you to move offscreen - +500 Doesn't work
    
        Text1.Text = "Icon Number " + Str(List1.ListIndex + 1) + "  is now hidden"
        Hidden(List1.ListIndex) = "Y" 'had a reason for this
    End If
FindIcons 'reload
Else
End If
End Sub

Private Sub LoadDisplaySettings()
'these should stay viable throughout the life of the program
 Dim a As Boolean
Dim i&
Dim n&
i = 0

Do While EnumDisplaySettings(0&, i&, DevM) > 0
     i = i + 1
Loop
'the value of i here is actually MORE
'than we need  but what the hey
ReDim ScrMode(i) As ScreenDisplayModes

i = 0
n = 0
'load our DevM structure
Do While EnumDisplaySettings(0&, i&, DevM) > 0
        
        If DevM.dmBitsPerPel >= 8 Then
          List2.AddItem Str(DevM.dmPelsWidth) + " *" + Str(DevM.dmPelsHeight) + "   " + Str(DevM.dmBitsPerPel) + " bit"
          ScrMode(n).bits = DevM.dmBitsPerPel
          ScrMode(n).nwidth = DevM.dmPelsWidth
          ScrMode(n).nheight = DevM.dmPelsHeight
          n = n + 1
        Else
     
        End If
   i = i + 1
Loop
End Sub
Private Sub ChangePixelDepthByScrRes(sss As String)
Dim RetVal
Dim shellString As String
'if you get error codes here then
'possibly quickres has been uninstalled
'supposed to work on osr2         win98?
'example given only for educational purposes
shellString = "Rundll.exe DeskCp16.dll,QUICKRES_RUNDLLENTRY "
shellString = shellString + sss
RetVal = Shell(shellString, 1)

End Sub
Private Sub Command12_Click() 'change resolution dynamically

If List2.ListIndex > -1 Then
    'Alright do it
    Dim i As Long
    WeDidIt = 1
    i = List2.ListIndex
 
    
    If Option1.Value = True Then 'use CDS_FORCE
        'call sub
        ChangePixelDepth ScrMode(i).bits, ScrMode(i).nwidth, ScrMode(i).nheight ' change pixel depth
        RefreshDesktopFully
    Else ' do it using Quickres dll
        'convert params to a string
        
        Dim tempvar As String
        tempvar = LTrim$(Str(ScrMode(i).nwidth)) + "x" + LTrim$(Str(ScrMode(i).nheight)) + "x" + LTrim$(Str(ScrMode(i).bits))
        'call sub
        ChangePixelDepthByScrRes tempvar
        PlaceIconsInPosition ScrMode(i).bits, ScrMode(i).nwidth, ScrMode(i).nheight ' restore icons to new positions
        RefreshDesktopFully
    End If

Else
End If

End Sub

Private Sub Command2_Click()
'refresh for Mark Icons
WroteToDeskTop& = 0
RefreshDesktop

End Sub

Private Sub Command3_Click()
'Swap icons randomly
MovedIcon& = 1
SwapIcon

End Sub

Private Sub Command4_Click() 'I am oo lazy to rename
'refresh from above
MovedIcon& = 0
RefreshIconPositions

End Sub

Private Sub Command5_Click()
'::::::::::::::::::::::::::: Save Icon Positions Sub
Form2.Show
Form1.Hide
End Sub

Private Sub Command6_Click()
'::::::::::::::::::::::::::: restore saved icons Sub
Form3.Show
End Sub
Public Sub ChangePixelDepth(pixeldepth2 As Integer, nwidth2 As Long, nheight2 As Long) ' change pixel depth
'********************************************
'ABSOLUTELY FORM1'S AUTOREDRAW MUST BE FALSE  - bugfix
'==========         ========== ====    =====
'********************************************
'********* WARNING ********
'Some of this code is not documented.
'It works fine in my machine , BUT I cannot guarantee
'that it will work well for you.
'I have tried to catch all the bugs before posting.
'The rest is up to you.
'UNSUPPORTED AND UNDOCUMENTED CODE
'**************************
'Take all visible objects that we have and hide them.
'Check the Subclass events in the NewWindowProc function
'Please read the code before changing anything.
'Comments abound!
':::::
FindIcons   'this is easier than writng a new routine
            'and my wife will tell you I am a lazy bugger
            'reload all vars and update current
            'icon positions

Picture1.Picture = LoadPicture() 'VERY IMPORTANT this line
'Resolves dibeng.dll blowing up
'also was painting icons on screen at resolution changeover
'4 Hrs of corrupted screen real estate later - oh well.
'what else would you rather be doing?


' all the hiding may not be necessary
'everything depends on the hardware acceleration set
'in the Control Panel - System - Graphics
'Set to low for better results
'this is NOT a great implementation
'needs lots of work really

Form1.AutoRedraw = False
'set it false here so that we KNOW that it is false
Text1.Text = "Changing mode to - " + Str(nwidth2) + " *" + Str(nheight2) + "   " + Str(pixeldepth2) + " bit"
Form1.Hide

'[[[[[[[[[[ here is the guts of it ]]]]]]]]]
Dim b&

'load DevM structure with what we want to do
    
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFLAGS
    DevM.dmPelsWidth = nwidth2
    DevM.dmPelsHeight = nheight2
    DevM.dmBitsPerPel = pixeldepth2
    DevM.dmSize = Len(DevM)

'/////////////////// IMPORTANT /////////////////////////
'****  EVERY ONE OF OUR OBJECTS IS NOW NOT VISIBLE  ****
'/////////////////// IMPORTANT /////////////////////////


  b = ChangeDisplaySettings(DevM, CDS_FORCE)
 
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'CDS_FORCE - this is NOT in the literature :-)
    'CDS_FORCE  = &H80000000  don't ask
    'A bloody great big huge HACK
    'Can't credit it, got it from a NewsGroup search
    'and no Name <sigh>  thanks mate
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                     
xxc = SendMessageByLong(HWND_BROADCAST, WM_SYSCOLORCHANGE, 0&, 0&)
     'tell ALL other windows what we are doing


'jump through some hoops
'better be nice and tell every window that we changed resolution

xxxd = SendMessageByLong(HWND_BROADCAST, WM_PALETTECHANGED, Me.hwnd, 0&)
'>>        >>           >>         >>                       ^ I did it
'we don't allow our main window to process this message when it returns
'Read documentation for WM_PALETTECHANGED for why

xxs = PostMessage(HWND_BROADCAST, WM_SYSCOLORCHANGE, 0&, 0&)


RefreshDesktopFully ' this time we want to erase the background
'respot our icons - checking for change in screen resolution size
PlaceIconsInPosition pixeldepth2, nwidth2, nheight2    ' restore icons to new positions
Form1.Show
End Sub
Private Sub Command7_Click() 'hide all the damn icons
Dim i As Long
For i = 0 To icount - 1

Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, i, _
    CLng((StartWidth + 50) + (StartHeight + 50) * &H10000))
    
    Hidden(i) = "Y"

Next i

Text1.Text = "ALL Icons are now hidden"
'well they should be
'TODO:
'Error trapping

'reload our position data
Picture1.Picture = LoadPicture()
FindIcons


End Sub
Private Sub CheckForSaved()

sss = Dir$(DirectoryName + "*.pps")
If sss = "" Then
'there are no save files  inform the user
MsgBox "There are no Icon configurations saved" + Chr$(13) + Chr$(10) + "Please choose Save under Current Icon Positions." + Chr$(13) + Chr$(10) + "This message will be shown at StartUp until there are saved configurations." + Chr$(13) + Chr$(10) + "You should have at least ONE configuration saved with all your Icons visible.", vbOKOnly + vbInformation


Else
End If
End Sub

Private Sub Command8_Click()
'put all the icons back on the desktop sub

PutAllIconsOnDeskTop
'reload our position data
FindIcons
Text1.Text = "ALL Icons are now visible"
'well they should be
'TODO:
'Error trapping
Picture1.Picture = LoadPicture()
End Sub

Private Sub Command9_Click()
'Hide \UnHide TaskBar
TaskBarHideUnhide
End Sub

Private Sub Form_Load()
'Ver 1.00 Created 11 April 1999
'locate Desktop icons

'Ver 2.00 Updated 26 June 1999
'major change - retrieve text


'Ver 3.00 Updated 2 August 1999
'programmers version
'major change - retrieve icon image
'added change resolution dynamically
'and restore icon pattern to changed resolution
'HIDE icons and restore (Suggested by Michael Hennessy)
'and a whole bunch of other stuff
':::::::::::::::::::::::::::::::::::



'Ver 4.00 will be a users version
'with reduced functionality.
'TODO:
'Passwords


'{{{{{{{{{{{{{{}}}}}}}}}}}}}}}}}}
'Show the little splash screen first
'somehow this seems to beat a little bug with the
'Show Desktop Button on the taskbar with IE 4.1 desktop update
'If this button is pressed just prior to this code loading
'we were getting a zero count back from LVM_GETTITEMCOUNT
'Doesn't seem to crop up if at least 1 window has painted
'and left the screen
'Small app notes in Module1.bas
'For best viewing results
'Compile and run from desktop shortcut for best use
'Nuff said
'::::::::::::::::::::::::::::::


'Stick our icon in the tray
Dim nd As NOTIFYICONDATA
  Dim lRet As Long
  With nd
    .cbSize = Len(nd)
    .hwnd = picHook.hwnd
    .uID = 1&
    .szTip = "DeskTop Icon Tools" & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
  End With
  lRet = Shell_NotifyIconA(NIM_ADD, nd)
'We may not put an icon on the Shell Tray if
'we want to run as a service
'Especially if this app changes
'into a full blown security app.

' Change this  if you change the .exe name
If App.EXEName = "DESK" Then ' are we in the compiled program?
'yep we are running desk.exe
InIde = 1
'*******************************
'WE ABSOLUTELY DON'T SUBCLASS IF WE ARE IN THE IDE
'Read the Warnings if you want to fiddle.
'Please!!
'*******************************
'Subclass our Form
OldWindowProc = SetWindowLong( _
        Me.hwnd, GWL_WNDPROC, _
        AddressOf NewWindowProc)
Else
'In the IDE
'Don't do any subclassing
'we won't detect some of the messages we need
'we don't have full functionality
InIde = 0
End If


Load frmSplash
frmSplash.Show
Me.Hide

'set some vars
MovedIcon& = 0
WroteToDeskTop& = 0
WeDidIt = 0

Pause 500
FindIcons 'where are the icons now ?
Unload frmSplash
Me.Show
'Me.gomakecoffee
If List1.ListCount > 0 Then
    List1.ListIndex = 0
    List1.SetFocus
    'draw the associated icon with the first icon name
    'in the listview
    'put it in our PictureBox
    f = ImageList_Draw(hImageList, IconNumber(0), Form1.Picture1.hdc, 0, 0, 0)
Else
    'shouldn't hit this, should have bailed before this
    Unload Form1
    End
End If


started = 1 ' we have started - need this later
LoadDisplaySettings

Text2.Text = Str(StartWidth) + " *" + Str(StartHeight) + "  " + Str(GetDeviceCaps(Me.hdc, BITSPIXEL)) + " bit"

CheckForSaved
Me.Refresh
Pause 1000




If InIde = 1 Then 'not in Ide

'*******  set our HOTKEY ******
'Ctrl-Shift-D
'*******  set our HOTKEY ******

'brings up our application regardless of
'whether we have the focus
'THIS ONLY WORKS WHEN NOT RUNNING IN IDE
'Can't receive our messages if we are in Ide
ssd& = RegisterHotKey(Me.hwnd, 1&, MOD_CONTROL Or MOD_SHIFT, &H44)
'see Case WM_HOTKEY in second module for entry point
'::::::::::::::::::::::::::::::
Form1.WindowState = 1


Else 'we are in the VB editor environment

Form1.WindowState = 0
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'IMPORTANT:
'Put all the icons back in place if we moved them
If icount > 0 And MovedIcon& = 1 Then RefreshIconPositions
'TODO:
'Have we mucked with the screen resolution ?


'Better cleanup if we marked the screen <g>
If WroteToDeskTop& = 1 Then RefreshDesktop

'TODO:
'Make sure that we are in the same screen resolution
'as when we started.
'At least ask user for input.
'I.E.
'Do you wish to restore resolution ?

'Unhook from the tray and destroy icon
Dim nd As NOTIFYICONDATA
  Dim iRet As Integer
  With nd
    .cbSize = Len(nd)
    .hwnd = picHook.hwnd
    .uID = 1&
  End With
  iRet = Shell_NotifyIconA(NIM_DELETE, nd)


If App.EXEName = "DESK" Then ' are we in the compiled program?

    'WARNING: This line is absolutely necessary if we have
    '         subclassed our Form
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, OldWindowProc)
    ssd& = UnregisterHotKey(Me.hwnd, 1)
Else 'we are in the VB Ide

End If

'make sure that the taskbar is visible on prog exit
lTaskBarHWND = FindWindow("Shell_TrayWnd", "")
lRet = ShowWindow(lTaskBarHWND, SW_SHOW)



Unload Form1 'ok , Bail
End

End Sub

Private Sub Form_Resize()

If started = 1 Then
    If Form1.WindowState = 1 And Form1.Visible = True Then
    
    Else
       
       If List1.ListCount > 0 Then
        List1.ListIndex = 0 'reset Listbox to first item
       Else
       End If
       
'+++  HAD to remove this
       ' Picture1.AutoRedraw = False 'needed to make this work
       'repaint the icon image
       ' f = ImageList_Draw(hImageList, IconNumber(List1.ListIndex), Form1.Picture1.hdc, 0, 0, 0)
       ' Picture1.AutoRedraw = True  'umm...
       'this is necessary because our picturebox will
       'not redraw
'+++  HAD to remove this
  
    End If
Else
End If

If Form1.WindowState = 1 Then 'did we minimize?

Form1.Hide
Form1.WindowState = 0
'this means that we are at normal size
'but hidden
'so if some other prog minimizes ALL windows
'we will disappear
'instead of leaving an ugly minimized form
'on the desktop
Else
End If




End Sub

Private Sub List1_Click()
' ok we have selected an icon in our listbox
' show it to us

Picture1.AutoRedraw = False 'don't ask me why
f = ImageList_Draw(hImageList, IconNumber(List1.ListIndex), Form1.Picture1.hdc, 0, 0, 0)
'places a copy of the associated image into our picturebox
'TODO:
'check return value
Picture1.AutoRedraw = True  'it works this way
                            'try changing it
End Sub

Private Sub mnuExit_Click()
Unload Form1
End
End Sub

Private Sub mnuOpen_Click()
Form1.WindowState = 0 'just to be sure
Form1.Show
      
End Sub

Private Sub picHook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 
'this is where we detect if our icon
'has been clicked on in the tray

 Static bRunning As Boolean
  Dim lMsg As Long
  lMsg = x / Screen.TwipsPerPixelX
  
     
  If Not (bRunning) Then 'avoid cascades
    bRunning = True
    Select Case lMsg
      
      Case WM_LBUTTONDBLCLK: 'double clicked
        Form1.WindowState = 0
        Form1.Show
      
      Case WM_LBUTTONDOWN:
      Case WM_LBUTTONUP:
      Case WM_RBUTTONDBLCLK:
      Case WM_RBUTTONDOWN:
      
      Case WM_RBUTTONUP: 'right click
        Me.PopupMenu mnuPop
        End Select
        bRunning = False
    End If
End Sub

