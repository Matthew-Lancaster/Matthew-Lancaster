VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   13785
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13755
   _ExtentX        =   24262
   _ExtentY        =   24315
   _Version        =   393216
   Description     =   "Dockable Find tool"
   DisplayName     =   "ExFind_D"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'The sections of this code to do with docking is
'based on code from Jason A. Fisher fisher@scapromo.com
'used in his CodeBrowser program
'I used this to avoid learning UserDocuments from the bottom up.
'Events handlers
Private WithEvents mobjCBEvts      As CommandBarEvents
Attribute mobjCBEvts.VB_VarHelpID = -1
'(I don't use these but have left them in for later experimental purposes
'''Private WithEvents mobjPrjEvts      As VBProjectsEvents
'''Private WithEvents mobjCmpEvts      As VBComponentsEvents
'''Private WithEvents mobjFCEvts       As FileControlEvents
'Module-level extensibility objects
'see notes below for what this is about
Private mWindow                    As Window
'you need a reference to the Microsoft Office Object Library for next line to work (or any add-in in fact)
'(I use ver 8.0 from office 97 but there are newer ones)
Private mobjMCBCtl                 As CommandBarControl
'Dockable add-in needs a GUID
'You MUST generate your own values for the guid string constant other wise you may clash with other tools
'Using a tool called Guidgen.exe, which is located in the \tools\idgen directory of Visual Basic.
'not installed automatically so time to dig out the disk.
Private Const mstrGuid             As String = "{703F17CC-ED79-464f-80FC-B082A5CDF2FC}"

Private Sub AddinInstance_Initialize()

  '<STUB> Reason:' Add_ins need this even if empty


End Sub

Private Sub AddinInstance_OnAddinsUpdate(custom() As Variant)

  '<STUB> Reason:' Add_ins need this even if empty
  '


End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)

  '<STUB> Reason:' Add_ins need this even if empty


End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, _
                                       ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
                                       ByVal AddInInst As Object, _
                                       custom() As Variant)

  'This event fires when the user loads the add-in from the Add-In Manager
  
  Dim Tmpstring1          As String
  Dim PreserveClipGraphic As Variant

  On Error GoTo ErrorHandler
  'Retain handle to the current instance of Visual Basic for later use
  Set VBInstance = Application
  'Hook project and components events
  '''  With VBInstance
  '''    Set mobjPrjEvts = .Events.VBProjectsEvents
  '''    Set mobjCmpEvts = .Events.VBComponentsEvents(Nothing)
  '''    Set mobjFCEvts = .Events.FileControlEvents(Nothing)
  '''  End With
  'Add a toolbar button to the end of the Standard toolbar
  With VBInstance.CommandBars("Add-Ins").Controls
    Set mobjMCBCtl = .Add(1, , , .Count + 1)
  End With
  'Hook event handler to receive the button's events
  Set mobjCBEvts = VBInstance.Events.CommandBarEvents(mobjMCBCtl)
  'Define how the button looks
  With mobjMCBCtl
    'We want a separator before the button
    '.BeginGroup = True 'I don't like to do this but uncomment it if you like
    'Set the caption for the button
    .Caption = AppDetails
    'This is a safer way of adding icons to menus as it
    'preserves the clipboards contents
    'Most Add-ins that use PasteFace are a pain because they
    'make copy and paste difficult/impossible.
    'NOTE if some other add-in doesn't use this technique
    'then you still lose your clipboard contents
    With Clipboard
      'set menu picture
      Tmpstring1 = .GetText
      PreserveClipGraphic = .GetData
      .SetData frmHelp.Picture1.Image
      mobjMCBCtl.PasteFace
      .Clear
      If IsObject(PreserveClipGraphic) Then
        .SetData PreserveClipGraphic
      End If
      If LenB(Tmpstring1) Then
        .SetText Tmpstring1
      End If
    End With
  End With
  'Convert the ActiveX document into a dockable tool window in the VB IDE
  'Uses the CreateToolWindow function
  'VB help doesn't explain this line very clearly so here's my explanation.
  'Set to an object of type Window (this can be private to the Add-in designer, it's only purpose is to
  '              allow you to make the add-in visible either when the menu button is clicked or during VB startup (see below)
  '1st parameter comes form this routine's parameters (don't touch)
  '2nd parameter is project name (up at top of project window on right (usually) plus a '.' connector plus the name of your UserDocument (bottom of project window)
  '3rp parameter the name you want to appear on tool (I use a routine to keep the Ver number up to date)
  '4th parameter a Guid number (you must generate a new one for each program DO NOT just cut and paste the one in the help file
  '              Use a tool called Guidgen.exe, which is located in the \tools\idgen directory of Visual Basic CD.
  '5th parameter The name of your Userdocument (or a variable holding a reference to it which is declared  As <Type = UserDocument's name>
  '              I used 'Public mobjDoc           As docfind' which I placed in a bas Module so that it was public to other parts of code
  Set mWindow = VBInstance.Windows.CreateToolWindow(AddInInst, "FindDoc.docFind", AppDetails, mstrGuid, mobjDoc)
  'Read Preferences from Registry (if they exist)
  mobjDoc.SettingsLoad
  mobjDoc.DoResize
  'This is what allows Find to act like part of VB and appear during VB's startup process
  'simple but not mentioned in any documentation I saw, just found by accident
  If bLaunchOnStart Then
    mWindow.Visible = True
    mobjDoc.Show
  End If
ExitSub:

Exit Sub

ErrorHandler:
  MsgBox "An error has occured!" & vbNewLine & Err.Number & ": " & Err.Description, vbCritical
  Resume ExitSub

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
                                          custom() As Variant)

  'This event fires when the user explicitly unloads the add-in
  
  Dim frm As Form

  On Error Resume Next
  'save properties
  'Persist Preferences to Registry
  mobjDoc.SettingsSave
  'Remove our button from the toolbar
  mobjMCBCtl.Delete
  'destroy the support forms
  'Added to v1.1.01 clean up properly
  For Each frm In Forms
    Unload frm
    Set frm = Nothing
  Next frm
  'Destroy the UserDocument
  Unload mobjDoc
  Set mobjDoc = Nothing
  'Destroy the add-in window
  Unload mWindow
  Set mWindow = Nothing
  On Error GoTo 0

End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)

  mobjDoc.ApplyChanges

End Sub

Private Sub AddinInstance_Terminate()

  '<STUB> Reason:' Add_ins need this even if empty


End Sub

Private Sub mobjCBEvts_Click(ByVal CommandBarControl As Object, _
                             handled As Boolean, _
                             CancelDefault As Boolean)

  'When the user clicks on the toolbar button, we want to display the add-in window

  mWindow.Visible = True
  mobjDoc.Show
  mobjDoc.Startup

End Sub

':) Roja's VB Code Fixer V1.1.2 (7/07/2003 2:02:43 AM) 22 + 157 = 179 Lines Thanks Ulli for inspiration and lots of code.

