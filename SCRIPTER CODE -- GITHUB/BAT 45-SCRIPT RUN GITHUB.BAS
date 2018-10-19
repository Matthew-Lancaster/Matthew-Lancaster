Attribute VB_Name = "Module1"
'--------------------------------------------------------------------------------
'    Component  : Module1
'    Project    : BAT 45-SCRIPT RUN GITHUB
'
'    Description: Sub Main
'
'    Author     : Matthew Lanacster _ Matt.Lan@Btinternet.com
'    Modified   : Sunday 04:05:20 Pm_07 October 2018
'--------------------------------------------------------------------------------


Sub Main()

    
'    On Error Resume Next
'    FILE1 = "C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.vbw"
'    If Dir(FILE1) <> "" Then
'        Kill FILE1
'    End If
'    On Error GoTo 0

    '-------------------------------------------------------------------
    ' FOR THIS CODE ENABLE _ MICROSOFT SCRIPTING RUNTIME _ IN REFERENCES
    '-------------------------------------------------------------------
    ' IT WILL DISPLAY THE DOS BOX PROMPT IN A SHOW WINDOW
    '-------------------------------------------------------------------
    
       
    ' Shell "CMD /K START """" /REALTIME /MAX """ + "C:\Program Files\Siber Systems\GoodSync\gsync.exe" + """" + " sync ""C SCRIPTOR __ y _ 7G __ GITHUB""", vbMaximizedFocus
    I = Shell("CMD /K " + """""" + "C:\Program Files\Siber Systems\GoodSync\gsync.exe" + """" + " " + "sync " + """" + "C SCRIPTOR __ y _ 7G __ GITHUB" + """""", vbMaximizedFocus)
    
    MsgBox I
    
    'MsgBox "CMD /K " + """" + "C:\Program Files\Siber Systems\GoodSync\gsync.exe" + """" + " sync " + """" + "C SCRIPTOR __ y _ 7G __ GITHUB" + """"

    ' EXPERIMENTATION FOUND THIS RESULT
    'CMD /K ""C:\Program Files\Siber Systems\GoodSync\gsync.exe" sync "C SCRIPTOR __ y _ 7G __ GITHUB""

    End
    
    If Command$ <> "" Then Exit Sub
    
    Dim RUN_EXE
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    RUN_EXE = "C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.BAT"
    objShell.Run """" + RUN_EXE + """", 1, False
    Set objShell = Nothing

End Sub




'------------------------------------
'Source 12 Links Open and Find Source
'------------------------------------


'----
'Running external command (VBS & WSH)
'https://www.virtualhelp.me/scripts/57-vb-script/424-running-external-command-vbs-wsh
'----

'2.Using Run

'If you want to run a program in a new process:

'object .Run(strCommand, [intWindowStyle], [bWaitOnReturn])

'intWindowStyle is an integer value indicating window style. Here's a table of styles:

'intWindowStyle Description
'0   Hides the window and activates another window.
'1   Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
'2   Activates the window and displays it as a minimized window.
'3   Activates the window and displays it as a maximized window.
'4   Displays a window in its most recent size and position. The active window remains active.
'5   Activates the window and displays it in its current size and position.
'6   Minimizes the specified window and activates the next top-level window in the Z order.
'7   Displays the window as a minimized window. The active window remains active.
'8   Displays the window in its current state. The active window remains active.
'9   Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
'10  Sets the show-state based on the state of the program that started the application.

'bWaitOnReturn an option to either wait for the process to return or continue without it (can be either true or false)
