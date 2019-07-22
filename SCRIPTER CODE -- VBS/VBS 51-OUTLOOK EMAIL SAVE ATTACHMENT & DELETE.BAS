'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 51-OUTLOOK EMAIL SAVE ATTACHMENT & DELETE.VBS
'# __
'# __ VBS 51-OUTLOOK EMAIL SAVE ATTACHMENT & DELETE.VBS
'# __
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __
'# 1st Work
'# START     TIME [ Fri 03-May-2019 17:55:00 ]
'# END FINAL TIME [ Fri 03-May-2019 18:10:00 ]
'# __
'# 1st Work Total Day Resume After
'# START     TIME [ Fri 03-May-2019 17:55:00 ]
'# END FINAL TIME [ Sat 04-May-2019 00:00:49 ]
'# __
'# Work -- SESSION 04 OF **
'# START     TIME [ Sun 05-May-2019 09:15:00 ]
'# END FINAL TIME [ Sun 05-May-2019 11:32:00 ]
'# __
'# Work -- SESSION 07 OF **
'# START     TIME [ Sun 11-Jun-2019 09:00:00 ]
'# END FINAL TIME [ Sun 11-May-2019 12:00:00 ]
'# __
'# Work -- SESSION 08 OF 08
'# START     TIME [ Wed 12-Jun-2019 10:05:50 ]
'# END FINAL TIME [ Wed 12-Jun-2019 11:40:00 ]
'# __
'# Time Span Begin to Ender
'# START     TIME [ Fri 03-May-2019 18:10:00 ]
'# END FINAL TIME [ Sun 11-May-2019 12:00:00 ]
'# __
'
' MORE DETAIL ABOUT SESSION INFO FOLLOW
'
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION
' -------------------------------------------------------------------
' IDEA GMAIL WITH ONLY 15 GIG KEEP FILL UP FULL WITH ATTACHMENT
' REQUIRE SAVE ATTACHMENT AND DELETE EMAIL AFTER A DATE
' BEEN SPEND SOME TIME CHECKER SOFTWARE TO DO THIS FROM
' Fri 03-May-2019 12:01:57
' ALL VERY EXPENSIVE SOFTWARE
' EASIER WAY IS USE OUTLOOK AND VB-SCRIPTOR
' IS SOME LEARN DONE
' IMAP FOLDER IS BEST WAY TO GO
' WHEN READ THE GMAIL
' THE FOLDER STRUCTURE APPEAR BUT ALL IS IN INBOX
' GET OVER LIMIT OF POP3 MAIL ONLY READ INBOX NOT WHAT IN FOLDER
' OOPS WELL TRY FIND VSB SCRIPT BUT SEEM HARD MUST USE VBA AS HERE
'
' WELL I INCLUDE IT IDEALLY AS A VBS BUT WITH EXTENSION .BAS
' IT BETTER SUIT TO RUN AS BAS RATHER THAN VBS

'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' -------------------------------------------------------------------
'# START     TIME [ Fri 03-May-2019 18:10:00 ]
'# END FINAL TIME [ Fri 03-May-2019 18:10:00 ]
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 002 __ CREATE DAY -- LATE NIGHT CODE
' & ALSO OTHER CODE MORE WAS 12:20 BED AND UP AGAIN MORE OTHER CODE AT 3:20 AM
' WITH JUST 3 HOUR ASLEEP PATTERN
' BEGIN FRIDAY DAY WAKE UP WAS 7 AM UNTIL SAT 04 MAY 12 12:20
' -------------------------------------------------------------------
'# START TIME [ Fri 03-May-2019 20:50:05 ]
'# END FINAL TIME [ Sat 04-May-2019 00:00:49 ]
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 003 __
' -------------------------------------------------------------------
' FEW MORE WORK THE NEXT DAY
' SEEM FAULT AS NOT ABLE TO READ A IMAP FOLDER
' SO USE POP3 FOR NOW
' ADDED ON EMPTY THE DELETED FOLDER AFTER USER AND BEFORE FUNNY
' YET TO SEE OF DELETE ON THIS CLIENT SIDE HAPPEN ON-LINE WEB SERVER
' TO USE THE CODE
' IT IS MORE LIKE VBA VB6 THAN VBSCRIPT
' BUT STILL SAME ALMOST
' THE IDE IN OUTLOOK IS QUITE A GOOD ONE
' FOR PROPER-ING SYNTAX AND THING
' AND USUAL HOT KEY TO RUN IT ALL
' LOOK UP MACRO SET SECURITY LOW SO ABLE RUN THE SCRIPT HERE
' WHEN GO TO VISUAL BASIC EDITOR
' PASTE THE CODE INTO THIS_OUTLOOK_SESSION ALL ONE WORD
' YOU CAN HAVE ANOTHER MODULE IF CODE BIGGER
' AND THEN LOOK AT MACRO
' EVERY ROUTINE YOU PUT IN WILL CAN AS MACRO TO RUN
' SEEM TO DO IN ORDER FIRST COME
' SO TOP MAIN ROUTINE COULD SUB CALL INTO A FEW OTHER
' HAD DIFFICULTY AS WANT TO SCHEDULE IN SOMEHOW AND ON
' FOUND ANSWER HERE
' ----
' Outlook's Rules and Alerts: Run a Script
' https://www.slipstick.com/outlook/rules/outlooks-rules-and-alerts-run-a-script/
' ----
' YOU GOT TO INSERT THIS SNIPPET INTO YOUR CODE AS MAIN EXECUTION ROUTINE CALL
' TO USER RUN BY RULE
' -------------------------------------------------------------------
' Public Sub ShowMessage(Item As Outlook.MailItem)
' Call MAIN
' End Sub
'
' EMAIL ARRIVE DOESN'T SHOW RUN SCRIPT WITH MACRO NOT SHOW THERE TO SELECT IN
' HAD TO FIND MY OWN CODE WHICH USED TO RUN ENTIRELY WITHIN VBA AS A STAND ALONE APP
' AS WEBSITE PROVIDE CODE MAY OF WORK ONE TIME BUT BIT BUGGY FOR CLEVER
' ANOTHER NICE THING TO LEARN
' JUST LIKE THE VB IDE ENVIRONMENT
' YOU CAN CHANGE CONTROL BUTTON CUSTOMIZE QUITE MANICALLY TO WHATEVER YOU WANTER
' AND HAVE THE SCRIPT TO RUN AS A BUTTON ANYTIME WANT
' REMEMBER TO EXIT PROPERLY LET IT DO IT CLOSE DOWN WHICH TAKE TIME EVERY NOW AND THEN
' OR ELSE ALL CUSTOMIZATION GONE & AND YOUR CODE UP-TO-DATE
' JUST DRAG AND DROP THEM CONTROL CUSTOMIZE
' AND SET ICON HAVE LABEL ALSO
' YOU CAN EVEN DRAG A MAIN MENU ITEM FROM PULL DOWN
' AND DRAG DROP AS A BUTTON THERE EASIER THAN VB6 WOULD BE
' -------------------------------------------------------------------
'# START     TIME [ Sat 04-May-2019 05:43:27 ]
'# END FINAL TIME [ Sat 04-May-2019 13:28:00 ]
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 004 __
' -------------------------------------------------------------------
' NEW 3RD DAY
' LOT MORE CODE TODAY
' ERROR CHECKER IS MORE IMPORTANT
' ERROR TRAPPING DOESN'T QUITE REPORT A LOT
' IF ERROR IT BREAK THE CODE AND JUMP BACK TO SOME PART BEFORE LIKE IN NEST LOOP
' MUST CHECK COUNT OF ITEM BEFORE ASK COLLECT THEM
' SOUND EFFECT INCLUDED RATHER THAN SET RULE
' WAS A DESIRE NIGHT BEFORE TO ALLOW_GO CODE RUN UNTIL FINISH
' AND THAT DOES HAPPEN WITH ERROR TRAP ALL CLEAR
' ANOTHER THING IS LESS SPAGHETTI CODE ON UNIQUE NAME
' AND A STOP THING
' WITH COMPARE SO THAT ANY IMAGE WITH SAME FILENAME BUT
' DIFFERENT CONTENT ARE CHECKED OVER
' THAT IS IMPORTANT BECAUSE ALL ATTACHMENT ARE COLLECTING
' BUT NOT WANT DELETE ANY EMAIL UNLESS WEEK OLDER
' MAYBE DEVISE A SYSTEM FOR THAT ONE AS DELAY EACH CHECK MAIL LATER
' WHEN GOT OVER THE INITIAL BACKLOG TO DO WHICH OKAY AT MOMENT
' NOT GOING TO BE SPEEDY FOR EACH MAIL ARRIVE CHECK ON TEN DAY WORTH
' OKAY FOR NOW WORKING -- GET ALL DONE
' -------------------------------------------------------------------
'# START     TIME [ Sun 05-May-2019 09:15:00 ]
'# END FINAL TIME [ Sun 05-May-2019 11:32:00 ]
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 005 __
' -------------------------------------------------------------------
' FOR FIRST TIME CODE IS ABLE TO LEFT RUNNING AND WORK WELL
' WE ADD NEW ROUTINE TO OPEN FILE ON DRIVE TO INDICATE WORKER IN ACTION
' OR ELSE THE OUTLOOK BEGIN RUN THE SCRIPT AND NEW MAIL ARRIVE VERY QUICKLY
' IS PROCESSOR THE SCRIPT TO RUN AGAIN
' NEW TEMP FILE ALLOW_GO SCRIPT TO RUN FULL BEFORE WANT AGAIN
' EVERYTHING RUN SMOOTH AND QUICKER NOW
' ALSO A CODE IN AUTOHOTKEY THAT DEL THE TEMP FILE IF OUTLOOK NOT RUNNER
' SOMEWAY THE TEMP FILE LEFT BEHIND WITH CRASH SHUTDOWN OF OUTLOOK
' -------------------------------------------------------------------
'# START     TIME [ Sat 25-May-2019 15:14:44 ] 1 HOUR
'# END FINAL TIME [ Sat 25-May-2019 16:14:44 ]
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 006 __
' -------------------------------------------------------------------
' ANOTHER DAY
' WORK IT ALL
' IS A FILE NOW WITH CONFIG OPTION
' SET FILE TO STOP RUN BY EDIT CHANGE THERE
' AND EDIT TO HAVE SOUND OR NOT
' -------------------------------------------------------------------
'# START     TIME [ Sun 26-May-2019 14:15:30 ] 1 HOUR
'# END FINAL TIME [ Sun 26-May-2019 14:58:00 ]
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 007 __
' -------------------------------------------------------------------
' STOP START BUTTON FRONT SCREEN WORK GOOD
' THAT TAKE AWAY FROM CONTROL IN A FILE
'
' MORE DEFINE OF HOW I WANT FILE RENAME DONE
' CREATE FOLDER TREE DONE PROPERLY
'
' ADD FOLDER SCANNER NOT TO BE DONE EACH TIME
' AND LEARN DIFFERENCE OF DEFAULT FOLDER SCAN DO AND PERSONAL FOLDER
' WHICH DEFAULT ARE PRESET ONE
' USE HARDCODER AND IF NUMBER FOR PERSONAL FOLDE CHANGE GIVE WARNING
' USE THE ROUTINE TO SET NEW NUMBER AGAIN ON YOUR OWN
' ALSO ADD MOVE EMAIL THAT HASN;T GOT ATTACHMENT TO ANOTHER FOLDER FOR THEM
' SO NOT CLOGGING THE INBOX _ TO DO LATER FIND WHAT ABLE DO DELETER
'
' PLAY ABOUT SOD
' -------------------------------------------------------------------
'# START     TIME [ Sun 11-Jun-2019 09:00:00 ]
'# END FINAL TIME [ Sun 11-Jun-2019 12:00:00 ]
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' SESSION 008 __
' -------------------------------------------------------------------
' MAKE THE STOP START ALLOW_GO BUTTON WORK BETTER
' NOW WE HAVE PROCESS PID NUMBER FOR OUTLOOK
' SO START UP RUN SMOOTHER
' SOMETIME YOU DO REQUEST STOP AND CLOSE OUTLOOK
' BUT WHEN RESUME YOU FORGET PRESS STOP
' AND WANT TO BEGIN WHEN REQUIR WANTER PRESS ALLOW_GO BUTTON FIRST
' THIS IS NOW DONE HERE
'
' IS GOOD IF SCHEDUAL IS SET TO RUN AUTO GO OPENING
' AND THEN THAT WILL BEGIN
' -------------------------------------------------------------------
'# START     TIME [ Wed 12-Jun-2019 10:05:50 ]
'# END FINAL TIME [ Wed 12-Jun-2019 11:40:00 ]
'#-------------------------------------------------------------------



' -------------------------------------------------------------------
' SOURCE HERE CREDIT -- JUST A LITTLE HELP DONE MANY THIS ONE BEFORE
' WAS ADDITIONAL HELPFUL READER
' -------------------------------------------------------------------
' Save Attachments to the Hard Drive
' https://www.slipstick.com/developer/save-attachments-to-the-hard-drive/
' -------------------------------------------------------------------

'# ------------------------------------------------------------------
' LOCATION ON-LINE
'# ------------------------------------------------------------------
' Matthew-Lancaster/VBS 51-OUTLOOK EMAIL SAVE ATTACHMENT & DELETE.BAS
' https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20VBS/VBS%2051-OUTLOOK%20EMAIL%20SAVE%20ATTACHMENT%20%26%20DELETE.BAS
' -------------------------------------------------------------------

Public SOME_PROCESS_DONE
Public SOME_ITEM_DELETER
Public SOUND_OFF_ON

Private Declare Function sndPlaySound32 _
    Lib "winmm.dll" _
    Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long

'sndPlaySound32 "C:\Windows\Media\Chimes.wav", 0&
'If the file cannot be found, Windows will play the default sound, unless this is disabled by using the value SND_NODEFAULT in the uFlags parameter (see below).
'
'The uFlags parameter specifies how the sound is to be played. You can combine multiple options by adding their respective values together and passing the sum in the uFlags parameter.
'
'Const SND_SYNC = &H0        ' (Default) Play the sound synchronously. Code execution
'                            ' pauses until sound is complete.
'
'Const SND_ASYNC = &H1       ' Play the sound asynchronously. Code execution
'                            ' does not wait for sound to complete.
'
'Const SND_NODEFAULT = &H2   ' If the specified sound is not found, do not play
'                            ' the default sound (no sound is played).
'
'Const SND_MEMORY = &H4      ' lpszSoundName is a memory file of the sound.
'                            ' Not used in VBA/VB6.
'
'Const SND_LOOP = &H8        ' Continue playing sound in a loop until the next
'                            ' call to sndPlaySound.
'
'Const SND_NOSTOP = &H10     ' Do not stop playing the current sound before playing
'                            ' the specified sound.

Sub MAIN()
    
' --------------------------------------------------
' DUMMY SETTER SO ABLE TO THROW THROUGHT THE ROUTINE
' --------------------------------------------------
Dim objMsg As Outlook.MailItem 'Object
' --------------------------------------------------
Call ShowMessage(objMsg)
' --------------------------------------------------


End Sub

Public Sub ShowMessage(Item As Outlook.MailItem)

' ----------------------------------------------------------------------------------------------
' BY THIS ROUTINE IS ONLY WAY ABLE TO USE OUTLOOK TO MAKE ROUTINE WITH ON EVENT MESSENGER ARRIVE
' AND STILL HAVE TO USE __ SUB MAIN __ FOR OTHER TYPE OF CALL
' ----------------------------------------------------------------------------------------------
' IT NOT GOOD TO CALL THIS ROUTINE MANUALLY
' IS DOESN'T SEEM TO HOLD AND DIM VARIABLE IS BETTER USE __ SUB MAIN
' ----------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------
' NEW ROUTINE CHUNK CODE -- IN 3 BYTE SIZE EDUCATION
' [ Wednesday 10:44:20 Am_12 June 2019 ]
' ---------------------------------------------------
' THIS IS A BIT BIGG ROUTINE MAYBE FOR MAIL FLYING IN
' HAVE A GO MAKE EASIER
' LESS BUTTON TO PRESS LESS QUESTION TO ANSWER-AH
' QUICKER START LIFT OFF
' ---------------------------------------------------
' 0001
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'OUTLOOK.EXE'")
' KNOCK UP THIS FOR USE IN MY OUTLOOK SCRIPT CODING
For Each objProcess In colProcesses
    If Not (IsNull(objProcess.ProcessID)) Then
        strScriptProcessID = objProcess.ProcessID
        ' MsgBox "MY PROCESS NUMBER PID IS -- " + CStr(strScriptProcessID)
    End If
Next

' 0002
strFileControlpath = "D:\GM OUTLOOK\File_Control_PID.txt"
If Dir(strFileControlpath) <> "" Then
    FR1 = FreeFile
    Open strFileControlpath For Input As FR1
        Line Input #FR1, LINE_INPUT
        strScriptProcessID_CHECK = Val(LINE_INPUT)
    Close #FR1
End If

' 0003
If strScriptProcessID_CHECK <> strScriptProcessID Then
    strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
    On Error Resume Next
    Kill strFileControlpath
    On Error GoTo 0
    
    strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_STOP.txt"
    On Error Resume Next
    Kill strFileControlpath
    On Error GoTo 0
    
    strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_GO.txt"
    FR1 = FreeFile
    Open strFileControlpath For Output As FR1
        Print #FR1, "FLAG FOR GO _ MESGBOX INFO _ ONE PER RUN OF OUTLOOK"
    Close #FR1
    ' Stop
    
End If
' ---------------------------------------------------------------------


strFileControlpath = "D:\GM OUTLOOK\File_Control_PID.txt"
FR1 = FreeFile
Open strFileControlpath For Output As FR1
    Print #FR1, strScriptProcessID
Close #FR1



strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
If Dir(strFileControlpath) <> "" Then
    FR1 = FreeFile
    Open strFileControlpath For Input As FR1
        Line Input #FR1, LINE_INPUT
    Close #FR1
    If LINE_INPUT = "STOP" Then Exit Sub
End If

strFileControlpath = "D:\GM OUTLOOK\File_Control.txt"
If Dir(strFileControlpath) <> "" Then
    FR1 = FreeFile
    Open strFileControlpath For Input As FR1
        Line Input #FR1, LINE_INPUT
        On Error Resume Next
        Line Input #FR1, SOUND_OFF_ON
        On Error GoTo 0
    Close #FR1
    If LINE_INPUT = "STOP" Then Exit Sub
End If


strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_STOP.txt"
If Dir(strFileControlpath) <> "" Then
    FR1 = FreeFile
    Open strFileControlpath For Input As FR1
        Line Input #FR1, LINE_INPUT
    Close #FR1
    If LINE_INPUT = "STOP" Then Exit Sub
End If



    
Call RemoveAllItemsAndFoldersInDeletedItems

If SOME_ITEM_DELETER = True Then
    If SOUND_OFF_ON <> "SOUND OFF" Then
        sndPlaySound32 "C:\Windows\Media\Windows Ringout.wav", 1&
    End If
End If

Call SaveAttachments_Good

SOME_ITEM_DELETER = False
If SOME_PROCESS_DONE = True Then
    Call RemoveAllItemsAndFoldersInDeletedItems
    If SOME_ITEM_DELETER = True Then
        If SOUND_OFF_ON <> "SOUND OFF" Then
            sndPlaySound32 "C:\Windows\Media\Windows Ringout.wav", 1&
        End If
    End If
End If

strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
On Error Resume Next
Kill strFileControlpath
On Error GoTo 0

End

End Sub

Sub RemoveAllItemsAndFoldersInDeletedItems()
 
Dim oDeletedItems As Outlook.Folder
Dim oFolders As Outlook.Folders
Dim oItems As Outlook.Items
Dim i As Long
'Obtain a reference to deleted items folder
Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
If oDeletedItems.Items.Count > 0 Then
    Set oItems = oDeletedItems.Items
    For i = oItems.Count To 1 Step -1
        oItems.Item(i).Delete
        If SOME_ITEM_DELETER = False Then
            If SOUND_OFF_ON <> "SOUND OFF" Then
                sndPlaySound32 "C:\Windows\Media\Windows Ringout.wav", 1&
            End If
        End If
        SOME_ITEM_DELETER = True
        X = X + 1
        If X Mod 10 = 0 Then DoEvents
    Next
    If oDeletedItems.Folders.Count > 0 Then
        Set oFolders = oDeletedItems.Folders
        For i = oFolders.Count To 1 Step -1
            oFolders.Item(i).Delete
            SOME_ITEM_DELETER = True
            X = X + 1
            If X Mod 10 = 0 Then DoEvents
        Next
    End If
End If
'----
'Delete All Items and Subfolders in the Deleted Items Folder | Microsoft Docs
'https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/delete-all-items-and-subfolders-in-the-deleted-items-folder
'----

End Sub


Sub Remove_All_INBOX()
 
Dim oINBOXItems As Outlook.Folder
Dim oFolders As Outlook.Folders
Dim oItems As Outlook.Items
Dim i As Long
'Obtain a reference to deleted items folder
Set oINBOXItems = Application.Session.GetDefaultFolder(olFolderInbox)
If oINBOXItems.Items.Count > 0 Then
    Set oItems = oINBOXItems.Items
    For i = oItems.Count To 1 Step -1
        oItems.Item(i).Delete
        If SOME_ITEM_DELETER = False Then
            ' If SOUND_OFF_ON <> "SOUND OFF" Then
                sndPlaySound32 "C:\Windows\Media\Windows Ringout.wav", 1&
            ' End If
        End If
        SOME_ITEM_DELETER = True
        X = X + 1
        If X Mod 10 = 0 Then DoEvents
    Next
'    If oINBOXItems.Folders.Count > 0 Then
'        Set oFolders = oINBOXItems.Folders
'        For i = oFolders.Count To 1 Step -1
'            oFolders.Item(i).Delete
'            SOME_ITEM_DELETER = True
'            X = X + 1
'            If X Mod 10 = 0 Then DoEvents
'        Next
'    End If
End If
'----
'Delete All Items and Subfolders in the Deleted Items Folder | Microsoft Docs
'https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/delete-all-items-and-subfolders-in-the-deleted-items-folder
'----

End Sub


Public Sub COMMAND_STOP()

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_STOP.txt"

If Dir(strFileControlpath) <> "" Then
    Set f = FSO.GetFile(strFileControlpath)
    FileModDate = f.DateLastModified
    Set f = Nothing
End If

strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_STOP.txt"
FR1 = FreeFile
Open strFileControlpath For Output As FR1
    Print #FR1, "STOP"
Close #FR1

Dim TEXT_1(100)
i = i + 1: TEXT_1(i) = "QUICK REPEAT BUTTON TIMER 10 SECOND INFO"
i = i + 1: TEXT_1(i) = "THIS BUTTON WILL STOP THE PROCESS OF ALL QUICK AS ABLE"
i = i + 1: TEXT_1(i) = "AND THEN YOU HAVE TO PRESS ALLOW_GO BUTTON AGAIN TO MAKE ANYTHING HAPPEN MAIN ROUTINE"
i = i + 1: TEXT_1(i) = "A CHECK IS MADE OUTLOOK HAS BEEN SHUT DONE"
i = i + 1: TEXT_1(i) = "BY PROCESS PID NUMBER SINCE CLOSE TO CLEAR THE FLAG THAT WAY"
i = i + 1: TEXT_1(i) = "----------------------------------------------"
i = i + 1: TEXT_1(i) = "IF THIS STOP BUTTON PRESS STOP WORK GO"
i = i + 1: TEXT_1(i) = "IF ALLOW_GO BUTTON PRESS IT WILL ENABLE RESUME WORK WHEN REQUEST"
i = i + 1: TEXT_1(i) = "----------------------------------------------"
i = i + 1: TEXT_1(i) = "IF OUTLOOK HAS JUST OPENED AND STOP PRESS BEFORE IT WILL STILL BE STOP UNLESS"
i = i + 1: TEXT_1(i) = "IF OUTLOOK HAS JUST OPENED THE PROCESS PID NUMBER WILL CHECK"
i = i + 1: TEXT_1(i) = "AND THEN ALLOW_GO WILL RESUME AS IF PRESS BUTTON"
i = i + 1: TEXT_1(i) = "THE PROCESS PID NUMBER IS CHECK FOR EVERY MAIL THAT ARRIVE"
i = i + 1: TEXT_1(i) = "HAS TO GET THERE BEFORE STOP CHECKER"
i = i + 1: TEXT_1(i) = "----------------------------------------------"
i = i + 1: TEXT_1(i) = "IS GOOD IF SCHEDUAL IS SET TO RUN AUTO GO OPENING"
i = i + 1: TEXT_1(i) = "AND THEN THAT WILL BEGIN"
i = i + 1: TEXT_1(i) = "----------------------------------------------"

For R = 1 To i
    H1 = H1 + TEXT_1(R) + vbCrLf
Next

sndPlaySound32 "C:\Windows\Media\Windows Error.wav", 1&

Dim oDeletedItems As Outlook.Folder
Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems)
If oDeletedItems.Items.Count = 0 Then
    If FileModDate > Now - TimeSerial(0, 0, 10) Then
        MsgBox H1, vbMsgBoxSetForeground
    End If
End If


End Sub

Public Sub COMMAND_GO()

Call COMMAND_ALLOW_GO

End Sub


Public Sub COMMAND_ALLOW_GO()

Dim TEXT_1(100)
i = i + 1: TEXT_1(i) = "ALL MAIN ROUTINE RUNNER ARE SET TO ALLOW_GO THINGS WILL HAPPEN FOR MAIN ROUTINE"
i = i + 1: TEXT_1(i) = "CAREFULL DON'T PRESS THIS BUTTON WHILE INCOMING MESSENGER ARE GO"
i = i + 1: TEXT_1(i) = "THIS IS THE PROBLEM WHEN EACH MSSAGE ARRIVE FROM EMAIL SYSTEM"
i = i + 1: TEXT_1(i) = "AND THEN THIS WHOLE MAIN ROUTINE IS CALLED AND IT TAKE A LONG TIME AND REALLY"
i = i + 1: TEXT_1(i) = "CONFUSE THING IF IT ABLE TO BE QUICKLY PUT ON AND THEN INTERRUPTED OFF"
i = i + 1: TEXT_1(i) = "AND PUT ON AGAIN AS QUICK AS EMAIL ARRIVER HTH"
i = i + 1: TEXT_1(i) = "----------------------------------------------"
i = i + 1: TEXT_1(i) = "IF THIS BUTTON PRESS RESUME WORK GO"
i = i + 1: TEXT_1(i) = "IF STOP BUTTON PRESS IT WILL STOP WORK"
i = i + 1: TEXT_1(i) = "IF OUTLOOK HAS JUST OPENED AND STOP PRESS BEFORE IT WILL STILL BE STOP UNLESS"
i = i + 1: TEXT_1(i) = "IF OUTLOOK HAS JUST OPENED THE PROCESS PID NUMBER WILL CHECK"
i = i + 1: TEXT_1(i) = "AND THEN ALLOW_GO WILL RESUME AS IF PRESS BUTTON"
i = i + 1: TEXT_1(i) = "THE PROCESS PID NUMBER IS CHECK FOR EVERY MAIL THAT ARRIVE"
i = i + 1: TEXT_1(i) = "HAS TO GET THERE BEFORE STOP CHECKER"
i = i + 1: TEXT_1(i) = "----------------------------------------------"
i = i + 1: TEXT_1(i) = "IS GOOD IF SCHEDUAL IS SET TO RUN AUTO GO OPENING"
i = i + 1: TEXT_1(i) = "AND THEN THAT WILL BEGIN"
i = i + 1: TEXT_1(i) = "----------------------------------------------"
i = i + 1: TEXT_1(i) = "MESSENGER HERE DISPLAY ONCE PER RUN OF OUTLOOK"
i = i + 1: TEXT_1(i) = "----------------------------------------------"

For R = 1 To i
    H1 = H1 + TEXT_1(R) + vbCrLf
Next

strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_GO.txt"
'Stop
If Dir(strFileControlpath) <> "" Then
    sndPlaySound32 "C:\Windows\Media\Windows Error.wav", 1&
    MsgBox H1, vbMsgBoxSetForeground
End If
    
strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_GO.txt"
On Error Resume Next
Kill strFileControlpath
On Error GoTo 0


strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_STOP.txt"
On Error Resume Next
Kill strFileControlpath
On Error GoTo 0

strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
On Error Resume Next
Kill strFileControlpath
If Err.Number > 0 Then
    ' FILE NOT FOUND
    ' MsgBox Err.Description
    ' Stop
    If Err.Description <> "File not found" Then
        A_TEST = A_TEST
        MsgBox Err.Description
        Stop
    End If
End If
On Error GoTo 0

sndPlaySound32 "C:\Windows\Media\Windows Error.wav", 1&


End Sub

Public Sub CLEAR_LOCK_FLAG()
strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
On Error Resume Next
Kill strFileControlpath
On Error GoTo 0

End Sub




Public Sub SaveAttachments_Good()


strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_STOP.txt"
If Dir(strFileControlpath) <> "" Then
    FR1 = FreeFile
    Open strFileControlpath For Input As FR1
        Line Input #FR1, LINE_INPUT
    Close #FR1
    strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
    On Error Resume Next
    Kill strFileControlpath
    On Error GoTo 0
    If LINE_INPUT = "STOP" Then Exit Sub
End If



' Init an Outlook Application object.
Dim objOL As New Outlook.Application

' ---------------------------------
' METHOD THROW ERROR IS NOT SAME
' GLAD HAD MY OLDER ARCHIVE OF CODE
' THESE THING TAKE FOREVER TO DEBUG
' ---------------------------------
' Dim objOL As Outlook.Application
' Dim objSelection As Outlook.Selection

Dim objAttachments As Outlook.Attachments
Dim objSelection As MAPIFolder
Dim objMsg As Outlook.MailItem 'Object
Dim oNameSpace As NameSpace

Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String

Dim RR

' Get the path
strFolderpath = "D:\GM OUTLOOK\"


strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
FR1 = FreeFile
Open strFileControlpath For Output As FR1
    Print #FR1, "STOP"
Close #FR1



' Set oFolder = oNameSpace.PickFolder
' Set oFolder = oNameSpace2.GetDefaultFolder(23)  'olFolderJunkFolder
' Set oFolder = oNameSpace3.GetDefaultFolder(3)   'olFolderDeleted
' Set oFolder = oNameSpace.GetDefaultFolder(6)    'olFolderInBox
' Set oParentFolder = oMAPI.Folders("nvr1000.roidsrim@gmail.com")
' Set oParentFolder = oMAPI.Folders(oNameSpace.GetDefaultFolder(0).Name)
' Set oParentFolder = oNameSpace.GetDefaultFolder(1)

Set oNameSpace = objOL.GetNamespace("MAPI")

Set oMAPI = GetObject("", "Outlook.application").GetNamespace("MAPI")

'Set oParentFolder = oMAPI.Folders(1)
'If oParentFolder.Folders.Count Then
'  For RR = 1 To oParentFolder.Folders.Count
'    If Trim(oParentFolder.Folders(RR).Name) = "Inbox" Then
'        ' MsgBox Trim(oParentFolder.Folders(RR).Name)
'        Exit For
'    End If
'  Next RR
'End If
'Set objSelection = oParentFolder.Folders(RR)

Set objSelection = oNameSpace.GetDefaultFolder(olFolderInbox)

Set DeletedFolder = oNameSpace.GetDefaultFolder(olFolderDeletedItems)

'Set oParentFolder = oMAPI.Folders(1)
'If oParentFolder.Folders.Count Then
'  For RR = 1 To oParentFolder.Folders.Count
'    If Trim(oParentFolder.Folders(RR).Name) = "NOT_ATTACHMENT" Then
'        MsgBox Str(RR) + " -- " + Trim(oParentFolder.Folders(RR).Name)
'        Exit For
'    End If
'  Next RR
'End If

' -----------------------------------------
' METHOD FOR FIND DEFAULT FOLDER -- NOT FOR PERSONAL MADE FOLDER
' -----------------------------------------
' oNameSpace.GetDefaultFolder.Folders.Count
' -----------------------------------------
If 1 = 2 Then
    For RR = 1 To 100
        On Error Resume Next
        Err.Clear
        Set NOT_ATTACHMENT_Folder = oNameSpace.GetDefaultFolder(RR)
        If Err.Number = 0 Then
            EXAMPLE = NOT_ATTACHMENT_Folder.Name
            If NOT_ATTACHMENT_Folder.Name = "NOT_ATTACHMENT" Then
              MsgBox Str(RR) + " -- " + NOT_ATTACHMENT_Folder.Name
              Exit For
            End If
        End If
        On Error GoTo 0
    Next RR
End If

Set oParentFolder = oMAPI.Folders(1)
If 1 = 2 Then
    If oParentFolder.Folders.Count Then
      For RR = 1 To oParentFolder.Folders.Count
        If oParentFolder.Folders(RR).Name = "NOT_ATTACHMENT" Then
            MsgBox Str(RR) + " -- " + oParentFolder.Folders(RR).Name
            Exit For
        End If
      Next RR
    End If
' Set NOT_ATTACHMENT_Folder = oParentFolder.Folders(RR)
End If

Set NOT_ATTACHMENT_Folder = oParentFolder.Folders(15)
If NOT_ATTACHMENT_Folder.Name <> "NOT_ATTACHMENT" Then
    MsgBox "THE ""NOT_ATTACHMENT"" FOLDER IS NOT CONFIGER TO CORRECT NUMBER SUPPOSED TO BE 15 -- STOP"
    Stop
End If

' MsgBox objSelection.Name
' MsgBox Str(objSelection.Items.Count)

' Check each selected item for attachments.
' For Each objMsg In objSelection

X_COUNT = objSelection.Items.Count
If X_COUNT > 0 Then
For RR = X_COUNT To 1 Step -1

    
    ' ----------------------------
    ' HAD A CHECK AT BEGIN ROUTINE
    ' AND ANOTHER HERE IN THE LOOP
    ' ----------------------------
    strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP_STOP.txt"
    If Dir(strFileControlpath) <> "" Then
        FR1 = FreeFile
        Open strFileControlpath For Input As FR1
            Line Input #FR1, LINE_INPUT
        Close #FR1
        strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
        On Error Resume Next
        Kill strFileControlpath
        On Error GoTo 0
        If LINE_INPUT = "STOP" Then Exit Sub
    End If



    Set objMsg = objSelection.Items.Item(RR)
    
    lngCount = objMsg.Attachments.Count
    If lngCount > 0 Then
        Set objAttachments = objMsg.Attachments
    End If
    A1 = Trim(objMsg.Subject)
    d1 = objMsg.ReceivedTime
    D2 = Format(d1, "YYYYMMDDHHMMSS")
    D3 = Format(d1, "YYYYMMDD")
    
    'X = X + 1
    'If X Mod 10 = 0 Then DoEvents

'        TRY TO GET EMAIL MSG ID LIKE SOME OTHER APP ABLE TO - NOT HERE
'        --------------------------------------------------------------
'        D3 = objMsg.EntryID
'        D3 = objMsg.ReceivedByEntryID
'        D3 = objMsg.ReceivedOnBehalfOfEntryID
'        D3 = objMsg.Class
'        D3 = objMsg.MessageClass
'        D3 = objMsg.ConversationIndex
'        MsgBox D3
    H3 = ""
    SET_GO = False
    If InStr(A1, "Alarm Message From Device EDVR") > 0 Then
        SET_GO = True
        H1 = Mid(A1, InStrRev(A1, "On Channel") + 11)
        H3 = "Motion Alarm On Device EDVR_"
        ' D2 = "" ' REMOVE THE TIME FOR THIS ONE AS IT IN .JPG ATTACHMENT
    End If
    ' Embedded Net DVR: Motion Detected On Channel A3
    If InStr(A1, "Embedded Net DVR: Motion Detected On Channel") > 0 Then
        SET_GO = True
        H1 = Mid(A1, InStrRev(A1, "On Channel") + 11)
        H1 = Replace(H1, " ", "_")
        ' H2 = Format(Val(Mid(A1, Len(A1) - 4)), "00")
        ' H3 = "DVR_A03_Motion_Detected"
        ' H3 = "DVR_" + H1 + H2 + "_Motion_Detected"
        H3 = "DVR_" + H1 + "_Motion_Detected_" + D2 + "_" + ""
    End If

    If InStr(A1, "Embedded Net DVR: sound intensity alarm On Channel") > 0 Then
        SET_GO = True
        H1 = Mid(A1, InStrRev(A1, "On Channel") + 11)
        H1 = Replace(H1, " ", "_")
        ' H2 = Format(Val(Mid(A1, Len(A1) - 4)), "00")
        ' H3 = "DVR_A03_Motion_Detected"
        ' H3 = "DVR_" + H1 + H2 + "_Motion_Detected"
        H3 = "DVR_" + H1 + "_Sound_Intensity_Detected_" + D2 + "_" + ""
    End If
    
    If lngCount = 0 Then
        SET_GO = False
    End If
    
    ' NOT WANTED FOR A MINUTE
    ' -----------------------
'        If D1 >= Now - 10 Then
'            SET_GO = False
'        End If
    
    'a1 = MyItems.Item(rrr).SenderName
    'a2 = MyItems.Item(rrr).SenderEmailAddress
    'a3 = Trim(MyItems.Item(rrr).Subject)
    'a4 = Trim(MyItems.Item(rrr).Body)


    ' Use the MsgBox command to troubleshoot. Remove it from the final code.
    ' MsgBox A1
    ' MsgBox Str(lngCount)
                
     ' GET THE ATTACHMENT FROM THEM ALL
    If SET_GO = True Then
        If lngCount > 0 Then

            ' ---------------------------------------------------
            ' Use a count down loop for removing items
            ' from a collection. Otherwise, the loop counter gets
            ' confused and only every other item is removed.
            ' ---------------------------------------------------
            For i = lngCount To 1 Step -1
                    X = X + 1
                    If X Mod 10 = 0 Then DoEvents
                    
                    ' Get the file name.
                    strFile = objAttachments.Item(i).FileName
                        
                    ' -----------------------------------------------------
                    ' HOW BIG WAS THAT STRING MANIPULATE
                    ' WITHOUT ERROR
                    ' -----------------------------------------------------
                    If InStr(A1, "Alarm Message From Device EDVR") > 0 Then
                        A2 = "A" + Mid(strFile, InStr(strFile, "_") - 1, 1)
                        strFolderpath_2 = strFolderpath + D3 + "\" + A2 + "\"
                        If Dir(strFolderpath_2, vbDirectory) = "" Then
                            I_RESULT = CreateFolderTree(strFolderpath_2)
                        End If
                        If Dir(strFolderpath_2, vbDirectory) = "" Then
                            MsgBox "PATH FAILED TO MAKE" + vbCrLf + strFolderpath_2
                        End If
                        
                        strFile_NAME = strFolderpath_2 + H3 + strFile
                                    
                    End If
                    
                    If InStr(A1, "Embedded Net DVR: Motion Detected On Channel") > 0 Then
                        strFolderpath_2 = strFolderpath + D3 + "\" + H1 + "\"

                        strFile_NAME = strFolderpath_2 + H3 + strFile
                        strFile_NAME_2 = H3 + strFile
                        
                        If Dir(strFolderpath_2, vbDirectory) = "" Then
                            I_RESULT = CreateFolderTree(strFolderpath_2)
                        End If
                        If Dir(strFolderpath_2, vbDirectory) = "" Then
                            MsgBox "PATH FAILED TO MAKE" + vbCrLf + strFolderpath_2
                        End If
                        If Dir(strFolderpath_2, vbDirectory) = "" Then
                            MsgBox "PATH FAILED TO MAKE" + vbCrLf + strFolderpath_2
                        End If
                    End If
                    
                    
                    If InStr(A1, "Embedded Net DVR: sound intensity alarm On Channel") > 0 Then
                        strFolderpath_2 = strFolderpath + D3 + " Sound Intensity Alarm\" + H1 + "\"

                        strFile_NAME = strFolderpath_2 + H3 + strFile
                        strFile_NAME_2 = H3 + strFile
                        strFile_TEMP = "D:\GM OUTLOOK\TEMP\" + strFile
                        
                        If Dir(strFolderpath_2, vbDirectory) = "" Then
                            I_RESULT = CreateFolderTree(strFolderpath_2)
                        End If
                        If Dir(strFolderpath_2, vbDirectory) = "" Then
                            MsgBox "PATH FAILED TO MAKE" + vbCrLf + strFolderpath_2
                        End If
                    End If
                    
                    
                    
                    If strFile_NAME <> "" Then
                        If Dir(strFile_NAME) <> "" Then
'                                    UNIQUE = UNIQUE + 1
                            If Dir("D:\GM OUTLOOK\TEMP", vbDirectory) = "" Then
                                On Error Resume Next
                                    MkDir "D:\GM OUTLOOK"
                                    MkDir "D:\GM OUTLOOK\TEMP"
                                On Error GoTo 0
                            End If
                            If Dir("D:\GM OUTLOOK\TEMP", vbDirectory) = "" Then
                                MsgBox "PATH FAILED TO MAKE" + vbCrLf + "D:\GM OUTLOOK\TEMP"
                            End If
                            
                            ' ---------------------------------------------------------------------------------
                            ' THE TEMP FILE IS WIRTTEN AND THEN COMPARE TO WHAT MIGHT ALREADY BE AT DESTINATION
                            ' FOR A DUPE-AH
                            ' MIGHT AS WELL TMEP IS SAME NAME ALWAYS COMPARE TO SEE WHAT CODE ABOVE
                            ' ---------------------------------------------------------------------------------
                            strFile_TEMP = "D:\GM OUTLOOK\TEMP\TEMP.JPG"
                            objAttachments.Item(i).SaveAsFile strFile_TEMP
                            SOME_PROCESS_DONE = True
                            
                            FR1 = FreeFile
                            Open strFile_TEMP For Binary As #FR1
                                PIC1_COMPARE = Input$(LOF(FR1), FR1)
                            Close #FR1
                            FR1 = FreeFile
                            A_TEST = strFile_NAME_2
                            Open strFile_NAME For Binary As #FR1
                                PIC2_COMPARE = Input$(LOF(FR1), FR1)
                            Close #FR1
                            If PIC2_COMPARE <> PIC1_COMPARE Then
                                RENAME = RENAME + 1
                                If RENAME < -1 Then
                                    MsgBox "PIC1_COMPARE<>PIC2_COMPARE" + vbCrLf + vbCrLf + strFile_TEMP + vbCrLf + strFile_NAME + "IMAGE NOT THE SAME AT COMPARE" + vbCrLf + "OUTLOOK SCRIPTOR -- NOW STOP" + vbCrLf + "WILL GET RENAME _DUPE_ AND TIMER"
                                    ' Shell "Explorer.exe /select," + strFile_NAME, vbMaximizedFocus
                                    ' Stop
                                End If
                                If InStr(strFile_NAME, "_Dupe_") > 0 Then
                                    ' Stop
                                    ' strFile_NAME_8 = strFile_NAME
                                    strFile_NAME = Mid(strFile_NAME, 1, InStrRev(strFile_NAME, "_Dupe_") - 1) + "_Dupe_" + Trim(Str(Timer)) + ".jpg"
                                    ' MsgBox strFile_NAME + vbCrLf + strFile_NAME_8
                                Else
                                '1 11
                                    strFile_NAME = Mid(strFile_NAME, 1, InStrRev(strFile_NAME, ".") - 1) + "_Dupe_" + Trim(Str(Timer)) + ".jpg"
                                End If
                                sndPlaySound32 "C:\Windows\Media\Windows Error.wav", 1&

                                ' MsgBox strFile_NAME
                            End If
                        End If
                        ' Save the attachment as a file.
                        objAttachments.Item(i).SaveAsFile strFile_NAME
                        Rem i = objAttachments.Item(i).DisplayName
                        SOME_PROCESS_DONE = True
                        
                        ' SOMETIME IS ERROR HERE A FILE DO EXIST CLICK OPEN WHEN CHECKER
                        ' RESTART AGAIN SUPPOSE
                        '[ Monday 11:54:30 Am_22 July 2019 ]
                        ' --------------------------------------------------------------
                        If FSO.FILEEXISTS(strFile_NAME) = False Then Stop
                        If Dir(strFile_NAME) = "" Then Stop
                    End If
            Next i
            
            ' DELETE THE MAIL WITH ATTACHEMTN USER OF 2 DAY OLDER
            If d1 = d1 Then '< Now - 1 Then
                ' Set objAttachments = Nothing
                On Error Resume Next
'                        objMsg.Delete
                'objMsg.Item(RR).Delete
                ' Set objMsg = objSelection.Items.Item(RR)
'                        FF = objMsg.To
'                        If FF <> "" Then
                 objMsg.Move DeletedFolder
                 SOME_PROCESS_DONE = True
'                        End If
                If Err.Number > 0 Then
                    Stop
                End If
                On Error GoTo 0
                
                RR_X = RR_X + 1
                If RR_X >= 10000 Then
                    GoTo ExitSub
                End If
            End If
        End If
    End If
    
    ' IF MESSAGE HAS NOT ATTACHMENT THAT I USE
    If SET_GO = False Then
        On Error Resume Next
        objMsg.Move NOT_ATTACHMENT_Folder
    End If
    
    X = X + 1
    If X Mod 10 = 0 Then
        DoEvents
    End If
    
Next
End If
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

'----
'Save Attachments to the Hard Drive
'https://www.slipstick.com/developer/save-attachments-to-the-hard-drive/
'----

strFileControlpath = "D:\GM OUTLOOK\File_Control_TEMP.txt"
On Error Resume Next
Kill strFileControlpath
On Error GoTo 0


End Sub


Private Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If nPos - 1 > 3 Then
            If Dir(Left$(sPath, nPos - 1), vbDirectory) = "" Then
                MkDir Left$(sPath, nPos - 1)
            End If
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
    CreateFolderTree = True
    If Dir(sPath, vbDirectory) = "" Then
        CreateFolderTree = False
    End If

    Exit Function

CreateFolderTreeError:
    Exit Function
End Function



' SOME EARLIER SOURCER
'
' -------------------------------------------------------------------
' SOURCE HERE
' -------------------------------------------------------------------
' vba - Remove email attachments after saving them and adding save location to email in Outlook - Stack Overflow
' https://stackoverflow.com/questions/26745239/remove-email-attachments-after-saving-them-and-adding-save-location-to-email-in
' -------------------------------------------------------------------

' LINK REFERENCE WHILE CODING
' ----
' Purchase Automatic Email Manager
' https://www.automatic-email-manager.com/order-automatic-email-manager
' --------
' GMIAL IMAP SERVER NAME - Google Search
' https://www.google.com/search?q=GMIAL+IMAP+SERVER+NAME
' --------
' VBSCRIPT FOR OUTLOOK TO DELETE OLDER MAIL - Google Search
' https://www.google.com/search?q=VBSCRIPT+FOR+OUTLOOK+TO+DELETE+OLDER+MAIL
' --------
' VBS OUTLOOK SAVE ATTACHMENT AND DELETE ON DATE - Google Search
' https://www.google.com/search?safe=active&ei=fG3MXIiCBfPgxgPxhauoAQ&q=VBS+OUTLOOK+SAVE+ATTACHMENT+AND+DELETE+ON+DATE&oq=VBS+OUTLOOK+SAVE+ATTACHMENT+AND+DELETE+ON+DATE&gs_l=psy-ab.3..33i160.3058.18916..19308...2.0..1.150.4598.61j4......0....1..gws-wiz.......0i71j35i304i39.zfpXIKNnuDc
' --------
' Save and Rename Outlook Email Attachments
' https://www.slipstick.com/developer/code-samples/save-rename-outlook-email-attachments/
' --------
' Save Attachments to the Hard Drive
' https://www.slipstick.com/developer/save-attachments-to-the-hard-drive/
' --------
' OutlookCode.com - Developer Learning Center for Microsoft Outlook
' http://www.outlookcode.com/
' --------
' Save and Delete Attachments from Outlook messages
' https://www.slipstick.com/developer/code-samples/save-and-delete-attachments/
' --------
' Attachment Management Tools for Outlook
' https://www.slipstick.com/addins/auto-process/attachment-management-tools-for-outlook/
' --------
' vba - Remove email attachments after saving them and adding save location to email in Outlook - Stack Overflow
' https://stackoverflow.com/questions/26745239/remove-email-attachments-after-saving-them-and-adding-save-location-to-email-in
' --------
' Outlook User's Solutions Center
' https://www.slipstick.com/outlook-user/
' --------
' Q and A - Office Vbscript for Outlook emails backup and cleanup
' https://gallery.technet.microsoft.com/office/Vbscript-for-Outlook-ec13f44a/view/Discussions#content
' --------
' Adding/Recreating a Mail Profile - HowTo-Outlook
' https://www.howto-outlook.com/faq/newprofile.htm
' --------
' vba outlook command to empty deleted items - Google Search
' https://www.google.com/search?safe=active&ei=7FjNXOeiNPHGxgO9r7fABw&q=vba+outlook+command+to+empty+deleted+items&oq=vba+outlook+command+to+empty+deleted+items&gs_l=psy-ab.3...171785.172157..172498...0.0..0.190.679.0j4......0....1..gws-wiz.......0i71.gBe2acgT688
' --------
' Can't empty deleted items folder - Outlook 2016 / Exchange Online Plan 1 - Microsoft Tech Community - 297841
' https://techcommunity.microsoft.com/t5/Office-365/Can-t-empty-deleted-items-folder-Outlook-2016-Exchange-Online/td-p/297841
' --------
' Delete All Items and Subfolders in the Deleted Items Folder | Microsoft Docs
' https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/delete-all-items-and-subfolders-in-the-deleted-items-folder
' --------
' Outlook's Rules and Alerts: Run a Script
' https://www.slipstick.com/outlook/rules/outlooks-rules-and-alerts-run-a-script/
' --------
' HOW MAKE OUTLOOK SEND / RECEIVE PROGREES AS A BUTTON - Google Search
' https://www.google.com/search?q=HOW+MAKE+OUTLOOK+SEND+%2F+RECEIVE+PROGREES+AS+A+BUTTON
' --------
' OUTLOOK HOW MAKE SCRIPT RUN UNTIL END NOT AS EVERY EMAIL ARRIVE - Google Search
' https://www.google.com/search?q=OUTLOOK+HOW+MAKE+SCRIPT+RUN+UNTIL+END+NOT+AS+EVERY+EMAIL+ARRIVE
' --------
' Rule using "run script" not working when new mail arrives
' https://social.technet.microsoft.com/Forums/en-US/5153ffae-fe92-48ab-ae1a-47ff545486b1/rule-using-run-script-not-working-when-new-mail-arrives?forum=outlook
' --------
' Playing A Sound In VBA
' http://www.cpearson.com/excel/PlaySound.aspx
' --------
' Free Sound Effect (SFX) Downloads for WavePad Sound Editor
' https://www.nch.com.au/wavepad/free-sfx-sound-effects-library.html?msclkid=0b951831240f1e6ca0c544ff805bfcc9
' --------
' VBA Code to permanently delete selected email | Outlook Forums by Slipstick.com
' https://forums.slipstick.com/threads/96078-vba-code-to-permanently-delete-selected-email/
' ----







