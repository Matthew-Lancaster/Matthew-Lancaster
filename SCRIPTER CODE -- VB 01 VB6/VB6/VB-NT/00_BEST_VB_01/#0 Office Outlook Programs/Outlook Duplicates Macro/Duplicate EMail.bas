Attribute VB_Name = "Module1"
'-----------------------------------------------------------------------
'  VBA Macros to delete « duplicates » entries in Microsoft Outlook
'
'  Author            : J.-C. Stritt
'  Last update       : 03-NOV-2002
'  First release     : 31-OCT-2002
'  Environment       : VBA for Outlook 2002
'  Operating system  : Windows XP
'
'  Goal              : - EmailsDeleteDuplicates delete email duplicates
'                      - CalendarDeleteDuplicates delete calendar duplicates
'                      - TaskDeleteDuplicates delete task duplicates
'
'  Remarks           : - first based on Microsoft Q294457 - OL2002
'                        "How to Programmatically Search a Folder Tree"
'
'                      - delete duplicates algorithms is based
'                        on a sort in a folder collection items
'                        and a string compare key builded so :
'                        Email = SenderName + Subject + ReceivedTime
'                        Calendar = Subject + StartDateTime + EndDateTime
'                        Task = Subject + Body + DueDate + DateCompleted
'                             + Importance + PercentComplete
'
'                        Note: string are limited to 40 characters
'
'                      - all these macros give you the possibility to
'                        pickup an entry folder. Then, the macro process
'                        each entry (delete if match) in the folder
'                        and recursivly to all subfolders (if exists)
'
'  (C) Copyright 2002 by Jean-Claude Stritt (jcsinfo@bluewin.ch)
'  Note : you can copy this code freely, but please keep these comments
'-----------------------------------------------------------------------
Option Explicit

'change if necessary these constants because the are Outlook language dependant
Const MailErasedFolder = "Deleted Items" 'language dependant
Const MailFieldToSort = "Received"       'language dependant
Const CalendarFieldToSort = "Start"
Const TaskFieldToSort = "Subject"
Const StrCompareLimit = 40               'check only first N characters for strings

Dim lCount As Long 'to count the deleted items


'add blanks at the ending of a string
Function AddBlanks(ByVal S As String, ByVal L As Integer) As String
  Dim i As Integer, Diff As Integer
  S = LTrim(S)
  Diff = L - Len(S)
  If Diff > 0 Then
    For i = 1 To Diff
      S = S + " "
    Next
  ElseIf Diff < 0 Then
    S = Left(S, L)
  End If
  AddBlanks = S
End Function





'the macro that delete the duplicates in emails folders
Sub EmailsDeleteDuplicates()
  Dim olApp As Outlook.Application
  Dim olSession As Outlook.NameSpace
  Dim olStartFolder As Outlook.MAPIFolder
  Dim strPrompt As String

  'initialize some global var
  lCount = 0

  'get a reference to the Outlook application and session.
  Set olApp = Application
  Set olSession = olApp.GetNamespace("MAPI")

  'ok to begin process ?
  If MsgBox("Ok to delete duplicate emails from a given pickup folder ?", vbYesNo + vbInformation) = vbYes Then
        
    'allow the user to pick the folder in which to start the search.
    Set olStartFolder = olSession.PickFolder
  
    'check to make sure user didn't cancel PickFolder dialog.
    If Not (olStartFolder Is Nothing) Then
      
      'process the first folder (and other by recursive calls to ProcessFolder)
      Call ProcessFolder(olStartFolder)
      MsgBox CStr(lCount) & " messages were deleted."
    
    End If
  End If
End Sub

'the process folder : each folder item is compared to the previous to delete duplicates
Sub ProcessFolder(CurrentFolder As Outlook.MAPIFolder)
  Dim i As Long
  Dim strLastKey As String, strNewKey As String
  Dim olNewFolder As Outlook.MAPIFolder
  Dim olTempItem As Object     'could be various item types
  Dim myItems As Outlook.Items 'a local copy of the collection
  
  'initialize last key string
  strLastKey = ""
   
  'copy the collection (it's obligatory for the sort) and sort them
  Set myItems = CurrentFolder.Items
  'Const MailFieldToSort = "Received"       'language dependant
  Call myItems.Sort("[" & MailFieldToSort & "]", True)
  
  'loop through the items in the current folder (backwards in this case of items to delete)
  For i = myItems.Count To 1 Step -1
     
    Set olTempItem = myItems(i)
 
    'process only mail items
    If TypeName(olTempItem) = "MailItem" Then
      With olTempItem
        strNewKey = AddBlanks(.SenderName, StrCompareLimit) _
                  & AddBlanks(.Subject, StrCompareLimit) _
                  & .ReceivedTime
        'Debug.Print strNewKey
        
        'check to see if a match is found
        If strNewKey = strLastKey Then
          
          'if you want debug then uncomment next line
          'Debug.Print strNewKey
             
          'delete the item (comment the next line if you want just debug)
          olTempItem.Delete
          
          'count deleted items
          lCount = lCount + 1
        End If
        
        'memorize last key found
        strLastKey = strNewKey
      End With
    End If
  
  Next

  'loop through and search each subfolder of the current folder.
  For Each olNewFolder In CurrentFolder.Folders
    If olNewFolder.Name <> MailErasedFolder Then
      Call ProcessFolder(olNewFolder)
    End If
  Next

End Sub





'the macro that delete the duplicates in calendar folder
Sub CalendarDeleteDuplicates()
  Dim olApp As Outlook.Application
  Dim olSession As Outlook.NameSpace
  Dim olStartFolder As Outlook.MAPIFolder
  Dim strPrompt As String

  'initialize some global var
  lCount = 0

  'get a reference to the Outlook application and session.
  Set olApp = Application
  Set olSession = olApp.GetNamespace("MAPI")

  'ok to begin process ?
  If MsgBox("Ok to delete duplicates in calendar ?", vbYesNo + vbInformation) = vbYes Then
        
    'allow the user to pick the folder in which to start the search.
    Set olStartFolder = olSession.PickFolder
  
    'check to make sure user didn't cancel PickFolder dialog.
    If Not (olStartFolder Is Nothing) Then
      
      'process the first folder (and other by recursive calls to ProcessFolder)
      Call ProcessFolder2(olStartFolder)
      MsgBox CStr(lCount) & " duplicates were deleted."
    
    End If
  End If
End Sub

'the process folder : each folder item is compared to the previous to delete duplicates
Sub ProcessFolder2(CurrentFolder As Outlook.MAPIFolder)
  Dim i As Long
  Dim strLastKey As String, strNewKey As String
  Dim olNewFolder As Outlook.MAPIFolder
  Dim olTempItem As Object     'could be various item types
  Dim myItems As Outlook.Items 'a local copy of the collection
  
  'initialize last key string
  strLastKey = ""
   
  'copy the collection (it's obligatory for the sort) and sort them
  Set myItems = CurrentFolder.Items
  Call myItems.Sort("[" & CalendarFieldToSort & "]", True)
  
  'loop through the items in the current folder (backwards in this case of items to delete)
  For i = myItems.Count To 1 Step -1
     
    Set olTempItem = myItems(i)
 
    'process only mail items
    If TypeName(olTempItem) = "AppointmentItem" Then
      With olTempItem
        strNewKey = AddBlanks(.Subject, StrCompareLimit) & " " _
                  & .Start & " " _
                  & .End
        'Debug.Print strNewKey
        
        'check to see if a match is found
        If strNewKey = strLastKey Then
          
          'if you want debug then uncomment next line
          'Debug.Print strNewKey
             
          'delete the item (comment the next line if you want just debug)
          myItems.IncludeRecurrences = False
          olTempItem.Delete
          
          'count deleted items
          lCount = lCount + 1
        End If
        
        'memorize last key found
        strLastKey = strNewKey
      End With
    End If
  
  Next

  'loop through and search each subfolder of the current folder.
  For Each olNewFolder In CurrentFolder.Folders
    If olNewFolder.Name <> MailErasedFolder Then
      Call ProcessFolder2(olNewFolder)
    End If
  Next

End Sub





'the macro that delete the duplicates in tasks folder
Sub TaskDeleteDuplicates()
  Dim olApp As Outlook.Application
  Dim olSession As Outlook.NameSpace
  Dim olStartFolder As Outlook.MAPIFolder
  Dim strPrompt As String

  'initialize some global var
  lCount = 0

  'get a reference to the Outlook application and session.
  Set olApp = Application
  Set olSession = olApp.GetNamespace("MAPI")

  'ok to begin process ?
  If MsgBox("Ok to delete duplicates in task folder ?", vbYesNo + vbInformation) = vbYes Then
        
    'allow the user to pick the folder in which to start the search.
    Set olStartFolder = olSession.PickFolder
  
    'check to make sure user didn't cancel PickFolder dialog.
    If Not (olStartFolder Is Nothing) Then
      
      'process the first folder (and other by recursive calls to ProcessFolder)
      Call ProcessFolder3(olStartFolder)
      MsgBox CStr(lCount) & " duplicates were deleted."
    
    End If
  End If
End Sub

'the process folder : each folder item is compared to the previous to delete duplicates
Sub ProcessFolder3(CurrentFolder As Outlook.MAPIFolder)
  Dim i As Long
  Dim strLastKey As String, strNewKey As String
  Dim olNewFolder As Outlook.MAPIFolder
  Dim olTempItem As Object     'could be various item types
  Dim myItems As Outlook.Items 'a local copy of the collection
  
  'initialize last key string
  strLastKey = ""
   
  'copy the collection (it's obligatory for the sort) and sort them
  Set myItems = CurrentFolder.Items
  Call myItems.Sort("[" & TaskFieldToSort & "]", True)
  
  'loop through the items in the current folder (backwards in this case of items to delete)
  For i = myItems.Count To 1 Step -1
     
    Set olTempItem = myItems(i)
 
    'process only mail items
    If TypeName(olTempItem) = "TaskItem" Then
      With olTempItem
        strNewKey = AddBlanks(.Subject, StrCompareLimit) & " " _
                  & AddBlanks(.Body, StrCompareLimit) & " " _
                  & .DueDate & " " _
                  & .DateCompleted & " " _
                  & .Importance & " " _
                  & .PercentComplete
        'Debug.Print strNewKey
        
        'check to see if a match is found
        If strNewKey = strLastKey Then
          
          'if you want debug then uncomment next line
          'Debug.Print strNewKey
             
          'delete the item (comment the next line if you want just debug)
          olTempItem.Delete
          
          'count deleted items
          lCount = lCount + 1
        End If
        
        'memorize last key found
        strLastKey = strNewKey
      End With
    End If
  
  Next

  'loop through and search each subfolder of the current folder.
  For Each olNewFolder In CurrentFolder.Folders
    If olNewFolder.Name <> MailErasedFolder Then
      Call ProcessFolder3(olNewFolder)
    End If
  Next

End Sub


