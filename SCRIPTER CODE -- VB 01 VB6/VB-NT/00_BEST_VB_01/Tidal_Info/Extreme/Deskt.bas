Attribute VB_Name = "PassModule"
Dim FILENAME_IN_USE_CHECK As String

'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Public Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetFocuses Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public PopBack

Type InfoRetrieved
  LenUserName As Long
  LenPassword As Long
  strUserName As String
  strPassword As String
End Type
Enum IsNewProfileNeeded
   yesitdoes = -1
   NoitDont = 0
End Enum
Public UserName(100) As Long
Public Password(100) As Long
Public LenUser As Long
Public LenPass As Long
Public PasswordInMemory As InfoRetrieved
Public Const UserProfile = "usrprofile.exl"
Public CreateProfile As Boolean

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
Private Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function


Public Function IsPasswordCorrect(ByVal lpPassword As String, lpUserName) As IsNewProfileNeeded
Dim NullReturn As Variant
 
 NullReturn = WriteUserNameToFile(lpUserName, lpPassword)
 IsPasswordCorrect = yesitdoes
Exit Function



   
   If lpPassword <> PasswordInMemory.strPassword Then
   IsPasswordCorrect = NoitDont
   Else
   IsPasswordCorrect = yesitdoes
   End If

   
   
End Function
Public Function DoesExist(ByVal lstrQuery As String) As Boolean
DoesExist = (Dir(lstrQuery) <> "")


End Function

Private Function GetUserNameFromFile() As InfoRetrieved
On Error GoTo ErrorHandler
Dim LenUserName As Long
Dim LenPassword As Long
Dim lngUserN(100) As Long
Dim lngPassN(100) As Long

'Dim lpStrUser, lpStrPassword,
Dim Letter As String
Dim lngShuffle As Long
    Open App.Path + "\" + UserProfile For Binary As #1
    
     Get #1, , LenUserName
     Get #1, , LenPassword
     Get #1, , lngUserN
     Get #1, , lngPassN
    
    Close #1
    
    For i = 1 To LenUserName
       lpStrUser = lpStrUser & Chr$((lngUserN(i) / 2))
    Next i
    For i = 1 To LenPassword
       lpStrPassword = lpStrPassword & Chr$(lngPassN(i) / 2)
       
    Next i
      
   
   With GetUserNameFromFile
     .LenPassword = LenPassword
     .LenUserName = LenUserName
     .strPassword = lpStrPassword
     .strUserName = lpstrusername
   End With
    Exit Function
    
ErrorHandler:
   frmPassLock.Hide
    MsgBox "Error:" & Err.Description
   End
   
End Function
Public Function WriteUserNameToFile(ByVal lstrUser As String, lstrPass As String)
On Error GoTo ErrorHandler

Dim lngShuffle As Long
Dim Letter As String

LenUser = Len(lstrUser)
LenPass = Len(lstrPass)

For i = 1 To LenUser
   Letter = Mid$(lstrUser, i, i + 1)
   lngShuffle = Asc(Letter)
   UserName(i) = lngShuffle * 2
Next i

For i = 1 To LenPass
   Letter = Mid$(lstrPass, i, i + 1)
   lngShuffle = Asc(Letter)
   Password(i) = lngShuffle * 2
Next i

Open App.Path + "\" + UserProfile For Binary As #1
  Put #1, , LenUser
  Put #1, , LenPass
  Put #1, , UserName
  Put #1, , Password
Close #1

    Exit Function
   
ErrorHandler:
   frmPassLock.Hide
   
   MsgBox "Error:" & error
   End
End Function

Public Sub ActivateForm()

   
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
     lngFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
       
    
    

    'SetWindowPos frmPassLock.hWnd, HWND_TOPMOST, 0, 0, 0, 0, lngFlags
    frmPassLock.SetFocus
End Sub



Public Sub InfFile()
      Dim UserFile As String
   
   
End Sub
Public Function CheckIfNewProfileIsNeeded() As IsNewProfileNeeded
If Not DoesExist(App.Path + "\" + UserProfile) Then
    MsgBox "No userprofile exists...One will be created, please type a username and password in the fields below", vbApplicationModal
    CheckIfNewProfileIsNeeded = yesitdoes
    Else
    CheckIfNewProfileIsNeeded = NoitDont
End If
End Function
Public Sub Main()
  
Dim DesktopdC As Long
Dim strName As String
Dim lngBuffer As Long
Dim INPNQuestion As IsNewProfileNeeded

App.title = "Extreme Lock Build: 2011" ' - Matta Ratta's Hutt Hutt"
'
LockDown = False
  
'ChDir App.Path
    
'INPNQuestion = CheckIfNewProfileIsNeeded()
'CreateProfile = (INPNQuestion = yesitdoes)
  
'If INPNQuestion <> yesitdoes Then PasswordInMemory = GetUserNameFromFile()
    
lpStrUser = GetUserName
lpStrPassword = " "
     
'strName = String$(255, 0)
'lngBuffer = GetUserName
'FirstPassLoad = False
Load frmPassLock

frmPassLock.txtUser.Text = lpStrUser
frmPassLock.Show
'frmPassLock.Hide

FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\00 Extreme Lock.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
FOLDER_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--"
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FolderExists(FOLDER_CHECK) = False Then
    i = CreateFolderTree(FOLDER_CHECK)
End If



FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
R2$ = Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " Chk"
'R3$ = Format$(frmPassLock.List1.ListCount + 1, "##0000") + "-" + R2$
Print #FR1, R2$
Close #FR1
  
End Sub
