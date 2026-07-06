Attribute VB_Name = "Module1"
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetFocuses Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

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

Public Function IsPasswordCorrect(ByVal lpPassword As String, lpUserName) As IsNewProfileNeeded
Dim NullReturn As Variant
If CreateProfile = True Then
 NullReturn = WriteUserNameToFile(lpUserName, lpPassword)
 IsPasswordCorrect = yesitdoes
Exit Function
End If
   
   If lpPassword <> PasswordInMemory.strPassword Then
   IsPasswordCorrect = NoitDont
   Else
   IsPasswordCorrect = yesitdoes
   End If

   
   
End Function
Public Function DoesExist(ByVal lstrQuery As String) As Boolean
DoesExist = (Dir(lstrQuery) <> "")


End Function

Public Function GetUserNameFromFile() As InfoRetrieved
On Error GoTo ErrorHandler
Dim LenUserName As Long
Dim LenPassword As Long
Dim lngUserN(100) As Long
Dim lngPassN(100) As Long

Dim lpStrUser, lpStrPassword, Letter As String
Dim lngShuffle As Long
    Open UserProfile For Binary As #1
    
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
   frmPassLock.Visible = False
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

Open UserProfile For Binary As #1
  Put #1, , LenUser
  Put #1, , LenPass
  Put #1, , UserName
  Put #1, , Password
Close #1

    Exit Function
   
ErrorHandler:
   frmPassLock.Visible = False
   
   MsgBox "Error:" & Error
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
       
    
    

    SetWindowPos frmPassLock.hWnd, HWND_TOPMOST, 0, 0, 0, 0, lngFlags
End Sub



Public Sub InfFile()
      Dim UserFile As String
   
   
End Sub
Public Function CheckIfNewProfileIsNeeded() As IsNewProfileNeeded
If Not DoesExist(UserProfile) Then
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
  App.Title = "Extreme Lock Build: 2011"
  INPNQuestion = CheckIfNewProfileIsNeeded()
  CreateProfile = (INPNQuestion = yesitdoes)
  
  If INPNQuestion <> yesitdoes Then PasswordInMemory = GetUserNameFromFile()
     
     
  strName = String$(255, 0)
  lngBuffer = GetUserName(strName, Len(strName))
  Load frmPassLock
  frmPassLock.txtUser.Text = strName
  frmPassLock.Show
 
  
End Sub
