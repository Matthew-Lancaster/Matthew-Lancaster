VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5916
   ClientLeft      =   1140
   ClientTop       =   1512
   ClientWidth     =   6516
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5916
   ScaleWidth      =   6516
   Begin RichTextLib.RichTextBox RTB_Txt_Keys 
      Height          =   1176
      Left            =   48
      TabIndex        =   9
      Top             =   2772
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   2074
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "HKEY_USERS"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Tag             =   "&H80000003"
      Top             =   960
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "HKEY_LOCAL_MACHINE"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Tag             =   "&H80000002"
      Top             =   720
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "HKEY_CURRENT_USER"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Tag             =   "&H80000001"
      Top             =   480
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "HKEY_CURRENT_MACHINE"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Tag             =   "&H80000005"
      Top             =   240
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "HKEY_CLASSES_ROOT"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Tag             =   "&H80000000"
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox txtKeys 
      Height          =   960
      Left            =   12
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1656
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Key:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  =============================================================
'# __ D:\VB6\VB-NT\00_Best_VB_01\READ REGISTRY KEY AND SUB\Get_Installed_App_Set.vbp
'# __
'# __ VB6 -- Get_Installed_App_Set.vbp
'#
'# BY Matthew __ Matt.Lan@Btinternet.com __
'# __
'# START       TIME [ Thu 31-May-2018 02:02:42 ]
'# TO          TIME [ Thu 31-May-2018 02:28:00 ]
'# __
'  =============================================================

'# ------------------------------------------------------------------
'# ------------------------------------------------------------------

'# ------------------------------------------------------------------
'SOURCE CODE CREDIT AUTHOR
'# ------------------------------------------------------------------
'----
'VB Helper: HowTo: Enumerate registry keys, subkeys, and values
'http://www.vb-helper.com/howto_list_registry_keys_and_values.html
'----
'# ------------------------------------------------------------------



Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Private Const ERROR_SUCCESS = 0&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = _
    ((STANDARD_RIGHTS_ALL Or _
    KEY_QUERY_VALUE Or _
    KEY_SET_VALUE Or _
    KEY_CREATE_SUB_KEY Or _
    KEY_ENUMERATE_SUB_KEYS Or _
    KEY_NOTIFY Or KEY_CREATE_LINK) And _
    (Not SYNCHRONIZE))
Private Const ERROR_NO_MORE_ITEMS = 259&

Private m_SelectedSection As Long

Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_RESOURCE_LIST = 8
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Private Const REG_SZ = 1
' Get the key information for this key and
' its subkeys.
Private Function GetKeyInfo(ByVal section As Long, ByVal key_name As String, ByVal indent As Integer) As String
Dim subkeys As Collection
Dim subkey_values As Collection
Dim subkey_num As Integer
Dim subkey_name As String
Dim subkey_value As String
Dim length As Long
Dim hKey As Long
Dim txt As String
Dim subkey_txt As String
Dim value_num As Long
Dim value_name_len As Long
Dim value_name As String
Dim reserved As Long
Dim value_type As Long
Dim value_string As String
Dim value_data(1 To 1024) As Byte
Dim value_data_len As Long
Dim i As Integer

    Set subkeys = New Collection
    Set subkey_values = New Collection
    
    ' Open the key.
    If RegOpenKeyEx(section, _
        key_name, _
        0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS _
    Then
        MsgBox "Error opening key."
        Exit Function
    End If

    ' Enumerate the key's values.
    value_num = 0
    Do
        value_name_len = 1024
        value_name = Space$(value_name_len)
        value_data_len = 1024

        If RegEnumValue(hKey, value_num, _
            value_name, value_name_len, 0, _
            value_type, value_data(1), value_data_len) _
                <> ERROR_SUCCESS Then Exit Do

        value_name = Left$(value_name, value_name_len)

        Select Case value_type
            Case REG_BINARY
                txt = txt & Space$(indent) & "> " & value_name & " = [binary]" & vbCrLf
            Case REG_DWORD
                value_string = "&H" & _
                    Format$(Hex$(value_data(4)), "00") & _
                    Format$(Hex$(value_data(3)), "00") & _
                    Format$(Hex$(value_data(2)), "00") & _
                    Format$(Hex$(value_data(1)), "00")
                txt = txt & Space$(indent) & "> " & value_name & " = " & value_string & vbCrLf
            Case REG_DWORD_BIG_ENDIAN
                txt = txt & Space$(indent) & "> " & value_name & " = [dword big endian]" & vbCrLf
            Case REG_DWORD_LITTLE_ENDIAN
                txt = txt & Space$(indent) & "> " & value_name & " = [dword little endian]" & vbCrLf
            Case REG_EXPAND_SZ
                txt = txt & Space$(indent) & "> " & value_name & " = [expand sz]" & vbCrLf
            Case REG_FULL_RESOURCE_DESCRIPTOR
                txt = txt & Space$(indent) & "> " & value_name & " = [full resource descriptor]" & vbCrLf
            Case REG_LINK
                txt = txt & Space$(indent) & "> " & value_name & " = [link]" & vbCrLf
            Case REG_MULTI_SZ
                txt = txt & Space$(indent) & "> " & value_name & " = [multi sz]" & vbCrLf
            Case REG_NONE
                txt = txt & Space$(indent) & "> " & value_name & " = [none]" & vbCrLf
            Case REG_RESOURCE_LIST
                txt = txt & Space$(indent) & "> " & value_name & " = [resource list]" & vbCrLf
            Case REG_RESOURCE_REQUIREMENTS_LIST
                txt = txt & Space$(indent) & "> " & value_name & " = [resource requirements list]" & vbCrLf
            Case REG_SZ
                value_string = ""
                For i = 1 To value_data_len - 1
                    value_string = value_string & Chr$(value_data(i))
                Next i
                txt = txt & Space$(indent) & "> " & value_name & " = """ & value_string & """" & vbCrLf
        End Select

        value_num = value_num + 1
    Loop

    ' Enumerate the subkeys.
    subkey_num = 0
    Do
        ' Enumerate subkeys until we get an error.
        length = 256
        subkey_name = Space$(length)
        If RegEnumKey(hKey, subkey_num, _
            subkey_name, length) _
                <> ERROR_SUCCESS Then
                Exit Do
                End If
        subkey_num = subkey_num + 1
        
        subkey_name = Left$(subkey_name, InStr(subkey_name, Chr$(0)) - 1)
        subkeys.Add subkey_name
    
        ' Get the subkey's value.
        length = 256
        subkey_value = Space$(length)
        If RegQueryValue(hKey, subkey_name, _
            subkey_value, length) _
            <> ERROR_SUCCESS _
        Then
            subkey_values.Add "Error"
        Else
            ' Remove the trailing null character.
            subkey_value = Left$(subkey_value, length - 1)
            subkey_values.Add subkey_value
        End If
    Loop
    
    ' Close the key.
    If RegCloseKey(hKey) <> ERROR_SUCCESS Then
        MsgBox "Error closing key."
    End If

    ' Recursively get information on the keys.
    For subkey_num = 1 To subkeys.Count
        subkey_txt = GetKeyInfo(section, key_name & "\" & subkeys(subkey_num), indent + 4)
        Debug.Print subkeys(subkey_num)
        
        'If subkeys(subkey_num) = "{FF66E9F6-83E7-3A3E-AF14-8DE9A809A6A4}" Then Stop
        
        txt = txt & Format(subkey_num, "0000 ") & String(70, "-") & vbCrLf & Space(indent) & _
            subkeys(subkey_num) & _
            ": " & subkey_values(subkey_num) & vbCrLf & _
            subkey_txt
    Next subkey_num

    GetKeyInfo = txt
End Function

Private Sub Command1_Click()

    'txtKeys.Text = Text1.Text & vbCrLf & _
        GetKeyInfo(m_SelectedSection, Text1.Text, 4)

    RTB_Txt_Keys.Text = Text1.Text & vbCrLf & _
        GetKeyInfo(m_SelectedSection, Text1.Text, 4)



End Sub

Private Sub Form_Load()
    Option1(2).Value = True
    Option1(3).Value = True
    Text1.Text = "RemoteAccess"
    Text1.Text = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
    'HKEY_LOCAL_MACHINE
    Text1.Text = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
End Sub

Private Sub Form_Resize()
Dim hgt As Single
Dim wid As Single

    hgt = ScaleHeight - txtKeys.Top
    If hgt < 120 Then hgt = 120
    txtKeys.Move 0, txtKeys.Top, _
        ScaleWidth, hgt

    wid = ScaleWidth - Text1.Left
    If wid < 120 Then wid = 120
    Text1.Width = wid
    
    RTB_Txt_Keys.Top = txtKeys.Top
    
    hgt = ScaleHeight - RTB_Txt_Keys.Top
    If hgt < 120 Then hgt = 120
    RTB_Txt_Keys.Move 0, RTB_Txt_Keys.Top, _
        ScaleWidth, hgt
    
    
End Sub


Private Sub Option1_Click(Index As Integer)
    ' Save the selected section number.
    m_SelectedSection = CLng(Option1(Index).Tag)
End Sub


