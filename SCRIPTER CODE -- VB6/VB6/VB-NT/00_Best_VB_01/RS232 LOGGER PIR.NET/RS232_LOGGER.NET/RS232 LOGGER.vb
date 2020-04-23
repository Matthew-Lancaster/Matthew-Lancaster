Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic

Friend Class DIALER
    Inherits System.Windows.Forms.Form
    ' ------------------------------------------------------------------------
    ' IF WANT TO DISPLAY TAIL.EXE AT BEGINNER THEN RESTORE THIS LINE BACK INNER
    ' FILE_NAME_4 = "ZZ RS232 FRONT DOOR LOGGER.txt"
    ' ------------------------------------------------------------------------
    ' Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimized
    ' ------------------------------------------------------------------------
    ' SEARCH ANY THIS TEXT CHUNK
    ' [ Friday 09:50:50 Am_26 July 2019 ]
    ' BE NICER IF START MINIMIZED
    ' ------------------------------------------------------------------------
    Dim ARRAY_PORT_RS___01() As String
    Dim ARRAY_PORT_RS_HASH() As String

    Dim FIRST_BOOT_CODE As Object

    Dim FSO As Object
    Dim TIMER_1_TIMER_RUN_ONCE As Object

    Dim COUNT_ERROR_ME_MSComm_PIR_PORTOPEN_PIR As Object
    Dim COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR As Object

    Dim DOOR_OPEN_HAPPEN As Object
    Dim FOLDER_NAME, FILE_NAME_4 As Object
    Dim X2, FILE_NAME, X1, A_NOW As Object
    Dim PROGRAM_LOAD As Object
    Public cProcesses As New clsCnProc
    Dim OLD_VAR_DSR_3, VAR_DSR_3 As Object
    Dim OLD_VAR_DSR_4, VAR_DSR_4 As Object


    Public I_N_TAIL As Object
    Public I_N_NOTEPAD As Object
    Public I_N_AUTOHOTKEY As Object


    Private Const MONITOR_ON As Short = -1
    Private Const MONITOR_OFF As Short = 2
    Private Const SC_MONITORPOWER As Integer = &HF170
    Private Const WM_SYSCOMMAND As Integer = &H112

    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    ' Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Any) As Integer
    ' Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam) As Integer
    ' ----
    ' [Example] Sendmessage API in .NET-VBForums 
    ' http://www.vbforums.com/showthread.php?393556-Example
    ' ----

    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    ' Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As IntPtr, ByVal hWnd2 As IntPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As IntPtr



    Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
    Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer

    Private Declare Function FindWindow2 Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Integer, ByVal lpWindowName As Integer) As Integer

    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
    Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
    Private Declare Function GetParent Lib "user32" (ByVal hWnd As Integer) As Integer
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Integer) As Integer
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer

    Private Const WM_CLOSE As Integer = &H10
    Private Const WM_USER As Integer = &H400
    Private Const WM_COMMAND As Integer = &H111
    Private Const GW_HWNDNEXT As Short = 2

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer


    Private Sub DIALER_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
        'If Me.WindowState <> vbNormal Then End
        '
        'End

    End Sub

    Private Sub DIALER_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'If Me.WindowState <> vbNormal Then End
        '
        'If KeyAscii = 27 Then End

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub GetSerialPortNames()
        '    ' Show all available COM ports.
        '    For Each sp As String In My.Computer.Ports.SerialPortNames
        '        ListBox1.Items.Add (sp)
        '    Next
    End Sub



    Public Sub DIALER_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'If App.PrevInstance = True Then End

        Dim I_N_TAIL

        '' Show all available COM ports.
        Dim AB As String
        Dim AD As Integer
        ReDim Preserve ARRAY_PORT_RS___01(2)
        ReDim Preserve ARRAY_PORT_RS_HASH(2)
        For Each sp As String In My.Computer.Ports.SerialPortNames
            If Hex(sp.GetHashCode) = "B9948936" Then
                ARRAY_PORT_RS___01(1) = Mid(sp, 4)
                ARRAY_PORT_RS_HASH(1) = Hex(sp.GetHashCode)
            End If
        Next
        For Each sp As String In My.Computer.Ports.SerialPortNames
            If Hex(sp.GetHashCode) = "76AB45BD" Then
                ARRAY_PORT_RS___01(2) = Mid(sp, 4)
                ARRAY_PORT_RS_HASH(2) = Hex(sp.GetHashCode)
            End If
        Next

        FIRST_BOOT_CODE = True

        'KILL ITSELF IN __.EXE KILL SOFTLY
        'WHILE ISIDE LEARN
        '---------------------------------
        'Dim VAR As Object
        'Dim PID As Integer
        'If IsIDE() = True Then
        '    PID = -1
        '    'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        '    'UPGRADE_WARNING: Couldn't resolve default property of object VAR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    VAR = cProcesses.GetEXEID(PID, My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe")
        '    If PID <> -1 Then
        '        'Call Process_HIGH_PRIORITY_CLASS(PID)
        '        'UPGRADE_WARNING: Couldn't resolve default property of object VAR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        VAR = cProcesses.Process_Kill(PID)
        '        Beep()
        '        End
        '    End If
        'End If

        ' I_N_TAIL = "C:\PStart\# NOT INSTALL REQUIRED\Tail\Tail.exe"
        ' If Dir(I_N_TAIL) = "" Then
        ' MsgBox "THE EXE FILE __ NOT EXIST" + vbCrLf + vbCrLf + I_N_TAIL
        ' Beep
        ' Exit Sub
        ' End If

        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\#Wave Sounds\DobFig22 01.WAV") Then
            Me.MMControl9.Notify = True
            Me.MMControl9.Wait = True
            Me.MMControl9.Shareable = False
            Me.MMControl9.DeviceType = "waveaudio"
            'ME.MMControl9.FileName = App.Path + "\SG_02_04.WAV"
            Me.MMControl9.FileName = My.Application.Info.DirectoryPath & "\#Wave Sounds\DobFig22 01.WAV"
            Me.MMControl9.Command = "Open"
            Me.MMControl9.Command = "prev"
            '    Me.MMControl9.Command = "Play"
            '    Do
            '    Loop Until MMControl9.Mode = 525
            Me.MMControl9.Command = "prev"
            Me.MMControl9.Command = "Play"
        End If

        FSO = CreateObject("Scripting.FileSystemObject")

        'UPGRADE_WARNING: Couldn't resolve default property of object PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PROGRAM_LOAD = True

        'UPGRADE_WARNING: Couldn't resolve default property of object OLD_VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        OLD_VAR_DSR_3 = -10
        'UPGRADE_WARNING: Couldn't resolve default property of object OLD_VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        OLD_VAR_DSR_4 = -10

        Call TIMER_1_Tick(TIMER_1, New System.EventArgs())

        If IsIDE() = True Then
            '    Me.Visible = True
            ' Me.ShowInTaskbar = True
        End If

        'Me.Visible = True
        '
        'If IsIDE = False Then
        '    Me.WindowState = vbMinimized
        'End If

        If IsIDE() = True Then
            TIMER_PIR.Interval = 1000
            TIMER_FRONT_DOOR.Interval = 1000
        End If

        TIMER_PIR.Enabled = True
        TIMER_FRONT_DOOR.Enabled = True

        On Error Resume Next

        Me.Hide()
        Me.Show()
        'Exit Sub

        'UPGRADE_WARNING: Couldn't resolve default property of object ME_WINDOWSTATE_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If ME_WINDOWSTATE_2 = 2 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ME_WINDOWSTATE_1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Me.WindowState = ME_WINDOWSTATE_1
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object ME_WINDOWSTATE_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If ME_WINDOWSTATE_2 = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ME_WINDOWSTATE_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ME_WINDOWSTATE_2 = 2
            If IsIDE() = True Then
                Me.WindowState = System.Windows.Forms.FormWindowState.Normal
            End If
        End If
        Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Label1.Width) + 400)

    End Sub


    Private Sub DIALER_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        'If Me.WindowState <> vbNormal Then End
        '
        'If KeyCode = 27 Then End
    End Sub

    'UPGRADE_WARNING: Event DIALER.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub DIALER_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        Dim ME_WINDOWSTATE As Object

        'UPGRADE_WARNING: Couldn't resolve default property of object ME_WINDOWSTATE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ME_WINDOWSTATE = Me.WindowState

    End Sub

    Private Sub DIALER_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Me.MMControl9.Command = "Close"

    End Sub

    Sub TIMER_1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TIMER_1.Tick
        Dim MSComm_DOOR_PortOpen As Object
        Dim INDEX As Object
        Dim R2 As Object
        Dim PIR_INDEX As Object
        Dim R As Object

        ' -------------------------------------
        ' IF CHANGE ANY PORT
        ' DON'T FORGET THE OBVIOUS
        ' USB PORT HAS TO BE SET POWER SAVE OFF
        ' USB SETTER NOT RS232 PORT
        ' Thu 30-Jan-2020 19:28:00
        ' -------------------------------------

        ' PIR ------------------ AND NEXT DOOR
        On Error Resume Next
        ' ARRAY_PORT_RS_HASH(INDEX_NUM) = Hex(sp.GetHashCode)
        Err.Clear()

        VAR_DSR_3 = -4
        Using com1 As IO.Ports.SerialPort = My.Computer.Ports.OpenSerialPort("COM" + ARRAY_PORT_RS___01(1), 9600)
            com1.DtrEnable = True
            VAR_DSR_3 = com1.DsrHolding
        End Using

        PIR_INDEX = Val(ARRAY_PORT_RS___01(1))

        ' NEXT PORT IS MY FRONT DOOR OPEN LOGGER
        VAR_DSR_4 = -4
        Err.Clear()
        Using com1 As IO.Ports.SerialPort = My.Computer.Ports.OpenSerialPort("COM" + ARRAY_PORT_RS___01(2), 9600)
            com1.DtrEnable = True
            VAR_DSR_4 = com1.DsrHolding
        End Using

        ' --------------------------------------
        ' NONE COMM PORT ALL 16 TESTER
        ' LEAVE HIGH TO KEEP LIGHT FOR SCREEN ON
        ' --------------------------------------
        If Val(ARRAY_PORT_RS___01(1)) = 0 Then
            VAR_DSR_3 = 0
        Else
            VAR_DSR_3 = True
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object MSComm_DOOR_PortOpen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Me.MSComm_PIR.PortOpen = False Or MSComm_DOOR_PortOpen = False Then
            TIMER_1.Interval = 20000

        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object TIMER_1_TIMER_RUN_ONCE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TIMER_1_TIMER_RUN_ONCE = True

    End Sub

    Private Sub MNU_RESET_FORM_Click()
        Dim EXIT_TRUE As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object EXIT_TRUE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        EXIT_TRUE = True
        Me.Close()
        System.Windows.Forms.Application.DoEvents()
        Reset()
        'UPGRADE_WARNING: Couldn't resolve default property of object EXIT_TRUE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        EXIT_TRUE = False
        'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        ' Load(Form2)
    End Sub


    Private Sub Timer_ERROR_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer_ERROR.Tick
        Dim MSG_1 As Object

        'UPGRADE_WARNING: Couldn't resolve default property of object MSG_1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MSG_1 = ""

        'UPGRADE_WARNING: Couldn't resolve default property of object COUNT_ERROR_ME_MSComm_PIR_PORTOPEN_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If COUNT_ERROR_ME_MSComm_PIR_PORTOPEN_PIR > 100 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object MSG_1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            MSG_1 = "COUNT_ERROR_ME_MSComm_PIR_PORTOPEN_PIR > 100 NOT OPEN PORT" & vbCrLf
            'UPGRADE_WARNING: Couldn't resolve default property of object COUNT_ERROR_ME_MSComm_PIR_PORTOPEN_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object MSG_1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            MSG_1 = MSG_1 + Str(COUNT_ERROR_ME_MSComm_PIR_PORTOPEN_PIR) + vbCrLf + vbCrLf
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR > 100 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object MSG_1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            MSG_1 = MSG_1 + "COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR > 100 NOT OPEN PORT" + vbCrLf
            'UPGRADE_WARNING: Couldn't resolve default property of object COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object MSG_1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            MSG_1 = MSG_1 + Str(COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR) + vbCrLf + vbCrLf
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object MSG_1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If MSG_1 <> "" Then
            '        Me.WindowState = vbNormal
            ' MsgBox MSG_1, vbMsgBoxSetForeground
            ' End
            '        Form2.Timer1.Enabled = True
        End If

    End Sub

    Private Sub TIMER_PIR_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TIMER_PIR.Tick
        Dim FR1 As Object
        Dim RESULT As Object
        Dim R As Object
        Dim VAR_DSR_5 As Object
        Dim K As Object
        Dim TTITTY As Object
        Dim STATE_PIR As Object
        FSO = CreateObject("Scripting.FileSystemObject")

        'UPGRADE_WARNING: Couldn't resolve default property of object TIMER_1_TIMER_RUN_ONCE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If TIMER_1_TIMER_RUN_ONCE = False Then Exit Sub

        On Error Resume Next
        'UPGRADE_WARNING: Couldn't resolve default property of object TTITTY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        TTITTY = "Port " & VB6.Format(Me.MSComm_PIR.CommPort, "00") & " ___ PIR"

        If Me.MSComm_PIR.PortOpen = False Then
            ' Debug.Print "PIR ____ " + Time$ + " Me.MSComm_PIR.PortOpen = False " + TTITTY
            'UPGRADE_WARNING: Couldn't resolve default property of object TTITTY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Label1.Text = "PIR ____ " & TimeString & " -- " + TTITTY
            Label3.Text = "PIR ____ " & TimeString & " -- Me.MSComm_PIR.PortOpen = False"
            '    COUNT_ERROR_ME_MSComm_PIR_PORTOPEN_PIR = COUNT_ERROR_ME_MSComm_PIR_PORTOPEN_PIR + 1
            '    Exit Sub
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        VAR_DSR_3 = Me.MSComm_PIR.DSRHolding

        'UPGRADE_WARNING: Couldn't resolve default property of object K. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        K = ""
        If Me.MSComm_PIR.PortOpen = False Then
            ' --------------------------------------
            ' NONE COMM PORT ALL 16 TESTER
            ' LEAVE HIGH TO KEEP LIGHT FOR SCREEN ON
            ' --------------------------------------
            'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VAR_DSR_3 = True
            'UPGRADE_WARNING: Couldn't resolve default property of object K. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            K = " ARTIFICAL"
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If VAR_DSR_3 = False Then
            'UPGRADE_WARNING: Couldn't resolve default property of object STATE_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            STATE_PIR = " Not Active"
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object STATE_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            STATE_PIR = " Active _____ "
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If VAR_DSR_3 = 0 Then
            VAR_DSR_4 = "FALSE"
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VAR_DSR_4 = "TRUE"
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Me.MSComm_PIR.DSRHolding Then
            VAR_DSR_5 = "TRUE"
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VAR_DSR_5 = "FALSE"
        End If
        ' Debug.Print "PIR ____ " + Time$ + " " + VAR_DSR_4 + STATE_PIR + TTITTY
        'UPGRADE_WARNING: Couldn't resolve default property of object TTITTY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Label1.Text = "PIR _____ " & TimeString & " -- " + TTITTY
        'UPGRADE_WARNING: Couldn't resolve default property of object K. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Label2.Text = "PIR _____ " & TimeString & " -- " + VAR_DSR_5 + K
        'UPGRADE_WARNING: Couldn't resolve default property of object STATE_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Label3.Text = "PIR _____ " & TimeString & " -- " + STATE_PIR

        If Err.Number > 0 Or Err.Number = 8002 Then
            TIMER_1.Enabled = True
            'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VAR_DSR_3 = True
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object OLD_VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If OLD_VAR_DSR_3 = VAR_DSR_3 Then Exit Sub

        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object OLD_VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        OLD_VAR_DSR_3 = VAR_DSR_3

        Dim AR(4) As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(1) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(2) = "\\2-ASUS-EEE\2_ASUS_EEE_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(3) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(4) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_01_C_DRIVE"

        On Error Resume Next
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If VAR_DSR_3 = True Then
            For R = 1 To UBound(AR)
                'UPGRADE_WARNING: Couldn't resolve default property of object FSO.FILEExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If FSO.FILEExists(FILE_NAME_PIR(R)) = False Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object FSO.FOLDERExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If FSO.FOLDERExists(FOLDER_NAME_PIR(R)) = False Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME_PIR(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object RESULT. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        RESULT = CreateFolderTree(FOLDER_NAME_PIR(R))
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FR1 = FreeFile()
                    'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_PIR(R). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FileOpen(FR1, FILE_NAME_PIR(R), OpenMode.Output)
                    'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_PIR(R). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Debug.Print("WRITE " + FILE_NAME_PIR(R))
                    'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FileClose(FR1)
                End If
                ' Debug.Print FILE_NAME_PIR(R)
                ' -----------------------------------------------------
                ' EXAMPLE FILENAME
                ' \\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer #NFS__.txt"
                ' \\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer #NFS__1-ASUS-X5DIJ.txt
                ' -----------------------------------------------------
            Next
        Else
            For R = 1 To UBound(AR)
                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_PIR(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Kill(FILE_NAME_PIR(R))
                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_PIR(R). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Debug.Print("DELETE " + FILE_NAME_PIR(R))
            Next
        End If

    End Sub

    Function FILE_NAME_PIR(ByRef INDEX As Object) As Object
        Dim AR(4) As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(1) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(2) = "\\2-ASUS-EEE\2_ASUS_EEE_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(3) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(4) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object INDEX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME_PIR = "Autokey -- 14-Brightness With Dimmer #NFS__" & Mid(AR(INDEX), 3, InStr(4, AR(INDEX), "\") - 3) & ".txt"
        'UPGRADE_WARNING: Couldn't resolve default property of object INDEX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FOLDER_NAME = AR(INDEX) + "\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY"
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_PIR
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME_PIR = FILE_NAME
    End Function

    Function FOLDER_NAME_PIR(ByRef INDEX As Object) As Object
        Dim AR(4) As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(1) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(2) = "\\2-ASUS-EEE\2_ASUS_EEE_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(3) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(4) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_01_C_DRIVE"
        'UPGRADE_WARNING: Couldn't resolve default property of object INDEX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME_PIR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FOLDER_NAME_PIR = AR(INDEX) + "\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY"
    End Function

    '                FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
    '                FR1 = FreeFile
    '                Open FILE_NAME For Output As #FR1
    '                Close #FR1
    '                Debug.Print "WRITE FILE_NAME " + FILE_NAME
    '                FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_8
    '                FR1 = FreeFile
    '                Open FILE_NAME For Output As #FR1
    '                Debug.Print "WRITE FILE_NAME " + FILE_NAME
    '                Close #FR1
    '
    'FILE_NAME_2 = "RS232 FRONT DOOR.txt"
    'FILE_NAME_8 = "RS232 FRONT DOOR OPEN.txt"
    'FILE_NAME_9 = "RS232 FRONT DOOR CLOSE.txt"
    'FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2


    Sub TIMER_FRONT_DOOR_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TIMER_FRONT_DOOR.Tick
        Dim Path_And_FileName As Object
        Dim NEXT_AFTER_PROGRAM_LOAD As Object
        Dim FR1 As Object
        Dim FileName As Object
        Dim IR As Object
        Dim KILL_DONE As Object
        Dim RESULT As Object
        Dim R As Object
        Dim PATH_2 As Object
        Dim FILE_NAME_9 As Object
        Dim FILE_NAME_8 As Object
        Dim FILE_NAME_2 As Object
        Dim VAR_DSR_5 As Object
        Dim TTITTY As Object

        Dim STRING_VAR As String
        Dim STATE_DOOR As Object

        'UPGRADE_WARNING: Couldn't resolve default property of object TIMER_1_TIMER_RUN_ONCE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If TIMER_1_TIMER_RUN_ONCE = False Then Exit Sub

        TTITTY = "Port " & VB6.Format(ARRAY_PORT_RS___01(1), "00") & " ___ DOOR"


        If Me.MSComm_DOOR.PortOpen = False Then
            '    Label4.Caption = "DOOR ___ " + Time$ + " -- " + TTITTY
            '    Label6.Caption = "DOOR ___ " + Time$ + " -- Me.MSComm_DOOR.PortOpen = False "
            'UPGRADE_WARNING: Couldn't resolve default property of object COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR = COUNT_ERROR_ME_MSComm_DOOR_PORTOPEN_DOOR + 1
            ' Exit Sub

        End If

        If Me.MSComm_DOOR.DSRHolding = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VAR_DSR_4 = True
            'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VAR_DSR_5 = "TRUE"
            'UPGRADE_WARNING: Couldn't resolve default property of object STATE_DOOR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            STATE_DOOR = "Open _______"
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VAR_DSR_5 = "FALSE"
            'UPGRADE_WARNING: Couldn't resolve default property of object STATE_DOOR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            STATE_DOOR = "Close ______"
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object TTITTY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Label4.Text = "DOOR ___ " & TimeString & " -- " + TTITTY
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Label5.Text = "DOOR ___ " & TimeString & " -- " + VAR_DSR_5
        'UPGRADE_WARNING: Couldn't resolve default property of object STATE_DOOR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Label6.Text = "DOOR ___ " & TimeString & " -- " + STATE_DOOR

        If Me.MSComm_DOOR.PortOpen = False Then
            Exit Sub
        End If

        On Error Resume Next
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        VAR_DSR_4 = Me.MSComm_DOOR.DSRHolding
        If Me.MSComm_DOOR.PortOpen = False Then
            'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VAR_DSR_4 = False
            TIMER_1.Enabled = True
        End If
        If Err.Number > 0 Or Err.Number = 8002 Then
            TIMER_1.Enabled = True
            'VAR_DSR_4 = 1
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object OLD_VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If OLD_VAR_DSR_4 = VAR_DSR_4 Then Exit Sub

        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object OLD_VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        OLD_VAR_DSR_4 = VAR_DSR_4

        ' MsgBox Str(R) + " -- " + Str(VAR_DSR_3)
        ' Debug.Print Str(R) + " -- " + Str(VAR_DSR_3)

        Dim AR(1) As Object
        'AR(1) = "\\1-asus-x5dij\1_asus_x5dij_01_c_drive"
        'AR(2) = "\\2-asus-eee\2_asus_eee_01_c_drive"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(1) = "\\4-asus-gl522vw\4_asus_gl522vw_01_c_drive"
        'UPGRADE_WARNING: Couldn't resolve default property of object AR(1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AR(1) = "C:"

        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME_2 = "RS232 FRONT DOOR.txt"
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_8. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME_8 = "RS232 FRONT DOOR OPEN.txt"
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_9. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME_9 = "RS232 FRONT DOOR CLOSE.txt"
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME_4 = "ZZ RS232 FRONT DOOR LOGGER.txt"
        'UPGRADE_WARNING: Couldn't resolve default property of object PATH_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PATH_2 = "VB6\VB-NT\00_Best_VB_01\Tidal_Info"



        On Error Resume Next
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_DSR_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If VAR_DSR_4 = True Then
            For R = 1 To UBound(AR)
                'UPGRADE_WARNING: Couldn't resolve default property of object DOOR_OPEN_HAPPEN. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                DOOR_OPEN_HAPPEN = True
                'UPGRADE_WARNING: Couldn't resolve default property of object PATH_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object AR(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FOLDER_NAME = AR(1) + "\SCRIPTOR DATA\" + PATH_2
                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
                'UPGRADE_WARNING: Couldn't resolve default property of object FSO.FOLDERExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If FSO.FOLDERExists(FOLDER_NAME) = False Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object RESULT. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RESULT = CreateFolderTree(FOLDER_NAME)
                End If
                'If DOOR_OPEN_HAPPEN = True Then
                'UPGRADE_WARNING: Couldn't resolve default property of object KILL_DONE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                KILL_DONE = False
                'UPGRADE_WARNING: Couldn't resolve default property of object STATE_DOOR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If InStr(UCase(STATE_DOOR), "OPEN") > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object FIRST_BOOT_CODE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If FIRST_BOOT_CODE = True Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object FIRST_BOOT_CODE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FIRST_BOOT_CODE = False
                        'UPGRADE_WARNING: Couldn't resolve default property of object FSO.FILEExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If FSO.FILEExists(FILE_NAME) = True Then
                            Me.WindowState = System.Windows.Forms.FormWindowState.Normal
                            'UPGRADE_WARNING: Couldn't resolve default property of object IR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            IR = MsgBox("FILE EXIST FOR FIRST RUN CODE" & vbCrLf & "CONFIRM OTHER CODE NOT LEFT BY ACCIDENT" & vbCrLf & "MIGHT NOT RUN CORRECT" & vbCrLf & "REMAIN OR LEAVE DELETE -- YES OR NOT", MsgBoxStyle.YesNo + MsgBoxStyle.MsgBoxSetForeground)
                            If IR = MsgBoxResult.No Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Kill(FileName)
                                'UPGRADE_WARNING: Couldn't resolve default property of object KILL_DONE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                KILL_DONE = True
                            End If
                        End If
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object KILL_DONE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object STATE_DOOR. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If InStr(UCase(STATE_DOOR), "OPEN") > 0 And KILL_DONE = False Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
                        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FR1 = FreeFile()
                        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FileOpen(FR1, FILE_NAME, OpenMode.Output)
                        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FileClose(FR1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Debug.Print("WRITE FILE_NAME " + FILE_NAME)
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_8. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_8
                        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FR1 = FreeFile()
                        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FileOpen(FR1, FILE_NAME, OpenMode.Output)
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Debug.Print("WRITE FILE_NAME " + FILE_NAME)
                        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FileClose(FR1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
                        'UPGRADE_WARNING: Couldn't resolve default property of object A_NOW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        A_NOW = Now
                        'UPGRADE_WARNING: Couldn't resolve default property of object PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If PROGRAM_LOAD = True Then
                            Call CHECK_ARCHIVE_LOGGER()
                            'UPGRADE_WARNING: Couldn't resolve default property of object PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            PROGRAM_LOAD = False
                            'UPGRADE_WARNING: Couldn't resolve default property of object X2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object X1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If X1 > X2 Then
                                Call WRITE_LOGGER_BEGIN()
                                Call WRITE_LOGGER_OPEN_INFO()
                            End If
                            'UPGRADE_WARNING: Couldn't resolve default property of object NEXT_AFTER_PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            NEXT_AFTER_PROGRAM_LOAD = True
                        End If
                        'UPGRADE_WARNING: Couldn't resolve default property of object NEXT_AFTER_PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If PROGRAM_LOAD = False And NEXT_AFTER_PROGRAM_LOAD = False Then
                            Call WRITE_LOGGER_OPEN_INFO()
                        End If

                        If FindWindow("", "Tidal Information...") = 0 Then

                            Me.MMControl9.Command = "prev"
                            Me.MMControl9.Command = "Play"
                            Do
                            Loop Until MMControl9.Mode = 525
                            Me.MMControl9.Command = "prev"
                            Me.MMControl9.Command = "Play"
                            Do
                            Loop Until MMControl9.Mode = 525
                            Me.MMControl9.Command = "prev"
                            Me.MMControl9.Command = "Play"

                            Shell("D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe", AppWinStyle.MinimizedNoFocus)
                            ' Debug.Print Str(Now)
                            ' TRY MAKE SURE ONLY RUN ONCE TO STARTER -- SEEM OKAY
                        End If
                    End If
                End If
            Next
        Else
            For R = 1 To UBound(AR)
                On Error Resume Next
                'UPGRADE_WARNING: Couldn't resolve default property of object PATH_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object AR(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FOLDER_NAME = AR(1) + "\SCRIPTOR DATA\" + PATH_2
                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Kill(FILE_NAME)

                'UPGRADE_WARNING: Couldn't resolve default property of object DOOR_OPEN_HAPPEN. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If DOOR_OPEN_HAPPEN = True Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object DOOR_OPEN_HAPPEN. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    DOOR_OPEN_HAPPEN = False
                    'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_9. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_9
                    'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FR1 = FreeFile()
                    'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FileOpen(FR1, FILE_NAME, OpenMode.Output)
                    'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Debug.Print("WRITE FILE_NAME " + FILE_NAME)
                    'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FileClose(FR1)
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4

                'UPGRADE_WARNING: Couldn't resolve default property of object A_NOW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                A_NOW = Now
                'UPGRADE_WARNING: Couldn't resolve default property of object PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If PROGRAM_LOAD = True Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    PROGRAM_LOAD = False
                    Call CHECK_ARCHIVE_LOGGER()
                    'UPGRADE_WARNING: Couldn't resolve default property of object X1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object X2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If X2 > X1 Then
                        Call WRITE_LOGGER_BEGIN()
                        Call WRITE_LOGGER_CLOSE_INFO()
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object NEXT_AFTER_PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    NEXT_AFTER_PROGRAM_LOAD = True
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object NEXT_AFTER_PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If PROGRAM_LOAD = False And NEXT_AFTER_PROGRAM_LOAD = False Then
                    Call WRITE_LOGGER_CLOSE_INFO()
                End If

                Me.MMControl9.Command = "prev"
                Me.MMControl9.Command = "Play"
                Do
                Loop Until MMControl9.Mode = 525
                Me.MMControl9.Command = "prev"
                Me.MMControl9.Command = "Play"
                Do
                Loop Until MMControl9.Mode = 525
                Me.MMControl9.Command = "prev"
                Me.MMControl9.Command = "Play"

            Next
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object NEXT_AFTER_PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If NEXT_AFTER_PROGRAM_LOAD = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object NEXT_AFTER_PROGRAM_LOAD. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            NEXT_AFTER_PROGRAM_LOAD = False
            'UPGRADE_WARNING: Couldn't resolve default property of object PATH_2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object AR(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FOLDER_NAME = AR(1) + "\SCRIPTOR DATA\" + PATH_2
            'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
            'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Path_And_FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Path_And_FileName = FILE_NAME
            'UPGRADE_WARNING: Couldn't resolve default property of object Path_And_FileName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If FindWinPart_ANY_STRING("Tail for Win32 - [Non-Workspace Files - " + Path_And_FileName + "]") = 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object I_N_TAIL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                If Dir(I_N_TAIL) <> "" Then


                    ' ------------------------------------------------------------------------
                    ' IF WANT TO DISPLAY TAIL.EXE AT BEGINNER THEN RESTORE THIS LINE BACK INNER
                    ' FILE_NAME_4 = "ZZ RS232 FRONT DOOR LOGGER.txt"
                    ' ------------------------------------------------------------------------
                    'Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimized
                End If
            End If
        End If


    End Sub

    'FOLDER_NAME = AR(1) + "\SCRIPTOR DATA\" + PATH_2
    'FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
    'A_NOW = Now
    'If PROGRAM_LOAD = True Then
    '    Call CHECK_ARCHIVE_LOGGER
    '    PROGRAM_LOAD = False
    '    If X1 > X2 Then
    '        Call WRITE_LOGGER_BEGIN
    '        Call WRITE_LOGGER_OPEN_INFO
    '    End If
    '    NEXT_AFTER_PROGRAM_LOAD = True
    'End If

    Sub WRITE_LOGGER_OPEN_INFO()
        Dim FR1 As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FR1 = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(FR1, FILE_NAME, OpenMode.Append)
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Debug.Print("WRITE FILE_NAME " + FILE_NAME)
        'UPGRADE_WARNING: Couldn't resolve default property of object A_NOW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(FR1, VB6.Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS -- HH:MM:SS AM/PM -- DDD") & " -- DOOR OPEN")
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(FR1)
    End Sub

    Sub WRITE_LOGGER_CLOSE_INFO()
        Dim FR1 As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME_4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FOLDER_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FR1 = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(FR1, FILE_NAME, OpenMode.Append)
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Debug.Print("WRITE FILE_NAME " + FILE_NAME)
        'UPGRADE_WARNING: Couldn't resolve default property of object A_NOW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(FR1, VB6.Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS -- HH:MM:SS AM/PM -- DDD") & " -- DOOR CLOSE")
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(FR1)
    End Sub


    Sub WRITE_LOGGER_BEGIN()
        Dim FR1 As Object
        Dim A2 As Object
        Dim A1 As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object A1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        A1 = " -----------------------------------------"
        'UPGRADE_WARNING: Couldn't resolve default property of object A2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        A2 = " -- RS232 LOGGER FRONT DOOR BEGIN"
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FR1 = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(FR1, FILE_NAME, OpenMode.Append)
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Debug.Print("WRITE FILE_NAME " + FILE_NAME)
        'UPGRADE_WARNING: Couldn't resolve default property of object A1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object A_NOW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(FR1, VB6.Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS") & " - -" & VB6.Format(A_NOW, "HH:MM:SS AMPM -- DDD") + A1)
        'UPGRADE_WARNING: Couldn't resolve default property of object A2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object A_NOW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(FR1, VB6.Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS") & " - -" & VB6.Format(A_NOW, "HH:MM:SS AMPM -- DDD") + A2)
        'UPGRADE_WARNING: Couldn't resolve default property of object A1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object A_NOW. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(FR1, VB6.Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS") & " - -" & VB6.Format(A_NOW, "HH:MM:SS AMPM -- DDD") + A1)
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(FR1)
    End Sub

    Sub CHECK_ARCHIVE_LOGGER()
        Dim LOF_MINUS As Object
        Dim FR1 As Object

        Dim STRING_VAR As String
        STRING_VAR = Space(2000)
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FR1 = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FILE_NAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(FR1, FILE_NAME, OpenMode.Binary)
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object LOF_MINUS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        LOF_MINUS = LOF(FR1) - 2000
        'UPGRADE_WARNING: Couldn't resolve default property of object LOF_MINUS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If LOF_MINUS < 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object LOF_MINUS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            LOF_MINUS = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            STRING_VAR = CStr(LOF(FR1))
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object LOF_MINUS. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        FileGet(FR1, STRING_VAR, LOF_MINUS)
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(FR1)
        'UPGRADE_WARNING: Couldn't resolve default property of object X1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X1 = InStrRev(STRING_VAR, "-- DOOR CLOSE")
        'UPGRADE_WARNING: Couldn't resolve default property of object X2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X2 = InStrRev(STRING_VAR, "-- DOOR OPEN")

        STRING_VAR = ""

    End Sub

    Function FindWinPart_ANY_STRING(ByRef TTF As Object) As Integer
        Dim ghj As String

        Dim Window_Title_String As Object
        Dim test_pid, test_hwnd, test_thread_id As Integer

        Dim cText As String

        FindWinPart_ANY_STRING = 0

        'Need this is you gonna use this procedure get from CIDRun ME and another one
        'Find the first window
        test_hwnd = FindWindow2(0, 0)

        Do While test_hwnd <> 0

            'UPGRADE_WARNING: Couldn't resolve default property of object Window_Title_String. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Window_Title_String = GetWindowTitle(test_hwnd)
            cText = Space(255)
            ghj = CStr(GetClassName(test_hwnd, cText, 255))

            'UPGRADE_WARNING: Couldn't resolve default property of object TTF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Window_Title_String. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If InStr(Window_Title_String, TTF) Then

                FindWinPart_ANY_STRING = test_hwnd
                Exit Function

            End If

            'retrieve the next window
            test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

        Loop

    End Function



    Function StripNulls(ByRef OriginalStr As String) As String
        'Removes NullStrings.
        If (InStr(OriginalStr, Chr(0)) > 0) Then
            OriginalStr = VB.Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
    End Function



    'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
    'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

    Function GetUserName() As String
        Dim UserName As New VB6.FixedLengthString(255)
        Call GetUserNameA(UserName.Value, 255)
        GetUserName = VB.Left(UserName.Value, InStr(UserName.Value, Chr(0)) - 1)
    End Function

    Function GetComputerName() As String
        Dim UserName As New VB6.FixedLengthString(255)
        Call GetComputerNameA(UserName.Value, 255)
        GetComputerName = VB.Left(UserName.Value, InStr(UserName.Value, Chr(0)) - 1)
    End Function




    Public Function CreateFolderTree(ByVal sPath As String) As Boolean
        Dim R As Object
        Dim nPos As Short

        On Error GoTo CreateFolderTreeError

        nPos = InStr(sPath, "\")
        If Mid(sPath, 1, 2) = "\\" Then
            nPos = 2
            For R = 1 To 3
                nPos = InStr(nPos + 1, sPath, "\")
            Next
        End If
        While nPos > 0

            '----------------------------------------------------------------------------
            'ROUTINE TAKEN FROM
            '----------------------------------------------------------------------------
            'SEND_TO_SCRIPT_IRFAR - Microsoft Visual Basic [design] - [mdlFileSys (Code)]
            'D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAR\#0 Send To Text List of Files and Sub Folders IRFAN.exe
            '----------------------------------------------------------------------------
            'MODIFIED A BIT DIR COMMAND REPLACE MORE COMPLEX WAY
            '----------------------------------------------------------------------------

            'If Not FolderExists(Left$(sPath, nPos - 1)) Then

            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            If Dir(VB.Left(sPath, nPos - 1), FileAttribute.Directory) = "" Then
                MkDir(VB.Left(sPath, nPos - 1))
            End If
            nPos = InStr(nPos + 1, sPath, "\")
        End While
        'If Not FolderExists(sPath) Then MkDir sPath
        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir(sPath, FileAttribute.Directory) = "" Then MkDir(sPath)

        CreateFolderTree = True
        Exit Function

CreateFolderTreeError:
        Exit Function
    End Function


    Function GetActiveWindow(ByVal ReturnParent As Boolean) As Integer
        Dim GetForegroundWindow As Object
        Dim i As Integer
        Dim j As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object GetForegroundWindow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = GetForegroundWindow
        If ReturnParent Then
            Do While i <> 0
                j = i
                i = GetParent(i)
            Loop
            i = j
        End If
        GetActiveWindow = i
    End Function

    'UPGRADE_NOTE: Handle was upgraded to Handle_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Function GetParentTitle(ByVal Handle_Renamed As Integer) As String
        Dim i As Integer
        Dim j As Integer
        Dim TTx1 As String
        i = Handle_Renamed
        Do While i <> 0
            j = i
            i = GetParent(i)
        Loop
        i = j
        TTx1 = GetWindowTitle(i)
        GetParentTitle = TTx1

    End Function

    Function GetWindowTitle(ByVal hWnd As Integer) As String
        Dim L As Integer
        Dim S As String
        L = GetWindowTextLength(hWnd)
        S = Space(L + 1)
        GetWindowText(hWnd, S, L + 1)
        GetWindowTitle = VB.Left(S, L)
    End Function
    Function GetWindowClass(ByVal hWnd As Integer) As String
        Dim Ret As Integer
        Dim sText As String
        sText = Space(255)
        Ret = GetClassName(hWnd, sText, 255)
        sText = VB.Left(sText, Ret)
        GetWindowClass = sText
    End Function



    '***********************************************
    '# Check, whether we are in the IDE
    Function IsIDE() As Boolean
        'IsIDE = False
        'Exit Function
        System.Diagnostics.Debug.Assert(Not TestIDE(IsIDE), "")
    End Function
    Private Function TestIDE(ByRef Test As Boolean) As Boolean
        Test = True
    End Function
    '***********************************************



    Private Sub Timer_TICKER_CHAIN_DOG_SHOW_RUN_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer_TICKER_CHAIN_DOG_SHOW_RUN.Tick
        Dim FR1 As Object
        Dim VAR_FILENAME As Object
        Dim RUN_HERE_ONCE As Object
        If Timer_TICKER_CHAIN_DOG_SHOW_RUN.Interval <> 2000 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object RUN_HERE_ONCE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RUN_HERE_ONCE = True
            Timer_TICKER_CHAIN_DOG_SHOW_RUN.Interval = 2000
        End If

        On Error Resume Next
        'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_FILENAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        VAR_FILENAME = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & "_APP_RUNNER_TICKER_#NFS_EX.TXT"

        ' If IsIDE = True And RUN_HERE_ONCE Then Debug.Print VAR_FILENAME

        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FR1 = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object VAR_FILENAME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(FR1, VAR_FILENAME, OpenMode.Output)
        'UPGRADE_WARNING: Couldn't resolve default property of object FR1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(FR1)
        On Error GoTo 0
    End Sub

    Private Sub Timer20_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer20.Tick
        Exit Sub
        Call MNU_RESET_FORM_Click()
    End Sub
End Class