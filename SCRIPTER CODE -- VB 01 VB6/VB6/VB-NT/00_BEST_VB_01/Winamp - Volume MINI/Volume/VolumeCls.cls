Imports System.Runtime.InteropServices

Public Class Sound

#Region "Comments"
    ' This class is responisble for all interactions for muting and controlling the master volume level

    ' There is no warranty associated with this code - if you use this and it blows up you machine.
    ' Well, that will teach you to write you own code in future, won't it ;-)

    ' Mark Dryden (aka Drydo@vbcity.com)
#End Region

#Region "Constants"
    Private Const MMSYSERR_NOERROR As Integer = 0
    Private Const MAXPNAMELEN As Integer = 32
    Private Const MIXER_LONG_NAME_CHARS As Integer = 64
    Private Const MIXER_SHORT_NAME_CHARS As Integer = 16
    Private Const MIXERCONTROL_CT_CLASS_FADER As Integer = &H50000000
    Private Const MIXERCONTROL_CT_UNITS_UNSIGNED As Integer = &H30000
    Private Const MIXERCONTROL_CT_UNITS_BOOLEAN As Integer = &H10000
    Private Const MIXERCONTROL_CT_CLASS_SWITCH As Integer = &H20000000
    Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST As Integer = &H0&
    Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS As Integer = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
    Private Const MIXERCONTROL_CONTROLTYPE_FADER As Integer = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
    Private Const MIXERCONTROL_CONTROLTYPE_VOLUME As Integer = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
    Private Const MIXER_GETLINEINFOF_COMPONENTTYPE As Integer = &H3&
    Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE As Integer = &H2
    Private Const MIXERCONTROL_CONTROLTYPE_BASS As Integer = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
    Private Const MIXERCONTROL_CONTROLTYPE_TREBLE As Integer = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
    Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER As Integer = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
    Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN As Integer = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_UNITS_BOOLEAN)
    Private Const MIXERCONTROL_CONTROLTYPE_MUTE As Integer = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
#End Region

#Region "Structs"

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERCONTROL
        <FieldOffset(0)> Public cbStruct As Integer           '  size in Byte of MIXERCONTROL
        <FieldOffset(4)> Public dwControlID As Integer        '  unique control id for mixer device
        <FieldOffset(8)> Public dwControlType As Integer      '  MIXERCONTROL_CONTROLTYPE_xxx
        <FieldOffset(12)> Public fdwControl As Integer         '  MIXERCONTROL_CONTROLF_xxx
        <FieldOffset(16)> Public cMultipleItems As Integer     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
        <FieldOffset(20), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MIXER_SHORT_NAME_CHARS)> Public szShortName As String ' * MIXER_SHORT_NAME_CHARS  ' short name of control
        <FieldOffset(36), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MIXER_LONG_NAME_CHARS)> Public szName As String '  * MIXER_LONG_NAME_CHARS ' Integer name of control
        <FieldOffset(100)> Public lMinimum As Integer           '  Minimum value
        <FieldOffset(104)> Public lMaximum As Integer           '  Maximum value
        <FieldOffset(108), MarshalAs(UnmanagedType.ByValArray, SizeConst:=11, ArraySubType:=UnmanagedType.AsAny)> Public reserved() As Integer      '  reserved structure space
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERCONTROLDETAILS
        <FieldOffset(0)> Public cbStruct As Integer       '  size in Byte of MIXERCONTROLDETAILS
        <FieldOffset(4)> Public dwControlID As Integer    '  control id to get/set details on
        <FieldOffset(8)> Public cChannels As Integer      '  number of channels in paDetails array
        <FieldOffset(12)> Public item As Integer           '  hwndOwner or cMultipleItems
        <FieldOffset(16)> Public cbDetails As Integer      '  size of _one_ details_XX struct
        <FieldOffset(20)> Public paDetails As IntPtr       '  pointer to array of details_XX structs
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERCONTROLDETAILS_UNSIGNED
        <FieldOffset(0)> Public dwValue As Integer        '  value of the control
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERLINECONTROLS
        <FieldOffset(0)> Public cbStruct As Integer       '  size in Byte of MIXERLINECONTROLS
        <FieldOffset(4)> Public dwLineID As Integer       '  line id (from MIXERLINE.dwLineID)
        <FieldOffset(8)> Public dwControl As Integer      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
        <FieldOffset(12)> Public cControls As Integer      '  count of controls pmxctrl points to
        <FieldOffset(16)> Public cbmxctrl As Integer       '  size in Byte of _one_ MIXERCONTROL
        <FieldOffset(20)> Public pamxctrl As IntPtr       '  pointer to first MIXERCONTROL array
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure MIXERLINE
        <FieldOffset(0)> Public cbStruct As Integer                '  size of MIXERLINE structure
        <FieldOffset(4)> Public dwDestination As Integer          '  zero based destination index
        <FieldOffset(8)> Public dwSource As Integer               '  zero based source index (if source)
        <FieldOffset(12)> Public dwLineID As Integer               '  unique line id for mixer device
        <FieldOffset(16)> Public fdwLine As Integer                '  state/information about line
        <FieldOffset(20)> Public dwUser As Integer                 '  driver specific information
        <FieldOffset(24)> Public dwComponentType As Integer        '  component type line connects to
        <FieldOffset(28)> Public cChannels As Integer              '  number of channels line supports
        <FieldOffset(32)> Public cConnections As Integer           '  number of connections (possible)
        <FieldOffset(36)> Public cControls As Integer              '  number of controls at this line
        <FieldOffset(40), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MIXER_SHORT_NAME_CHARS)> Public szShortName As String  ' * MIXER_SHORT_NAME_CHARS
        <FieldOffset(56), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MIXER_LONG_NAME_CHARS)> Public szName As String ' * MIXER_LONG_NAME_CHARS
        <FieldOffset(120)> Public dwType As Integer
        <FieldOffset(124)> Public dwDeviceID As Integer
        <FieldOffset(128)> Public wMid As Integer
        <FieldOffset(132)> Public wPid As Integer
        <FieldOffset(136)> Public vDriverVersion As Integer
        <FieldOffset(168), MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst:=MAXPNAMELEN)> Public szPname As String ' * MAXPNAMELEN
    End Structure

#End Region

#Region "API Declarations"

    Private Declare Function mixerOpen Lib "winmm.dll" _
    (ByRef phmx As Integer, <MarshalAs(UnmanagedType.U4)> ByVal uMxId As Integer, ByVal dwCallback As Integer, ByVal dwInstance As Integer, ByVal fdwOpen As Integer) As Integer

    Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" _
        (<MarshalAs(UnmanagedType.I4)> ByVal hmxobj As Integer, ByRef pmxl As MIXERLINE, ByVal fdwInfo As Integer) As Integer

    Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" _
        (<MarshalAs(UnmanagedType.I4)> ByVal hmxobj As Integer, ByRef pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Integer) As Integer

    Private Declare Function mixerSetControlDetails Lib "winmm.dll" (<MarshalAs(UnmanagedType.I4)> ByVal hmxobj As Integer, _
        ByRef pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Integer) As Integer

#End Region

    Public Sub SetVolume(ByVal Level As Integer)
        ' Sets the volume to a specific percentage as passed through
        Dim hmixer As Integer
        Dim volCtrl As New MIXERCONTROL
        Dim lngReturn As Integer
        Dim lngVolSetting As Integer

        ' Obtain the hmixer struct
        lngReturn = mixerOpen(hmixer, 0, 0, 0, 0)

        ' Error check
        If lngReturn <> 0 Then Exit Sub

        ' Obtain the volumne control object
        Call GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
            MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)

        ' Then determine the value of the volume
        lngVolSetting = CType(volCtrl.lMaximum * (Level / 100), Integer)

        ' Then set the volume
        SetVolumeControl(hmixer, volCtrl, lngVolSetting)
    End Sub

    Public Sub SetSound(ByVal boolMute As Boolean)
        ' This routine sets the volume setting of the current unit depending on the value passed through
        Dim hmixer As Integer
        Dim volCtrl As New MIXERCONTROL
        Dim lngReturn As Integer
        Dim lngVolSetting As Integer

        ' Obtain the hmixer struct
        lngReturn = mixerOpen(hmixer, 0, 0, 0, 0)

        ' Error check
        If lngReturn <> 0 Then Exit Sub

        ' Obtain the volumne control object
        Call GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
            MIXERCONTROL_CONTROLTYPE_MUTE, volCtrl)

        ' Then determine the value of the volume
        If boolMute Then
            ' Mute
            lngVolSetting = 1
        Else
            ' Turn the sound on
            lngVolSetting = 0
        End If

        ' Then set the volume
        SetVolumeControl(hmixer, volCtrl, lngVolSetting)
    End Sub

    Private Function GetVolumeControl(ByVal hmixer As Integer, ByVal componentType As Integer, ByVal ctrlType As Integer, _
        ByRef mxc As MIXERCONTROL) As Boolean
        ' Obtains an appropriate pointer and info for the volume control

        ' [Note: original source taken from MSDN http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q178456&]

        ' This function attempts to obtain a mixer control. Returns True if successful.
        Dim mxlc As New MIXERLINECONTROLS
        Dim mxl As New MIXERLINE
        Dim rc As Integer, pmem As IntPtr

        mxl.cbStruct = Marshal.SizeOf(mxl)
        mxl.dwComponentType = componentType

        ' Obtain a line corresponding to the component type
        rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)

        If (MMSYSERR_NOERROR = rc) Then
            mxlc.cbStruct = Marshal.SizeOf(mxlc)
            mxlc.dwLineID = mxl.dwLineID
            mxlc.dwControl = ctrlType
            mxlc.cControls = 1
            mxlc.cbmxctrl = Marshal.SizeOf(mxc)

            ' Allocate a buffer for the control
            pmem = Marshal.AllocHGlobal(Marshal.SizeOf(mxc))
            mxlc.pamxctrl = pmem

            mxc.cbStruct = Marshal.SizeOf(mxc)

            ' Get the control
            rc = mixerGetLineControls(hmixer, _
                                      mxlc, _
                                      MIXER_GETLINECONTROLSF_ONEBYTYPE)

            If (MMSYSERR_NOERROR = rc) Then

                mxc = CType(Marshal.PtrToStructure(mxlc.pamxctrl, GetType(MIXERCONTROL)), MIXERCONTROL)

                Marshal.FreeHGlobal(pmem)
                Return True
            End If
            Marshal.FreeHGlobal(pmem)
            Exit Function
        End If

        Return False
    End Function

    Private Function SetVolumeControl(ByVal hmixer As Integer, _
    ByVal mxc As MIXERCONTROL, _
                            ByVal volume As Integer) As Boolean
        ' Sets the volumne from the pointer of the object passed through

        ' [Note: original source taken from MSDN http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q178456&]

        'This function sets the value for a volume control. Returns True if successful

        Dim mxcd As MIXERCONTROLDETAILS
        Dim vol As MIXERCONTROLDETAILS_UNSIGNED
        Dim rc As Integer

        Dim hptr As IntPtr

        mxcd.item = 0
        mxcd.dwControlID = mxc.dwControlID
        mxcd.cbStruct = Marshal.SizeOf(mxcd)
        mxcd.cbDetails = Marshal.SizeOf(vol)

        hptr = Marshal.AllocHGlobal(Marshal.SizeOf(vol))

        ' Allocate a buffer for the control value buffer
        mxcd.paDetails = hptr
        mxcd.cChannels = 1
        vol.dwValue = volume

        Marshal.StructureToPtr(vol, hptr, False)

        ' Set the control value
        rc = mixerSetControlDetails(hmixer, _
                                   mxcd, _
                                   0)

        Marshal.FreeHGlobal(hptr)

        If (MMSYSERR_NOERROR = rc) Then
            Return True
        Else
            Return False
        End If
    End Function
End Class