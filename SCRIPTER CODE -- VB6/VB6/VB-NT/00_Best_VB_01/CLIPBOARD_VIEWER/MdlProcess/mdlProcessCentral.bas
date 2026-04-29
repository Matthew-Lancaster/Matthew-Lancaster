Attribute VB_Name = "PROCESSCENTRAL"
'Attribute VB_Name = "mdlProcess"
'COMMON CODE



'ADD COMPONENT
'2ND ON DOWN OF COMMON CONTROLS
'Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
'ENABLE USE OF LIST VIEW

'FOR USE OF KILL COMMAND ONLY COULD BE EASIER


Option Explicit



Public Idle_Timer_Proc

Private cProcesses As New clsCnProc

'// 60% of this constants are not used by program, mostly because this is free-source
'// version, many of the stuffs here do not exist in msdn or api viewer, so I've left them
'// so you could use them, if you need.

Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, hModule As Long, ByVal cb As Long, cbNeeded As Long) As Long




Public process() As PROCESSENTRY32
Public Module() As MODULEENTRY32
Public Thread() As THREADENTRY32
Public ProcessID(0 To 999) As Long

Public Const DELETE As Long = &H10000
Public Const READ_CONTROL As Long = &H20000
Public Const WRITE_DAC As Long = &H40000
Public Const WRITE_OWNER As Long = &H80000
Public Const SYNCHRONIZE As Long = &H100000
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL As Long = &HFFFF
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const GENERIC_EXECUTE As Long = &H20000000
Public Const GENERIC_ALL As Long = &H10000000

Public Const MAX_PATH As Long = 260


Public Const EXCEPTION_NONCONTINUABLE As Long = &H1
Public Const EXCEPTION_MAXIMUM_PARAMETERS As Long = 15


Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_DEVICE As Long = &H40
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Public Const FILE_ATTRIBUTE_SPARSE_FILE As Long = &H200
Public Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400
Public Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Public Const FILE_ATTRIBUTE_OFFLINE As Long = &H1000
Public Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED As Long = &H2000
Public Const FILE_ATTRIBUTE_ENCRYPTED As Long = &H4000

Public Const FILE_SHARE_READ As Long = &H1
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const FILE_SHARE_DELETE As Long = &H4

Public Const IMAGE_DOS_SIGNATURE As Long = &H4D5A
Public Const IMAGE_OS2_SIGNATURE As Long = &H4E45
Public Const IMAGE_OS2_SIGNATURE_LE As Long = &H4C45
Public Const IMAGE_NT_SIGNATURE As Long = &H50450000

Public Const IMAGE_SIZEOF_FILE_HEADER As Long = 20


Public Const KEY_CREATE_LINK As Long = &H20
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_EVENT As Long = &H1
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_WRITE As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE As Long = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_WOW64_32KEY As Long = &H200
Public Const KEY_WOW64_64KEY As Long = &H100
Public Const KEY_WOW64_RES As Long = &H300

Public Const MAXIMUM_PROCESSORS As Long = 32

Public Const PROCESS_TERMINATE As Long = &H1
Public Const PROCESS_CREATE_THREAD As Long = &H2
Public Const PROCESS_SET_SESSIONID As Long = &H4
Public Const PROCESS_VM_OPERATION As Long = &H8
Public Const PROCESS_VM_READ As Long = &H10
Public Const PROCESS_VM_WRITE As Long = &H20
Public Const PROCESS_DUP_HANDLE As Long = &H40
Public Const PROCESS_CREATE_PROCESS As Long = &H80
Public Const PROCESS_SET_QUOTA As Long = &H100
Public Const PROCESS_SET_INFORMATION As Long = &H200
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Public Const PROCESSOR_ARCHITECTURE_INTEL As Long = 0
Public Const PROCESSOR_ARCHITECTURE_MIPS As Long = 1
Public Const PROCESSOR_ARCHITECTURE_ALPHA As Long = 2
Public Const PROCESSOR_ARCHITECTURE_PPC As Long = 3
Public Const PROCESSOR_ARCHITECTURE_SHX As Long = 4
Public Const PROCESSOR_ARCHITECTURE_ARM As Long = 5
Public Const PROCESSOR_ARCHITECTURE_IA64 As Long = 6
Public Const PROCESSOR_ARCHITECTURE_ALPHA64 As Long = 7
Public Const PROCESSOR_ARCHITECTURE_MSIL As Long = 8
Public Const PROCESSOR_ARCHITECTURE_AMD64 As Long = 9
Public Const PROCESSOR_ARCHITECTURE_IA32_ON_WIN64 As Long = 10
Public Const PROCESSOR_ARCHITECTURE_UNKNOWN As Long = &HFFFF

Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Public Const REG_DWORD_BIG_ENDIAN As Long = 5
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8
Public Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9
Public Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10
Public Const REG_QWORD As Long = 11
Public Const REG_QWORD_LITTLE_ENDIAN As Long = 11

Public Const REG_OPTION_RESERVED As Long = 0
Private Const REG_OPTION_NON_VOLATILE As Long = 0
Public Const REG_OPTION_VOLATILE As Long = 1
Public Const REG_OPTION_CREATE_LINK As Long = 2
Public Const REG_OPTION_BACKUP_RESTORE As Long = 4
Public Const REG_OPTION_OPEN_LINK As Long = &H8
Public Const REG_LEGAL_OPTION As Long = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE Or REG_OPTION_OPEN_LINK)

Public Const SE_CREATE_TOKEN_NAME As String = "SeCreateTokenPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN_NAME As String = "SeAssignPrimaryTokenPrivilege"
Public Const SE_LOCK_MEMORY_NAME As String = "SeLockMemoryPrivilege"
Public Const SE_INCREASE_QUOTA_NAME As String = "SeIncreaseQuotaPrivilege"
Public Const SE_UNSOLICITED_INPUT_NAME As String = "SeUnsolicitedInputPrivilege"
Public Const SE_MACHINE_ACCOUNT_NAME As String = "SeMachineAccountPrivilege"
Public Const SE_TCB_NAME As String = "SeTcbPrivilege"
Public Const SE_SECURITY_NAME As String = "SeSecurityPrivilege"
Public Const SE_TAKE_OWNERSHIP_NAME As String = "SeTakeOwnershipPrivilege"
Public Const SE_LOAD_DRIVER_NAME As String = "SeLoadDriverPrivilege"
Public Const SE_SYSTEM_PROFILE_NAME As String = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME_NAME As String = "SeSystemtimePrivilege"
Public Const SE_PROF_SINGLE_PROCESS_NAME As String = "SeProfileSingleProcessPrivilege"
Public Const SE_INC_BASE_PRIORITY_NAME As String = "SeIncreaseBasePriorityPrivilege"
Public Const SE_CREATE_PAGEFILE_NAME As String = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT_NAME As String = "SeCreatePermanentPrivilege"
Public Const SE_BACKUP_NAME As String = "SeBackupPrivilege"
Public Const SE_RESTORE_NAME As String = "SeRestorePrivilege"
Public Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
Public Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Public Const SE_AUDIT_NAME As String = "SeAuditPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT_NAME As String = "SeSystemEnvironmentPrivilege"
Public Const SE_CHANGE_NOTIFY_NAME As String = "SeChangeNotifyPrivilege"
Public Const SE_REMOTE_SHUTDOWN_NAME As String = "SeRemoteShutdownPrivilege"
Public Const SE_UNDOCK_NAME As String = "SeUndockPrivilege"
Public Const SE_SYNC_AGENT_NAME As String = "SeSyncAgentPrivilege"
Public Const SE_ENABLE_DELEGATION_NAME As String = "SeEnableDelegationPrivilege"
Public Const SE_MANAGE_VOLUME_NAME As String = "SeManageVolumePrivilege"

Public Const SE_PRIVILEGE_ENABLED_BY_DEFAULT As Long = &H1
Public Const SE_PRIVILEGE_ENABLED As Long = &H2
Public Const SE_PRIVILEGE_USED_FOR_ACCESS As Long = &H80000000

Public Const THREAD_TERMINATE As Long = &H1
Public Const THREAD_SUSPEND_RESUME As Long = &H2
Public Const THREAD_GET_CONTEXT As Long = &H8
Public Const THREAD_SET_CONTEXT As Long = &H10
Public Const THREAD_SET_INFORMATION As Long = &H20
Public Const THREAD_QUERY_INFORMATION As Long = &H40
Public Const THREAD_SET_THREAD_TOKEN As Long = &H80
Public Const THREAD_IMPERSONATE As Long = &H100
Public Const THREAD_DIRECT_IMPERSONATION As Long = &H200
Public Const THREAD_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)

Public Const THREAD_BASE_PRIORITY_LOWRT As Long = 15
Public Const THREAD_BASE_PRIORITY_MAX As Long = 2
Public Const THREAD_BASE_PRIORITY_MIN As Long = -2
Public Const THREAD_BASE_PRIORITY_IDLE As Long = -15

Public Const TIME_ZONE_ID_UNKNOWN As Long = 0
Public Const TIME_ZONE_ID_STANDARD As Long = 1
Public Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Public Const TOKEN_ASSIGN_PRIMARY As Long = &H1
Public Const TOKEN_DUPLICATE As Long = &H2
Public Const TOKEN_IMPERSONATE As Long = &H4
Public Const TOKEN_QUERY As Long = &H8
Public Const TOKEN_QUERY_SOURCE As Long = &H10
Public Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Public Const TOKEN_ADJUST_GROUPS As Long = &H40
Public Const TOKEN_ADJUST_DEFAULT As Long = &H80
Public Const TOKEN_ADJUST_SESSIONID As Long = &H100
Public Const TOKEN_ALL_ACCESS_P As Long = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Public Const TOKEN_ALL_ACCESS_NT As Long = (TOKEN_ALL_ACCESS_P Or TOKEN_ADJUST_SESSIONID)
Public Const TOKEN_ALL_ACCESS As Long = (TOKEN_ALL_ACCESS_P)
Public Const TOKEN_READ As Long = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
Public Const TOKEN_WRITE As Long = (STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Public Const TOKEN_EXECUTE As Long = STANDARD_RIGHTS_EXECUTE

Public Const VER_NT_WORKSTATION As Long = &H1
Public Const VER_NT_DOMAIN_CONTROLLER As Long = &H2
Public Const VER_NT_SERVER As Long = &H3

Public Const VER_PLATFORM_WIN32s As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public Const VER_SUITE_SMALLBUSINESS As Long = &H1
Public Const VER_SUITE_ENTERPRISE As Long = &H2
Public Const VER_SUITE_BACKOFFICE As Long = &H4
Public Const VER_SUITE_COMMUNICATIONS As Long = &H8
Public Const VER_SUITE_TERMINAL As Long = &H10
Public Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20
Public Const VER_SUITE_EMBEDDEDNT As Long = &H40
Public Const VER_SUITE_DATACENTER As Long = &H80
Public Const VER_SUITE_SINGLEUSERTS As Long = &H100

Public Const STATUS_WAIT_0 As Long = &H0
Public Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Public Const STATUS_USER_APC As Long = &HC0
Public Const STATUS_TIMEOUT As Long = &H102
Public Const STATUS_PENDING As Long = &H103
Public Const DBG_EXCEPTION_HANDLED As Long = &H10001
Public Const DBG_CONTINUE As Long = &H10002
Public Const STATUS_SEGMENT_NOTIFICATION As Long = &H40000005
Public Const DBG_TERMINATE_THREAD As Long = &H40010003
Public Const DBG_TERMINATE_PROCESS As Long = &H40010004
Public Const DBG_CONTROL_C As Long = &H40010005
Public Const DBG_CONTROL_BREAK As Long = &H40010008
Public Const STATUS_GUARD_PAGE_VIOLATION As Long = &H80000001
Public Const STATUS_DATATYPE_MISALIGNMENT As Long = &H80000002
Public Const STATUS_BREAKPOINT As Long = &H80000003
Public Const STATUS_SINGLE_STEP As Long = &H80000004
Public Const DBG_EXCEPTION_NOT_HANDLED As Long = &H80010001
Public Const STATUS_ACCESS_VIOLATION As Long = &HC0000005
Public Const STATUS_IN_PAGE_ERROR As Long = &HC0000006
Public Const STATUS_INVALID_HANDLE As Long = &HC0000008
Public Const STATUS_NO_MEMORY As Long = &HC0000017
Public Const STATUS_ILLEGAL_INSTRUCTION As Long = &HC000001D
Public Const STATUS_NONCONTINUABLE_EXCEPTION As Long = &HC0000025
Public Const STATUS_INVALID_DISPOSITION As Long = &HC0000026
Public Const STATUS_ARRAY_BOUNDS_EXCEEDED As Long = &HC000008C
Public Const STATUS_FLOAT_DENORMAL_OPERAND As Long = &HC000008D
Public Const STATUS_FLOAT_DIVIDE_BY_ZERO As Long = &HC000008E
Public Const STATUS_FLOAT_INEXACT_RESULT As Long = &HC000008F
Public Const STATUS_FLOAT_INVALID_OPERATION As Long = &HC0000090
Public Const STATUS_FLOAT_OVERFLOW As Long = &HC0000091
Public Const STATUS_FLOAT_STACK_CHECK As Long = &HC0000092
Public Const STATUS_FLOAT_UNDERFLOW As Long = &HC0000093
Public Const STATUS_INTEGER_DIVIDE_BY_ZERO As Long = &HC0000094
Public Const STATUS_INTEGER_OVERFLOW As Long = &HC0000095
Public Const STATUS_PRIVILEGED_INSTRUCTION As Long = &HC0000096
Public Const STATUS_STACK_OVERFLOW As Long = &HC00000FD
Public Const STATUS_CONTROL_C_EXIT As Long = &HC000013A
Public Const STATUS_FLOAT_MULTIPLE_FAULTS As Long = &HC00002B4
Public Const STATUS_FLOAT_MULTIPLE_TRAPS As Long = &HC00002B5
Public Const STATUS_REG_NAT_CONSUMPTION As Long = &HC00002C9
Public Const STATUS_SXS_EARLY_DEACTIVATION As Long = &HC015000F
Public Const STATUS_SXS_INVALID_DEACTIVATION As Long = &HC0150010

Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Boolean, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Boolean
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean
Public Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetProcessAffinityMask Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpProcessAffinityMask As Long, ByRef lpSystemAffinityMask As Long) As Boolean
Public Declare Function GetProcessIoCounters Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpIoCounters As IO_COUNTERS) As Boolean
Public Declare Function GetProcessTimes Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Boolean
Public Declare Function GetProcessVersion Lib "kernel32.dll" (ByVal ProcessID As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function GetThreadTimes Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Boolean
Public Declare Function IsBadReadPtr Lib "kernel32.dll" (ByRef lp As Any, ByVal ucb As Long) As Boolean
Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Boolean
Public Declare Function lstrlen Lib "kernel32.dll" (ByVal lpString As Any) As Long
Public Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Boolean
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Public Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean
Public Declare Function SetSystemTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME) As Boolean
Public Declare Function SetThreadIdealProcessor Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwIdealProcessor As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long, ByVal nPriority As Long) As Boolean
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32.dll" (ByVal lpTopLevelExceptionFilter As Long) As Long
Public Declare Function SetVolumeLabel Lib "kernel32.dll" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Boolean
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Boolean
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Public Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean


Public Const CREATE_NEW As Long = 1
Public Const CREATE_ALWAYS As Long = 2
Public Const OPEN_EXISTING As Long = 3
Public Const OPEN_ALWAYS As Long = 4
Public Const TRUNCATE_EXISTING As Long = 5

Public Const DEBUG_PROCESS As Long = &H1
Public Const DEBUG_ONLY_THIS_PROCESS As Long = &H2
Public Const CREATE_SUSPENDED As Long = &H4
Public Const DETACHED_PROCESS As Long = &H8
Public Const CREATE_NEW_CONSOLE As Long = &H10
Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const IDLE_PRIORITY_CLASS As Long = &H40
Public Const HIGH_PRIORITY_CLASS As Long = &H80
Public Const REALTIME_PRIORITY_CLASS As Long = &H100
Public Const CREATE_NEW_PROCESS_GROUP As Long = &H200
Public Const CREATE_UNICODE_ENVIRONMENT As Long = &H400
Public Const CREATE_SEPARATE_WOW_VDM As Long = &H800
Public Const CREATE_SHARED_WOW_VDM As Long = &H1000
Public Const CREATE_FORCEDOS As Long = &H2000
Public Const BELOW_NORMAL_PRIORITY_CLASS As Long = &H4000
Public Const ABOVE_NORMAL_PRIORITY_CLASS As Long = &H8000
Public Const CREATE_BREAKAWAY_FROM_JOB As Long = &H1000000

Public Const WAIT_IO_COMPLETION As Long = STATUS_USER_APC
Public Const STILL_ACTIVE As Long = STATUS_PENDING
Public Const EXCEPTION_ACCESS_VIOLATION As Long = STATUS_ACCESS_VIOLATION
Public Const EXCEPTION_DATATYPE_MISALIGNMENT As Long = STATUS_DATATYPE_MISALIGNMENT
Public Const EXCEPTION_BREAKPOINT As Long = STATUS_BREAKPOINT
Public Const EXCEPTION_SINGLE_STEP As Long = STATUS_SINGLE_STEP
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED As Long = STATUS_ARRAY_BOUNDS_EXCEEDED
Public Const EXCEPTION_FLT_DENORMAL_OPERAND As Long = STATUS_FLOAT_DENORMAL_OPERAND
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO As Long = STATUS_FLOAT_DIVIDE_BY_ZERO
Public Const EXCEPTION_FLT_INEXACT_RESULT As Long = STATUS_FLOAT_INEXACT_RESULT
Public Const EXCEPTION_FLT_INVALID_OPERATION As Long = STATUS_FLOAT_INVALID_OPERATION
Public Const EXCEPTION_FLT_OVERFLOW As Long = STATUS_FLOAT_OVERFLOW
Public Const EXCEPTION_FLT_STACK_CHECK As Long = STATUS_FLOAT_STACK_CHECK
Public Const EXCEPTION_FLT_UNDERFLOW As Long = STATUS_FLOAT_UNDERFLOW
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO As Long = STATUS_INTEGER_DIVIDE_BY_ZERO
Public Const EXCEPTION_INT_OVERFLOW As Long = STATUS_INTEGER_OVERFLOW
Public Const EXCEPTION_PRIV_INSTRUCTION As Long = STATUS_PRIVILEGED_INSTRUCTION
Public Const EXCEPTION_IN_PAGE_ERROR As Long = STATUS_IN_PAGE_ERROR
Public Const EXCEPTION_ILLEGAL_INSTRUCTION As Long = STATUS_ILLEGAL_INSTRUCTION
Public Const EXCEPTION_NONCONTINUABLE_EXCEPTION As Long = STATUS_NONCONTINUABLE_EXCEPTION
Public Const EXCEPTION_STACK_OVERFLOW As Long = STATUS_STACK_OVERFLOW
Public Const EXCEPTION_INVALID_DISPOSITION As Long = STATUS_INVALID_DISPOSITION
Public Const EXCEPTION_GUARD_PAGE As Long = STATUS_GUARD_PAGE_VIOLATION
Public Const EXCEPTION_INVALID_HANDLE As Long = STATUS_INVALID_HANDLE
Public Const CONTROL_C_EXIT As Long = STATUS_CONTROL_C_EXIT

Public Const DONT_RESOLVE_DLL_REFERENCES As Long = &H1
Public Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2
Public Const LOAD_WITH_ALTERED_SEARCH_PATH As Long = &H8
Public Const LOAD_IGNORE_CODE_AUTHZ_LEVEL As Long = &H10

Public Const DOCKINFO_UNDOCKED As Long = &H1
Public Const DOCKINFO_DOCKED As Long = &H2
Public Const DOCKINFO_USER_SUPPLIED As Long = &H4
Public Const DOCKINFO_USER_UNDOCKED As Long = (DOCKINFO_USER_SUPPLIED Or DOCKINFO_UNDOCKED)
Public Const DOCKINFO_USER_DOCKED As Long = (DOCKINFO_USER_SUPPLIED Or DOCKINFO_DOCKED)

Public Const DRIVE_UNKNOWN As Long = 0
Public Const DRIVE_NO_ROOT_DIR As Long = 1
Public Const DRIVE_REMOVABLE As Long = 2
Public Const DRIVE_FIXED As Long = 3
Public Const DRIVE_REMOTE As Long = 4
Public Const DRIVE_CDROM As Long = 5
Public Const DRIVE_RAMDISK As Long = 6

Public Const FILE_BEGIN As Long = 0
Public Const FILE_CURRENT As Long = 1
Public Const FILE_END As Long = 2

Public Const FILE_ENCRYPTABLE As Long = 0
Public Const FILE_IS_ENCRYPTED As Long = 1
Public Const FILE_SYSTEM_ATTR As Long = 2
Public Const FILE_ROOT_DIR As Long = 3
Public Const FILE_SYSTEM_DIR As Long = 4
Public Const FILE_UNKNOWN As Long = 5
Public Const FILE_SYSTEM_NOT_SUPPORT As Long = 6
Public Const FILE_USER_DISALLOWED As Long = 7
Public Const FILE_READ_ONLY As Long = 8
Public Const FILE_DIR_DISALLOWED As Long = 9

Public Const FILE_TYPE_UNKNOWN As Long = &H0
Public Const FILE_TYPE_DISK As Long = &H1
Public Const FILE_TYPE_CHAR As Long = &H2
Public Const FILE_TYPE_PIPE As Long = &H3
Public Const FILE_TYPE_REMOTE As Long = &H8000

Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Public Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Public Const FORMAT_MESSAGE_FROM_STRING As Long = &H400
Public Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800
Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = &H2000
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF


Public Const HW_PROFILE_GUIDLEN As Long = 39

Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const INVALID_FILE_SIZE As Long = &HFFFFFFFF
Public Const INVALID_SET_FILE_POINTER As Long = -1

Public Const MAX_PROFILE_LEN As Long = 80

Public Const MAX_COMPUTERNAME_LENGTH As Long = 31

Public Const MAXLONG As Long = &H7FFFFFFF

Public Const SEM_FAILCRITICALERRORS As Long = &H1
Public Const SEM_NOGPFAULTERRORBOX As Long = &H2
Public Const SEM_NOALIGNMENTFAULTEXCEPT As Long = &H4
Public Const SEM_NOOPENFILEERRORBOX As Long = &H8000

Public Const THREAD_PRIORITY_LOWEST As Long = THREAD_BASE_PRIORITY_MIN
Public Const THREAD_PRIORITY_BELOW_NORMAL As Long = (THREAD_PRIORITY_LOWEST + 1)
Public Const THREAD_PRIORITY_NORMAL As Long = 0
Public Const THREAD_PRIORITY_HIGHEST As Long = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_ABOVE_NORMAL As Long = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_ERROR_RETURN As Long = (MAXLONG)
Public Const THREAD_PRIORITY_TIME_CRITICAL As Long = THREAD_BASE_PRIORITY_LOWRT
Public Const THREAD_PRIORITY_IDLE As Long = THREAD_BASE_PRIORITY_IDLE

Public Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF


Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Public Type HW_PROFILE_INFO
    dwDockInfo As Long
    szHwProfileGuid As String * HW_PROFILE_GUIDLEN
    szHwProfileName As String * MAX_PROFILE_LEN
End Type

Public Type SYSTEM_INFO
    dwOemID As Long 'Union
    'WORD wProcessorArchitecture
    'WORD wReserved

    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Public Type SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    Reserved1 As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName As String * 64
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName As String * 64
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
 
'Public Enum COMPUTER_NAME_FORMAT
'    ComputerNameNetBIOS
'    ComputerNameDnsHostname
'    ComputerNameDnsDomain
'    ComputerNameDnsFullyQualified
'    ComputerNamePhysicalNetBIOS
'    ComputerNamePhysicalDnsHostname
'    ComputerNamePhysicalDnsDomain
'    ComputerNamePhysicalDnsFullyQualified
'    ComputerNameMax
'End Enum
'


'
'Public Type ADMINISTRATOR_POWER_POLICY
'    MinSleep As SYSTEM_POWER_STATE
'    MaxSleep As SYSTEM_POWER_STATE
'    MinVideoTimeout As Long
'    MaxVideoTimeout As Long
'    MinSpindownTimeout As Long
'    MaxSpindownTimeout As Long
'End Type

Public Type BATTERY_REPORTING_SCALE
    Granularity As Long
    Capacity As Long
End Type

Public Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    ExceptionRecord As Long
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(0 To EXCEPTION_MAXIMUM_PARAMETERS - 1) As Long
End Type

Public Type EXCEPTION_POINTERS
    ExceptionRecord As EXCEPTION_RECORD
    ContextRecord As Long
End Type

Public Type IMAGE_DOS_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(0 To 3)   As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(0 To 9) As Integer
    e_lfanew As Long
End Type

Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type IMAGE_OS2_HEADER
    ne_magic As Integer
    ne_ver As Byte
    ne_rev As Byte
    ne_enttab As Integer
    ne_cbenttab As Integer
    ne_crc As Long
    ne_flags As Integer
    ne_autodata As Integer
    ne_heap As Integer
    ne_stack As Integer
    ne_csip As Long
    ne_sssp As Long
    ne_cseg As Integer
    ne_cmod As Integer
    ne_cbnrestab As Integer
    ne_segtab As Integer
    ne_rsrctab As Integer
    ne_restab As Integer
    ne_modtab As Integer
    ne_imptab As Integer
    ne_nrestab As Long
    ne_cmovent As Integer
    ne_align As Integer
    ne_cres As Integer
    ne_exetyp As Byte
    ne_flagsothers As Byte
    ne_pretthunks As Integer
    ne_psegrefbytes As Integer
    ne_swaparea As Integer
    ne_expver As Integer
End Type
  
Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Type LUID
    lowpart As Long
    highpart As Long
End Type

Public Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Public Type IO_COUNTERS
    ReadOperationCount As LARGE_INTEGER
    WriteOperationCount As LARGE_INTEGER
    OtherOperationCount As LARGE_INTEGER
    ReadTransferCount As LARGE_INTEGER
    WriteTransferCount As LARGE_INTEGER
    OtherTransferCount As LARGE_INTEGER
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Type PROCESSOR_POWER_INFORMATION
    Number As Long
    MaxMhz As Long
    CurrentMhz As Long
    MhzLimit As Long
    MaxIdleState As Long
    CurrentIdleState As Long
End Type

Public Type SYSTEM_BATTERY_STATE
    AcOnLine As Byte
    BatteryPresent As Byte
    Charging As Byte
    Discharging As Byte
    Spare1(0 To 3) As Byte
    MaxCapacity As Long
    RemainingCapacity As Long
    Rate As Long
    EstimatedTime As Long
    DefaultAlert1 As Long
    DefaultAlert2 As Long
End Type

'Public Type SYSTEM_POWER_CAPABILITIES
'    PowerButtonPresent As Byte
'    SleepButtonPresent As Byte
'    LidPresent As Byte
'    SystemS1 As Byte
'    SystemS2 As Byte
'    SystemS3 As Byte
'    SystemS4 As Byte
'    SystemS5 As Byte
'    HiberFilePresent As Byte
'    FullWake As Byte
'    VideoDimPresent As Byte
'    ApmPresent As Byte
'    UpsPresent As Byte
'    ThermalControl As Byte
'    ProcessorThrottle As Byte
'    ProcessorMinThrottle As Byte
'    ProcessorMaxThrottle As Byte
'    spare2(0 To 3) As Byte
'    DiskSpinDown As Byte
'    spare3(0 To 7) As Byte
'    SystemBatteriesPresent As Byte
'    BatteriesAreShortTerm As Byte
'    BatteryScale(0 To 2) As BATTERY_REPORTING_SCALE
'    AcOnLineWake As SYSTEM_POWER_STATE
'    SoftLidWake As SYSTEM_POWER_STATE
'    RtcWake As SYSTEM_POWER_STATE
'    MinDeviceWakeState As SYSTEM_POWER_STATE
'    DefaultLowLatencyWake As SYSTEM_POWER_STATE
'End Type

Public Type SYSTEM_POWER_INFORMATION
    MaxIdlenessAllowed As Long
    Idleness As Long
    TimeRemaining As Long
    CoolingMode As Long
End Type

Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(64) As LUID_AND_ATTRIBUTES
End Type


'Public Enum POWER_INFORMATION_LEVEL
'    SystemPowerPolicyAc = 0
'    SystemPowerPolicyDc = 1
'    VerifySystemPolicyAc = 2
'    VerifySystemPolicyDc = 3
'    SystemPowerCapabilities = 4
'    SystemBatteryState = 5
'    SystemPowerStateHandler = 6
'    ProcessorStateHandler = 7
'    SystemPowerPolicyCurrent = 8
'    AdministratorPowerPolicy = 9
'    SystemReserveHiberFile = 10
'    ProcessorInformation = 11
'    SystemPowerInformation = 12
'    ProcessorStateHandler2 = 13
'    LastWakeTime = 14
'    LastSleepTime = 15
'    SystemExecutionState = 16
'    SystemPowerStateNotifyHandler = 17
'    ProcessorPowerPolicyAc = 18
'    ProcessorPowerPolicyDc = 19
'    VerifyProcessorPowerPolicyAc = 20
'    VerifyProcessorPowerPolicyDc = 21
'    ProcessorPowerPolicyCurrent = 22
'End Enum

'Public Enum SYSTEM_POWER_STATE
'    PowerSystemUnspecified = 0
'    PowerSystemWorking = 1
'    PowerSystemSleeping1 = 2
'    PowerSystemSleeping2 = 3
'    PowerSystemSleeping3 = 4
'    PowerSystemHibernate = 5
'    PowerSystemShutdown = 6
'    PowerSystemMaximum = 7
'End Enum

Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Heap32First Lib "kernel32.dll" (ByRef lphe As HEAPENTRY32, ByVal th32ProcessID As Long, ByVal th32HeapID As Long) As Boolean
Public Declare Function Heap32ListFirst Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lphl As HEAPLIST32) As Boolean
Public Declare Function Heap32ListNext Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lphl As HEAPLIST32) As Boolean
Public Declare Function Heap32Next Lib "kernel32.dll" (ByRef lphe As HEAPENTRY32) As Boolean
Public Declare Function Module32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpme As MODULEENTRY32) As Boolean
Public Declare Function Module32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpme As MODULEENTRY32) As Boolean
Public Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lppe As PROCESSENTRY32) As Boolean
Public Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lppe As PROCESSENTRY32) As Boolean
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean


Public Const HF32_DEFAULT As Long = 1
Public Const HF32_SHARED As Long = 2

Public Const LF32_FIXED As Long = &H1
Public Const LF32_FREE As Long = &H2
Public Const LF32_MOVEABLE As Long = &H4

Public Const MAX_MODULE_NAME32 As Long = 255

Public Const TH32CS_SNAPHEAPLIST As Long = &H1
Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const TH32CS_SNAPTHREAD As Long = &H4
Public Const TH32CS_SNAPMODULE As Long = &H8
Public Const TH32CS_SNAPALL As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT As Long = &H80000000
    
    
Public Type HEAPENTRY32
    dwSize As Long
    hHandle As Long
    dwAddress As Long
    dwBlockSize As Long
    dwFlags As Long
    dwLockCount As Long
    dwResvd As Long
    th32ProcessID As Long
    th32HeapID As Long
End Type

Public Type HEAPLIST32
    dwSize As Long
    th32ProcessID As Long
    th32HeapID As Long
    dwFlags As Long
End Type

Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256    'MAX_MODULE_NAME32 + 1
    szExePath As String * MAX_PATH
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID  As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type





Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const sLocation As String = "mdlProcess"



'Public Const ERROR_MOD_NOT_FOUND As Long = 126




'-----------------------------------------------------
'-----------------------------------------------------
'                    End of Declarations
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------




'
'
'
'Public Sub Err_Dll(ErrorNum As Long, ErrorDesc As String, Source As String, SubOrFunction As String)
'    File_WriteTo "ERROR: " & ErrorNum & " at " & Source & "\" & SubOrFunction & " >>> " & ErrorDesc
'End Sub
'Public Sub Err_Vb(ErrorNum As Long, ErrorDesc As String, Source As String, SubOrFunction As String)
'    File_WriteTo "VB ERROR: " & ErrorNum & " at " & Source & "\" & SubOrFunction & " >>> " & ErrorDesc
'End Sub
'Public Sub File_WriteTo(Text As String)
'    '// Allways use this in programs:
' '   Open App.Path & "\PROGRAM.LOG" For Append As #1
' '       Print #1, Text
' '   Close #1
'End Sub



Public Function Heap32_Enum(ByRef Heap() As HEAPENTRY32, ByVal lProcessID As Long, ByVal lHeapID As Long) As Long
On Error GoTo VB_Error
    
    '// Here, we get all the heaps, put them in heap entry (32 !!!)
    ReDim Heap(0)
    
    Dim HEAPENTRY32 As HEAPENTRY32
    Dim lHeap As Long
    
    HEAPENTRY32.dwSize = Len(HEAPENTRY32)
    If Heap32First(HEAPENTRY32, lProcessID, lHeapID) = False Then
        Heap32_Enum = -1
         'Call Err_Dll(Err.LastDllError, "Heap32First failed", sLocation, "Heap32_Enum")
        Exit Function
    Else
        ReDim Heap(lHeap)
        Heap(lHeap) = HEAPENTRY32
    End If
    
    Do
        If Heap32Next(HEAPENTRY32) = False Then
            Exit Do
        Else
            lHeap = lHeap + 1
            ReDim Preserve Heap(lHeap)
            Heap(lHeap) = HEAPENTRY32
        
        
        
        End If
    Loop
    
    
    Heap32_Enum = lHeap
    
Exit Function
VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Heap32_Enum"
Resume Next
End Function



Public Function Heap32List_Enum(ByRef HeapList() As HEAPLIST32, ByVal lProcessID As Long) As Long

On Error GoTo VB_Error

    '// Now we list them in heaplist, and  all the heaps are in your hands
    '// But this stuff is not included in program, his main purpose is to find
    '// active processes.
    
    ReDim HeapList(0)
    
    Dim HEAPLIST32 As HEAPLIST32
    Dim hSnapShot As Long
    Dim lHeapList As Long
    Dim I
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPHEAPLIST, lProcessID) ': If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "D", sLocation & "\Heap32List_Enum", "CreateToolhelp32Snapshot")
    
    HEAPLIST32.dwSize = Len(HEAPLIST32)
    If Heap32ListFirst(hSnapShot, HEAPLIST32) = False Then
        Heap32List_Enum = -1
    
         'Call Err_Dll(Err.LastDllError, "Heap32First failed", sLocation, "Heap32List_Enum")
        
        'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Heap32List_Enum")
        I = CloseHandle(hSnapShot)
        Exit Function
    Else
        ReDim HeapList(lHeapList)
        HeapList(lHeapList) = HEAPLIST32
    End If
    
    Do
        If Heap32ListNext(hSnapShot, HEAPLIST32) = False Then
            Exit Do
        Else
            lHeapList = lHeapList + 1
            ReDim Preserve HeapList(lHeapList)
            HeapList(lHeapList) = HEAPLIST32
            
            End If
        
    Loop
    
    'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Heap32List_Enum")
    I = CloseHandle(hSnapShot)
    
    Heap32List_Enum = lHeapList
    
Exit Function

VB_Error:
    
'    Err_Vb Err.Number, Err.Description, sLocation, "Heap32_ListEnum"

Resume Next
End Function

Public Function Module32_Enum(ByRef Module() As MODULEENTRY32, Optional ByVal lProcessID As Long) As Long
On Error GoTo VB_Error

    '// To get modules
    
    ReDim Module(0)
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = True Then
        Dim MODULEENTRY32 As MODULEENTRY32
        Dim hSnapShot As Long
        Dim lModule As Long
        Dim I
        
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, lProcessID)
        'If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "CreateToolHelp32Snapshoot failed", sLocation, "Module32_Enum")
        
        
        MODULEENTRY32.dwSize = Len(MODULEENTRY32)
        If Module32First(hSnapShot, MODULEENTRY32) = False Then
            Module32_Enum = -1
              'Call Err_Dll(Err.LastDllError, "Module32First failed", sLocation, "Module32_Enum")
            
            'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Module32_Enum")
            I = CloseHandle(hSnapShot)
            
            Exit Function
        Else
            ReDim Module(lModule)
            Module(lModule) = MODULEENTRY32
        End If
        
        Do
            If Module32Next(hSnapShot, MODULEENTRY32) = False Then
                Exit Do
            Else
                
                lModule = lModule + 1
                ReDim Preserve Module(lModule)
                Module(lModule) = MODULEENTRY32
            End If
        Loop
        
        'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Module32_Enum")   'Call Error_API(Err.LastDllError, sLocation & "\Module32_Enum", "CloseHandle")
        I = CloseHandle(hSnapShot)
        
        Module32_Enum = lModule
    End If
    
Exit Function
VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Module32_Enum"
Resume Next
End Function



Public Function Process32_Enum(ByRef process() As PROCESSENTRY32) As Long
On Error GoTo VB_Error

    '// Get the most wanted, processes

    ReDim process(0)
    
    Dim PROCESSENTRY32 As PROCESSENTRY32
    Dim hSnapShot As Long
    Dim lProcess As Long
    'Dim we$
    Dim I
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    'If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "CreateToolHelp32Snapshoot failed ::: INVALID_HANDLE_VALUE", sLocation, "Process32_Enum")
    
    PROCESSENTRY32.dwSize = Len(PROCESSENTRY32)
    If Process32First(hSnapShot, PROCESSENTRY32) = False Then
        Process32_Enum = -1
          
        'Call Err_Dll(Err.LastDllError, "Process32First failed", sLocation, "Process32_Enum")
        
        'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Process32_Enum")
        I = CloseHandle(hSnapShot) = False
        Exit Function
    Else
        ReDim process(lProcess)
        process(lProcess) = PROCESSENTRY32
        'we$ = PROCESSENTRY32.
    End If

    Do
        If Process32Next(hSnapShot, PROCESSENTRY32) = False Then
            Exit Do
        Else
        
            lProcess = lProcess + 1
            ReDim Preserve process(lProcess)
            process(lProcess) = PROCESSENTRY32
            'we$ = PROCESSENTRY32.szExeFile
        End If
    Loop
    
    'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Process32_Enum")   '(Err.LastDllError, sLocation & "\Process32_Enum", "CloseHandle")
    I = CloseHandle(hSnapShot)
    
    Process32_Enum = lProcess
    
Exit Function

VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Process32_Enum"
Resume Next
End Function

Public Function Thread32_Enum(ByRef Thread() As THREADENTRY32, ByVal lProcessID As Long) As Long
On Error GoTo VB_Error
    
    '// I'm tired, just ask me...
    
    ReDim Thread(0)
    
    Dim THREADENTRY32 As THREADENTRY32
    Dim hSnapShot As Long
    Dim lThread As Long
    Dim I
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcessID)
    'If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "CreateToolHelp32Snapshoot ::: INVALID_HANDLE_VALUE failed", sLocation, "Thread32_Enum")
    
    THREADENTRY32.dwSize = Len(THREADENTRY32)
    If Thread32First(hSnapShot, THREADENTRY32) = False Then
        Thread32_Enum = -1
          'Call Err_Dll(Err.LastDllError, "Thread32First failed", sLocation, "Thread32_Enum")
        
        'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Thread32_Enum")
        I = CloseHandle(hSnapShot)
        Exit Function
    Else
        ReDim Thread(lThread)
        Thread(lThread) = THREADENTRY32
    End If
    
    Do
        If Thread32Next(hSnapShot, THREADENTRY32) = False Then
            Exit Do
        Else
            lThread = lThread + 1
            ReDim Preserve Thread(lThread)
            Thread(lThread) = THREADENTRY32
        End If
    Loop
    
    'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Thread32_Enum")   'Call Error_API(Err.LastDllError, sLocation & "\Thread32_Enum", "CloseHandle")
    I = CloseHandle(hSnapShot)
    
    Thread32_Enum = lThread
    
Exit Function
VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Thread32_Enum"
Resume Next
End Function




Public Function Function_Exist(ByVal sModule As String, ByVal sFunction As String) As Boolean
On Error GoTo VB_Error
    '// Just stuff for checking funtions.
    Dim hHandle As Long
    Dim I
    
    '// Get the handle
    hHandle = GetModuleHandle(sModule)
    
    If hHandle = 0 Then
        '// Raise not found error
        'If Err.LastDllError <> ERROR_MOD_NOT_FOUND Then Call Err_Dll(Err.LastDllError, "GetModuleHandle failed ::: MOD_NOT_FOUND", sLocation, "Function_Exist")
        
        hHandle = LoadLibraryEx(sModule, 0&, 0&)
        ': If hHandle = 0 Then Call Err_Dll(Err.LastDllError, "LoadLibraryEx failed", sLocation, "Function_Exist")
        
        '// Now see if there is adress
        If GetProcAddress(hHandle, sFunction) = 0 Then
              'Call Err_Dll(Err.LastDllError, "GetProcAdress failed", sLocation, "Function_Exist")
            Function_Exist = False
        Else
            Function_Exist = True
        End If
        
        'If FreeLibrary(hHandle) = False Then Call Err_Dll(Err.LastDllError, "FreeLibrary failed", sLocation, "Function_Exist")
        I = FreeLibrary(hHandle)
    Else
        If GetProcAddress(hHandle, sFunction) = 0 Then
              'Call Err_Dll(Err.LastDllError, "GetProcAdress failed", sLocation, "Function_Exist")
            Function_Exist = Function_Exist
        Else
            Function_Exist = True
        End If
    End If
    
Exit Function
VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Function_Exist"
Resume Next
End Function



Public Sub List_ActiveProcess()
    '// Same
    Dim RR
    Dim lCount  As Long
    Dim pFile   As String
'    Dim NodeX   As Node
    Dim eer
    Dim Ew$
    Dim VirusActive
    VirusActive = 0
    
    lCount = Process32_Enum(process())
    
    Dim I As Long
    'With TreeView
    '    .Nodes.Clear
        For I = 0 To lCount
           ' Set NodeX = .Nodes.Add(, , CStr("PROCESS_" & I), FileName_Parse(Process(I).szExeFile), 1)
            ProcessID(I) = process(I).th32ProcessID
            eer = Module32_Enum(Module(), ProcessID(I))
         '   NodeX.Tag = Process(I).szExeFile
        
        'ert$=("PROCESS_" & O, tvwChild, CStr("MODULE_" & O & "NUM_" & I), Module(I).szModule, 2)
        Ew$ = process(I).szExeFile
        
        'Ew$ = Module(i).szExePath & "\" & Module(i).szModule
        'Ew$ = Module(i).szExePath & "\" & Module(i).szModule
        'Ew$ = process(i).

        RR = ProcessID(I)
        'CID_Run_Me.List1.AddItem Format$(rr, "0000") + "." + Format$(O, "0000") + "." + Format$(i, "0000 ") + Ew$
        'ProcLB.List1.AddItem Format$(rr, "0000 ") + Ew$


'ew$ = Process(i).szExeFile
'If InStr(ew$, "ccApp.exe") Then
'virusactive = 1
'End If



        Next I
    'End With
End Sub



Public Sub List_ActiveModules()
    '// At the end we have all the modules, listed in the form
        
    Dim RR

    Dim lCount As Long
    Dim pCount As Long
    Dim pFile As String
'    Dim NodeX As Node
    Dim I As Long
    Dim O As Long
    Dim lSize As Long
    
    Dim Cmdv As Long
    Dim Ew$
    
    Dim Ebuy As Long
    Dim OTSS
    
    pCount = Process32_Enum(process())
    
    For O = 0 To pCount
        lCount = Module32_Enum(Module(), ProcessID(O))
        For I = 0 To lCount
              '  Set NodeX = .Nodes.Add("PROCESS_" & O, tvwChild, CStr("MODULE_" & O & "NUM_" & I), Module(I).szModule, 2)
             '   NodeX.Tag = Module(I).szExePath & "\" & Module(I).szModule
        Ew$ = Module(I).szExePath & "\" & Module(I).szModule
        RR = ProcessID(O)
        'CID_Run_Me.List3.AddItem Format$(rr, "0000") + "." + Format$(O, "0000") + "." + Format$(i, "0000 ") + Ew$
        'ProcLB.List1.AddItem Format$(rr, "0000 ") + Ew$
        
        
        'Curprochwnd = Module(i).hModule
        
        OTSS = ProcessID(O)
        
        
        'Call Perfect(Ew$, 0)
        'If NoMusic = 1 Then FlingGrater1$ = Ew$
        'If NoMonOff = 1 Then FlingGrater2$ = Ew$
        Next I
    Next O
'If NoMusic = 1 And Cmdv = 1 Then NoMusic = 0
'If Ebuy = 0 Then Ebuyer = 0
End Sub

'Public Sub List_ActiveThreads(ListView As ListView, P_ID As Long)
Public Sub List_ActiveThreads(vbNullString, P_ID As Long)
    '// A little different
    
    Dim lCount As Long
    Dim I As Long
'    Dim ItemX As ListItem
    
    lCount = Thread32_Enum(Thread(), ProcessID(P_ID - 1))
    
'    With ListView
'        .ListItems.Clear
'        For i = 0 To lCount
'            If Thread(i).th32OwnerProcessID = ProcessID(P_ID - 1) Then
'                Set ItemX = ListView.ListItems.Add(, , Thread(i).th32OwnerProcessID, , 4)
'           '     ItemX.SubItems(1) = Thread(I).th32ThreadID
'           '     ItemX.SubItems(2) = Thread(I).cntUsage
'            End If
'        Next i
'    End With
End Sub

'
'Public Sub Process_Kill(P_ID As Long)
'    '// Kill the wanted process
'
'    'EXAMPLE ---- Call Process_Kill(PID)
'
'    Dim hProcess As Long
'    Dim lExitCode As Long
'    Dim XI As Long
'
'    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
'
'    'If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
'
'    'XI = GetExitCodeProcess(hProcess, lExitCode) = False
'
'    'If GetExitCodeProcess(hProcess, lExitCode) = False Then Call Err_Dll(Err.LastDllError, "GetExitCodeProcess failed", sLocation, "Kill_Process")
'
'    XI = TerminateProcess(hProcess, lExitCode) = False
'
''    If TerminateProcess(hProcess, lExitCode) = False Then Call Err_Dll(Err.LastDllError, "TerminateProcess failed", sLocation, "Kill_Process")
'
'    XI = CloseHandle(hProcess) = False
'
'    'If CloseHandle(hProcess) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Kill_Process")
'
'End Sub


'
'
'Public Function Process_Kill(P_ID As Long) As Long
'    '// Kill the wanted process
'
'    'EXAMPLE ---- VAR = cProcesses.Process_Kill(PID)
'
'    'MODIFIED TO ADD 201204182300 HOUR ---
'
'
'    Dim hProcess As Long
'    Dim lExitCode As Long
'    Dim XI As Long
'
'    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
'
'
'    XI = TerminateProcess(hProcess, lExitCode) = False
'
'    Call CloseHandle(hProcess)
'
'End Function
'




Public Sub Thread_Suspend(T_ID As Long)
    
    On Error GoTo VB_Error
    
        Dim hThread As Long
        Dim lSuspendCount As Long
        Dim I
        
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
        'If hThread = 0 Then Err_Dll Err.LastDllError, "OpenThread failed", sLocation, "Suspend_Thread"  'Call Error_API(Err.LastDllError, sLocation & "\cmdSuspend_Click", "OpenThread")
        
        lSuspendCount = SuspendThread(hThread)
        ': If lSuspendCount = -1 Then Err_Dll Err.LastDllError, "Suspend failed", sLocation, "Suspend_Thread"
        
        'If CloseHandle(hThread) = False Then Err_Dll Err.LastDllError, "CloseThread failed", sLocation, "Suspend_Thread"
        I = CloseHandle(hThread) '= False Then Err_Dll Err.LastDllError, "CloseThread failed", sLocation, "Suspend_Thread"
        
    Exit Sub
VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Thread_Suspend"
    Resume Next
End Sub

Public Sub Thread_Resume(T_ID As Long)
    On Error GoTo VB_Error
    
        Dim hThread As Long
        Dim lSuspendCount As Long
        Dim I
        
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
        'If hThread = 0 Then Err_Dll Err.LastDllError, "OpenThread failed", sLocation, "Suspend_Thread"  'Call Error_API(Err.LastDllError, sLocation & "\cmdSuspend_Click", "OpenThread")
    
        lSuspendCount = ResumeThread(hThread)
        ': If lSuspendCount = -1 Then Err_Dll Err.LastDllError, "OpenThread failed", sLocation, "Suspend_Thread"
        
        'If CloseHandle(hThread) = False Then Err_Dll Err.LastDllError, "CloseThread failed", sLocation, "Suspend_Thread"
        I = CloseHandle(hThread)
        
    Exit Sub
VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Thread_Resume"
    Resume Next
End Sub

Public Sub Process_Idle(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    SetPriorityClass hProcess, NORMAL_PRIORITY_CLASS
    
End Sub

Public Sub Process_Low(P_ID As Long)
    
    'MsgBox "PROCESS LOW COMMON CODE"
    
    'Exit Sub
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    
    
    If Idle_Timer_Proc < Now Then
        'SetPriorityClass hProcess, BELOW_NORMAL_PRIORITY_CLASS
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    Else
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    End If
End Sub

'Private Sub Process_Low_Low(P_ID As Long)
'
'    Dim hProcess As Long
'    Dim lExitCode As Long
'
'    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
'
'    SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
'
'End Sub

'Private Sub Process_Low_BELOW_NORM(P_ID As Long)
'
'    Dim hProcess As Long
'    Dim lExitCode As Long
'
'    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
'
'    SetPriorityClass hProcess, BELOW_NORMAL_PRIORITY_CLASS
'
'End Sub

'Private Sub Process_ABOVE_NORMAL_PRIORITY_CLASS(P_ID As Long)
'
'    Dim hProcess As Long
'    Dim lExitCode As Long
'
'    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
'
'    SetPriorityClass hProcess, ABOVE_NORMAL_PRIORITY_CLASS
'
'End Sub

Public Sub Process_BELOW_NORMAL_PRIORITY_CLASS(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    
    SetPriorityClass hProcess, BELOW_NORMAL_PRIORITY_CLASS

End Sub
Public Sub Process_NORMAL_PRIORITY_CLASS(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    
    SetPriorityClass hProcess, NORMAL_PRIORITY_CLASS

End Sub


Function CloseProgram(Optional ByVal ProcessHandle As Long, Optional ByVal ThreadHandle As Long, Optional CloseThread As Boolean = True) As Boolean
  
  Dim ReturnValue As Long
  Dim ExitCode    As Long

  If CloseThread = True Then
    ReturnValue = GetExitCodeThread(ThreadHandle, ExitCode)
    If ReturnValue <> 0 Then
  '    TerminateProcess ProcessHandle, ExitCode
       TerminateThread ThreadHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
    
  Else
    ReturnValue = GetExitCodeProcess(ProcessHandle, ExitCode)
    If ReturnValue <> 0 Then
   '   TerminateThread ThreadHandle, ExitCode
       TerminateProcess ProcessHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
  End If
  Exit Function
  
End Function


    

Public Sub closewindow(P_ID As Long)
    '// Kill the wanted process
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    '  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID): If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    
    '  If GetExitCodeProcess(hProcess, lExitCode) = False Then Call Err_Dll(Err.LastDllError, "GetExitCodeProcess failed", sLocation, "Kill_Process")
    '  If TerminateProcess(hProcess, lExitCode) = False Then Call Err_Dll(Err.LastDllError, "TerminateProcess failed", sLocation, "Kill_Process")
    
    ' If CloseHandle(hProcess) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Kill_Process")
End Sub


Public Sub Process_LOW_PRIORITY_CLASS(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    If Idle_Timer_Proc < Now Then
        'SetPriorityClass hProcess, BELOW_NORMAL_PRIORITY_CLASS
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    Else
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    End If
End Sub



Public Sub Process_HIGH_PRIORITY_CLASS(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    
    SetPriorityClass hProcess, HIGH_PRIORITY_CLASS

End Sub
    




Public Sub Process_REALTIME_PRIORITY_CLASS(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    
    SetPriorityClass hProcess, REALTIME_PRIORITY_CLASS


End Sub
    






