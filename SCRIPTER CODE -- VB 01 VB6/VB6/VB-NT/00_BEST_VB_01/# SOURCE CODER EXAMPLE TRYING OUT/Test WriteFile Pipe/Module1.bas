Attribute VB_Name = "Module1"
Option Explicit

   Public Const FILE_ATTRIBUTE_NORMAL = &H80
   Public Const FILE_FLAG_NO_BUFFERING = &H20000000
   Public Const FILE_FLAG_WRITE_THROUGH = &H80000000

   Public Const PIPE_ACCESS_DUPLEX = &H3
   Public Const PIPE_READMODE_MESSAGE = &H2
   Public Const PIPE_TYPE_MESSAGE = &H4
   Public Const PIPE_WAIT = &H0

   Public Const INVALID_HANDLE_VALUE = -1

   Public Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)
   Public Const SECURITY_DESCRIPTOR_REVISION = (1)

   Type SECURITY_ATTRIBUTES
           nLength As Long
           lpSecurityDescriptor As Long
           bInheritHandle As Long
   End Type

   Public Const GMEM_FIXED = &H0
   Public Const GMEM_ZEROINIT = &H40
   Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

   Declare Function GlobalAlloc Lib "kernel32" ( _
      ByVal wFlags As Long, ByVal dwBytes As Long) As Long
   Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
   Declare Function CreateNamedPipe Lib "kernel32" Alias _
      "CreateNamedPipeA" ( _
      ByVal lpName As String, _
      ByVal dwOpenMode As Long, _
      ByVal dwPipeMode As Long, _
      ByVal nMaxInstances As Long, _
      ByVal nOutBufferSize As Long, _
      ByVal nInBufferSize As Long, _
      ByVal nDefaultTimeOut As Long, _
      lpSecurityAttributes As Any) As Long

   Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal dwRevision As Long) As Long

   Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" ( _
      ByVal pSecurityDescriptor As Long, _
      ByVal bDaclPresent As Long, _
      ByVal pDacl As Long, _
      ByVal bDaclDefaulted As Long) As Long

   Declare Function ConnectNamedPipe Lib "kernel32" ( _
      ByVal hNamedPipe As Long, _
      lpOverlapped As Any) As Long

   Declare Function DisconnectNamedPipe Lib "kernel32" ( _
      ByVal hNamedPipe As Long) As Long

   Declare Function WriteFile Lib "kernel32" ( _
      ByVal hFile As Long, _
      lpBuffer As Any, _
      ByVal nNumberOfBytesToWrite As Long, _
      lpNumberOfBytesWritten As Long, _
      lpOverlapped As Any) As Long

   Declare Function ReadFile Lib "kernel32" ( _
      ByVal hFile As Long, _
      lpBuffer As Any, _
      ByVal nNumberOfBytesToRead As Long, _
      lpNumberOfBytesRead As Long, _
      lpOverlapped As Any) As Long

   Declare Function FlushFileBuffers Lib "kernel32" ( _
      ByVal hFile As Long) As Long

   Declare Function CloseHandle Lib "kernel32" ( _
      ByVal hObject As Long) As Long
                    

