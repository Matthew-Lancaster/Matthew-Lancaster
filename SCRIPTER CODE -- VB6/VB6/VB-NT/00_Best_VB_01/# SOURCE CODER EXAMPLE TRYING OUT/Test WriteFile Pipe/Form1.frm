VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
   Private Const szPipeName = "\\.\pipe\bigtest"
   Private Const BUFFSIZE = 20000
   Private BigBuffer(BUFFSIZE) As Byte, pSD As Long
   Private sa As SECURITY_ATTRIBUTES
   Private hPipe As Long

   Private Sub Form_load()
      Dim i As Long, dwOpenMode As Long, dwPipeMode As Long
      Dim res As Long, nCount As Long, cbnCount As Long
      For i = 0 To BUFFSIZE - 1       'Fill an array of numbers
         BigBuffer(i) = i Mod 256
      Next i

      'Create the NULL security token for the pipe
      pSD = GlobalAlloc(GPTR, SECURITY_DESCRIPTOR_MIN_LENGTH)
      res = InitializeSecurityDescriptor(pSD, SECURITY_DESCRIPTOR_REVISION)
      res = SetSecurityDescriptorDacl(pSD, -1, 0, 0)
      sa.nLength = LenB(sa)
      sa.lpSecurityDescriptor = pSD
      sa.bInheritHandle = True

      'Create the Named Pipe
      dwOpenMode = PIPE_ACCESS_DUPLEX Or FILE_FLAG_WRITE_THROUGH
      dwPipeMode = PIPE_WAIT Or PIPE_TYPE_MESSAGE Or PIPE_READMODE_MESSAGE
      hPipe = CreateNamedPipe(szPipeName, dwOpenMode, dwPipeMode, _
                              10, 10000, 2000, 10000, sa)

      Do  'Wait for a connection, block until a client connects
         res = ConnectNamedPipe(hPipe, ByVal 0)

         'Read/Write data over the pipe
         cbnCount = 4

         res = ReadFile(hPipe, nCount, LenB(nCount), cbnCount, ByVal 0)

         If nCount <> 0 Then

            If nCount > BUFFSIZE Then 'Client requested nCount bytes
               nCount = BUFFSIZE      'but only send up to 20000 bytes
            End If
            'Write the number of bytes requested
            res = WriteFile(hPipe, BigBuffer(0), nCount, cbnCount, ByVal 0)
            'Make sure the write is finished
            res = FlushFileBuffers(hPipe)
         End If

         'Disconnect the NamedPipe
         res = DisconnectNamedPipe(hPipe)
      Loop Until nCount = 0

      'Close the pipe handle
      CloseHandle hPipe
      GlobalFree (pSD)
      End
   End Sub

