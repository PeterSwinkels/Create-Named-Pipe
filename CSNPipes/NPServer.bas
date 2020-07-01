Attribute VB_Name = "NPServerModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants, functions, and structures used by this program.
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const PIPE_ACCESS_DUPLEX As Long = &H3&
Private Const PIPE_NOWAIT As Long = &H1&
Private Const PIPE_TYPE_MESSAGE As Long = &H4&
Private Const PIPE_WAIT As Long = &H0&
Private Const PIPE_UNLIMITED_INSTANCES As Long = &HFF&

Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function ConnectNamedPipe Lib "Kernel32.dll" (ByVal hNamedPipe As Long, lpOverlapped As Any) As Long
Private Declare Function CreateNamedPipeA Lib "Kernel32.dll" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As Any) As Long
Private Declare Function ReadFile Lib "Kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "Kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

'This procedure is executed when this program started.
Public Sub Main()
Dim Buffer() As Byte
Dim BytesWritten As Long
Dim Message As String
Dim PipeH As Long
Dim PipeName As String

   Message = InputBox$("Enter a message to be sent:", , "Hi.")
   PipeName = "\\.\pipe\namedpipe"
   If Not PipeH = INVALID_HANDLE_VALUE Then
      PipeH = CheckForError(CreateNamedPipeA(PipeName, PIPE_ACCESS_DUPLEX, PIPE_WAIT Or PIPE_TYPE_MESSAGE, PIPE_UNLIMITED_INSTANCES, CLng(&H400&), CLng(&H400&), CLng(30000), ByVal CLng(0)), "CrNP")
      CheckForError ConnectNamedPipe(PipeH, ByVal CLng(0)), "CNP"
      Buffer() = Message
      CheckForError WriteFile(PipeH, Buffer(0), UBound(Buffer()) - LBound(Buffer()), BytesWritten, ByVal CLng(0)), "WF"
      CheckForError CloseHandle(PipeH), "CH"
   End If
End Sub


