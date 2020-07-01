Attribute VB_Name = "NPClientModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants and functions used by this program.
Private Declare Function CallNamedPipeA Lib "Kernel32.dll" (ByVal lpNamedPipeName As String, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, ByVal nTimeOut As Long) As Long

'This procedure is executed when this program started.
Public Sub Main()
Dim BytesRead As Long
Dim InputBuffer(&H0& To &H400&) As Byte
Dim OutputBuffer() As Byte
Dim PipeName As String

   ReDim OutputBuffer(&H0& To &H400&) As Byte
   PipeName = "\\.\pipe\namedpipe"
   CheckForError CallNamedPipeA(PipeName, InputBuffer(0), UBound(InputBuffer()) - LBound(InputBuffer()), OutputBuffer(0), UBound(OutputBuffer()) - LBound(OutputBuffer()), BytesRead, CLng(30000)), "CllNP"

   If BytesRead > &H0 Then
      ReDim Preserve OutputBuffer(LBound(OutputBuffer()) To BytesRead) As Byte
      MsgBox "The following data was recieved: " & vbCr & """" & CStr(OutputBuffer()) & ".""", vbInformation
   Else
      MsgBox "No was data received", vbExclamation
   End If
End Sub


