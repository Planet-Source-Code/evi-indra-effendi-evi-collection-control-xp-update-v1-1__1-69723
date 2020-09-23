Attribute VB_Name = "ModGlobal"
Option Explicit

Private Declare Function ComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Property Get GetComputerName() As String
Dim sBuffer As String
Dim lAns As Long

On Error GoTo Error

sBuffer = Space$(255)
lAns = ComputerName(sBuffer, 255)

If lAns <> 0 Then
   GetComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
End If
    
Error:
End Property
