Attribute VB_Name = "modSpecialFolders"
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Function WinDir() As String
WinDir = String(255, Chr$(0))
GetWindowsDirectory WinDir, 255
WinDir = Left(WinDir, InStr(1, WinDir, Chr$(0)) - 1)
End Function
Function SysDir() As String
SysDir = String(255, Chr$(0))
GetSystemDirectory SysDir, 255
SysDir = Left(SysDir, InStr(1, SysDir, Chr$(0)) - 1)
End Function
Function TempDir() As String
TempDir = String(255, Chr$(0))
GetTempPath 255, TempDir
TempDir = Left(TempDir, InStr(1, TempDir, Chr$(0)) - 1)
End Function


