Attribute VB_Name = "modFileFunctions"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long

Const SW_SHOWNORMAL = 1

Function StartDoc(ByVal DocName As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "", SW_SHOWNORMAL)
End Function

Function XGetSize(ByVal Filename As String) As Long
On Error GoTo ErrorOccured
Dim freef
freef = FreeFile
XGetSize = VBA.FileSystem.FileLen(Filename)
Exit Function
ErrorOccured:
XGetSize = -1
End Function
Function XGetFileName(ByVal Filename As String) As String
Dim i
i = InStrRev(Filename, "\")
If i > 0 Then
XGetFileName = Mid(Filename, i + 1)
End If
End Function
Function XGetParentFolder(ByVal Filename As String) As String
Dim i
i = InStrRev(Filename, "\")
If i > 0 Then
XGetParentFolder = Left(Filename, i)
End If
End Function

Function XBuildPath(ByVal sPath As String, sFileName As String)
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
XBuildPath = sPath & sFileName
End Function
Function IsDirectoryExist(ByVal SomePath As String) As Boolean
On Error GoTo ErrorOccured
ChDir SomePath
IsDirectoryExist = True
Exit Function
ErrorOccured:
IsDirectoryExist = False
End Function


Function IsFileExist(ByVal Filename As String) As Boolean
On Error GoTo ErrorOccured
Dim freef
freef = FreeFile
Open Filename For Input As freef: Close freef
IsFileExist = True
Exit Function
ErrorOccured:
IsFileExist = False
End Function

Sub DeleteFile(ByVal Filename)
On Error Resume Next
Kill Filename
End Sub
