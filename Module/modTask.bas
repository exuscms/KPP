Attribute VB_Name = "modTask"
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Function getalltopwindows(ByVal hWnd As Long, ByVal lParam As Long) As Long

Dim foregroundwindow As Long
Dim textlen As Long
Dim windowtext As String
Dim svar As Long
Static lastwindowtext As String

foregroundwindow = hWnd

textlen = GetWindowTextLength(foregroundwindow) + 1

windowtext = Space(textlen)
svar = GetWindowText(foregroundwindow, windowtext, textlen)
windowtext = Left(windowtext, Len(windowtext) - 1)

If windowtext = "" Then GoTo slask

If IsWindowVisible(foregroundwindow) > 0 Then

frmTask.lstTask.AddItem windowtext
frmTask.lstTask.ItemData(frmTask.lstTask.NewIndex) = foregroundwindow
lastwindowtext = windowtext

Else

frmTask.lstTask.AddItem windowtext
frmTask.lstTask.ItemData(frmTask.lstTask.NewIndex) = foregroundwindow
lastwindowtext = windowtext

End If
slask:

getalltopwindows = 1
End Function



