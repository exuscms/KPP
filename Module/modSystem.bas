Attribute VB_Name = "modSystem"
Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Boolean) As Long
'Setupapi.dll
Public Declare Function SetupPromptReboot Lib "setupapi.dll" (ByRef FileQueue As Long, ByVal Owner As Long, ByVal ScanOnly As Long) As Long
'Shell32.dll
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Winmm.dll
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
'Ntdll.dll
Public Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal bEnablePrivilege As Long, ByVal IsThreadPrivilege As Long, ByRef PreviousValue As Long) As Long
'Advapi32.dll
Public Declare Function InitializeAcl Lib "advapi32.dll" (ByRef pAcl As ACL, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long

Public Const ACL_REVISION As Long = 2
Public Const DACL_SECURITY_INFORMATION As Long = &H4&
Public Const SECURITY_DESCRIPTOR_REVISION As Long = 1

Public Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type
Public Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long

Private Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As ACL
    Dacl As ACL
End Type

Public Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, ByRef pDacl As Any, ByVal bDaclDefaulted As Long) As Long
Public Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Public Declare Function RegCreateKey& Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, lphKey As Long)
Public Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpszSubKey As String, ByVal fdwType As Long, ByVal lpszValue As String, ByVal dwLength As Long)
Public Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
'User32.dll
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
'User32
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000

Public Type DEVMODE
   dmDeviceName As String * CCDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ExitWindows Lib "user32" Alias "ExitWindowsEx" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Type PointAPI
        x As Long
        y As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'Kernel32.dll
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function BeepAPI Lib "kernel32.dll" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)

Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Public Declare Sub RaiseException Lib "kernel32.dll" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, ByRef lpArguments As Long)
Public Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'Kernel32
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function IsSystemResumeAutomatic Lib "kernel32" () As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function AllocConsole Lib "kernel32" () As Long
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Public Declare Function FreeConsole Lib "kernel32" () As Long
Public Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Public Const STD_OUTPUT_HANDLE = -11&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const MAX_PATH = 256&
Public Const REG_SZ = 1
Public CS As CREATESTRUCT

Public Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Const HWND_BROADCAST& = &HFFFF&
Public Const WM_SYSCOMMAND& = &H112&
Public Const SC_MONITORPOWER& = &HF170&
Public Const MONITOR_ON& = &HFFFFFFFF
Public Const MONITOR_OFF& = 2&
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Const INVALID_HANDLE_VALUE = -1

Public hHook As Long
Public Codes As String
Public SendInfo As Boolean

Public Function GetMainSerialNumber() As String
    Dim VolumeSerialNumber As Long
    If GetVolumeInformation("C:\", vbNullString, 0&, VolumeSerialNumber, ByVal 0&, ByVal 0&, vbNullString, 0&) Then
        GetMainSerialNumber = Left$(Hex$(VolumeSerialNumber), 4) & "-" & Mid$(Hex$(VolumeSerialNumber), 5, 4)
    Else
        GetMainSerialNumber = "0000-0000"
    End If
End Function

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long
With lParam

If nCode < 0 Or .dwExtraInfo = 33 Then
LowLevelKeyboardProc = CallNextHookEx(hHook, nCode, wParam, lParam)
Exit Function
End If

If .flags = 0 Then
Codes = Chr(.vkCode)

End If
End With
End Function

Public Sub MPControl(ByVal Directory As String, ByVal Play As Boolean)
    On Error GoTo Errors:
 Dim strCon As String

       strCon = IIf(Play = True, "play mp", "close mp")
       
       Call mciSendString("open " & Chr(34) & Directory & Chr(34) & " alias mp wait", vbNullString, 128, 0)
       
       Call mciSendString(strCon, vbNullString, 128, 0)
Errors:
frmMain.Debugs (Directory & "이 존재하지 않거나 지원되지 않는 파일")
End Sub

Public Sub Playa(ByVal Directory As String, ByVal Play As Boolean)
    On Error GoTo Errors:
 Dim strCon As String

       strCon = IIf(Play = True, "play mp", "close mp")
       
       Call mciSendString("open " & Chr(34) & Directory & Chr(34) & " alias mp wait", vbNullString, 128, 0)
       
       Call mciSendString(strCon, vbNullString, 128, 0)
Errors:
End Sub

Public Sub TurnOnMonitor()
    SendMessage HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_ON
End Sub

Public Sub TurnOffMonitor()
    SendMessage HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_OFF
End Sub

Public Function SetFolderTime(strFolderName As String, aryTime() As Long) As Boolean
    Dim SetTime(2)     As SYSTEMTIME
    Dim TimeCreate     As FILETIME
    Dim TimeLastAccess As FILETIME
    Dim TimeLastModify As FILETIME
    Dim SA             As SECURITY_ATTRIBUTES
    Dim hDir As Long
    Dim BLK  As Boolean
    
    
    '## 생성일자 설정
    With SetTime(0)
        .wYear = aryTime(0, 0)
        .wMonth = aryTime(0, 1)
        .wDay = aryTime(0, 2)
        .wDayOfWeek = aryTime(0, 3)
        .wHour = aryTime(0, 4)
        .wMinute = aryTime(0, 5)
        .wSecond = aryTime(0, 6)
    End With
    
    '## 최종수정일자 설정
    With SetTime(1)
        .wYear = aryTime(1, 0)
        .wMonth = aryTime(1, 1)
        .wDay = aryTime(1, 2)
        .wDayOfWeek = aryTime(1, 3)
        .wHour = aryTime(1, 4)
        .wMinute = aryTime(1, 5)
        .wSecond = aryTime(1, 6)
    End With
    
    '## 최종 엑세스일자 설정
    With SetTime(2)
        .wYear = aryTime(2, 0)
        .wMonth = aryTime(2, 1)
        .wDay = aryTime(2, 2)
        .wDayOfWeek = aryTime(2, 3)
        .wHour = aryTime(2, 4)
        .wMinute = aryTime(2, 5)
        .wSecond = aryTime(2, 6)
    End With
    
    hDir = CreateFile(strFolderName, _
                      GENERIC_WRITE, _
                      FILE_SHARE_READ, _
                      SA, _
                      OPEN_EXISTING, _
                      FILE_FLAG_BACKUP_SEMANTICS, _
                      0)
    
    If hDir <> INVALID_HANDLE_VALUE Then
        Call SystemTimeToFileTime(SetTime(0), TimeCreate)
        Call SystemTimeToFileTime(SetTime(1), TimeLastModify)
        Call SystemTimeToFileTime(SetTime(2), TimeLastAccess)
        
        SetFolderTime = SetFileTime(hDir, TimeCreate, TimeLastAccess, TimeLastModify)
        
        Call CloseHandle(hDir)
    End If
End Function


Private Function SetDirProtect(ByVal strTarget As String, bFlag As Boolean) As Boolean
 Dim SEC As SECURITY_DESCRIPTOR
 Dim ACL As ACL
 
   On Error GoTo Err
       
    Call InitializeSecurityDescriptor(SEC, SECURITY_DESCRIPTOR_REVISION)
    
    If bFlag Then
        Call InitializeAcl(ACL, Len(ACL), ACL_REVISION)
        Call SetSecurityDescriptorDacl(SEC, True, ACL, True)
    Else
        Call SetSecurityDescriptorDacl(SEC, True, ByVal 0&, True)
    End If
    
    If SetFileSecurity(strTarget, DACL_SECURITY_INFORMATION, SEC) Then
        SetDirProtect = True
    Else
        GoTo Err
    End If
    Exit Function
   
Err:
   SetDirProtect = False
   
   On Error GoTo 0
   
End Function

'Public Sub GetDiskSpace(ByVal strRoot As String, ByRef types As Long)
'    Dim curByte(2) As Currency
'    Call GetDiskFreeSpaceEx(strRoot & ":\", curByte(0), curByte(1), curByte(2))
'    If types = 0 Then
'        GetDiskSpace = Int(curByte(1) / 100)               '## 전체 디스크 공간
'    ElseIf types = 1 Then
'        GetDiskSpace = Int(curByte(0) / 100)                '## 사용가능한 디스크 공간
'    ElseIf types = 2 Then
'        GetDiskSpace = Int((curByte(1) - curByte(2)) / 100)  '## 이미 사용한 디스크 공간
'    End If
'End Sub


Public Sub SetDescDSP(ByVal x As Long, ByVal y As Long)
  Dim DM As DEVMODE
  
    Call EnumDisplaySettings(0&, 0&, DM)
    
    With DM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        .dmPelsWidth = x
        .dmPelsHeight = y
    End With
    Call ChangeDisplaySettings(DM, CDS_TEST)
End Sub

