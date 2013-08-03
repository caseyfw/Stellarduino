Attribute VB_Name = "ProcessPriority"
Option Explicit

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByVal lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_SET_INFORMATION As Long = &H200
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const BELOW_NORMAL_PRIORITY_CLASS = 16384
Private Const ABOVE_NORMAL_PRIORITY_CLASS = 32768
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100

Public Enum ProcessPriorities
    ppidle = IDLE_PRIORITY_CLASS
    ppbelownormal = BELOW_NORMAL_PRIORITY_CLASS
    ppAboveNormal = ABOVE_NORMAL_PRIORITY_CLASS
    ppNormal = NORMAL_PRIORITY_CLASS
    ppHigh = HIGH_PRIORITY_CLASS
    ppRealtime = REALTIME_PRIORITY_CLASS
End Enum

Public Enum ThreadPriority
    THREAD_PRIORITY_LOWEST = -2
    THREAD_PRIORITY_BELOW_NORMAL = -1
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_HIGHEST = 2
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    THREAD_PRIORITY_TIME_CRITICAL = 15
    THREAD_PRIORITY_IDLE = -15
End Enum

Public Function ProcessPriorityGet(Optional ByVal processId As Long, Optional ByVal hwnd As Long) As Long
    Dim hProc As Long
    Const fdwAccess As Long = PROCESS_QUERY_INFORMATION

    If processId = 0 Then
        If hwnd <> 0 Then
            Call GetWindowThreadProcessId(hwnd, processId)
        Else
            processId = GetCurrentProcessId()
        End If
    End If

    hProc = OpenProcess(fdwAccess, 0&, processId)
    ProcessPriorityGet = GetPriorityClass(hProc)
    Call CloseHandle(hProc)

End Function

Public Function ProcessPrioritySet(Optional ByVal processId As Long, Optional ByVal hwnd As Long, _
                    Optional ByVal priority As ProcessPriorities = NORMAL_PRIORITY_CLASS) As Long

    Dim hProc As Long
    Const fdwAccess1 As Long = PROCESS_QUERY_INFORMATION Or PROCESS_SET_INFORMATION
    Const fdwAccess2 As Long = PROCESS_QUERY_INFORMATION

    If processId = 0 Then
        If hwnd <> 0 Then
            Call GetWindowThreadProcessId(hwnd, processId)
        Else
            processId = GetCurrentProcessId()
        End If
    End If

    hProc = OpenProcess(fdwAccess1, 0&, processId)
    If hProc Then
        Call SetPriorityClass(hProc, priority)
    Else
        ' enable return of current priority setting.
        hProc = OpenProcess(fdwAccess2, 0&, processId)
    End If

    ProcessPrioritySet = GetPriorityClass(hProc)
    
    Call CloseHandle(hProc)

End Function

Public Function ProcessThreadPrioritySet(Optional ByVal priority As ThreadPriority = THREAD_PRIORITY_NORMAL) As ThreadPriority
    Dim hThread As Long
    Dim rc As Long

    hThread = GetCurrentThread()
    rc = SetThreadPriority(hThread, priority)
    ProcessThreadPrioritySet = GetThreadPriority(hThread)
End Function

Public Sub ReadProcessPriority()
    Dim tmptxt As String
    Dim priority As Long
    
    tmptxt = HC.oPersist.ReadIniValue("ProcessPrioirty")
    Select Case tmptxt
        Case "0"
            priority = ProcessPrioritySet(, , ppNormal)
        Case "1"
            priority = ProcessPrioritySet(, , ppAboveNormal)
        Case "2"
            priority = ProcessPrioritySet(, , ppHigh)
        Case "3"
            priority = ProcessPrioritySet(, , ppRealtime)
        Case Else
            priority = ProcessPrioritySet(, , ppNormal)
    End Select

End Sub
