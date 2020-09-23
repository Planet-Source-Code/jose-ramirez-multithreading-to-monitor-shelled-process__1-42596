Attribute VB_Name = "Module1"
Option Explicit

'Threading stuff
Public Const INFINITE As Long = &HFFFFFFFF
Public Const WAIT_TIMEOUT As Long = 258&
Public Const STATUS_WAIT_0 As Long = &H0
Public Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

'Synchronization stuff
Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Public Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Public Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


'ShellExecute stuff
Private Const SEE_MASK_FLAG_DDEWAIT As Long = &H100
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40

Public Const SW_NORMAL As Long = 1

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpstrVerb As Long
    lpstrFile As Long
    lpstrParameters As Long
    lpstrDirectory As Long
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpstrClass As Long
    hkeyClass As Long
    dwHotKey As Long
    hIconOrMon As Long
    hProcess As Long
End Type

Public Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long


'Window message stuff
Public Const WM_SETTEXT As Long = &HC

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Misc stuff
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Function ThreadStartExe(ByVal lpThreadData As Long) As Long

Dim udtData As ThreadDataType
Dim lResp As Long
Dim bFinished As Boolean
Dim bExit As Boolean
Dim lThreadID As Long
Dim strMessage As String
Dim sei As SHELLEXECUTEINFO

    CopyMemory udtData, ByVal lpThreadData, Len(udtData)
    If (WaitForSingleObject(hEventNoMessages, 0) = WAIT_TIMEOUT) Then
        strMessage = "Starting process..."
        SendMessage udtData.hWndEdit, WM_SETTEXT, 0, ByVal strMessage
    End If
    With sei
        .cbSize = Len(sei)
        .fMask = SEE_MASK_FLAG_DDEWAIT Or SEE_MASK_NOCLOSEPROCESS
        .hwnd = udtData.hwnd
        .lpstrVerb = udtData.Action
        .lpstrFile = udtData.ExeName
        .nShow = SW_NORMAL
    End With
    lResp = ShellExecuteEx(sei)
    strMessage = "Process started"
    SendMessage udtData.hWndEdit, WM_SETTEXT, 0, ByVal strMessage
    If (lResp = 0) Then
        ThreadStartExe = 1 '1 indicates error
    Else
        If (sei.hInstApp <= 32) Then
            ThreadStartExe = 1
        Else
            bFinished = False
            bExit = False
            Do While (Not (bExit) And Not (bFinished))
                bExit = (WaitForSingleObject(hEventExit, 0) = WAIT_OBJECT_0)
                bFinished = (WaitForSingleObject(sei.hProcess, 200) = WAIT_OBJECT_0)
            Loop
            If bExit Then
                TerminateProcess sei.hProcess, 0
            End If
            CloseHandle sei.hProcess
            ThreadStartExe = IIf(bFinished, 0, 2)
        End If
    End If
    SetEvent hEventNoProcess
    If (WaitForSingleObject(hEventNoMessages, 0) = WAIT_TIMEOUT) Then
        strMessage = "No current process"
        SendMessage udtData.hWndEdit, WM_SETTEXT, 0, ByVal strMessage
    End If
End Function
