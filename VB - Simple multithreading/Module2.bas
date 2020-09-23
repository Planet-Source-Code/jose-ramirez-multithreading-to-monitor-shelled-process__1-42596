Attribute VB_Name = "Module2"
Option Explicit

'Thread data
Type ThreadDataType
    hwnd As Long
    hWndEdit As Long
    ExeName As Long
    Action As Long
End Type

'Thread handle
Public hThreadExe As Long

'Event handle to indicate the StartExe thread to kill the process and exit
Public hEventExit As Long
'Event handle to signal there is no process
Public hEventNoProcess As Long
'Event handle to signal the exe thread that no further messages should be sent to the
'main thread
Public hEventNoMessages As Long
