VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch and monitor process"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frProcess 
      Caption         =   "Process to launch"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox ddAction 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cbBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox tbFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "C:\Winnt\notepad.exe"
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Action:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "File name:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cbStop 
      Caption         =   "Stop process"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox tbStatus 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "No external process"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton cbStart 
      Caption         =   "Start process"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrOptions As Variant

Private udtData As ThreadDataType
Private arrExe() As Byte
Private arrAction() As Byte

Private Sub cbBrowse_Click()
    With cdDialog
        .CancelError = True
        .DialogTitle = "Select document to shell"
        .filename = tbFilename.Text
        .Filter = "Executable files (*.exe)|*.exe|Office documents (*.xls, *.doc, *.ppt)|*.xls; *.doc; *.ppt|All files (*.*)|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly
        On Error Resume Next
        .ShowOpen
        If (Err.Number <> 0) Then
            Exit Sub
        End If
        tbFilename.Text = .filename
    End With
End Sub

Private Sub cbStart_Click()
    cbStart.Enabled = False
    frProcess.Enabled = False
    LaunchApp tbFilename.Text, ddAction.Text
    cbStop.Enabled = True
End Sub

Private Sub LaunchApp(ByVal strApp As String, ByVal strAction As String)

Dim lThreadID As Long

    arrExe() = StrConv(strApp & vbNullChar, vbFromUnicode)
    arrAction() = StrConv(strAction & vbNullChar, vbFromUnicode)
    With udtData
        .ExeName = VarPtr(arrExe(0))
        .Action = VarPtr(arrAction(0))
        .hwnd = Me.hwnd
        .hWndEdit = tbStatus.hwnd
    End With
    If (hEventExit = 0) Then
        hEventExit = CreateEvent(0, 1, 0, "ExitStartExe")
    Else
        ResetEvent hEventExit
    End If
    If (hEventNoProcess = 0) Then
        hEventNoProcess = CreateEvent(0, 1, 0, "SignalNoShellProcess")
    Else
        ResetEvent hEventNoProcess
    End If
    If (hEventNoMessages = 0) Then
        hEventNoMessages = CreateEvent(0, 1, 0, "SignalNoMessages")
    Else
        ResetEvent hEventNoMessages
    End If
    hThreadExe = CreateThread(0, 0, AddressOf ThreadStartExe, udtData, 0, lThreadID)
    If (hThreadExe = 0) Then
        MsgBox "Could not create thread."
    End If
End Sub

Private Sub KillExe()
    If (hEventNoMessages <> 0) Then
        SetEvent hEventNoMessages
    End If
    If (hEventExit <> 0) Then
        SetEvent hEventExit
    End If
    If (hThreadExe <> 0) Then
        WaitForSingleObject hThreadExe, INFINITE
        CloseHandle hThreadExe
        CloseHandle hEventNoMessages
        CloseHandle hEventExit
        CloseHandle hEventNoProcess
        hThreadExe = 0
        hEventNoMessages = 0
        hEventExit = 0
        hEventNoProcess = 0
    End If
End Sub

Private Sub cbStop_Click()
    cbStop.Enabled = False
    KillExe
    EnableThreadControls
End Sub

Private Sub Form_Load()

Dim lCount As Long

    arrOptions = Array("edit", "explore", "find", "open", "print", "properties")
    With ddAction
        For lCount = LBound(arrOptions) To UBound(arrOptions)
            .AddItem arrOptions(lCount)
        Next lCount
        .Text = "open"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillExe
End Sub

Private Sub tbStatus_Change()
    If (hEventNoProcess <> 0) Then
        If (WaitForSingleObject(hEventNoProcess, 0) = WAIT_OBJECT_0) Then
            Beep
            EnableThreadControls
        End If
    End If
End Sub

Private Sub EnableThreadControls()
    cbStop.Enabled = False
    cbStart.Enabled = True
    frProcess.Enabled = True
End Sub
