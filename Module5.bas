Attribute VB_Name = "Module5"
Option Explicit

Private Const STARTF_USESHOWWINDOW = &H1
'Private Const STARTF_USESIZE = &H2
'Private Const STARTF_USEPOSITION = &H4
'Private Const STARTF_USECOUNTCHARS = &H8
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Const HWND_BOTTOM = 1 '将窗口置于窗口列表底部
Private Const HWND_TOP = 0 '将窗口置于Z序列的顶部；Z序列代表在分级结构中，窗口针对一个给定级别的窗口显示的顺序
Private Const HWND_TOPMOST = -1 '将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const HWND_NOTOPMOST = -2 '将窗口置于列表顶部，并位于任何最顶部窗口的后面
Private Const SWP_NOMOVE = &H2 '保持当前位置 (x和y设定将被忽略)
Private Const SWP_NOSIZE = &H1 '保持当前大小 (cx和cy会被忽略)

Private Const MAX_TOOLTIP As Integer = 64
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_INFO = &H10
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1

Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_HIDE = 0
Private Const SW_RESTORE = 9

Private Type STARTUPINFO
          cb   As Long
          lpReserved   As String
          lpDesktop   As String
          lpTitle   As String
          dwX   As Long
          dwY   As Long
          dwXSize   As Long
          dwYSize   As Long
          dwXCountChars   As Long
          dwYCountChars   As Long
          dwFillAttribute   As Long
          dwFlags   As Long
          wShowWindow   As Integer
          cbReserved2   As Integer
          lpReserved2   As Long
          hStdInput   As Long
          hStdOutput   As Long
          hStdError   As Long
End Type
Private Type PROCESS_INFORMATION
          hProcess   As Long
          hThread   As Long
          dwProcessID   As Long
          dwThreadID   As Long
End Type

'Private Type NOTIFYICONDATA
' cbSize As Long
' hwnd As Long
' uID As Long
' uFlags As Long
' uCallbackMessage As Long
' hIcon As Long
' szTip As String * MAX_TOOLTIP
'End Type
Private Type NOTIFYICONDATA2 'Extended version
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutAndVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type


Private Declare Function CreateProcessA Lib "kernel32 " (ByVal lpApplicationName As Long, _
          ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal _
          lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal _
          dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal _
          lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, _
          lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32 " (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function showWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

'Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwmainssage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwmainssage As Long, lpData As NOTIFYICONDATA2) As Long

Private nfIconData As NOTIFYICONDATA2

Public Sub ShellAndWait(cmdline As String, Optional visible As Boolean = False) '同步的Shell
 Dim NameOfProc As PROCESS_INFORMATION, NameStart As STARTUPINFO, X As Long
 With NameStart
'  .dwFlags = STARTF_USESIZE + STARTF_USEPOSITION
   .dwFlags = STARTF_USESHOWWINDOW
  If Not visible Then
   '.dwXSize = 0: .dwYSize = 0: .dwX = -32767: .dwY = -32767
   .wShowWindow = SW_HIDE
  End If
'  .dwXCountChars = 0: .dwYCountChars = 0
 End With
 NameStart.cb = LenB(NameStart)
 X = CreateProcessA(0&, cmdline, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
 X = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
 X = CloseHandle(NameOfProc.hProcess)
End Sub

Public Sub setWindowTopMost(hwnd As Long, Optional atTop As Boolean = True)
 If hwnd < 0 Then Exit Sub
 If atTop Then
  SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
 Else
  SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
 End If
End Sub
Sub doShow()
 showWindow FormPanel.hwnd, SW_RESTORE
 setWindowTopMost FormPanel.hwnd
 FormPanel.Top = FormPanel.windowTop
 FormPanel.Left = FormPanel.windowLeft
 FormPanel.SetFocus
 Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
 nfIconData.hwnd = 0
End Sub
Sub hideToTray()
 With nfIconData
    .hwnd = FormPanel.hwnd
    .uID = FormPanel.Icon
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .uCallbackMessage = WM_MDIACTIVATE
    .hIcon = FormPanel.Icon.handle
    .szTip = "WiChat: " & nowID & vbCrLf & "State: " & stateToString(nowState) & vbNullChar
    .cbSize = LenB(nfIconData)
 End With
 Shell_NotifyIcon NIM_ADD, nfIconData
 FormPanel.Hide
End Sub
Sub TrayBalloon(ByVal sBaloonText As String, ByVal sBallonTitle As String)
 If nfIconData.hwnd = 0 Then Exit Sub
 With nfIconData
    .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
    .dwInfoFlags = NIF_INFO
    .szInfoTitle = sBallonTitle & vbNullChar
    .szInfo = sBaloonText & vbNullChar
 End With
 Shell_NotifyIcon NIM_MODIFY, nfIconData
End Sub
