Attribute VB_Name = "Module1"
Option Explicit
#If IS_TEST Then
Public Const VER As String = "1.12 - Test Version"
Public Const QueryDevice As Integer = 255
#Else
Public Const VER As String = "1.12"
Public Const QueryDevice As Integer = 4
#End If


Public Const MaxIDLen As Integer = 8
Public Const SessionLen As Integer = 16
Public Const KeyLen As Integer = 16
Public Const MaxOfflineMsg As Integer = 64
Public Const MaxNotation As Integer = 16
Public Const MaxMsgBlock As Long = 50# * 1024 - 16
Public Const MaxCapableFileSize As Long = 1024# * 1024
'Unit:Byte

Private Const MaxSetting As Integer = 31, MaxUserSetting As Integer = 31
Public Const MaxNote As Integer = 64, MaxSession As Integer = 8, MaxSessionKey As Integer = 128
Public Const MaxSessionTime As Long = 1200 'In second
Public Const MaxEmotion As Integer = 66

Private Const RESPONSE_RES_OK As Integer = 0
Private Const RESPONSE_RES_NOT_EXIST As Integer = 1
Private Const RESPONSE_RES_SIZE_TOO_LARGE As Integer = 2
Private Const RESPONSE_RES_EOF As Integer = 3
Private Const RESPONSE_RES_OUT_RANGE As Integer = 4

Private Const Action_Acc_Change As Integer = 1
Private Const Action_Fri_Change As Integer = 2
Private Const Action_Msg_Exchange As Integer = 3
Private Const Action_Msg_Get_List As Integer = 4
Private Const Action_Msg_Get_Key As Integer = 5

Public Const blankPage As String = "about:blank"
Public Const textboxName As String = "tC"
Private Const nullKey As String = "0000000000000000"
Public Const TimeFormat As String = "yyyy-mm-dd hh:mm:ss"

Public Enum State
 none = 0
 onLine = 1
 Offline = 2
 Busy = 4
 Hide = 5
End Enum
Public Enum NoteEvent
 none = 0
 FriendAdd = 1
 FriendDelete = 2
 GotMsg = 3
End Enum

Public Type fontStyle
 family As String
 size As Integer
 color As String
 basic As Integer '4*4 bit indicates Delete,Underline,Italic,Bold
 align As Integer '0=none;1=left;2=center;3=right
End Type
Public Type Notification
 type As NoteEvent
 time As Date
 source As String
 destination As String
 handle As Long
End Type
Public Type Session
 id As String
 active As Boolean
 cache As String
 input As String
End Type
Private Type sessionKey
 id As String
 key As String
 updateTime As String
End Type
Public Type IDInfo
 id As String
 note As String
 offlineMsg As String
End Type

Private WM_TASKBARCREATED As Long
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONUP = &H205
'Private Const WM_CLOSE = &H10
'Private Const WM_QUIT = &H12
'Private Const WM_DESTROY = &H2
Private Const WM_HOTKEY = &H312
Public Const WM_MDIACTIVATE = &H222

Private Const GWL_WNDPROC = (-4)
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const MAX_PATH As Integer = 260


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
'注册系统级热键所必需

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public nowID As String
Public nowSession As String
Public nowState As State
Public sessionValid As Long
Public nowOfflineMsg As String

Private sessionKey As String, keySalt As String

Private sessionList(MaxSession) As Session, sessionCount As Integer
Private sessionKeyList(MaxSessionKey) As sessionKey, pSessionKeyList As Integer 'Sender's Key
Private sessionKeyList2(MaxSessionKey) As sessionKey, pSessionKeyList2 As Integer 'Receiver's Key
Private sessionKeyLoaded As Boolean

Private noteStack(MaxNote) As Notification, pNoteStack As Integer, nowHandle As Long
Private tempWebPage As String

Private config(MaxSetting) As String, userConfig(MaxSetting) As String
'Config index:
'1=Last ID
'2=Last Programme State
'3=Chating Record Path
'4=Chating Record Path with EXE

'User config index:
'1=Last Login State
'2=Last unhandled session(split in ';')
'3=Preferred font family
'4=Preferred font size
'5=Preferred font color
'6=Preferred font basic
'7=Preferred font align
'8=Preferred send shortcut key
'9=Preferred setting when capturing screen
'10=Preferred visualized notifying method

Private preWinProc As Long



Public Property Get lastID() As String
 lastID = config(1)
End Property
Public Property Let lastID(value As String)
 config(1) = value
End Property
Public Property Get lastState() As String
 lastState = Val(config(2))
End Property
Public Property Get recordPath(Optional fillEmpty As Boolean = True) As String
 If config(4) = "1" And fillEmpty Then
 #If IS_WINE Then
  recordPath = AppPath
 #Else
  recordPath = App.path
 #End If
 Else
  recordPath = config(3)
 End If
End Property
Public Property Let recordPath(Optional fillEmpty As Boolean = True, value As String)
 If value = vbNullString Then
  config(3) = vbNullString: config(4) = "1"
 Else
  config(3) = TrimEnd(Trim(value), "\")
  config(4) = "0"
 End If
End Property

Public Property Get lastUserState() As State
 lastUserState = Val(userConfig(1))
End Property
Public Property Let lastUserState(value As State)
 userConfig(1) = str(value)
End Property
Public Property Get lastUserFont() As fontStyle
 lastUserFont.family = userConfig(3)
 lastUserFont.size = Val(userConfig(4))
 lastUserFont.color = userConfig(5)
 lastUserFont.basic = Val(userConfig(6))
 lastUserFont.align = Val(userConfig(7))
End Property
Public Property Let lastUserFont(value As fontStyle)
 userConfig(3) = value.family
 userConfig(4) = str(value.size)
 userConfig(5) = value.color
 userConfig(6) = str(value.basic)
 userConfig(7) = str(value.align)
End Property
Public Property Get lastUserSendOperation() As Integer
 lastUserSendOperation = Val(userConfig(8))
End Property
Public Property Let lastUserSendOperation(value As Integer)
 userConfig(8) = str(value)
End Property
Public Property Get lastUserCaptureHide() As Integer
 lastUserCaptureHide = Val(userConfig(9))
End Property
Public Property Let lastUserCaptureHide(value As Integer)
 userConfig(9) = str(value)
End Property
Public Property Get lastUserVisualNotification() As Integer
 lastUserVisualNotification = Val(userConfig(10))
End Property
Public Property Let lastUserVisualNotification(value As Integer)
 userConfig(10) = str(value)
End Property

Public Function Note_count() As Integer
 Note_count = pNoteStack
End Function
Public Function Note_peek(handle As Integer) As Notification
 Dim temp As Notification, i As Integer
 temp.type = NoteEvent.none
 For i = 1 To pNoteStack
  If noteStack(i).handle = handle Then
   temp = noteStack(i)
   Exit For
  End If
 Next i
 Note_peek = temp
End Function
Private Sub Note_Remove(index As Integer)
 Dim j As Integer
 If pNoteStack < 1 Then Exit Sub
 For j = index + 1 To pNoteStack
    noteStack(j - 1) = noteStack(j)
 Next j
 pNoteStack = pNoteStack - 1
End Sub
Public Function Note_get() As Notification()
 Dim temp() As Notification, i As Integer
 ReDim temp(0)
 temp(0).type = NoteEvent.none
 Note_get = temp
 If pNoteStack <= 0 Then Exit Function
 ReDim temp(pNoteStack)
 For i = 1 To pNoteStack
  temp(i) = noteStack(i)
 Next i
 Note_get = temp
End Function
Private Function Note_push(value As Notification)
 Note_push = False
 If pNoteStack >= MaxNote Then Exit Function
 Dim i As Integer
 For i = 1 To pNoteStack
  If noteStack(i).source = value.source And noteStack(i).destination = value.destination And noteStack(i).type = value.type Then Exit Function
 Next i
 pNoteStack = pNoteStack + 1
 noteStack(pNoteStack) = value
End Function
Private Sub Note_clear(Optional kind As NoteEvent = NoteEvent.none)
 Dim i As Integer, j As Integer
 For i = 1 To pNoteStack
  If noteStack(pNoteStack).type = kind Or kind = NoteEvent.none Then
   For j = i + 1 To pNoteStack
    noteStack(j - 1) = noteStack(j)
   Next j
   pNoteStack = pNoteStack - 1
  End If
 Next i
End Sub

Public Function Session_Exist(id As String) As Boolean
 Session_Exist = True
 Dim i As Integer
 For i = 1 To MaxSession
  If sessionList(i).id = id Then Exit For
 Next i
 If i > MaxSession Then Session_Exist = False
End Function
Public Function Session_Count() As Integer
 Session_Count = sessionCount
End Function
Public Function Session_Now() As String
 Session_Now = sessionList(getActiveSession).id
End Function
Public Sub Session_List(buffer() As String)
 Dim i As Integer, j As Integer
 ReDim buffer(MaxSession)
 j = -1
 For i = 1 To MaxSession
  If sessionList(i).id <> vbNullString Then
   j = j + 1
   buffer(j) = sessionList(i).id
  End If
 Next i
 If j > -1 Then ReDim Preserve buffer(j) Else Erase buffer
End Sub


Public Sub Main()
#If IS_TEST Then
  showMsg "This is a test version. No guarantee of data security was provided.", "Warning!", vbExclamation
#End If
#If IS_DEBUG Then
 test
#End If

 WM_TASKBARCREATED = RegisterWindowMessage("TaskbarCreated")
 App.TaskVisible = False
 
 reset
 loadSettings
 Module3.init
 
 Load FormMain
 FormMain.show
End Sub
Sub reset()
 nowID = vbNullString
 nowSession = vbNullString
 nowState = 0
 sessionValid = 0
 
 Dim i As Integer
 For i = 0 To MaxNote
  noteStack(i).type = NoteEvent.none
 Next i
 pNoteStack = 0
 For i = 0 To MaxSession
  sessionList(i).id = vbNullString
 Next i
 sessionCount = 0
 
 resetTempWebPage
End Sub
Sub Destroy()
 Module3.clear True
 resetTempWebPage
 End
End Sub
Private Sub loadSettings()
 Dim filen As Integer, fileName As String, i As Integer
 resetSettings
 fileName = TrimEnd(App.path, "\") & "\wichat.dat"
 If Not fileExists(fileName) Then Exit Sub
 filen = FreeFile
 Open fileName For Random Access Read As #filen Len = 128
  For i = 1 To MaxSetting
   Get #filen, i, config(i)
  Next i
 Close #filen
End Sub
Private Sub saveSettings()
 Dim filen As Integer, i As Integer
 filen = FreeFile
 Open TrimEnd(App.path, "\") & "\wichat.dat" For Random Access Write As #filen Len = 128
  For i = 1 To MaxSetting
   Put #filen, i, config(i)
  Next i
 Close #filen
End Sub
Private Sub resetSettings()
 Dim i As Integer
 For i = 0 To MaxSetting
  config(i) = vbNullString
 Next i
End Sub
Public Sub loadUserSettings()
 Dim filen As Integer, fileName As String, i As Integer
 resetUserSettings
 checkUserDirectory nowID
 filen = FreeFile
 fileName = recordPath & "\record\" & nowID & "\setting.dat"
 If fileExists(fileName) Then
  Open fileName For Random Access Read As #filen Len = 128
   For i = 1 To MaxUserSetting
    Get #filen, , userConfig(i)
   Next i
  Close #filen
 End If
End Sub
Public Sub saveUserSettings()
 Dim filen As Integer, i As Integer
 filen = FreeFile
 If Not checkUserDirectory(nowID) Then Exit Sub
 Open recordPath & "\record\" & nowID & "\setting.dat" For Random Access Write As #filen Len = 128
  For i = 1 To MaxUserSetting
   Put #filen, , userConfig(i)
  Next i
 Close #filen
 saveSessionKey
End Sub
Private Sub resetUserSettings()
 Dim i As Integer
 For i = 0 To MaxUserSetting
  userConfig(i) = vbNullString
 Next i
 userConfig(1) = 1
 userConfig(3) = "Times New Roman"
 userConfig(4) = 15
 userConfig(5) = "000000"
 userConfig(6) = 0
 userConfig(7) = 1
 userConfig(8) = 1
 userConfig(9) = 1
 userConfig(10) = 1
End Sub
Public Sub loadSessionFile()
 Dim filen As Integer, i As Integer, temp As String, char As String * 1, Count As Integer
 filen = FreeFile
 checkUserDirectory nowID
 temp = recordPath & "\record\" & nowID & "\session.dat"
 If Not fileExists(temp) Then Exit Sub
 Open temp For Binary Access Read As #filen
   i = 0
   Do
    temp = String(MaxIDLen, vbNullChar)
    Get #filen, , temp
    temp = trimNull(temp)
    If Len(temp) < 8 And Len(temp) > 1 Then i = i + 1: sessionList(i).id = temp
    If EOF(filen) Then Exit Do
    
    Get #filen, , char
    If Asc(char) = 1 Then sessionList(i).active = True Else sessionList(i).active = False
    If EOF(filen) Then Exit Do
    
    Get #filen, , char
    
    temp = vbNullString
    Count = 0
    Do
     Get #filen, , char
     If char = Chr(7) Then Count = Count + 1
     If Count > 3 Then temp = Left(temp, Len(temp) - 3): Exit Do
     temp = temp & char
    Loop Until EOF(filen)
    sessionList(i).input = temp
    If EOF(filen) Then Exit Do
    
    temp = vbNullString
    Count = 0
    Do
     Get #filen, , char
     If char = Chr(7) Then Count = Count + 1
     If Count > 3 Then temp = Left(temp, Len(temp) - 3): Exit Do
     temp = temp & char
    Loop Until EOF(filen)
    sessionList(i).cache = temp
   Loop Until EOF(filen) Or i >= MaxSession
 Close #filen
 sessionCount = i
End Sub
Public Sub saveSessionFile()
 Dim filen As Integer, i As Integer, temp As String
 filen = FreeFile
 If Not checkUserDirectory(nowID) Then Exit Sub
 If FormPanel.bindTextBox Then sessionList(getActiveSession).input = FormPanel.textArea.innerText
 temp = recordPath & "\record\" & nowID & "\session.dat"
 If fileExists(temp) Then Kill temp
 Open temp For Binary Access Write As #filen
  For i = 1 To MaxSession
   If sessionList(i).id <> vbNullString Then
    Put #filen, , formatID(sessionList(i).id)
    If sessionList(i).active Then temp = Chr(1) & vbNullChar Else temp = Chr(2) & vbNullChar
    Put #filen, , temp
    temp = sessionList(i).input & String(4, Chr(7))
    Put #filen, , temp
    temp = sessionList(i).cache & String(4, Chr(7))
    Put #filen, , temp
   End If
  Next i
 Close #filen
End Sub
Private Sub loadSessionKey()
 Dim filen As Integer, i As Integer, p As Long
 Dim temp As String, char As Byte, buffer() As Byte, buffer2() As Byte
 
 If sessionKeyLoaded Then Exit Sub
 If Not msgLogin Then Exit Sub
 
 filen = FreeFile
 checkUserDirectory nowID
 temp = recordPath & "\record\" & nowID & "\session2.dat"
 If Not fileExists(temp) Then Exit Sub
 If FileLen(temp) < 8 Then Exit Sub
 ReDim buffer(FileLen(temp) - 1)
 Open temp For Binary Access Read As #filen
  Get #filen, , buffer
 Close #filen
 sessionKeyLoaded = True
 If Not byteArrayIsDimed(buffer) Then Exit Sub
 If (UBound(buffer) + 1) Mod 8 > 0 Then ReDim Preserve buffer(CInt((UBound(buffer) + 1) \ 8) * 8 - 1)
 cDecrypt buffer, keySalt, buffer2
 i = 0
 p = 0
 Do While p < UBound(buffer2)
    temp = vbNullString
    Do
     char = buffer2(p): p = p + 1
     If char = 0 Then Exit Do
     temp = temp & Chr(char)
    Loop Until p > UBound(buffer2)
    If checkID(temp) Then
     pSessionKeyList = pSessionKeyList + 1
     sessionKeyList(pSessionKeyList).id = temp
    Else
     Exit Do
    End If
    If p > UBound(buffer2) Then Exit Do
    
    temp = vbNullString
    Do
     char = buffer2(p): p = p + 1
     If char = 0 Then Exit Do
     temp = temp & Chr(char)
    Loop Until p > UBound(buffer2)
    sessionKeyList(pSessionKeyList).key = temp
    If p > UBound(buffer2) Then Exit Do
    
    temp = vbNullString
    Do
     char = buffer2(p): p = p + 1
     If char = 0 Then Exit Do
     temp = temp & Chr(char)
    Loop Until p > UBound(buffer2)
    sessionKeyList(pSessionKeyList).updateTime = temp
    If pSessionKeyList >= MaxSessionKey Then Exit Do
 Loop
   
 Do While p < UBound(buffer2)
    temp = vbNullString
    Do
     char = buffer2(p): p = p + 1
     If char = 0 Then Exit Do
     temp = temp & Chr(char)
    Loop Until p > UBound(buffer2)
    If checkID(temp) Then
     pSessionKeyList2 = pSessionKeyList2 + 1
     sessionKeyList2(pSessionKeyList2).id = temp
    Else
     Exit Do
    End If
    If p > UBound(buffer2) Then Exit Do
    
    temp = vbNullString
    Do
     char = buffer2(p): p = p + 1
     If char = 0 Then Exit Do
     temp = temp & Chr(char)
    Loop Until p > UBound(buffer2)
    sessionKeyList2(pSessionKeyList2).key = temp
    If p > UBound(buffer2) Then Exit Do
    
    temp = vbNullString
    Do
     char = buffer2(p): p = p + 1
     If char = 0 Then Exit Do
     temp = temp & Chr(char)
    Loop Until p > UBound(buffer2)
    sessionKeyList2(pSessionKeyList2).updateTime = temp
    If pSessionKeyList2 >= MaxSessionKey Then Exit Do
 Loop
End Sub
Private Sub saveSessionKey()
 If Not sessionKeyLoaded = True Then Exit Sub
 If Not msgLogin Then Exit Sub '
 'These above are VERY IMPORTANT! Otherwise session key file will be cleared if no loading/login failed.
 
 Dim filen As Integer, i As Integer, temp As String, buffer() As Byte, buffer2() As Byte
 filen = FreeFile
 If Not checkUserDirectory(nowID) Then Exit Sub
 For i = 1 To MaxSessionKey
   If sessionKeyList(i).id <> vbNullString Then
    temp = temp & sessionKeyList(i).id & vbNullChar & sessionKeyList(i).key & vbNullChar & Format(sessionKeyList(i).updateTime, TimeFormat) & vbNullChar
   End If
 Next i
 temp = temp & vbNullChar
 For i = 1 To MaxSessionKey
   If sessionKeyList2(i).id <> vbNullString Then
    temp = temp & sessionKeyList2(i).id & vbNullChar & sessionKeyList2(i).key & vbNullChar & Format(sessionKeyList2(i).updateTime, TimeFormat) & vbNullChar
   End If
 Next i
 stringToBytesW temp, buffer
 cEncrypt buffer, keySalt, buffer2
 setByteArraySafe buffer
  temp = recordPath & "\record\" & nowID & "\session2.dat"
 If fileExists(temp) Then Kill temp
 Open temp For Binary Access Write As #filen
  Put #filen, , buffer2
 Close #filen
End Sub
Sub saveAll()
 saveSettings
 saveUserSettings
 saveSessionFile
End Sub
Function getNewHandle() As Long
 nowHandle = nowHandle Mod 32767 + 1
 getNewHandle = nowHandle
End Function
Private Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'WndProc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam): Exit Function
 Select Case msg
  Case WM_MDIACTIVATE
   If lParam = WM_LBUTTONDBLCLK Then
    doShow
   ElseIf lParam = WM_RBUTTONUP Then
    FormPanel.PopupMenu FormPanel.menuTray
   End If
  Case WM_HOTKEY     '如果拦截到热键标志常数
      
  Case WM_TASKBARCREATED
   If FormPanel.visible = False Then hideToTray
  Case Else
   WndProc = CallWindowProc(preWinProc, hwnd, msg, wParam, lParam)
 End Select
End Function

Public Sub monitor(active As Boolean)
 If active Then
  Dim temp As Long
  temp = GetWindowLong(FormPanel.hwnd, GWL_WNDPROC)
  If preWinProc <> getFunctionAddr(AddressOf WndProc) Then preWinProc = temp '仅当监控函数地址改变时才更新preWinProc
  SetWindowLong FormPanel.hwnd, GWL_WNDPROC, AddressOf WndProc
  setHotKey True
 Else
  If preWinProc <> 0 Then SetWindowLong FormPanel.hwnd, GWL_WNDPROC, preWinProc
  setHotKey False
 End If
End Sub
Public Sub setHotKey(active As Boolean)

End Sub

Public Function showMsg(msg As String, Optional ByVal Title As String = "", Optional dialogType As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
 If Title = "" Then
  Title = "WiChat "
  If (dialogType And vbExclamation) > 0 Then
   Title = Title & "Warning"
  ElseIf (dialogType And vbCritical) > 0 Then
   Title = Title & "Error"
  Else
   Title = Title & "Notice"
   dialogType = dialogType Or vbInformation
  End If
 End If
 showMsg = MsgBox(msg, dialogType, Title)
End Function

Public Function verifyAccount(ByVal id As String, ByVal PW As String) As Integer
 verifyAccount = 0
 If Not (checkID(id) And checkPW(PW)) Then
  showMsg "ID/Password format error.", , vbInformation
  Exit Function
 End If
 Dim tempKey As String, tempString As String, tempID As String, i As Integer
 Dim bufferIn() As Byte, bufferOut() As Byte, p As Integer
 
 tempKey = LeftB(genKey(vbNullString, True), KeyLen)
 tempString = formatID(id)
 stringToBytesW tempString, bufferIn
 stringToBytesW PW, bufferOut
 PW = LeftB(getSHA1(bufferOut), KeyLen)
 encode bufferIn, tempKey, bufferOut, MaxIDLen - 1
 bytesToStringA bufferOut, tempID, 0, 6  'new ID
 tempID = tempID & ChrB(0)
 stringToBytesW tempString, bufferOut
 bytesToStringA bufferOut, tempString, 0, UBound(bufferOut)
 fuse tempString, tempKey, tempString
 
 ReDim bufferIn(1)
 bufferIn(0) = QueryDevice
 bufferIn(1) = 1
 stringToBytesA tempString, bufferOut
 mergeBytes bufferIn, bufferOut
 stringToBytesA tempKey, bufferOut
 mergeBytes bufferIn, bufferOut
 
 If Not sendRequest(1, "/Account/log/login.php", bufferIn, bufferOut) Then GoTo networkError
 If bufferOut(0) = RESPONSE_DEVICE_UNSUPPORTED Then GoTo versionError
 If Not bufferOut(0) = RESPONSE_SUCCESS Then GoTo unknownError
 
 
 stringToBytesA tempKey, bufferIn
 bytesToStringA bufferOut, tempString, 1, KeyLen
 encode bufferIn, tempString, bufferOut, KeyLen, True
 tempKey = bytesToStringW(bufferOut, 0, UBound(bufferOut))
 stringToBytesA PW, bufferOut
 Module3.clear
 Module3.loadKey tempKey
 mergeBytes Module3.source, bufferOut
 Module3.encrypt
 
 ReDim bufferIn(1)
 bufferIn(0) = QueryDevice
 bufferIn(1) = 2
 stringToBytesA tempID, bufferOut
 mergeBytes bufferIn, bufferOut
 mergeBytes bufferIn, Module3.target
 ReDim bufferOut(0)
 bufferOut(0) = State.onLine
 mergeBytes bufferIn, bufferOut
 
 If Not sendRequest(1, "/Account/log/login.php", bufferIn, bufferOut) Then GoTo networkError
 If Not bufferOut(0) = RESPONSE_SUCCESS Then GoTo verifyError
 bytesToStringA bufferOut, nowSession, 1, KeyLen 'Session key
 sessionValid = bytesToInt(bufferOut, KeyLen + 1) '2 bytes
 nowState = AscB(bytesToStringW(bufferOut, KeyLen + 3, KeyLen + 3))
 Module3.clear False
 removeBytes bufferOut, 0, KeyLen + 4
 mergeBytes Module3.source, bufferOut
 Module3.decrypt
 bytesToStringA Module3.target, nowOfflineMsg, 0, UBound(Module3.target)
 nowOfflineMsg = trimNull(nowOfflineMsg)
 
 stringToBytesW tempKey, bufferIn
 bytesToStringA bufferIn, tempKey, 0, UBound(bufferIn)
 stringToBytesA nowSession, bufferIn
 encode bufferIn, tempKey, bufferOut, KeyLen, True
 sessionKey = bytesToStringW(bufferOut, 0, UBound(bufferOut))
 nowID = id
 
 setStringSafe PW
 verifyAccount = 1
 Exit Function
 
networkError:
 setStringSafe PW
 verifyAccount = -1
 FormMain.showState "Network error.", vbRed
 Exit Function
 
verifyError:
 setStringSafe PW
 verifyAccount = 0
 FormMain.showState "Verification failed.", vbRed
 Exit Function
 
versionError:
 setStringSafe PW
 FormMain.showState "Please check your WiChat version.", vbRed
 verifyAccount = -1
 Exit Function
 
unknownError:
 setStringSafe PW
 FormMain.showState "Unknown error occurs. Cannot login.", vbRed
 verifyAccount = -1
End Function
Public Function changeState(newState As State) As Boolean
 changeState = False
 If nowState = State.none Then
  showMsg "Please log in first!", , vbCritical
  Exit Function
 End If
 Dim bufferIn() As Byte, bufferOut() As Byte
 ReDim bufferIn(1)
 bufferIn(0) = 8: bufferIn(1) = newState
 If Not exchangeData(bufferIn, bufferOut, Action_Acc_Change) Then
  showMsg "Network error. Cannot change state.", , vbCritical
  Exit Function
 End If
 If UBound(bufferOut) < 1 Or bufferOut(0) <> RESPONSE_SUCCESS Then
  showMsg "Server error. Cannot change state.", , vbCritical
 Else
  nowState = bufferOut(1)
 End If
 Select Case newState
  Case 1
   FormPanel.menuHomeStateOnline.Checked = True
   FormPanel.menuTrayStateOnline.Checked = True
  Case 2
   FormPanel.menuHomeStateOffline.Checked = True
   FormPanel.menuTrayStateOffline.Checked = True
  Case 4
   FormPanel.menuHomeStateBusy.Checked = True
   FormPanel.menuTrayStateBusy.Checked = True
  Case 5
   FormPanel.menuHomeStateHide.Checked = True
   FormPanel.menuTrayStateHide.Checked = True
 End Select
 FormPanel.addTask taskUpdateState
End Function
Public Function changePW(oldPW As String, newPW As String) As Boolean
 changePW = False
 If Not (checkPW(oldPW) And checkPW(newPW)) Then
  showMsg "Password format error.", , vbInformation
  Exit Function
 End If
 Dim bufferIn() As Byte, bufferOut() As Byte
 ReDim bufferIn(1)
 bufferIn(0) = 9: bufferIn(1) = Int(Rnd * 256)
 stringToBytesW oldPW, bufferOut
 oldPW = LeftB(getSHA1(bufferOut), KeyLen)
 stringToBytesW newPW, bufferOut
 newPW = LeftB(getSHA1(bufferOut), KeyLen)
 stringToBytesA oldPW, bufferOut
 mergeBytes bufferIn, bufferOut
 stringToBytesA newPW, bufferOut
 mergeBytes bufferIn, bufferOut
 If Not exchangeData(bufferIn, bufferOut, Action_Acc_Change) Then
  showMsg "Network error. Cannot change password.", , vbCritical
  GoTo changeFinally
 End If
 If bufferOut(0) <> RESPONSE_SUCCESS Then
  showMsg "Password error. Cannot change password.", , vbCritical
 Else
  Erase bufferIn
  stringToBytesA nowSession, bufferIn
  FormPanel.addTask taskType.taskChangeSession
  showMsg "Password successfully changed.", , vbInformation
 End If
changeFinally:
 setStringSafe oldPW
 setStringSafe newPW
End Function
Public Function changeSession() As Boolean
 changeSession = False
  Dim bufferIn() As Byte, bufferOut() As Byte
 ReDim bufferIn(1)
 bufferIn(0) = 5: bufferIn(1) = Int(Rnd * 256)
 If Not exchangeData(bufferIn, bufferOut, Action_Acc_Change) Then
  FormPanel.addTask taskLogOut
  Exit Function
 End If
 If bufferOut(0) <> RESPONSE_SUCCESS Then
  FormPanel.addTask taskLogOut
 Else
  bytesToStringA bufferOut, nowSession, 1, SessionLen
  stringToBytesA nowSession, bufferIn
  stringToBytesW sessionKey, bufferOut
  bytesToStringA bufferOut, sessionKey, 0, UBound(bufferOut)
  encode bufferIn, sessionKey, bufferOut, KeyLen, True
  sessionKey = bytesToStringW(bufferOut, 0, UBound(bufferOut))
  changeSession = True
 End If
End Function
Public Function changeMsg(msg As String) As Boolean
 changeMsg = False
 Dim temp As String
 temp = LeftB(msg, MaxOfflineMsg - 2) & vbNullChar
 Dim bufferIn() As Byte, bufferOut() As Byte
 ReDim bufferIn(1)
 bufferIn(0) = 7: bufferIn(1) = Int(Rnd * 256)
 stringToBytesW "<MSG>", bufferOut
 mergeBytes bufferIn, bufferOut
 stringToBytesA temp, bufferOut
 mergeBytes bufferIn, bufferOut
 stringToBytesW "</MSG>", bufferOut
 mergeBytes bufferIn, bufferOut
 If Not exchangeData(bufferIn, bufferOut, Action_Acc_Change) Then
  showMsg "Network error. Cannot change offline message.", , vbCritical
  Exit Function
 End If
 If bufferOut(0) <> RESPONSE_SUCCESS Then
  showMsg "Server error. Cannot change offline message.", , vbCritical
 Else
  nowOfflineMsg = LeftB(msg, MaxOfflineMsg - 2)
  FormPanel.addTask taskUpdateState
  changeMsg = True
 End If
End Function
Public Function changeNote(id As String, notation As String) As Boolean
 changeNote = False
 Dim temp As String
 temp = LeftB(notation, MaxNote - 2) & vbNullChar
 Dim bufferIn() As Byte, bufferOut() As Byte
 ReDim bufferIn(1)
 bufferIn(0) = 12: bufferIn(1) = Int(Rnd * 256)
 stringToBytesW formatID(id), bufferOut
 mergeBytes bufferIn, bufferOut
 intToBytes bufferOut, LenB(temp)
 mergeBytes bufferIn, bufferOut
 stringToBytesA temp, bufferOut
 mergeBytes bufferIn, bufferOut
 If Not exchangeData(bufferIn, bufferOut, Action_Acc_Change) Then
  showMsg "Network error. Cannot change note.", , vbCritical
  Exit Function
 End If
 If bufferOut(0) <> RESPONSE_SUCCESS Then
  showMsg "Server error. Cannot change note.", , vbCritical
 Else
  changeNote = True
  FormPanel.addTask taskUpdateFriList
 End If
End Function
Public Sub updateFriendList(ByRef list As ListView)
 list.ListItems.clear
 
 Dim bufferIn() As Byte, bufferOut() As Byte
 ReDim bufferIn(1)
 bufferIn(0) = 1: bufferIn(1) = 2
 If Not exchangeData(bufferIn, bufferOut, Action_Fri_Change) Then Exit Sub
 If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Sub
 
 Dim temp As String, p1 As Long, p2 As Long, pE As Long
 temp = bytesToStringW(bufferOut, 2, UBound(bufferOut))
 p1 = InStr(temp, "<IDList t=c>"): pE = InStr(p1 + 1, temp, "</IDList>")
 Dim tempID As String, tempNote As Notification, Count As Long
 If p1 > 0 And pE > 13 Then
  Count = 0
  Do
   p1 = InStr(p1 + 1, temp, "<ID s="): p2 = InStr(p1 + 1, temp, "</ID>")
   If p1 < 1 Or p2 < 1 Or p1 > pE Then Exit Do
   tempID = trimNull(Mid(temp, p1 + 8, p2 - p1))
   If tempID <> vbNullString Then list.ListItems.Add , , vbNullString, , getStateImageIndex(Val(Mid(temp, p1 + 6, 1)))
   If Mid(temp, p1 + 6, 1) = "1" Then Count = Count + 1
   With list.ListItems(list.ListItems.Count).ListSubItems
    .Add , , tempID
    .Add , , tempID
   End With
  Loop
  FormPanel.labelFriendList.Caption = "Friend List (" & Count & "/" & list.ListItems.Count & ")"
 End If
 
 If Not list.SelectedItem Is Nothing Then list.SelectedItem.Selected = False
 Note_clear FriendAdd
 Note_clear FriendDelete
 p1 = InStr(temp, "<IDList t=b>"): pE = InStr(p1 + 1, temp, "</IDList>")
 If p1 > 0 And pE > 13 Then
  Do
   p1 = InStr(p1 + 1, temp, "<ID>"): p2 = InStr(p1 + 1, temp, "</ID>")
   If p1 < 1 Or p2 < 1 Or p1 > pE Then Exit Do
   tempID = trimNull(Mid(temp, p1 + 4, p2 - p1 - 4))
   If tempID <> vbNullString Then
    With tempNote
     .type = FriendDelete
     .source = tempID
     .destination = nowID
     .time = Now
     .handle = getNewHandle
     Note_push tempNote
    End With
   End If
  Loop
 End If
 p1 = InStr(temp, "<IDList t=w>"): pE = InStr(p1 + 1, temp, "</IDList>")
 If p1 > 0 And pE > 13 Then
  Do
   p1 = InStr(p1 + 1, temp, "<ID>"): p2 = InStr(p1 + 1, temp, "</ID>")
   If p1 < 1 Or p2 < 1 Or p1 > pE Then Exit Do
   tempID = trimNull(Mid(temp, p1 + 4, p2 - p1 - 4))
   If tempID <> vbNullString Then
    With tempNote
     .type = FriendAdd
     .source = tempID
     .destination = nowID
     .time = Now
     .handle = getNewHandle
     Note_push tempNote
    End With
   End If
  Loop
 End If
  
 ReDim bufferIn(1)
 bufferIn(0) = 11: bufferIn(1) = Rnd * 255
 temp = "<IDList>"
 For p1 = 1 To list.ListItems.Count
  temp = temp & "<ID>" & formatID(list.ListItems(p1).SubItems(1)) & "</ID>"
 Next p1
 temp = temp & "</IDList>"
 stringToBytesW temp, bufferOut
 mergeBytes bufferIn, bufferOut
 If Not exchangeData(bufferIn, bufferOut, Action_Fri_Change) Then Exit Sub
 If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Sub
 
 Dim p3 As Long, p4 As Long
 temp = bytesToStringW(bufferOut, 2, UBound(bufferOut))
 p1 = InStr(temp, "<MList>"): pE = InStr(p1 + 1, temp, "</MList>")
 If p1 > 0 And pE > 13 Then
  Count = 0
  Do
   p1 = InStr(p1 + 1, temp, "<ID>"): p2 = InStr(p1 + 1, temp, "</ID>")
   If p1 < 1 Or p2 < 1 Or p1 > pE Then Exit Do
   tempID = trimNull(Mid(temp, p1 + 4, p2 - p1 - 4))
   If tempID <> vbNullString Then
    Count = Count + 1
    p3 = InStr(p1 + 1, temp, "<NOTE>"): p4 = InStr(p3 + 1, temp, "</NOTE>")
    If p3 < 1 Or p4 < 1 Or p4 > pE Then Exit Do
    bytesToStringA bufferOut, tempID, p3 + 7, p4 - 1
    tempID = trimNull(fixStringTail(tempID))
    With list.ListItems(Count)
     If Len(tempID) > 0 Then
      .ListSubItems.Item(2) = trimNull(tempID) & "(" & .ListSubItems.Item(1) & ")"
     End If
     .ListSubItems.Item(2).ToolTipText = .ListSubItems.Item(2) & "[" & stateToString(getStateForImageIndex(.SmallIcon)) & "]"
    End With
   End If
  Loop
 End If
End Sub
Public Function addFriend(friendID As String) As Boolean
 addFriend = False
 Dim bufferIn() As Byte, bufferOut() As Byte, temp As String
 ReDim bufferIn(1)
 bufferIn(0) = 2: bufferIn(1) = Int(Rnd * 256)
 stringToBytesW "<IDList><ID>" & formatID(friendID) & "</ID></IDList>", bufferOut
 mergeBytes bufferIn, bufferOut
 If Not exchangeData(bufferIn, bufferOut, Action_Fri_Change) Then Exit Function
 If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Function
 temp = bytesToStringW(bufferOut, 2, UBound(bufferOut))
 Dim p1 As Long, p2 As Long, pE As Long
 p1 = InStrB(temp, "<IDList t=f>"): pE = InStrB(temp, "</IDList>")
 If p1 > 0 And pE > 13 Then
  Do
   If p1 < 1 Or p2 < 1 Or p1 > pE Then Exit Do
   p1 = InStr(p1 + 1, temp, "<ID>"): p2 = InStr(p1 + 1, temp, "</ID>")
   If trimNull(Mid(temp, p1 + 4, p2 - p1 - 4)) = friendID Then Exit Function
  Loop
 End If
 addFriend = True
End Function
Public Function delFriend(friendID As String) As Boolean
 delFriend = False
 Dim bufferIn() As Byte, bufferOut() As Byte, temp As String
 ReDim bufferIn(1)
 bufferIn(0) = 3: bufferIn(1) = Int(Rnd * 256)
 stringToBytesW "<IDList><ID>" & formatID(friendID) & "</ID></IDList>", bufferOut
 mergeBytes bufferIn, bufferOut
 If Not exchangeData(bufferIn, bufferOut, Action_Fri_Change) Then Exit Function
 If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Function
 temp = bytesToStringW(bufferOut, 2, UBound(bufferOut))
 Dim p1 As Long, p2 As Long, pE As Long
 p1 = InStr(temp, "<IDList t=f>"): pE = InStr(temp, "</IDList>")
 If p1 > 0 And pE > 13 Then
  Do
   p1 = InStr(p1 + 1, temp, "<ID>"): p2 = InStr(p1 + 1, temp, "</ID>")
   If p1 < 1 Or p2 < 1 Or p1 > pE Then Exit Do
   If Mid(temp, p1 + 4, p2 - p1) = friendID Then Exit Function
  Loop
 End If
 delFriend = True
End Function
Public Function getFriendInfo(friendID As String) As IDInfo
 getFriendInfo.id = vbNullString
 getFriendInfo.offlineMsg = vbNullString
 Dim bufferIn() As Byte, bufferOut() As Byte, temp As String
 ReDim bufferIn(1)
 bufferIn(0) = 10: bufferIn(1) = Int(Rnd * 256)
 stringToBytesW "<IDList><ID>" & formatID(friendID) & "</ID></IDList>", bufferOut
 mergeBytes bufferIn, bufferOut
 If Not exchangeData(bufferIn, bufferOut, Action_Fri_Change) Then Exit Function
 If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Function
 getFriendInfo.id = friendID
 temp = bytesToStringW(bufferOut, 2, UBound(bufferOut))
 Dim p1 As Long, p2 As Long, pE As Long
 p1 = InStr(temp, "<MList>"): pE = InStr(p1 + 1, temp, "</MList>")
 If p1 > 0 And pE > 15 Then
  Do
   p1 = InStr(p1 + 1, temp, "<ID>"): p2 = InStr(p1 + 1, temp, "</ID>")
   If p1 < 1 Or p2 < 1 Or p1 > pE Then Exit Do
   If trimNull(Mid(temp, p1 + 4, p2 - p1)) = friendID Then
    p1 = InStr(p1 + 1, temp, "<MSG>"): p2 = InStr(p1 + 1, temp, "</MSG>")
    getFriendInfo.offlineMsg = trimNull(stringWToA(Mid(temp, p1 + 5, p2 - p1)))
   End If
   Exit Do
  Loop
 End If
End Function
Public Function createSession(id As String) As Boolean
 createSession = False
 If Not checkID(id) Then Exit Function
 Dim i As Integer
 If sessionCount >= MaxSession Then Exit Function
 For i = 1 To MaxSession
  If sessionList(i).id = id Then Exit For
 Next i
 If i > MaxSession Then
  For i = 1 To MaxSession
   If sessionList(i).id = vbNullString Then Exit For
  Next i
  If i > MaxSession Then Exit Function 'Should never be here
 Else
  createSession = True
  Exit Function
 End If
 sessionCount = sessionCount + 1
 With sessionList(i)
  .id = id
  .cache = vbNullString
  .input = vbNullString
 End With

 createSession = True
End Function
Public Function loadSessionContent(id As String, textLog As WebBrowser, textInput As WebBrowser) As Boolean
 Dim i As Integer, j As Integer, received As Boolean
 For i = 1 To MaxSession
  If sessionList(i).id = id Then Exit For
 Next i
 If i > MaxSession Then Exit Function
 
 j = getActiveSession
 If FormPanel.bindTextBox Then sessionList(j).input = FormPanel.textArea.innerText
 sessionList(j).active = False
 
 received = False
 For j = 1 To pNoteStack
  If noteStack(j).source = id And noteStack(j).type = GotMsg Then
   Note_Remove j
   received = True
  End If
 Next j
 If received And id <> nowID Then receiveMessage id
 
 Dim filen As Integer, buffer() As Byte
 filen = FreeFile
 resetTempWebPage
 Open tempWebPage For Binary Access Write As #filen
  getHTMLHeaderA buffer, 0
  Put #filen, 1, buffer
  UnicodeToUtf8 render(sessionList(i).cache), buffer
  Put #filen, , buffer
  getHTMLFooterA buffer, 0
  Put #filen, , buffer
 Close #filen
 textLog.Navigate "file:///" & tempWebPage & "?r=" & Rnd
 DoEvents

 resetTempWebPage
 Open tempWebPage For Binary Access Write As #filen
  getHTMLHeaderA buffer, 1
  Put #filen, 1, buffer
  UnicodeToUtf8 sessionList(i).input, buffer
  Put #filen, , buffer
  getHTMLFooterA buffer, 1
  Put #filen, , buffer
 Close #filen
 textInput.Navigate "file:///" & tempWebPage & "?r=" & Rnd
 DoEvents
 FormPanel.changeFont
 Dim doc As IHTMLDocument2, body As IHTMLElement2
 Set doc = textLog.Document
 If Not doc Is Nothing Then
  Set body = doc.body
  If Not body Is Nothing Then doc.parentWindow.scrollTo 0, doc.body.scrollHeight
 End If
 sessionList(i).active = True
 loadSessionContent = True
End Function
Public Sub closeSession(textInput As WebBrowser)
 Dim i As Integer
 i = getActiveSession
 If i < 1 Then Exit Sub
 sessionList(i).active = False
 sessionList(i).id = vbNullString
 sessionCount = sessionCount - 1
End Sub
Public Sub clearSession(textLog As WebBrowser)
 Dim i As Integer
 i = getActiveSession
 If i < 1 Then Exit Sub
 sessionList(i).cache = vbNullString
 Dim buffer() As Byte, filen As Integer
 filen = FreeFile
 resetTempWebPage
 Open tempWebPage For Binary Access Write As #filen
  getHTMLHeaderA buffer, 0
  Put #filen, 1, buffer
  UnicodeToUtf8 render(sessionList(i).cache), buffer
  Put #filen, , buffer
  getHTMLFooterA buffer, 0
  Put #filen, , buffer
 Close #filen
 textLog.Navigate "file:///" & tempWebPage
End Sub
Public Function isNowSessionEmpty() As Integer '0=All empty; 1=Content not empty; 2=Input not empty
 isNowSessionEmpty = 0
 Dim i As Integer
 i = getActiveSession
 If i < 1 Then Exit Function
 If sessionList(i).cache <> vbNullString Then isNowSessionEmpty = isNowSessionEmpty + 1
 If sessionList(i).input <> vbNullString Then isNowSessionEmpty = isNowSessionEmpty + 2
End Function
Public Sub clearInput(textInput As WebBrowser)
 Dim filen As Integer, buffer() As Byte
 filen = FreeFile
 resetTempWebPage
 Open tempWebPage For Binary Access Write As #filen
  getHTMLHeaderA buffer, 1
  Put #filen, 1, buffer
  getHTMLFooterA buffer, 1
  Put #filen, , buffer
 Close #filen
 textInput.Navigate "file:///" & tempWebPage
 DoEvents
 FormPanel.changeFont
 filen = getActiveSession
 If filen < 1 Then Exit Sub
 sessionList(filen).input = vbNullString
End Sub
Public Function sendMessage(textLog As WebBrowser, textInput As WebBrowser, decoration As fontStyle) As Boolean
 sendMessage = False
 Dim i As Long, j As Long, id As String, key As String
 Dim temp As String, buffer() As Byte
 If FormPanel.bindTextBox Then temp = translate(FormPanel.textArea.innerText)
 If Trim(temp) = vbNullString Then Exit Function
 
 i = getActiveSession
 If i < 1 Then Exit Function
 id = sessionList(i).id
 temp = addSenderInfo(temp, decoration)


 If Not msgLogin Then Exit Function
 loadSessionKey
 If sessionList(i).id <> nowID Then
  Dim bufferIn() As Byte, bufferOut() As Byte
  For j = 1 To MaxSessionKey
   If sessionKeyList(j).id = id Then Exit For
  Next j
  Randomize Timer
  If j > MaxSessionKey Then 'Create sending key positively
   key = Left(genKey, KeyLen)
   
   ReDim bufferIn(1)
   bufferIn(0) = 2: bufferIn(1) = 2
   stringToBytesW formatID(id), bufferOut
   mergeBytes bufferIn, bufferOut
   longToBytes bufferOut, 32
   mergeBytes bufferIn, bufferOut
   ReDim bufferOut(1)
   bufferOut(0) = 1: bufferOut(1) = 0
   mergeBytes bufferIn, bufferOut
   
   ReDim bufferOut(8)
   bufferOut(0) = 127: bufferOut(1) = 255:   bufferOut(2) = 127: bufferOut(3) = 255
   bufferOut(4) = QueryDevice: bufferOut(5) = 3
   
   bufferOut(6) = Asc(Rnd * 255): bufferOut(7) = Asc(Rnd * 255): bufferOut(8) = Asc(Rnd * 255)
   mergeBytes bufferIn, bufferOut
   stringToBytesA DecToHex(key), bufferOut
'   bufferOut(2 + (bufferOut(0) + bufferOut(1)) Mod 14) = Asc(Int(Rnd * 10))
   mergeBytes bufferIn, bufferOut
   ReDim bufferOut(6)
   For j = 0 To 6: bufferOut(j) = Asc(Rnd * 255): Next j
   mergeBytes bufferIn, bufferOut
   
   If Not exchangeData(bufferIn, bufferOut, Action_Msg_Exchange) Then Exit Function
   If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Function
   pSessionKeyList = pSessionKeyList Mod MaxSessionKey + 1
   With sessionKeyList(pSessionKeyList)
    .id = id
    .key = key
    .updateTime = Now
   End With
  Else
   key = sessionKeyList(j).key
  End If
  
  Erase bufferOut
  If Not dataXMLize(temp, buffer) Then Exit Function
  
  If Int(Rnd * 22) = 0 And UBound(buffer) < 512 Then 'Reset sending key positively & randomly
   For j = 1 To MaxSessionKey
    If sessionKeyList(j).id = id Then Exit For
   Next j
   If j <= MaxSessionKey Then
    sessionKeyList(j).key = Left(genKey, KeyLen)
    sessionKeyList(j).updateTime = Now
   Else
    Exit Function 'Should never reach here
   End If
   dataXMLize "<O><I>CK</I><V>" & sessionKeyList(j).key & "</V></O>", bufferOut
  End If
  If byteArrayIsDimed(bufferOut) Then j = UBound(bufferOut) + 1 Else j = 0
  j = j + UBound(buffer) + 1
  stringToBytesW "<L>" & j & "</L>", bufferIn
  mergeBytes bufferOut, bufferIn, False
  mergeBytes bufferOut, buffer
  If (UBound(bufferOut) + 1) Mod 8 > 0 Then ReDim Preserve bufferOut(UBound(bufferOut) + 8 - (UBound(bufferOut) + 1) Mod 8)
  cEncrypt bufferOut, stringWToA(key), buffer
  
  
  If UBound(buffer) < MaxMsgBlock Then
   ReDim bufferIn(1)
   bufferIn(0) = 2: bufferIn(1) = 2
   stringToBytesW formatID(id), bufferOut
   mergeBytes bufferIn, bufferOut
   longToBytes bufferOut, UBound(buffer) + 1
   mergeBytes bufferIn, bufferOut
   ReDim bufferOut(1)
   bufferOut(0) = 1: bufferOut(1) = 0
   mergeBytes bufferIn, bufferOut
   mergeBytes bufferIn, buffer
   FormPanel.showProcess True, "Sending Message..."
   If Not exchangeData(bufferIn, bufferOut, Action_Msg_Exchange) Then FormPanel.showProcess False: Exit Function
   FormPanel.showProcess False
  Else
   j = 0
   Do
    FormPanel.showProcess True, "Sending Message...", j / UBound(buffer) * 100
    If j + MaxMsgBlock >= UBound(buffer) Then
     ReDim Module3.source(UBound(buffer) - j)
     For i = 0 To UBound(buffer) - j
      Module3.source(i) = buffer(i + j)
     Next i
     i = 1
    Else
     ReDim Module3.source(MaxMsgBlock - 1)
     For i = 0 To MaxMsgBlock - 1
      Module3.source(i) = buffer(i + j)
     Next i
     i = 0
    End If
    ReDim bufferIn(1)
    bufferIn(0) = 2: bufferIn(1) = 2
    stringToBytesW formatID(id), bufferOut
    mergeBytes bufferIn, bufferOut
    longToBytes bufferOut, UBound(Module3.source) + 1
    mergeBytes bufferIn, bufferOut
    ReDim bufferOut(1)
    bufferOut(0) = i: bufferOut(1) = 1
    mergeBytes bufferIn, bufferOut
    mergeBytes bufferIn, Module3.source
    If Not exchangeData(bufferIn, bufferOut, Action_Msg_Exchange) Then FormPanel.showProcess False: Exit Function
    j = j + MaxMsgBlock
   Loop Until i = 1
   FormPanel.showProcess False
  End If
  If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Function
 End If
 setStringSafe key
 
 i = getActiveSession
 sessionList(i).cache = sessionList(i).cache & temp
 Dim doc As IHTMLDocument2, tempDIV As HTMLDivElement
 Set doc = textLog.Document
 Set tempDIV = doc.createElement("div")
 tempDIV.innerHTML = render(temp)
 tempDIV.className = "m"
 doc.body.appendChild tempDIV
 doc.parentWindow.scrollTo 0, doc.body.scrollHeight
 sendMessage = True
End Function
Public Function getMessageList() As Boolean
 getMessageList = False
 If Not msgLogin Then Exit Function
 Dim bufferIn() As Byte, bufferOut() As Byte
 ReDim bufferIn(1)
 bufferIn(0) = 1: bufferIn(1) = Int(Rnd * 256)
 If Not exchangeData(bufferIn, bufferOut, Action_Msg_Get_List) Then Exit Function
 If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Function
 Dim p1 As Long, p2 As Long, pE As Long, temp As String, tempNote As Notification
 temp = bytesToStringW(bufferOut, 2, UBound(bufferOut))
 p1 = InStr(temp, "<IDList t=v>"): pE = InStr(temp, "</IDList>")
 If p1 > 0 And pE > 11 Then
  Do
   p1 = InStr(p1 + 1, temp, "<ID>"): p2 = InStr(p1 + 1, temp, "</ID>")
   If p1 < 1 Or p2 < 1 Or p1 > pE Then Exit Do
   tempNote.source = trimNull(Mid(temp, p1 + 4, p2 - p1 - 4))
   tempNote.destination = nowID
   tempNote.type = NoteEvent.GotMsg
   tempNote.handle = getNewHandle
   tempNote.time = Now
   Note_push tempNote
  Loop
 End If
 getMessageList = True
End Function
Public Function receiveMessage(id As String) As Boolean
 receiveMessage = False
 
 If Not Session_Exist(id) Then
  If Not createSession(id) Then Exit Function
 End If
 
 If nowState = State.none Or nowState = State.Offline Then Exit Function

 If Not msgLogin Then Exit Function
 loadSessionKey
 
 Dim bufferIn() As Byte, bufferOut() As Byte, buffer() As Byte, i As Integer, j As Integer
 ReDim bufferIn(1)
 bufferIn(0) = 1
 If id = "10000" Then bufferIn(1) = 1 Else bufferIn(1) = 2
 stringToBytesW formatID(id), bufferOut
 mergeBytes bufferIn, bufferOut
 longToBytes bufferOut, MaxMsgBlock
 mergeBytes bufferIn, bufferOut
 ReDim bufferOut(1)
 bufferOut(0) = 1: bufferOut(1) = 0
 mergeBytes bufferIn, bufferOut
 FormPanel.showProcess True, "Receiving Message..."
 If Not exchangeData(bufferIn, bufferOut, Action_Msg_Exchange) Then FormPanel.showProcess False: Exit Function
 FormPanel.showProcess False
 If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Function
 If bufferOut(1) = RESPONSE_RES_NOT_EXIST Then Exit Function
 If bufferOut(1) = RESPONSE_RES_SIZE_TOO_LARGE Then
  Dim l As Long
  l = bytesToLong(bufferOut, 2) \ MaxMsgBlock
  i = 0
  Do
   FormPanel.showProcess True, "Receiving Message...", i / l * 100
   bufferIn(UBound(bufferIn) - 1) = 0: bufferIn(UBound(bufferIn)) = i
   If Not exchangeData(bufferIn, bufferOut, Action_Msg_Exchange) Then FormPanel.showProcess False: Exit Function
   If bufferOut(0) = RESPONSE_SUCCESS Then
    If bufferOut(1) = RESPONSE_RES_EOF Then
     i = -1
    Else
     If bufferOut(1) <> RESPONSE_RES_OK Then FormPanel.showProcess False: Exit Function
    End If
    removeBytes bufferOut, 0, 2
    mergeBytes buffer, bufferOut
   Else
    FormPanel.showProcess False
    Exit Function
   End If
   If i = -1 Then Exit Do
   i = i + 1
  Loop
  FormPanel.showProcess False
 Else
  copyBytes buffer, bufferOut, 2, UBound(bufferOut)
 End If
 Erase bufferOut
 
 'Allocate index of ID required.
 For i = 1 To MaxSessionKey
   If sessionKeyList2(i).id = id Then Exit For
  Next i
 If i > MaxSessionKey Then
   pSessionKeyList2 = pSessionKeyList2 Mod MaxSessionKey + 1
   sessionKeyList2(pSessionKeyList2).id = id
   i = pSessionKeyList2
 End If
  
 Dim p As Long, p2 As Long, temp As String
 Erase bufferOut
 p = 0
 Do
  temp = sessionKeyList2(i).key
  ReDim bufferIn(3)
  bufferIn(0) = 127: bufferIn(1) = 255: bufferIn(2) = 127: bufferIn(3) = 255
  p2 = inBytes(buffer, bufferIn, p)
  Erase bufferIn
  If p2 >= 0 Then 'Got Operation Signal
   If p2 > 0 Then
      copyBytes bufferIn, buffer, p, p2 - 1
   End If
   If buffer(p2 + 5) = 3 Then 'Is notification of reseting receiving key
     bytesToStringA buffer, sessionKeyList2(i).key, p2 + 9, p2 + 9 + KeyLen - 1
     sessionKeyList2(i).key = HexToDec(sessionKeyList2(i).key)
     sessionKeyList2(i).updateTime = Now
   ElseIf buffer(p2 + 5) = 4 Then
     For j = 1 To MaxSessionKey
      If sessionKeyList(j).id = id Then Exit For
     Next j
     If j <= MaxSessionKey Then 'Is request of reseting sending key
      sessionKeyList(j).id = vbNullString
      setStringSafe (sessionKeyList(j).key)
     End If
   End If
  Else
   Exit Do
  End If
  If peelAdditionalInfo(bufferIn, temp, bufferOut) Then sessionKeyList2(i).updateTime = Now
  p = p2 + 32
 Loop
 
 Module3.clear
 If p < UBound(buffer) Then
  copyBytes bufferIn, buffer, p, UBound(buffer)
  If peelAdditionalInfo(bufferIn, sessionKeyList2(i).key, Module3.source) Then sessionKeyList2(i).updateTime = Now
 End If
 If byteArrayIsDimed(Module3.source) Then mergeBytes bufferOut, Module3.source
 
 If sessionKeyList2(i).key = nullKey Then FormPanel.addTask taskRebuildConnection
 
  
'   If p1 > 0 Then
'    Module3.clear
'    copyBytes Module3.source, buffer, 0, p1 - 1
'    cDecrypt Module3.source, sessionKeyList2(i).key, Module3.target
'    If isReasonable(Module3.target) Then mergeBytes bufferOut, Module3.target Else sessionKeyList2(i).key = nullKey
'   End If
'  End If
'  Do While p1 >= 0
'   sessionKeyList2(i).key = bytesToStringW(buffer, p1 + 6, p1 + 21)
'   sessionKeyList2(i).updateTime = Now
'   p2 = inBytes(buffer, bufferIn, p1 + 1)
'   If p2 < 0 Then p2 = UBound(buffer)
'   Module3.clear
'   copyBytes Module3.source, buffer, p1 + 24, p2
''   Dim j As Integer
''   For j = 0 To 9
''    Mid(sessionKeyList2(i).key, 3 + (Asc(Left(sessionKeyList2(i).key, 1)) + Asc(Mid(sessionKeyList2(i).key, 2, 1))) Mod 14, 1) = j
'    cDecrypt Module3.source, sessionKeyList2(i).key, Module3.target
''    If isReasonable(Module3.target) Then Exit For
''   Next j
'    If isReasonable(Module3.target) Then mergeBytes bufferOut, Module3.target Else sessionKeyList2(i).key = nullKey
'   If p2 < UBound(buffer) Then p1 = p2 Else Exit Do 'Same as p1=-1
'  Loop
' Else
'  If i > MaxSessionKey Then Exit Function
'  Module3.clear
'  mergeBytes Module3.source, buffer
'  cDecrypt Module3.source, sessionKeyList2(i).key, Module3.target
'  If isReasonable(Module3.target) Then mergeBytes bufferOut, Module3.target Else sessionKeyList2(i).key = nullKey
' End If
 
 temp = dataUnxmlize(bufferOut, nowID)
 fixSenderInfo temp
 For i = 1 To MaxSession
  If sessionList(i).id = id Then Exit For
 Next i
 If i <= MaxSession Then sessionList(i).cache = sessionList(i).cache & temp
 receiveMessage = True
End Function
Private Function peelAdditionalInfo(src() As Byte, ByRef key As String, dest() As Byte) As Boolean 'Indicate wheather key has been updated
 peelAdditionalInfo = False
 Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long, p6 As Long, p7 As Long
 Dim buffer() As Byte, buffer2() As Byte, temp As String, l As Long
 Do While byteArrayIsDimed(src)
   If Not cDecrypt(src, stringWToA(key), buffer) Then key = nullKey: Exit Do
  
   If Not isReasonable(buffer) Then key = nullKey: Exit Do
  
   stringToBytesW "<L>", buffer2:  p1 = inBytes(buffer, buffer2)
   If p1 < 0 Then Exit Do
   stringToBytesW "</L>", buffer2:  p2 = inBytes(buffer, buffer2, p1 + 1)
   If p2 < 0 Then Exit Do
   l = CLng(bytesToStringW(buffer, p1 + 3, p2 - 1))
   If l <= 0 Then Exit Do
   stringToBytesW "<O>", buffer2:  p3 = inBytes(buffer, buffer2, p2 + 1)
   If p3 - p2 > 4 Then p3 = -1: GoTo con2
   stringToBytesW "</O>", buffer2:  p4 = inBytes(buffer, buffer2, p3 + 1)
   If p4 < 0 Then p4 = -1: GoTo con2
   temp = bytesToStringW(buffer, p3 + 3, p4 - 1)
   p7 = 0
   Do
    p5 = InStr(p7 + 1, temp, "<I>"): If p5 < 1 Then Exit Do
    p6 = InStr(p5 + 1, temp, "</I>"): If p6 < 1 Then Exit Do
    Select Case Mid(temp, p5 + 3, p6 - p5 - 3)
     Case "CK"
      p5 = InStr(p5 + 1, temp, "<V>"): If p5 < 1 Then GoTo con3
      p7 = InStr(p5 + 1, temp, "</V>"): If p7 < 1 Then Exit Do
      If p7 - p5 - 3 > 0 Then
       key = Mid(temp, p5 + 3, p7 - p5 - 3)
       peelAdditionalInfo = True
      End If
     Case Else
    End Select
con3:
   Loop
con2:
   If p3 < 0 Then
    copyBytes buffer2, buffer, p2 + 4, p2 + 3 + l
   Else
    copyBytes buffer2, buffer, p4 + 4, p2 + 3 + l
   End If
   removeBytes src, 0, p2 + 4 + l + (8 - (p2 + 4 + l) Mod 8) Mod 8
   mergeBytes dest, buffer2
  Loop
End Function
Public Function fixBrokenConnection()
 Dim buffer() As Byte, bufferIn() As Byte, bufferOut() As Byte
 Dim i As Integer, j As Integer
 For i = 1 To MaxSessionKey
  If sessionKeyList2(i).id <> vbNullString And sessionKeyList2(i).key = nullKey Then
    'Notify changing sending key passively
    
    ReDim bufferIn(1)
    bufferIn(0) = 2: bufferIn(1) = 2
    stringToBytesW formatID(sessionKeyList2(i).id), bufferOut
    mergeBytes bufferIn, bufferOut
    longToBytes bufferOut, 32
    mergeBytes bufferIn, bufferOut
    ReDim bufferOut(1)
    bufferOut(0) = 1: bufferOut(1) = 0
    mergeBytes bufferIn, bufferOut
    
    ReDim bufferOut(31)
    bufferOut(0) = 127: bufferOut(1) = 255:   bufferOut(2) = 127: bufferOut(3) = 255
    bufferOut(4) = QueryDevice: bufferOut(5) = 4
    For j = 6 To 31: bufferOut(j) = Asc(Rnd * 255): Next j
    mergeBytes bufferIn, bufferOut
    
    exchangeData bufferIn, bufferOut, Action_Msg_Exchange
    sessionKeyList2(i).id = vbNullString
  End If
 Next i
End Function
Private Function msgLogin() As Boolean
 msgLogin = False
 If nowState = State.none Or nowState = State.Offline Then Exit Function
 If sessionKeyList(0).key <> vbNullString Then msgLogin = True: Exit Function
 
 Dim bufferIn() As Byte, bufferOut() As Byte, tempString As String
 ReDim bufferIn(1)
 bufferIn(0) = QueryDevice: bufferIn(1) = 3
 stringToBytesA nowSession, bufferOut
 mergeBytes bufferIn, bufferOut
 tempString = Left(genKey, KeyLen)
 stringToBytesW tempString, bufferOut
 mergeBytes bufferIn, bufferOut
 If Not sendRequest(2, "/Record/log/login.php", bufferIn, bufferOut) Then Exit Function
 If bufferOut(0) <> RESPONSE_SUCCESS Then Exit Function
 bytesToStringA bufferOut, keySalt, 2, 17
 stringToBytesW tempString, bufferOut
 bytesToStringA bufferOut, tempString, 0, UBound(bufferOut)
 stringToBytesW sessionKey, bufferOut
 encode bufferOut, tempString, bufferIn, KeyLen, True
 sessionKeyList(0).key = bytesToStringW(bufferIn, 0, UBound(bufferIn))
 msgLogin = True
End Function
Private Function exchangeData(Data() As Byte, buffer() As Byte, object As Integer) As Boolean
 exchangeData = False
 Dim temp() As Byte, bufferIn() As Byte, tempString As String
 
 Erase buffer
 Select Case object
  Case Action_Acc_Change, Action_Fri_Change
   stringToBytesA nowSession, buffer
   getCRC32 Data, temp
   mergeBytes buffer, temp
   Module3.clear
   mergeBytes Module3.source, Data
   Module3.loadKey sessionKey
   Module3.encrypt
   mergeBytes bufferIn, buffer
   mergeBytes bufferIn, Module3.target
   If Not sendRequest(1, "/Account/acc/action.php", bufferIn, temp) Then Exit Function
  Case Action_Msg_Exchange
   stringToBytesA nowSession, buffer
   getCRC32 Data, temp
   mergeBytes buffer, temp
   Module3.clear
   mergeBytes Module3.source, Data
   Module3.loadKey sessionKeyList(0).key
   Module3.encrypt
   mergeBytes bufferIn, buffer
   mergeBytes bufferIn, Module3.target
   If Not sendRequest(2, "/Record/query/action.php", bufferIn, temp) Then Exit Function
  Case Action_Msg_Get_List
   stringToBytesA nowSession, buffer
   getCRC32 Data, temp
   mergeBytes buffer, temp
   Module3.clear
   mergeBytes Module3.source, Data
   Module3.loadKey sessionKeyList(0).key
   Module3.encrypt
   mergeBytes bufferIn, buffer
   mergeBytes bufferIn, Module3.target
   If Not sendRequest(2, "/Record/query/get.php", bufferIn, temp) Then Exit Function
  Case Else
   Exit Function
 End Select
 
 If UBound(temp) < 4 Then Exit Function
 bytesToStringA temp, tempString, 0, 3
 removeBytes temp, 0, 4
 Module3.clear False
 mergeBytes Module3.source, temp
 If Not Module3.decrypt Then Exit Function
 getCRC32 Module3.target, temp
 stringToBytesA tempString, bufferIn
 If Not bytesSame(bufferIn, temp) Then Exit Function
 Erase buffer, bufferIn
 mergeBytes buffer, Module3.target
 setStringSafe tempString
 exchangeData = True
End Function
Private Sub resetTempWebPage()
 If tempWebPage = vbNullString Then
  tempWebPage = TrimEnd(TrimEnd(App.path, "\"), "/") & "\~wichat.tmp"
 End If
 resetFile tempWebPage
End Sub
Public Function checkID(ByVal str As String) As Boolean
 checkID = False
 If Len(str) < 1 Then Exit Function
 Dim i As Integer
 For i = 1 To Len(str)
  If Mid(str, i, 1) < "0" Or Mid(str, i, 1) > "9" Then Exit Function
 Next i
 str = str & vbNullChar
 If Len(str) > MaxIDLen Then Exit Function
 checkID = True
End Function
Private Function checkPW(ByVal str As String) As Boolean
 If Len(str) > 16 Then Exit Function
 Dim i As Integer
 For i = 1 To Len(str)
  If Not (Mid(str, i, 1) >= "0" And Mid(str, i, 1) <= "9" Or Mid(str, i, 1) >= "A" And Mid(str, i, 1) <= "Z" Or Mid(str, i, 1) >= "a" And Mid(str, i, 1) <= "z") Then Exit Function
 Next i
 checkPW = True
End Function
Private Function formatID(ByRef id As String) As String 'Unicode String Only
  formatID = id & String(MaxIDLen - Len(id), Chr(0))
End Function
Private Function isReasonable(ByRef str() As Byte) As Boolean
 isReasonable = False
 Static i As Integer
 If Not byteArrayIsDimed(str) Then Exit Function
 If UBound(str) < 16 Then Exit Function
 For i = 0 To 7
  If str(i) < 32 Or str(i) > 126 Then Exit Function
 Next i
 isReasonable = True
End Function
Private Function getActiveSession() As Integer
 Static i As Integer
 For i = 1 To MaxSession
  If sessionList(i).active = True Then getActiveSession = i: Exit For
 Next i
 If i > MaxSession Then getActiveSession = 0
End Function
Private Function getFunctionAddr(ByRef FuncProc As Long) As Long
 getFunctionAddr = FuncProc
End Function


#If IS_DEBUG Then

Private Sub test()
' Dim a() As Byte, b() As Byte
' 'stringToBytesW String(32, Chr(1)), a
' stringToBytesW "123", a
' encode a, "9699946964066660", b, 7, False
' End
'Module3.loadKey "2323893283855600124312"
'stringToBytesW "Haha!", Module3.source
'Module3.decrypt
'mergeBytes Module3.target, Module3.source
'Module3.decrypt
'MsgBox bytesToStringW(Module3.target, 0, UBound(Module3.target))
'stringToBytesW "9182783573352291", a
'Dim temp As String
'stringToBytesW "3698437218020163", b
'bytesToStringA b, temp, 0, UBound(b)
'encode a, temp, b, 16, True
'MsgBox bytesToStringW(b, 0, UBound(b))
'End
'Dim c As String, d As String
'd = "00000000000000000000"
'fuse ByVal ("5157459" & vbNullChar), d, c
'fuse_R c, d, c
'MsgBox c
End Sub
Public Sub sprint(ByVal str As String)
 Dim a() As Byte, i As Integer, l As Integer
 l = LenB(str)
 ReDim a(l - 1)
 For i = 0 To l - 1
  a(i) = AscB(MidB(str, i + 1, 1))
 Next i
 str = ""
 For i = 0 To l - 1
  str = str & hex(a(i)) & " "
 Next i
 Debug.Print str
End Sub

#End If
