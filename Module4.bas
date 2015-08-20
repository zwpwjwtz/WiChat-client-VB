Attribute VB_Name = "Module4"
Option Explicit

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_DONT_CACHE = &H4000000
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000

#If Not IS_WINE Then

Private Const FEATURE_DISABLE_NAVIGATION_SOUNDS = 21
Private Const SET_FEATURE_ON_THREAD = 1
Private Const SET_FEATURE_ON_PROCESS = 2
'Private Const SET_FEATURE_IN_REGISTRY = 4
Private Const SET_FEATURE_ON_THREAD_LOCALMACHINE = 8
Private Const SET_FEATURE_ON_THREAD_INTRANET = 16
Private Const SET_FEATURE_ON_THREAD_TRUSTED = 32
Private Const SET_FEATURE_ON_THREAD_INTERNET = 64
Private Const SET_FEATURE_ON_THREAD_RESTRICTED = 128

#End If

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInternetHandle As Long) As Boolean
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal lpszServerName As String, ByVal nProxyPort As Integer, ByVal lpszUsername As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hInternetSession As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, ByVal lpszReferer As String, ByVal lpszAcceptTypes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Integer
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As Long, ByVal lOptionalLength As Long) As Boolean
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal buffer As Long, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function GetLastError Lib "kernel32" () As Long

#If Not IS_WINE Then
Private Declare Function CoInternetSetFeatureEnabled Lib "urlmon.dll" (FeatureEntry As Long, dwFlags As Long, fEnable As Boolean) As Long
#End If


Private Const AcceptAll As String * 4 = "*/*" & vbNullChar
Private Const BrowserAgent As String * 12 = "Mozilla/4.0" & vbNullChar
Private Const MaxRequestCount As Integer = 2

#If IS_LOCAL_SERVER Then
Private Const RootServer As String = "127.0.0.1."
#Else
Private Const RootServer As String = "dns.wichat.org"
#End If
Private AccServerList() As String
Private RecServerList() As String

Private Const QueryHeader As String * 8 = "WiChatCQ"
Private Const QueryHeaderLen As Integer = 8
Private Const QueryGetAcc As Integer = 1
Private Const QueryGetRec As Integer = 2
Private Const ResponseHeader As String * 8 = "WiChatSR"
Private Const ResponseHeaderLen As Integer = 8


Public Const RESPONSE_NONE As Integer = 0
Public Const RESPONSE_SUCCESS As Integer = 1
Public Const RESPONSE_BUSY As Integer = 2
Public Const RESPONSE_INVALID As Integer = 3
Public Const RESPONSE_DEVICE_UNSUPPORTED = 4
Public Const RESPONSE_FAILED As Integer = 5
Public Const RESPONSE_IN_MAINTANANCE As Integer = 8

Private iBuffer() As Byte, oBuffer() As Byte


Public Function init(Optional refresh As Boolean = False) As Boolean
 Static hasInited As Boolean
 If hasInited And Not refresh Then init = True: Exit Function
 init = False
 
 #If Not IS_WINE Then
 CoInternetSetFeatureEnabled FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_THREAD, True
 CoInternetSetFeatureEnabled FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_PROCESS, True
 CoInternetSetFeatureEnabled FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_THREAD_LOCALMACHINE, True
 CoInternetSetFeatureEnabled FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_THREAD_INTRANET, True
 CoInternetSetFeatureEnabled FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_THREAD_TRUSTED, True
 CoInternetSetFeatureEnabled FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_THREAD_INTERNET, True
 CoInternetSetFeatureEnabled FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_THREAD_INTERNET, True
 #End If
 
 Dim r As Integer
 r = getServerList
 Select Case r
  Case 0
   hasInited = True
   init = True
  Case 1
   showMsg "Cannot connect to server. Please check your network.", , vbExclamation
  Case 2
   showMsg "Server error." & vbCrLf & "Please contact technical support for more detail.", , vbCritical
  Case 3
   showMsg "WiChat service is currently unavailable." & vbCrLf & "Please try again later.", , vbExclamation
  Case 4
   showMsg "This version of WiChat is no longer supported." & vbCrLf & "Please try another version.", , vbCritical
  Case Else
   showMsg "Unknown error occurs." & vbCrLf & "Please contact technical support for more detail.", , vbCritical
 End Select
End Function
Public Function getServerList() As Integer '0=OK; 1=No Networking; 2=Server No Response; 3=Server Error; 4=Version Error;-1=Other Unknown Error

 getServerList = 0
 Dim received As Long, i As Integer
 Dim tempString As String, p As Long
 
 stringToBytesW QueryHeader, iBuffer
 ReDim Preserve iBuffer(UBound(iBuffer) + 2), oBuffer(1023)
 iBuffer(QueryHeaderLen) = QueryDevice
 iBuffer(QueryHeaderLen + 1) = QueryGetAcc
 
 received = HttpRequest(RootServer, 80, "/Root/query/index.php", "POST", VarPtr(iBuffer(0)), UBound(iBuffer) + 1, oBuffer)
 If received < 1 Then getServerList = 1: Exit Function
 Select Case oBuffer(ResponseHeaderLen)
  Case RESPONSE_NONE, RESPONSE_BUSY
   getServerList = 2: Exit Function
  Case RESPONSE_SUCCESS
   If oBuffer(ResponseHeaderLen + 1) < 0 Then getServerList = 3: Exit Function
   ReDim AccServerList(oBuffer(ResponseHeaderLen + 1) - 1)
   AccServerList = Split(bytesToStringW(oBuffer, ResponseHeaderLen + 2, UBound(oBuffer)), vbNullChar, oBuffer(ResponseHeaderLen + 1))
   If Not stringArrayIsDimed(AccServerList) Then getServerList = 3: Exit Function
'   tempString = bytesToString(oBuffer, 2, UBound(oBuffer))
'   p = 1: i = -1
'   Do While p < Len(tempString)
'    i = i + 1
'    p = InStr(p, tempString, vbNullChar)
'    If p = 0 Then p = Len(tempString)
'    AccServerList(i) = Left(tempString, p)
'    tempString = Right(Len(tempString) - p)
'   Loop
'   If i >= 0 Then
'    ReDim Preserve AccServerList(i)
'   Else
'    getServerList = 3
'   End If
  Case RESPONSE_DEVICE_UNSUPPORTED
   getServerList = 4: Exit Function
  Case Else
   getServerList = -1: Exit Function
 End Select
 
 iBuffer(QueryHeaderLen + 1) = QueryGetRec
 received = HttpRequest(RootServer, 80, "/Root/query/index.php", "POST", VarPtr(iBuffer(0)), UBound(iBuffer) + 1, oBuffer)
 If received < 1 Then getServerList = 1: Exit Function
 Select Case oBuffer(ResponseHeaderLen)
  Case RESPONSE_NONE, RESPONSE_INVALID
   getServerList = 2: Exit Function
  Case RESPONSE_SUCCESS
   If oBuffer(ResponseHeaderLen + 1) < 0 Then getServerList = 3: Exit Function
   ReDim RecServerList(oBuffer(ResponseHeaderLen + 1) - 1)
   RecServerList = Split(bytesToStringW(oBuffer, ResponseHeaderLen + 2, UBound(oBuffer)), vbNullChar, oBuffer(ResponseHeaderLen + 1))
   If Not stringArrayIsDimed(RecServerList) Then getServerList = 3: Exit Function
  Case RESPONSE_BUSY, RESPONSE_IN_MAINTANANCE
   getServerList = 3: Exit Function
  Case Else
   getServerList = -1: Exit Function
 End Select
 'For i = 0 To UBound(AccServerList): MsgBox (AccServerList(i)): Next i
End Function

Public Function sendRequest(serverID As Integer, URL As String, content() As Byte, buffer() As Byte) As Boolean
 sendRequest = False
 If Not init Then Exit Function
 Dim server As String
 If serverID = 1 Then
  If Not stringArrayIsDimed(AccServerList) Then Exit Function
  server = trimNull(AccServerList(0))
 Else
  If Not stringArrayIsDimed(RecServerList) Then Exit Function
  server = trimNull(RecServerList(0))
 End If
 Erase iBuffer
 mergeBytes iBuffer, content
 addHeader iBuffer, QueryHeader
 
 Static received As Long, Count As Integer
 ReDim buffer(MaxBlockSize)
 #If IS_DEBUG Then
   received = HttpRequest(server, 80, URL, "POST", VarPtr(iBuffer(0)), UBound(iBuffer) + 1, buffer)
 #Else
 For Count = 1 To MaxRequestCount
  received = HttpRequest(server, 80, URL, "POST", VarPtr(iBuffer(0)), UBound(iBuffer) + 1, buffer)
  If received > 0 Then Exit For
 Next Count
 #End If
 If received > ResponseHeaderLen Then
  ReDim Preserve buffer(received - 1)
  If bytesToStringW(buffer, 0, ResponseHeaderLen - 1) <> ResponseHeader Then Exit Function
  removeBytes buffer, 0, ResponseHeaderLen
  sendRequest = True
 End If
End Function

Private Function HttpRequest(strHostName As String, intPort As Integer, strUrl As String, strMethod As String, bytePostData As Long, ByVal lngPostDataLen As Long, byteReceive() As Byte) As Long
    Static hInternet As Long, hConnect As Long, hRequest As Long
    Static lngNumberOfBytesRead As Long, buffer(1023) As Byte
    Static blnRet As Boolean, p As Long, c As Integer
    
    HttpRequest = 0
    
    '打开一个Session会话
    hInternet = InternetOpen(BrowserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hInternet = 0 Then
        'MsgBox "hInternetOpen函数调用失败！"
        GoTo Ret0
    End If
    
    '连接服务器
    hConnect = InternetConnect(hInternet, strHostName, intPort, vbNullString, "HTTP/1.1", INTERNET_SERVICE_HTTP, 0, 0)
    If hConnect = 0 Then
        'MsgBox "InternetConnect函数调用失败！"
        GoTo Ret0
    End If
     
    '创建一个请求
    hRequest = HttpOpenRequest(hConnect, strMethod, strUrl, "HTTP/1.1", vbNullString, VarPtr(AcceptAll), INTERNET_FLAG_DONT_CACHE, 0)
    If hRequest = 0 Then
        'MsgBox "HttpOpenRequest函数调用失败！"
        GoTo Ret0
    End If
    
    blnRet = HttpSendRequest(hRequest, vbNullString, 0, bytePostData, lngPostDataLen)
    If blnRet = False Then
        'MsgBox "HttpSendRequest函数调用失败！"
        GoTo Ret0
    End If
    
    p = 0
    Do
        blnRet = InternetReadFile(hRequest, VarPtr(buffer(0)), 1024, lngNumberOfBytesRead)
        If blnRet = False Or Not CBool(lngNumberOfBytesRead) Then Exit Do
        If p + lngNumberOfBytesRead > UBound(byteReceive) Then Exit Do
        For c = 0 To lngNumberOfBytesRead - 1
         byteReceive(p + c) = buffer(c)
        Next c
        p = p + lngNumberOfBytesRead
    Loop
    HttpRequest = p
Ret0:
    If hRequest Then Call InternetCloseHandle(hRequest)
    If hConnect Then Call InternetCloseHandle(hConnect)
    If hInternet Then Call InternetCloseHandle(hInternet)
End Function


Private Sub addHeader(buffer() As Byte, header As String) 'ANSI only
 Dim l As Integer
 l = Len(header)
 If l = 0 Then Exit Sub
 ReDim Preserve buffer(UBound(buffer) + l)
 Dim i As Long
 For i = UBound(buffer) To l Step -1
  buffer(i) = buffer(i - l)
 Next i
 For i = 1 To l
  buffer(i - 1) = AscB(Mid(header, i, 1))
 Next i
End Sub
