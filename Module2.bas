Attribute VB_Name = "Module2"
Option Explicit
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Private Const CP_UTF8 = 65001
Public Const LOCAL_EMTION_DEFAULT_PATH = "\WiChatRes\emotions\"
Private Const LOCAL_PICTURE_DEAFULT_IMAGE = "\WiChatRes\default_image.png"
Private Const LOCAL_PICTURE_DEAFULT_FILE = "\WiChatRes\default_file.png"

Public Sub peelNull(ByRef byteArray() As Byte)
 If UBound(byteArray) > 1 Then
    Static i As Long, j As Long
    j = 1
    For i = 2 To UBound(byteArray) Step 2
        byteArray(j) = byteArray(i)
        j = j + 1
    Next i
 End If
 ReDim Preserve byteArray(CInt((UBound(byteArray) - 1) / 2))
End Sub
Public Function trimNull(strString As String) As String 'Perfect and efficient Function. Great!
    Static l As Long
    l = InStr(1, strString, Chr(0))
    If l = 1 Then
        trimNull = ""
    ElseIf l > 0 Then
        trimNull = Left$(strString, l - 1)
    Else
        trimNull = strString
    End If
End Function
Public Function addQuo(ByRef str As String)
 addQuo = """" & str & """"
End Function
Public Function getFileName(fullURL As String) As String
 Static flag As Boolean, c As Integer
 fullURL = TrimEnd(fullURL, "\")
 fullURL = TrimEnd(fullURL, "/")
 For c = Len(fullURL) To 1 Step -1
  If Mid$(fullURL, c, 1) = "\" Or Mid$(fullURL, c, 1) = "/" Then flag = True: Exit For
 Next c
 If flag <> True Then getFileName = fullURL Else getFileName = Right(fullURL, Len(fullURL) - c)
End Function
Function getFilePath(fullURL As String) As String
 Dim i As Integer, j As Integer
 If Len(fullURL) < 1 Then Exit Function
 i = InStrRev(fullURL, "/")
 j = InStrRev(fullURL, "\")
 If i < j And i > 0 Or j = 0 Then j = i
 If j > 1 Then
  getFilePath = Left(fullURL, j - 1)
  If InStr(getFilePath, "\") = 0 Then getFilePath = getFilePath & "\"
 Else
  getFilePath = fullURL
 End If
End Function
Public Function TrimEnd(str As String, Optional TrimChar As String = vbNullChar) As String 'TrimNulls的变体
 Static l As Long: l = Len(str)
 Do While l > 0
  If InStr(TrimChar, Mid(str, l, 1)) <= 0 Then Exit Do
  l = l - 1
 Loop
 TrimEnd = Left(str, l)
End Function
Public Function fixStringTail(ByRef str As String) As String
 If LenB(str) Mod 2 > 0 Then fixStringTail = LeftB(str, LenB(str) - 1) Else fixStringTail = str
End Function
Public Sub setStringSafe(ByRef str As String)
 str = String(LenB(str), vbNullChar)
 str = vbNullString
End Sub
Public Sub setByteArraySafe(ByRef var() As Byte)
 If Not byteArrayIsDimed(var) Then Exit Sub
 Dim i As Integer
 For i = LBound(var) To UBound(var)
  var(i) = 0
 Next i
End Sub
Public Sub bytesToStringA(ByRef byteArray() As Byte, ByRef buffer As String, lowerBound As Long, upperBound As Long) 'ANSI string
 Static i As Long, temp As String
 If upperBound - lowerBound < 1 Or upperBound < 0 Then
  buffer = vbNullString
  Exit Sub
 End If
 temp = vbNullString
 For i = lowerBound To upperBound
  temp = temp & ChrB((byteArray(i)))
 Next i
 buffer = temp
End Sub
Public Function bytesToStringW(ByRef byteArray() As Byte, lowerBound As Long, upperBound As Long) As String 'Unicode string
 Static temp As String, i As Long
 If upperBound - lowerBound < 0 Or upperBound < 0 Then
  bytesToStringW = vbNullString
  Exit Function
 End If
 temp = vbNullString
 For i = lowerBound To upperBound
  temp = temp & ChrW((byteArray(i)))
 Next i
 bytesToStringW = temp
End Function
Public Function stringToBytesW(str As String, buffer() As Byte) As Boolean 'Used by VB original String
 stringToBytesW = False
 If Len(str) = 0 Then Exit Function Else ReDim buffer(Len(str) - 1)
 Static i As Long
 For i = 1 To Len(str)
  buffer(i - 1) = AscB(Mid(str, i, 1))
 Next i
 stringToBytesW = True
End Function
Public Function stringToBytesA(str As String, buffer() As Byte) As Boolean 'Used by received string
 stringToBytesA = False
 If LenB(str) = 0 Then Exit Function Else ReDim buffer(LenB(str) - 1)
 Static i As Long
 For i = 1 To LenB(str)
  buffer(i - 1) = AscB(MidB(str, i, 1))
 Next i
 stringToBytesA = True
End Function
Public Function stringAToW(str As String, lowerBound As Long, upperBound As Long) As String
 stringAToW = vbNullString
 If lowerBound < 0 Or upperBound > LenB(str) Or upperBound < lowerBound Then Exit Function
 Static i As Long
 For i = 0 To upperBound - lowerBound
  stringAToW = stringAToW & Chr(AscB(MidB(str, lowerBound + i, 1)))
 Next i
End Function
Public Function stringWToA(str As String) As String
 Static i As Long
 Dim temp As String
 For i = 1 To Len(str)
  temp = temp & ChrB(AscB(Mid(str, i, 1)))
 Next i
 If Len(str) Mod 2 > 0 Then temp = temp & vbNullChar
 stringWToA = temp
End Function
Public Sub mergeBytes(dest() As Byte, src() As Byte, Optional append As Boolean = True)
 If Not byteArrayIsDimed(src) Then Exit Sub
 Static i As Long, l As Long
 
 If byteArrayIsDimed(dest) Then
  l = UBound(dest) + 1
  ReDim Preserve dest(UBound(dest) + UBound(src) + 1)
 Else
  l = 0
  ReDim dest(UBound(src))
  If Not append Then
   For i = 0 To UBound(src)
   dest(i) = src(i)
   Next i
  End If
 End If
 
 If append Then
  For i = 0 To UBound(src)
   dest(l + i) = src(i)
  Next i
 Else
  For i = UBound(dest) To UBound(dest) - l + 1 Step -1
   dest(i) = dest(i - UBound(src) - 1)
  Next i
  For i = 0 To UBound(src)
   dest(i) = src(i)
  Next i
 End If
End Sub
Public Sub copyBytes(ByRef dest() As Byte, ByRef src() As Byte, lowerBound As Long, upperBound As Long)
 If Not byteArrayIsDimed(src) Then Exit Sub
 If lowerBound < 0 Or upperBound < lowerBound Then Exit Sub
 Static i As Long, j As Long
 ReDim dest(upperBound - lowerBound)
 j = 0
 For i = lowerBound To upperBound
  dest(j) = src(i)
  j = j + 1
 Next i
End Sub
Public Function removeBytes(ByRef src() As Byte, pos As Long, Length As Long) As Boolean
 removeBytes = False
 If Not byteArrayIsDimed(src) Then Exit Function
 If pos + Length - 1 < UBound(src) Then
  Static i As Long
  For i = pos To UBound(src) - Length
   src(i) = src(i + Length)
  Next i
  ReDim Preserve src(UBound(src) - Length)
  removeBytes = True
 Else
  Erase src
  removeBytes = True
 End If
End Function
Public Function bytesSame(ByRef var1() As Byte, ByRef var2() As Byte) As Boolean
 bytesSame = False
 If Not (byteArrayIsDimed(var1) Or byteArrayIsDimed(var2)) Then Exit Function
 If UBound(var1) - LBound(var1) <> UBound(var2) - LBound(var2) Then Exit Function
 Static i As Long, d As Long
 d = LBound(var2) - LBound(var1)
 For i = LBound(var1) To UBound(var1)
  If var1(i) <> var2(i + d) Then Exit Function
 Next i
 bytesSame = True
End Function
Public Function UnicodeToUtf8(UCS As String, buffer() As Byte) As Boolean
 UnicodeToUtf8 = False
 Dim lLength As Long, lBufferSize As Long, lResult As Long

 lLength = Len(UCS)
 'If lLength = 0 Then Exit Function
 lBufferSize = lLength * 3 + 1
 ReDim buffer(lBufferSize - 1)
 lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UCS), lLength, buffer(0), lBufferSize, vbNullString, 0)
 If lResult <> 0 Then
  lResult = lResult - 1
  ReDim Preserve buffer(lResult)
  UnicodeToUtf8 = True
 Else
  ReDim buffer(0)
 End If
End Function

Public Function bytesToInt(var() As Byte, pos As Long) As Integer
 bytesToInt = 0
 If Not byteArrayIsDimed(var) Then Exit Function
 If UBound(var) < pos + 1 Then Exit Function
 bytesToInt = CInt(var(pos)) + &H100 * var(pos + 1)
End Function
Public Function bytesToLong(var() As Byte, pos As Long) As Long
 bytesToLong = 0
 If Not byteArrayIsDimed(var) Then Exit Function
 If UBound(var) < pos + 3 Then Exit Function
 bytesToLong = var(pos) + &H100& * var(pos + 1) + &H10000 * var(pos + 2) + &H1000000 * var(pos + 3)
End Function
Public Sub intToBytes(buffer() As Byte, value As Integer)
 ReDim buffer(1)
 buffer(0) = value And &HFF
 buffer(1) = (value \ &H100) And &HFF
End Sub
Public Sub longToBytes(buffer() As Byte, value As Long)
 ReDim buffer(3)
 buffer(0) = value And &HFF
 buffer(1) = (value \ &H100) And &HFF
 buffer(2) = (value \ &H10000) And &HFF
 buffer(3) = (value \ &H1000000) And &HFF
End Sub
Public Function render(ByVal content As String) As String 'To WebBrowser
 Dim p1 As Long, p2 As Long, temp As String, tempFile As String
 Dim c1 As Long, c2 As Long
 'Deal with line breaker
 p1 = 0
 Do
  p2 = p1
  p1 = InStr(p2 + 1, content, vbCrLf)
  If p1 > 0 And p1 < Len(content) Then temp = temp & Mid(content, p2 + 1, p1 - p2 - 1) & "<p>&nbsp;</p>" Else Exit Do
 Loop
 content = Mid(content, p2 + 1, Len(content) - p1)
 'Deal with emotion
 Do
  p1 = InStr(p1 + 1, content, "[/em")
  If p1 = 0 Then Exit Do
  p2 = InStr(p1, content, "]")
  If p2 = 0 Then Exit Do
  If p2 - p1 > 7 Then GoTo con1
   Select Case Val(Mid(content, p1 + 4, p2 - p1 - 4))
   Case 1 To MaxEmotion
    temp = App.path & LOCAL_EMTION_DEFAULT_PATH & Val(Mid(content, p1 + 4, p2 - p1 - 4)) & ".gif"
   Case Else
    temp = ""
  End Select
  content = Left(content, p1 - 1) & "<img src=""file:///" & temp & """ />" & Right(content, Len(content) - p2)
con1:
 Loop
 Do
  p1 = InStr(p1 + 1, content, "[/f=")
  If p1 < 2 Then Exit Do
  p2 = InStr(p1, content, "/]")
  If p2 < 2 Then Exit Do
  tempFile = Mid(content, p1 + 4, p2 - p1 - 4)
  c1 = InStrRev(content, "<div class=s", p1 - 1): c2 = InStrRev(content, "<div class=r", p1 - 1)
  If c1 > 0 And c2 < c1 Then
   temp = "You sent file """ & getFileName(tempFile) & """ to him/her."
  Else
   temp = "He/she sent file """ & getFileName(tempFile) & """ to you."
  End If
  content = Left(content, p1 - 1) & "<div style=""width:300px;border:1px solid;""><img src=""file:///" & App.path & LOCAL_PICTURE_DEAFULT_FILE & """style=""float:left;"" /><div style=""float:right"">" & temp & " <a href=""file:///" & tempFile & """ target=_blank>View the file</a></div></div>" & Right(content, Len(content) - p2 - 1)
 Loop
 Do
  p1 = InStr(p1 + 1, content, "[/i=")
  If p1 = 0 Then Exit Do
  p2 = InStr(p1, content, "/]")
  If p2 = 0 Then Exit Do
  temp = Mid(content, p1 + 4, p2 - p1 - 4)
  If Not fileExists(temp) Then temp = App.path & LOCAL_PICTURE_DEAFULT_IMAGE
  content = Left(content, p1 - 1) & "<img src=""file:///" & temp & """ />" & Right(content, Len(content) - p2 - 1)
 Loop
 render = content
End Function
Public Function translate(ByVal str As String) As String
 Dim p1 As Long, p2 As Long, temp As String
 
 p1 = 0
 Do
  p2 = p1
  p1 = InStr(p2 + 1, str, "<")
  If p1 > 0 Then temp = temp & Mid(str, p2 + 1, p1 - p2 - 1) & "&lt;" Else Exit Do
 Loop
 str = temp & Right(str, Len(str) - p2)
 
 temp = vbNullString
 p1 = 0
 Do
  p2 = p1
  p1 = InStr(p2 + 1, str, ">")
  If p1 > 0 Then temp = temp & Mid(str, p2 + 1, p1 - p2 - 1) & "&rt;" Else Exit Do
 Loop
 str = temp & Right(str, Len(str) - p2)
 
' temp = vbNullString
' p1 = 0
' Do
'  p2 = p1
'  p1 = InStr(p2 + 1, str, " ")
'  If p1 > 0 Then temp = temp & Mid(str, p2 + 1, p1 - p2 - 1) & "&nbsp;" Else Exit Do
' Loop
' str = temp & Right(str, Len(str) - p2)

 translate = str
End Function
Public Function dataXMLize(ByVal str As String, ByRef dest() As Byte) As Boolean
 dataXMLize = False
 Erase dest
 Dim p As Long, p1 As Long, p2 As Long, temp As String, temp2 As String
 Dim buffer() As Byte, buffer2() As Byte, filen As Integer
 p = -1: p1 = -1
 stringToBytesA StrConv(str, vbFromUnicode), buffer
 If UBound(buffer) Mod 2 = 0 Then ReDim Preserve buffer(UBound(buffer) + 1)
 'If LenB(str) Mod 2 > 0 Then str = str & vbNullChar
 
 Erase dest
 Do
  p2 = p1
  stringToBytesW "[/", buffer2:  p1 = inBytes(buffer, buffer2, p2 + 1)
  stringToBytesW "/]", buffer2:  p2 = inBytes(buffer, buffer2, p1 + 1)
  If p1 < 1 Or p2 < 1 Then Exit Do
  copyBytes buffer2, buffer, p + 1, p1 - p - 2 'Previous data before <D>
  mergeBytes dest, buffer2
  bytesToStringA buffer, temp2, p1 + 4, p2 - 1 'File Info
  temp2 = StrConv(temp2, vbUnicode)
  Select Case Chr(buffer(p1 + 2))
   Case "i"
    If Not fileExists(temp2) Then
     stringToBytesW "<D t=i l=-1 >", buffer2
     mergeBytes dest, buffer2
    Else
     stringToBytesW "<D t=i l=" & FileLen(temp2) & vbNullChar & " >", buffer2
     mergeBytes dest, buffer2
     filen = FreeFile
     ReDim buffer2(FileLen(temp2))
     Open temp2 For Binary Access Read As #filen
     Get #filen, , buffer2
     Close #filen
     mergeBytes dest, buffer2
    End If
    stringToBytesW "</D>", buffer2
    mergeBytes dest, buffer2
    p = p2 + 1
   Case "f"
    If Not fileExists(temp2) Then
     stringToBytesW "<D t=f l=-1 >", buffer2
     mergeBytes dest, buffer2
    Else
     stringToBytesW "<D t=f l=" & FileLen(temp2) & " n=", buffer2
     mergeBytes dest, buffer2
     stringToBytesA StrConv(getFileName(temp2) & vbNullChar, vbFromUnicode), buffer2
     mergeBytes dest, buffer2
     stringToBytesW ">", buffer2
     mergeBytes dest, buffer2
     filen = FreeFile
     ReDim buffer2(FileLen(temp2))
     Open temp2 For Binary Access Read As #filen
     Get #filen, , buffer2
     Close #filen
     mergeBytes dest, buffer2
    End If
    stringToBytesW "</D>", buffer2
    mergeBytes dest, buffer2
    p = p2 + 1
   Case Else
  End Select
 Loop
 copyBytes buffer2, buffer, p + 1, UBound(buffer) 'Last data after </D>
 mergeBytes dest, buffer2
 dataXMLize = True
End Function
Public Function dataUnxmlize(ByRef src() As Byte, relatedID As String) As String
 dataUnxmlize = vbNullString
 If Not byteArrayIsDimed(src) Then Exit Function
 Dim p As Long, p1 As Long, p2 As Long, p3 As Long, l As Long, temp As String
 Dim buffer() As Byte, temp2 As String, tempFile As String
 p = -1: temp = vbNullString
 Do
  stringToBytesW "<D t=", buffer:  p1 = inBytes(src, buffer, p + 1)
  stringToBytesW ">", buffer:  p2 = inBytes(src, buffer, p1 + 1)
  If p1 < 0 Or p2 < 0 Then Exit Do
  bytesToStringA src, temp2, p1 + 5, p2
  temp2 = StrConv(temp2, vbUnicode)
  l = InStr(1, temp2, "n=")
  If l > 0 Then
   tempFile = Mid(temp2, l + 2, InStr(l + 2, temp2, vbNullChar) - l - 2)
  Else
   tempFile = vbNullString
  End If
  l = InStr(1, temp2, "l=")
  l = Val(Mid(temp2, l + 2, InStr(l + 2, temp2, " ") - l - 2))
  stringToBytesW "</D>", buffer:  p3 = inBytes(src, buffer, p2 + l)
  If p3 > 0 Then
   If l > 0 Then
    Select Case Left(temp2, 1)
     Case "f"
      copyBytes buffer, src, p2 + 1, p2 + l
      tempFile = recordPath & "\record\" & relatedID & "\cache\" & tempFile
      bytesToStringA src, temp2, p + 1, p1 - 1
      temp = temp & temp2
      If extractFile(buffer, tempFile) Then temp = temp & stringWToA("[/f=") & StrConv(tempFile, vbFromUnicode) & stringWToA("/]")
     Case "i"
      copyBytes buffer, src, p2 + 1, p2 + l
      tempFile = recordPath & "\record\" & relatedID & "\cache\" & getSHA1(buffer, True) & "." & getImageType(buffer)
      bytesToStringA src, temp2, p + 1, p1 - 1
      temp = temp & temp2
      If extractFile(buffer, tempFile) Then temp = temp & stringWToA("[/i=") & StrConv(tempFile, vbFromUnicode) & stringWToA("/]")
     Case Else
      bytesToStringA src, temp2, p + 1, p1 - 1
      temp = temp & temp2
    End Select
   End If
   p = p3 + 3
  Else
   Exit Do
  End If
 Loop
 bytesToStringA src, temp2, p + 1, UBound(src)
 temp = temp & temp2
 dataUnxmlize = StrConv(temp, vbUnicode)
End Function
Public Function inBytes(ByRef src() As Byte, toFind() As Byte, Optional start As Long = 0) As Long
 inBytes = -1
 If start < 0 Or Not byteArrayIsDimed(src) Or Not byteArrayIsDimed(toFind) Then Exit Function

 Dim i As Long, j As Long
 For i = start To UBound(src) - UBound(toFind)
  For j = 0 To UBound(toFind)
   If src(i + j) <> toFind(j) Then Exit For
  Next j
  If j > UBound(toFind) Then inBytes = i: Exit For
 Next i
End Function
Public Function addSenderInfo(ByRef str As String, decoration As fontStyle) As String
 Dim temp1 As String, temp2 As String
 If decoration.family <> "" Then temp1 = "face=""" & decoration.family & """"
 If decoration.color <> "" Then temp1 = temp1 & "color=""#" & decoration.color & """"
 If decoration.size > 0 Then temp1 = temp1 & intToSize(decoration.size)
 If temp1 <> vbNullString Then temp1 = "<font " & temp1 & ">": temp2 = "</font>"
 If (decoration.basic And &H1) Then temp1 = "<b>" & temp1: temp2 = temp2 & "</b>"
 If (decoration.basic And &H10) Then temp1 = "<i>" & temp1: temp2 = temp2 & "</i>"
 If (decoration.basic And &H100) Then temp1 = "<u>" & temp1: temp2 = temp2 & "</u>"
 If (decoration.basic And &H1000) Then temp1 = "<del>" & temp1: temp2 = temp2 & "</del>"
 Select Case decoration.align
  Case 1
   temp1 = " align=left>" & temp1
  Case 2
   temp1 = " align=center>" & temp1
  Case 3
   temp1 = " align=right>" & temp1
  Case Else
   temp1 = ">" & temp1
 End Select
 addSenderInfo = "<div class=s><b>" & nowID & "</b>&nbsp;&nbsp;" & Format(Now, TimeFormat) & "</div><div class=c" & temp1 & str & temp2 & "</div>"
End Function
Public Function colorToHex(value As String, Optional reverse As Boolean = False) As String
 If Not reverse Then
  Select Case value
  Case "Red"
   colorToHex = "FF0000"
  Case "Green"
   colorToHex = "00FF33"
  Case "Blue"
   colorToHex = "0066FF"
  Case "Yellow"
   colorToHex = "F5FB15"
  Case "Black"
   colorToHex = "000000"
  Case "White"
   colorToHex = "FFFFFF"
  Case "Grey"
   colorToHex = "999999"
  Case "Purple"
   colorToHex = "993399"
  Case "Orange"
   colorToHex = "FF6600"
  End Select
 Else
  Select Case value
  Case "FF0000"
   colorToHex = "Red"
  Case "00FF33"
   colorToHex = "Green"
  Case "0066FF"
   colorToHex = "Blue"
  Case "F5FB15"
   colorToHex = "Yellow"
  Case "000000"
   colorToHex = "Black"
  Case "FFFFFF"""
   colorToHex = "White"
  Case "999999"
   colorToHex = "Grey"
  Case "993399"
   colorToHex = "Purple"
  Case "FF6600"
   colorToHex = "Orange"
  End Select
 End If
End Function
Public Function intToSize(value As Integer) As String
 If value = 14 Then
  intToSize = ""
 Else
  intToSize = "size="""
  If value - 14 > 0 Then
   intToSize = intToSize & "+" & value - 14
  Else
   intToSize = intToSize & "-" & 14 - value
  End If
  intToSize = intToSize & """"
 End If
End Function
Public Function getStateImageIndex(stateNumber As Integer, Optional displayHide As Boolean = False) As Integer
 Select Case stateNumber
  Case State.onLine
   getStateImageIndex = 8
  Case State.Offline
   getStateImageIndex = 9
  Case State.Busy
   getStateImageIndex = 11
  Case State.Hide
   If displayHide Then getStateImageIndex = 10 Else getStateImageIndex = 9
  Case Else
   getStateImageIndex = 9
 End Select
End Function
Public Function getStateForImageIndex(index As Integer) As State
 Select Case index
  Case 8
   getStateForImageIndex = onLine
  Case 9
   getStateForImageIndex = Offline
  Case 11
   getStateForImageIndex = Busy
  Case Else
   getStateForImageIndex = Offline
 End Select
End Function
Public Sub getHTMLHeaderA(buffer() As Byte, docType As Integer) '0=Content Window;1=Input Window
 Dim temp() As Byte
 Erase buffer
 stringToBytesW "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd""><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title></title>", temp
 mergeBytes buffer, temp
 Select Case docType
  Case 0
   stringToBytesW "<style>div,span,ul,li{marin:0px;padding:0px;list-style-type:none;font-size:15px}p{text-align:left}img{border:0px;max-width:100%;height:auto}.s{color:#339966}.r{color:#0066CC}.a{color:#000000}.c{margin:0px 0px 15px 5px;width:350px;word-wrap:break-word}</style></head><body style=""width:360px;overflow:hidden;"">", temp
  Case 1
   stringToBytesW "<style>div,span,ul,li{marin:0px;padding:0px;list-style-type:none}p{text-align:left;}</style></head><body style=""width:98%;margin:2px;overflow:hidden""><textarea type=text id=" & textboxName & " cols=38 rows=3 style=""width:100%;overflow:hidden;border:0px;"">", temp
 End Select
 mergeBytes buffer, temp
End Sub
Public Sub getHTMLFooterA(buffer() As Byte, docType As Integer)
 Dim temp() As Byte
 Erase buffer
 Select Case docType
  Case 0
   stringToBytesW "</body></html>", temp
  Case 1
   stringToBytesW "</textarea></body></html>", temp
 End Select
 mergeBytes buffer, temp
End Sub
Public Sub fixSenderInfo(ByRef str As String)
 Dim i As Long
 Do
  i = InStr(i + 1, str, "<div class=s>")
  If i > 0 Then Mid(str, i + 11, 1) = "r" Else Exit Do
 Loop
End Sub
Public Function byteArrayIsDimed(ByRef arr() As Byte) As Boolean
 If SafeArrayGetDim(arr) > 0 Then byteArrayIsDimed = True Else byteArrayIsDimed = False
End Function
Public Function stringArrayIsDimed(ByRef arr() As String) As Boolean
If SafeArrayGetDim(arr) > 0 Then stringArrayIsDimed = True Else stringArrayIsDimed = False
End Function
Public Function stateToString(var As State) As String
 Select Case var
  Case State.Busy
   stateToString = "Busy"
  Case State.Hide
   stateToString = "Hide"
  Case State.Offline
   stateToString = "Offline"
  Case State.onLine
   stateToString = "Online"
  Case Else
   stateToString = vbNullString
 End Select
End Function
'Public Function getTempFileName() As String
' getTempFileName = "~wc" & Hour(Now) & Minute(Now) & Second(Now) & ".tmp"
'End Function

Private Function extractFile(ByRef src() As Byte, fileName As String) As Boolean
 extractFile = False
 If Not byteArrayIsDimed(src) Or fileName = vbNullString Then Exit Function
 If Dir(getFilePath(fileName)) = "" Then
  If Not createDirectory(getFilePath(fileName)) Then Exit Function
 End If
 Dim filen As Integer
 filen = FreeFile
 Open fileName For Binary Access Write As #filen
  Put #filen, , src
 Close #filen
 extractFile = True
End Function
Public Function resetFile(ByRef fileName As String, Optional fillZeroSize As Long = 0) As Boolean
 resetFile = False
 If fileName = "" Then Exit Function
 If fileExists(fileName) Then ShellAndWait "cmd /c del /F /Q " & addQuo(fileName), False
' Open FileName For Binary Access Write As #255
'  Do While fillZeroSize > 0
'   Put #255, , 0
'   fillZeroSize = fillZeroSize - 1
'  Loop
' Close #255
 resetFile = True
End Function
Public Function createDirectory(dirName As String) As Boolean
 ShellAndWait "cmd /c md " & addQuo(dirName), False
 createDirectory = dirExists(dirName)
End Function
Public Function fileExists(ByRef fileName As String) As Boolean
 On Error GoTo existError
 fileExists = False
 If fileName = vbNullString Then Exit Function
 If FileLen(fileName) > 0 Then
  fileExists = True
 Else
  If Dir(fileName, vbArchive + vbHidden + vbSystem + vbReadOnly) = getFileName(fileName) Then fileExists = True
 End If
 Exit Function
existError:
End Function
Public Function dirExists(ByVal dirName As String) As Boolean
 On Error GoTo existError
 dirName = TrimEnd(TrimEnd(dirName, "/"), "\")
 dirExists = False
 If dirName = vbNullString Then Exit Function
 If FileLen(dirName) > 0 Then Exit Function
 'FIXME:目录名比较时的大小写问题？
 If Dir(dirName, vbDirectory + vbArchive + vbHidden + vbSystem + vbReadOnly) = getFileName(dirName) Then dirExists = True
existError:
End Function
Public Function checkUserDirectory(ByRef id As String) As Boolean
 checkUserDirectory = False
 If Not dirExists(recordPath & "\record\" & id) Then
  If Not createDirectory(recordPath & "\record\" & id) Then Exit Function
 End If
 checkUserDirectory = True
End Function
Public Function getScreenShotFileName(suffix As String) As String
 getScreenShotFileName = vbNullString
 Dim dirName As String
 dirName = recordPath & "\record\" & nowID & "\cache"
 If Not dirExists(dirName) Then
  If Not createDirectory(dirName) Then Exit Function
 End If
 getScreenShotFileName = dirName & "\WiChatCap_" & Format(Now, "yyyymmddhhmmss") & suffix
End Function
Public Function formatFileSuffix(fileName As String, suffix As String) As String
 If LCase(Right(fileName, Len(suffix))) <> LCase(suffix) Then formatFileSuffix = fileName & suffix Else formatFileSuffix = fileName
End Function

Public Function URLEncode(ByRef strUrl As String) As String '
 Dim i As Long
 Dim tempStr As String
 For i = 1 To Len(strUrl)
  If Asc(Mid(strUrl, i, 1)) < 0 Then
   tempStr = "%" & Right(CStr(hex(Asc(Mid(strUrl, i, 1)))), 2)
   tempStr = "%" & Left(CStr(hex(Asc(Mid(strUrl, i, 1)))), Len(CStr(hex(Asc(Mid(strUrl, i, 1))))) - 2) & tempStr
   URLEncode = URLEncode & tempStr
  ElseIf (Asc(Mid(strUrl, i, 1)) >= 65 And Asc(Mid(strUrl, i, 1)) <= 90) Or (Asc(Mid(strUrl, i, 1)) >= 97 And Asc(Mid(strUrl, i, 1)) <= 122) Then
   URLEncode = URLEncode & Mid(strUrl, i, 1)
  Else
   URLEncode = URLEncode & "%" & hex(Asc(Mid(strUrl, i, 1)))
  End If
 Next i
End Function
  
Public Function URLDecode(ByRef strUrl As String) As String
 Dim i As Long
  
 If InStr(strUrl, "%") = 0 Then URLDecode = strUrl: Exit Function
  
 For i = 1 To Len(strUrl)
  If Mid(strUrl, i, 1) = "%" Then
   If Val("&H" & Mid(strUrl, i + 1, 2)) > 127 Then
    URLDecode = URLDecode & Chr(Val("&H" & Mid(strUrl, i + 1, 2) & Mid(strUrl, i + 4, 2)))
   i = i + 5
   Else
    URLDecode = URLDecode & Chr(Val("&H" & Mid(strUrl, i + 1, 2)))
    i = i + 2
   End If
  Else
   URLDecode = URLDecode & Mid(strUrl, i, 1)
  End If
 Next i
End Function

Public Function getImageType(ByRef src() As Byte) As String
 getImageType = vbNullString
 If UBound(src) < 4 Then Exit Function
 
 Dim FileHeader(4) As Byte, i As Integer
 For i = 0 To 4
  FileHeader(i) = src(i)
 Next i
 
    If (FileHeader(0) = 66) And (FileHeader(1) = 77) Then
      getImageType = "bmp"
      Exit Function
    End If
    If (FileHeader(0) = 255) And (FileHeader(1) = 216) Then
      getImageType = "jpg"
      Exit Function
    End If
    If (FileHeader(0) = 71) And (FileHeader(1) = 73) And (FileHeader(2) = 70) Then
      getImageType = "gif"
      Exit Function
    End If
    If (FileHeader(0) = 137) And (FileHeader(1) = 80) Then
      getImageType = "png"
      Exit Function
    End If
    If (FileHeader(0) = 73) And (FileHeader(1) = 73) Or (FileHeader(0) = 77) And (FileHeader(1) = 77) Then
      getImageType = "tiff"
      Exit Function
    End If
End Function
