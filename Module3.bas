Attribute VB_Name = "Module3"
Option Explicit

#If IS_WINE Then
Private Declare Function CRC32 Lib "C:\windows\system32\wichat.dll" (ByVal pBufferIn As Long, ByVal inLength As Long, ByVal pBufferOut As Long, ByVal outLength As Long) As Boolean
Private Declare Function SHA1 Lib "C:\windows\system32\wichat.dll" (ByVal pBufferIn As Long, ByVal inLength As Long, ByVal pBufferOut As Long, ByVal outLength As Long) As Boolean
Private Declare Function BlowFish Lib "C:\windows\system32\wichat.dll" (ByVal pBufferIn As Long, ByVal inLength As Long, ByVal pKey As Long, ByVal pBufferOut As Long, ByVal outLength As Long, ByVal mode As Long) As Boolean
#Else
Private Declare Function CRC32 Lib "wichat.dll" (ByVal pBufferIn As Long, ByVal inLength As Long, ByVal pBufferOut As Long, ByVal outLength As Long) As Boolean
Private Declare Function SHA1 Lib "wichat.dll" (ByVal pBufferIn As Long, ByVal inLength As Long, ByVal pBufferOut As Long, ByVal outLength As Long) As Boolean
Private Declare Function BlowFish Lib "wichat.dll" (ByVal pBufferIn As Long, ByVal inLength As Long, ByVal pKey As Long, ByVal pBufferOut As Long, ByVal outLength As Long, ByVal mode As Long) As Boolean
#End If

Public Const MinSeedLength = 8
Public Const MinKeyLength As Integer = 16
Public Const MaxKeyLength As Integer = 128
Public Const MaxBlockSize As Long = 64000
Private Const defaultSeed As String = "2HthEwoi@goTiuhfdg91/TGS#RT4iOI32J2G0"
Private Const ENC_KEY_SEED As String = "9foiw2H$$GDSF#T$GS0fdWERWeG1u032FHd39f0wlog23a11idlk4IJG0gGRR0"
Private Const ENC_DELTA_DEFAULT As String = "`-jvDj34hjG]vb 0-r 32-ug11`JWaepoj 1#@f12?#"
Private nineCell(8) As Integer
Private key As String

Public source() As Byte, target() As Byte

Private Function checkDigits(number As String) As Boolean
 Dim i As Integer, flag As Boolean
 flag = True
 For i = 1 To Len(number)
  If Mid$(number, i, 1) > "9" Or Mid$(number, i, 1) < "0" Then flag = False: Exit For
 Next i
 checkDigits = flag
End Function

Private Function hasFactor(number As Integer) As Integer
 If number = 0 Then hasFactor = 0: Exit Function
 number = Abs(number)
 Dim Count As Integer, temp As Integer
 Count = 0
 For temp = 1 To number
  If number / temp = Int(number / temp) Then Count = Count + 1
 Next temp
 hasFactor = Count
End Function

Public Function genKey(Optional seed As String = vbNullString, Optional hex As Boolean = False) As String
 If Len(seed) < MinSeedLength Then seed = seed & defaultSeed
 Dim bit As Integer, temp As String, i As Integer
 Do
  bit = Int(Rnd * MaxKeyLength + 1)
 Loop Until bit > MinKeyLength And hasFactor(bit) > 4
 Dim p As Integer, l As Integer
 p = 1
 l = Len(seed)
 Randomize Timer
 If hex Then
   For i = 1 To bit
   Select Case Int(Rnd * 4 + 1)
   Case 1
    temp = temp & ChrB((Asc(Mid$(seed, p, 1)) + Int(Rnd * 10 + 1)) Mod 256)
   Case 2
    temp = temp & ChrB(Abs((Asc(Mid$(seed, p, 1)) - Int(Rnd * 10 + 1)) Mod 256))
   Case 3
    temp = temp & ChrB((Asc(Mid$(seed, p, 1)) * Int(Rnd * 10 + 1)) Mod 256)
   Case 4
    temp = temp & ChrB((Asc(Mid$(seed, p, 1)) Mod Int(Rnd * 10 + 1)) Mod 256)
   End Select
   p = Int(Rnd * l) + 1
  Next i
 Else
  For i = 1 To bit
   Select Case Int(Rnd * 4 + 1)
   Case 1
    temp = temp & (Asc(Mid$(seed, p, 1)) + Int(Rnd * 10 + 1)) Mod 10
   Case 2
    temp = temp & Abs((Asc(Mid$(seed, p, 1)) - Int(Rnd * 10 + 1)) Mod 10)
   Case 3
    temp = temp & (Asc(Mid$(seed, p, 1)) * Int(Rnd * 10 + 1)) Mod 10
   Case 4
    temp = temp & (Asc(Mid$(seed, p, 1)) Mod Int(Rnd * 10 + 1)) Mod 10
   End Select
   p = Int(Rnd * l) + 1
  Next i
 End If
 genKey = temp
End Function
Public Sub fuse(ByVal str As String, delta As String, ByRef buffer As String, Optional base As Integer = 128)
  Dim i As Integer, j As Integer
  If LenB(delta) < 8 Then delta = ENC_DELTA_DEFAULT
  j = 1
  buffer = vbNullString
  For i = 1 To LenB(str)
   buffer = buffer & ChrB((AscB(MidB(str, i, 1)) + AscB(MidB(delta, j, 1)) * 3 + base) Mod 256)
   j = j Mod LenB(delta) + 1
  Next i
End Sub
Public Sub fuse_R(ByVal str As String, delta As String, ByRef buffer As String, Optional base As Integer = 128)
  Dim i As Integer, j As Integer
  If LenB(delta) < 8 Then delta = ENC_DELTA_DEFAULT
  j = 1
  buffer = vbNullString
  For i = 1 To LenB(str)
   buffer = buffer & ChrB((256 + (AscB(MidB(str, i, 1)) - AscB(MidB(delta, j, 1)) * 3 - base) Mod 256) Mod 256)
   j = j Mod LenB(delta) + 1
  Next i
End Sub
Public Function encrypt() As Boolean
 encrypt = False
 If Not byteArrayIsDimed(source) Then Exit Function
 If UBound(source) > MaxBlockSize Then Exit Function
 
 Dim keyArray() As Integer
 If Not buildKey(keyArray) Then Exit Function
 
 Dim il As Long, ol As Long, pBI As Long, pBO As Long
 Dim tempChar As String, tempByte As Byte
 Dim pK As Integer, kl As Integer
 kl = UBound(keyArray) + 1
 pK = 1

    'On Error GoTo encryptError
    
    tempByte = 0: tempChar = vbNullString
 
    il = UBound(source)
    ol = Int(il * 1.25) + 1
    ReDim target(ol)
     
    pBO = 0
    For pBI = 0 To il
     tempChar = shift(Format(source(pBI), "000"), keyArray(pK, 0) + keyArray(pK, 1))
     Select Case keyArray(pK, 0)
     Case 0
        tempChar = rotate(tempChar, keyArray(pK, 1))
     Case 1
        tempChar = invert(tempChar, keyArray(pK, 1))
     Case 2
        tempChar = reflect(tempChar, keyArray(pK, 1))
     Case 3
        tempChar = rotoreflect(tempChar, keyArray(pK, 1))
     End Select
     target(pBO) = tempByte * 4 ^ ((4 - pBI Mod 4) Mod 4) Or Val(tempChar) \ 4 ^ (pBI Mod 4 + 1)
     tempByte = (4 ^ (pBI Mod 4 + 1) - 1) And Val(tempChar)
     pK = (pK + source(pBI)) Mod kl
     pBO = pBO + 1
     If pBO Mod 5 = 4 Then
      target(pBO) = tempByte
      tempByte = 0
      pBO = pBO + 1
     End If
    Next pBI
    If pBI Mod 4 > 0 And pBO <= ol Then target(pBO) = 4 ^ (3 - (pBI - 1) Mod 4) * tempByte
    
  Erase keyArray
  encrypt = True
  Exit Function
encryptError:
  Erase target
  showMsg "Internal error occurs.", , vbCritical
End Function
Public Function decrypt() As Boolean
 decrypt = False
 If Not byteArrayIsDimed(source) Then Exit Function
 If UBound(source) > MaxBlockSize Then Exit Function
 
 Dim keyArray() As Integer
 If Not buildKey(keyArray) Then Exit Function
 
 Dim il As Long, ol As Long, pBI As Long, pBO As Long
 Dim tempChar As String, tempVar As Integer
 Dim pK As Integer, kl As Integer
 kl = UBound(keyArray) + 1
 pK = 1

  
   'On Error GoTo decryptError
   tempVar = 0: tempChar = vbNullString
   
   il = UBound(source)
   ol = Int((il + 5) / 1.25) - 4
   
   ReDim target(ol)
   
    For pBI = 0 To il - 1
     tempVar = (source(pBI) And (4 ^ (4 - pBI Mod 5) - 1)) * 4 ^ ((pBI + 1) Mod 5) Or source(pBI + 1) \ 4 ^ (3 - pBI Mod 5)
     tempChar = shift(Format(tempVar, "000"), -keyArray(pK, 0) - keyArray(pK, 1))
     Select Case keyArray(pK, 0)
     Case 0
        tempChar = rotate_R(tempChar, keyArray(pK, 1))
     Case 1
        tempChar = invert_R(tempChar, keyArray(pK, 1))
     Case 2
        tempChar = reflect_R(tempChar, keyArray(pK, 1))
     Case 3
        tempChar = rotoreflect_R(tempChar, keyArray(pK, 1))
     End Select
     target(pBO) = Val(tempChar)
     pK = (pK + target(pBO)) Mod kl
     pBO = pBO + 1
     If pBI Mod 5 = 3 Then pBI = pBI + 1
    Next pBI
    
  Erase keyArray
  decrypt = True
  Exit Function
decryptError:
  Erase target
  showMsg "Internal error occurs.", , vbCritical
End Function

Public Sub clear(Optional clearKey As Boolean = True)
 Erase source
 Erase target
 If clearKey Then setStringSafe key
End Sub
Public Function loadKey(keyString As String) As Boolean
 key = Trim(keyString)
 If checkKey Then
  loadKey = True
 Else
  loadKey = False
  setStringSafe key
 End If
End Function
Private Function checkKey() As Boolean
 checkKey = False
 If key = "" Or Len(key) < MinKeyLength Or Len(key) > MaxKeyLength Then Exit Function
 If Not checkDigits(key) Then Exit Function
 checkKey = True
End Function
Private Function buildKey(keyArray() As Integer) As Boolean
 On Error GoTo buildError
 Dim Count As Integer, i As Integer
 Count = 0
 ReDim keyArray(Int(Len(key) / 2) - 1, 1)
 For i = 1 To (UBound(keyArray) + 1) * 2 Step 2
  keyArray(Count, 0) = Val(Mid$(key, i, 1)) Mod 4
  keyArray(Count, 1) = Val(Mid$(key, i + 1, 1))
  Count = Count + 1
 Next i
 buildKey = True
 Exit Function
buildError:
 buildKey = False
End Function


Sub init()
 nineCell(0) = 1
 nineCell(1) = 2
 nineCell(2) = 3
 nineCell(3) = 6
 nineCell(4) = 9
 nineCell(5) = 8
 nineCell(6) = 7
 nineCell(7) = 4
 nineCell(8) = 5
End Sub

Private Function rotate(number As String, angle As Integer) As String
 Static shift As Integer, value As String, i As Integer, j As Integer
 If angle < 0 Then
  shift = (8 + (angle Mod 8)) Mod 8
 Else
  shift = angle Mod 8
  If shift = 0 Then
   rotate = number
   Exit Function
  End If
 End If
 For i = 1 To Len(number)
  value = Val(Mid(number, i, 1))
  If value = 0 Then
   rotate = rotate & "0"
  ElseIf value = 5 Then
   rotate = rotate & "5"
  Else
     For j = 0 To 7
      If nineCell(j) = value Then Exit For
     Next j
     rotate = rotate & nineCell((j + shift) Mod 8)
  End If
 Next i
End Function
Private Function rotate_R(number As String, angle As Integer) As String
 Static shift As Integer, value As String, i As Integer, j As Integer
 If angle < 0 Then
  shift = Abs((8 - angle) Mod 8)
 Else
  shift = 8 - angle Mod 8
  If shift = 8 Then
   rotate_R = number
   Exit Function
  End If
 End If
 For i = 1 To Len(number)
  value = Val(Mid(number, i, 1))
  If value = 0 Then
   rotate_R = rotate_R & "0"
  ElseIf value = 5 Then
   rotate_R = rotate_R & "5"
  Else
     For j = 0 To 7
      If nineCell(j) = value Then Exit For
     Next j
     rotate_R = rotate_R & nineCell((j + shift) Mod 8)
  End If
 Next i
End Function
Private Function invert(number As String, mirror As Integer) As String
 Static temp As String, value As Integer, i As Integer
 temp = vbNullString
 For i = 1 To Len(number)
  value = Val(Mid(number, i, 1))
  If value = 0 Then
   value = 5
  ElseIf value = 5 Then
   value = 0
  Else
   value = 10 - value
  End If
  temp = temp & value
 Next i
 invert = reflect(temp, mirror)
End Function
Private Function invert_R(number As String, mirror As Integer) As String
 Static temp As String, value As Integer, i As Integer
 temp = reflect_R(number, mirror)
 For i = 1 To Len(temp)
  value = Val(Mid(temp, i, 1))
  If value = 0 Then
   value = 5
  ElseIf value = 5 Then
   value = 0
  Else
   value = 10 - value
  End If
  invert_R = invert_R & value
 Next i
End Function
Private Function reflect(number As String, mirror As Integer) As String
 Static shift As Integer, value As Integer, i As Integer
 shift = mirror Mod 4
 For i = 1 To Len(number)
  value = Val(Mid(number, i, 1))
  If value <> 9 Then
   Select Case shift
    Case 0
     value = ((value \ 3 + 2) Mod 3) * 3 + value Mod 3
    Case 1
     value = (value \ 3) * 3 + (value + 1) Mod 3
    Case 2
     value = (value \ 3) * 3 + (value + 2) Mod 3
    Case 3
     value = ((value \ 3 + 1) Mod 3) * 3 + value Mod 3
   End Select
  End If
  reflect = reflect & value
 Next i
End Function
Private Function reflect_R(number As String, mirror As Integer) As String
 Static shift As Integer, value As Integer, i As Integer
 shift = 3 - mirror Mod 4
 For i = 1 To Len(number)
  value = Val(Mid(number, i, 1))
  If value <> 9 Then
   Select Case shift
    Case 0
     value = ((value \ 3 + 2) Mod 3) * 3 + value Mod 3
    Case 1
     value = (value \ 3) * 3 + (value + 1) Mod 3
    Case 2
     value = (value \ 3) * 3 + (value + 2) Mod 3
    Case 3
     value = ((value \ 3 + 1) Mod 3) * 3 + value Mod 3
   End Select
  End If
  reflect_R = reflect_R & value
 Next i
End Function
Private Function rotoreflect(number As String, mixedVar As Integer) As String
 rotoreflect = reflect(rotate(number, mixedVar), mixedVar)
End Function
Private Function rotoreflect_R(number As String, mixedVar As Integer) As String
 rotoreflect_R = rotate_R(reflect_R(number, mixedVar), mixedVar)
End Function
Private Function shift(number As String, digit As Integer) As String
 If digit >= 0 Then
  digit = digit Mod Len(number)
  shift = Right(number, digit) & Left(number, Len(number) - digit)
 Else
  digit = (-digit) Mod Len(number)
  shift = Right(number, Len(number) - digit) & Left(number, digit)
 End If
End Function

Public Sub encode(ByRef var() As Byte, key As String, ByRef buffer() As Byte, Optional outLen As Long = 0, Optional dec As Boolean = False)
'Should use ANSI key only
 Dim Length As Long, temp As Integer, tempString As String
 Dim i As Long, bufferIn() As Byte, bufferOut(19) As Byte
 If LenB(key) < MinKeyLength Then key = ENC_KEY_SEED
 Length = UBound(var) + 1
 If outLen < 1 Then
     If (Length < 1) Then outLen = 16 Else outLen = Length
 End If
 mergeBytes bufferIn, var
 Do
        SHA1 VarPtr(bufferIn(0)), Length, VarPtr(bufferOut(0)), 20
        mergeBytes bufferIn, bufferOut, False
        Length = UBound(bufferIn) + 1
        If dec Then
         For i = 0 To Length - 1
          bufferIn(i) = (CInt(bufferIn(i)) + AscB(MidB(key, (i + 22) Mod LenB(key) + 1, 1))) Mod 10 + 48
         Next i
        Else
         For i = 0 To Length - 1
          bufferIn(i) = (CInt(bufferIn(i)) + AscB(MidB(key, (i + 22) Mod LenB(key) + 1, 1))) Mod 256
         Next i
        End If
 Loop Until Length >= outLen
 ReDim Preserve bufferIn(outLen - 1)
 Erase buffer
 mergeBytes buffer, bufferIn
End Sub
'Public Function expectedLen(ByRef source() As Byte) As Long
' Dim i As Long
' i = Int(UBound(source) * 1.25) + 1
' If i > MaxBlockSize Then expectedLen = 0 Else expectedLen = i
'End Function

Public Function getCRC32(ByRef source() As Byte, ByRef result() As Byte) As Boolean '4 bytes Output
  getCRC32 = False
  If Not byteArrayIsDimed(source()) Then Exit Function
  ReDim result(3)
  getCRC32 = CRC32(VarPtr(source(LBound(source))), UBound(source) - LBound(source) + 1, VarPtr(result(0)), 4)
End Function
Public Function getSHA1(source() As Byte, Optional toDec As Boolean = False) As String  '20 bytes Output
 getSHA1 = vbNullString
 If Not byteArrayIsDimed(source()) Then Exit Function
 Dim result(19) As Byte, i As Integer, target As String
 If SHA1(VarPtr(source(LBound(source))), UBound(source) - LBound(source) + 1, VarPtr(result(0)), 20) = False Then Exit Function
 target = vbNullString
 If toDec Then
  For i = 0 To 19
   target = target & Chr(result(i) Mod 10 + 48)
  Next i
 Else
  For i = 0 To 19
   target = target & ChrB(result(i))
  Next i
 End If
 getSHA1 = target
End Function

Public Function cEncrypt(var() As Byte, ByRef key As String, res() As Byte) As Boolean
 cEncrypt = False
 If LenB(key) < MinKeyLength Or LenB(key) > MaxKeyLength Then Exit Function
 If (UBound(var) + 1) Mod 8 <> 0 Then ReDim Preserve var(UBound(var) + 8 - (UBound(var) + 1) Mod 8)
 ReDim res(UBound(var))
 Dim keyArray() As Byte
 stringToBytesA LeftB(key, 56), keyArray
 BlowFish VarPtr(var(0)), UBound(var) + 1, VarPtr(keyArray(0)), VarPtr(res(0)), UBound(res) + 1, 1
 setByteArraySafe keyArray
 cEncrypt = True
End Function
Public Function cDecrypt(var() As Byte, ByRef key As String, res() As Byte) As Boolean
 cDecrypt = False
 If LenB(key) < MinKeyLength Or LenB(key) > MaxKeyLength Then Exit Function
 If (UBound(var) + 1) Mod 8 <> 0 Then Exit Function
 ReDim res(UBound(var))
 Dim keyArray() As Byte
 stringToBytesA LeftB(key, 56), keyArray
 BlowFish VarPtr(var(0)), UBound(var) + 1, VarPtr(keyArray(0)), VarPtr(res(0)), UBound(res) + 1, 2
 setByteArraySafe keyArray
 cDecrypt = True
End Function
Public Function DecToHex(var As String) As String
 DecToHex = vbNullString
 Dim i As Integer
 For i = 1 To Len(var)
  DecToHex = DecToHex & ChrB((Asc(Mid(var, i, 1)) - 48) * 24 + Rnd * 22)
 Next i
End Function
Public Function HexToDec(var As String) As String
 HexToDec = vbNullString
 Dim i As Integer
 For i = 1 To LenB(var)
  HexToDec = HexToDec & Chr(AscB(MidB(var, i, 1)) \ 24 + 48)
 Next i
End Function
