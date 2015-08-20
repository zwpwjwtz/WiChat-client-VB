Attribute VB_Name = "Module6"
Private Const DIB_RGB_COLORS = 0
Private Const OBJ_BITMAP = 7

Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Private Type BITMAPFILEHEADER
        bfType(0 To 1) As Byte
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RECT
        Left  As Long
        Top  As Long
        Right  As Long
        Bottom  As Long
End Type

  Private Type GUID
        Data1   As Long
        Data2   As Integer
        Data3   As Integer
        Data4(0 To 7)       As Byte
  End Type
    
  Private Type GdiplusStartupInput
        GdiplusVersion   As Long
        DebugEventCallback   As Long
        SuppressBackgroundThread   As Long
        SuppressExternalCodecs   As Long
  End Type
    
  Private Type EncoderParameter
        GUID   As GUID
        NumberOfValues   As Long
        type   As Long
        value   As Long
  End Type
    
  Private Type EncoderParameters
        Count   As Long
        Parameter   As EncoderParameter
  End Type

Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
'Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal fileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long


Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long


Public Function SaveBMP(ByVal hDC As Long, fileName As String) As Boolean
    SaveBMP = False
    Dim hBitmap As Long
    hBitmap = GetCurrentObject(hDC, OBJ_BITMAP) '取得位图
    If hBitmap = 0 Then Exit Function
    
    Dim bm As BITMAP
    If GetObject(hBitmap, Len(bm), bm) = 0 Then Exit Function '得到位图信息
    
    Dim bmih As BITMAPINFOHEADER
    bmih.biSize = Len(bmih)
    bmih.biWidth = bm.bmWidth
    bmih.biHeight = bm.bmHeight
    bmih.biBitCount = 24
    bmih.biPlanes = 1
    bmih.biSizeImage = ((bmih.biWidth * 3 + 3) And &H7FFFFFFC) * bmih.biHeight '计算大小
    
    ReDim MapData(1 To bmih.biSizeImage) As Byte
    If GetDIBits(hDC, hBitmap, 0, bmih.biHeight, MapData(1), bmih, DIB_RGB_COLORS) = 0 Then Exit Function '取得位图数据
    
    Dim hF As Integer
    hF = FreeFile(1)
    
    On Error Resume Next
    Open fileName For Binary As hF
    If Err.number Then hF = -1
    On Error GoTo 0
    If hF = -1 Then Exit Function
    
    Dim bmfh As BITMAPFILEHEADER
    bmfh.bfType(0) = Asc("B")
    bmfh.bfType(1) = Asc("M")
    bmfh.bfOffBits = Len(bmfh) + Len(bmih)
    Put hF, , bmfh
    
    Put hF, , bmih
    
    Put hF, , MapData
    
    Close hF
    
    SaveBMP = True
    
End Function
    
Public Function SaveJPG(ByVal hBitmap As Long, ByVal fileName As String, Optional ByVal quality As Byte = 80) As Boolean
  SaveJPG = False
  If hBitmap = 0 Then Exit Function
  
  Dim tSI     As GdiplusStartupInput
  Dim lRes     As Long
  Dim lGDIP     As Long
  Dim lBitmap     As Long
    
        '   Initialize   GDI+
        tSI.GdiplusVersion = 1
        lRes = GdiplusStartup(lGDIP, tSI)
          
        If lRes = 0 Then
          
              '   Create   the   GDI+   bitmap
              '   from   the   image   handle
              lRes = GdipCreateBitmapFromHBITMAP(hBitmap, 0, lBitmap)
          
              If lRes = 0 Then
                    Dim tJpgEncoder     As GUID
                    Dim tParams     As EncoderParameters
                      
                    '   Initialize   the   encoder   GUID
                    CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
                
                    '   Initialize   the   encoder   parameters
                    tParams.Count = 1
                    With tParams.Parameter     '   Quality
                          '   Set   the   Quality   GUID
                          CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB3505E7EB}"), .GUID
                          .NumberOfValues = 1
                          .type = 1
                          .value = VarPtr(quality)
                    End With
                      
                    '   Save   the   image
                    lRes = GdipSaveImageToFile(lBitmap, StrPtr(fileName), tJpgEncoder, tParams)
                                                              
                    '   Destroy   the   bitmap
                    GdipDisposeImage lBitmap
                      
              End If
                
              '   Shutdown   GDI+
              GdiplusShutdown lGDIP
              SaveJPG = True
    
        End If
          
  End Function

Public Sub GetRGBColors(ByVal RGBColor As Long, ByRef RedColor As Long, ByRef GreenColor As Long, ByRef BlueColor As Long)
        RedColor = RGBColor Mod 256
        GreenColor = (RGBColor \ &H100) Mod 256
        BlueColor = (RGBColor \ &H10000) Mod 256
End Sub
Public Sub loadScreenToWindow(hDC As Long)
 Dim SourceDC As Long
 SourceDC = CreateDC("DISPLAY" & vbNullChar, 0, 0, 0)
 BitBlt hDC, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, SourceDC, 0, 0, vbSrcCopy
 DeleteDC SourceDC
End Sub
Public Function ScreenCaptureToFile(Left As Long, Top As Long, Right As Long, Bottom As Long, fileName As String, fileType As Integer) As Boolean 'Type:0=None;1=BMP;2=JPG
        ScreenCaptureToFile = False
        If fileType = 0 Then Exit Function
        Dim rWidth   As Long, rHeight           As Long
        Dim SourceDC   As Long, DestDC          As Long
        Dim BHandle   As Long
        rWidth = Right - Left
        rHeight = Bottom - Top
        SourceDC = FormScreenShot.hDC
        DestDC = CreateCompatibleDC(SourceDC)
        BHandle = CreateCompatibleBitmap(SourceDC, rWidth, rHeight)
        SelectObject DestDC, BHandle
        BitBlt DestDC, 0, 0, rWidth, rHeight, SourceDC, Left, Top, vbSrcCopy
        
        If fileType = 2 Then
         If SaveJPG(BHandle, fileName) Then ScreenCaptureToFile = True
        Else
         If SaveBMP(DestDC, fileName) Then ScreenCaptureToFile = True
        End If
        DeleteDC DestDC
End Function
