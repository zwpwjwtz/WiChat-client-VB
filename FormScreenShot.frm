VERSION 5.00
Begin VB.Form FormScreenShot 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "Save To File"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   240
      Left            =   -1000
      Top             =   -1000
      Width           =   255
   End
   Begin VB.Label labelColor 
      Caption         =   "(0,0,0)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label labelPos 
      Caption         =   "0 x 0"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "FormScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OriginalX    As Single        '区域起点X坐标
Private OriginalY    As Single        '区域起点的Y坐标
Private NewX    As Single
Private NewY    As Single
Private Status    As Integer              '当前状态（正在选择区域或者拖动区域）
Private rc    As RECT                    '区域的范围

Private Sub Command1_Click()
 savePic
 Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Command3_Click()
 savePicToFile
 Unload Me
End Sub

Private Sub Form_Load()
        setWindowTopMost Me.hwnd
        showFrame False
        Sleep 200
        DoEvents
        Me.AutoRedraw = True
        Screen.MousePointer = vbCrosshair                '  将光标改为十字型
        loadScreenToWindow Me.hDC
        Me.WindowState = 2
'        Shape1.Width = Screen.Width
'        Shape1.Height = Screen.Height
'        Shape1.Top = 0
'        Shape1.Left = 0
        Status = 1                   '绘图状态
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        If KeyAscii = vbKeyEscape Then Unload Me
End Sub


Private Sub Form_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
        showFrame False
        If Status = 1 Then                            '如果是抓取状态
                Shape1.visible = True
                Shape1.Width = 0
                Shape1.Height = 0
                OriginalX = X
                OriginalY = Y                                    '起点坐标
                Shape1.Left = OriginalX
                Shape1.Top = OriginalY
        Else                                                          '如果鼠标点在画好的选区内，则移动画好的选区
                rc.Left = Shape1.Left
                rc.Right = Shape1.Left + Shape1.Width
                rc.Top = Shape1.Top
                rc.Bottom = Shape1.Top + Shape1.Height
                If PtInRect(rc, X, Y) Then                  '如果按下的点位于区域内
                        NewX = X
                        NewY = Y                                          '则移动区域
                Else                                                      '否则重新画一个区域
                        Shape1.Width = 0
                        Shape1.Height = 0
                        OriginalX = X
                        OriginalY = Y
                        Shape1.Left = OriginalX
                        Shape1.Top = OriginalY
                        Status = 1                            '状态恢复到抓取
                End If
        End If
End Sub



Private Sub Form_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
        If Button = 1 Then
                If Status = 1 Then Status = 2
                OriginalX = Shape1.Left          '更新OriginalX，因为选择区域时可能会出现shape的right点大于left点
                OriginalY = Shape1.Top
        End If
        showFrame True
        moveFrame
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
        Static RGBColor   As Long, Red    As Long, Green    As Long, Blue    As Long

        RGBColor = GetPixel(Me.hDC, X, Y)
        GetRGBColors RGBColor, Red, Green, Blue
        labelColor.Caption = "RGB:( " & Red & ", " & Green & ", " & Blue & ") "
        If Button = 1 Then
                Shape1.visible = False
                If Status = 1 Then                           '如果是绘图状态
                        Screen.MousePointer = 2
                        If X > OriginalX And Y > OriginalY Then                              '根据鼠标位置调整shape1的大小和位置
                                Shape1.Move OriginalX, OriginalY, X - OriginalX, Y - OriginalY
                        ElseIf X < OriginalX And Y > OriginalY Then
                                Shape1.Move X, OriginalY, OriginalX - X, Y - OriginalY
                        ElseIf X > OriginalX And Y < OriginalY Then
                                Shape1.Move OriginalX, Y, X - OriginalX, OriginalY - Y
                        ElseIf X < OriginalX And Y < OriginalY Then
                                Shape1.Move X, Y, OriginalX - X, OriginalY - Y
                        End If
                        labelPos.Caption = "Scale: " & Shape1.Width & " x " & Shape1.Height                                 '显示当前区域的大小
                Else                                                              '如果是移动状态
                        Screen.MousePointer = 5
                        Shape1.Left = OriginalX - (NewX - X)
                        Shape1.Top = OriginalY - (NewY - Y)
                        If Shape1.Left < 0 Then Shape1.Left = 0             '使区域不超过屏幕
                        If Shape1.Top < 0 Then Shape1.Top = 0
                        If Shape1.Left + Shape1.Width > Screen.Width / 15 Then Shape1.Left = Screen.Width / 15 - Shape1.Width
                        If Shape1.Top + Shape1.Height > Screen.Height / 15 Then Shape1.Top = Screen.Height / 15 - Shape1.Height
                        moveFrame
                End If
                Shape1.visible = True
        End If
End Sub
Private Sub moveFrame()
 Dim X As Single, Y As Single
 X = Shape1.Left + Shape1.Width
 Y = Shape1.Top + Shape1.Height
 If (X + 289) * Screen.TwipsPerPixelX > Screen.Width Then X = Screen.Width / Screen.TwipsPerPixelX - 289
 If (Y + 49) * Screen.TwipsPerPixelY > Screen.Height Then Y = Screen.Height / Screen.TwipsPerPixelY - 49
 Shape2.Left = X: Shape2.Top = Y
 labelPos.Left = X + 8: labelPos.Top = Y + 8
 labelColor.Left = labelPos.Left: labelColor.Top = Y + 24
 Command1.Left = X + 136: Command1.Top = labelPos.Top
 Command2.Left = X + 232: Command2.Top = labelPos.Top
 Command3.Left = X + 168: Command3.Top = labelPos.Top
End Sub
Private Sub showFrame(show As Boolean)
 Shape2.visible = show
 labelPos.visible = show
 labelColor.visible = show
 Command1.visible = show
 Command2.visible = show
 Command3.visible = show
End Sub

Private Sub Form_DblClick()
   If PtInRect(rc, NewX, NewY) Then savePic
End Sub
Private Sub savePic()
 showFrame False
 Shape1.visible = False
 Sleep 200                                                   '有时候没有这两句会使得shape1也显示在截取的区域里
 DoEvents
 Dim tempFile As String
 tempFile = getScreenShotFileName(".jpg")
 If Not ScreenCaptureToFile(Shape1.Left, Shape1.Top, Shape1.Left + Shape1.Width, Shape1.Top + Shape1.Height, tempFile, 2) Then
  showMsg "Cannot make a screen shot.", , vbCritical
 Else
  FormPanel.addToText "[/i=" & tempFile & "/]"
 End If
End Sub
Private Sub savePicToFile()
 ModuleAddIn1.Title = "Save screen shooting as"
 ModuleAddIn1.InitDir = App.path
 ModuleAddIn1.fileName = "WiChatCap_" & Format(Now, "yyyymmddhhmmss")
 ModuleAddIn1.Filter = "Bitmap File(*.bmp)|(*.bmp)|JPEG File(*.jpg)|(*.jpg)"
 ModuleAddIn1.ExtentionIndex = 2
 ModuleAddIn1.ShowSave Me
 If ModuleAddIn1.fileName = vbNullString Then Exit Sub
 If ModuleAddIn1.ExtentionIndex = 1 Then
  ModuleAddIn1.fileName = formatFileSuffix(ModuleAddIn1.fileName, ".bmp")
 Else
  ModuleAddIn1.fileName = formatFileSuffix(ModuleAddIn1.fileName, ".jpg")
 End If
 showFrame False
 Shape1.visible = False
 Sleep 200                                                   '有时候没有这两句会使得shape1也显示在截取的区域里
 DoEvents
 If Not ScreenCaptureToFile(Shape1.Left, Shape1.Top, Shape1.Left + Shape1.Width, Shape1.Top + Shape1.Height, ModuleAddIn1.fileName, ModuleAddIn1.ExtentionIndex) Then
   showMsg "Cannot save screen shooting to file.", , vbCritical
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Screen.MousePointer = 0
 FormPanel.show
End Sub
