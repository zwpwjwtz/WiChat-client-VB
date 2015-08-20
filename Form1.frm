VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WiChat!"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5295
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox textPW 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox textID 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label labelState 
      Caption         =   "Welcome to Wichat!"
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Account:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label labelVer 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "WiChat!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private inited As Boolean
'Private humanExit As Boolean

Private Sub Command1_Click()
 textID.text = Trim(textID.text)
 textPW.text = Trim(textPW.text)
 If textID.text = vbNullString Then showMsg "Please input ID number!", , vbCritical: Exit Sub
 If Not checkID(textID.text) Then showMsg "ID invalid!", , vbCritical: Exit Sub
 If textPW.text = vbNullString Then showMsg "Please input your password!", , vbCritical: Exit Sub
 showState "Connecting to network..."
 Select Case verifyAccount(textID.text, textPW.text)
  Case 1
   setStringSafe textID.text
   Me.visible = False
   DoEvents
   Module1.lastID = textID.text
   Load FormPanel
   FormPanel.show
   Unload Me
  Case 0
   showMsg "ID or password error. Please try again.", , vbCritical
   textPW.text = vbNullString
  Case Else
  
 End Select
End Sub

Private Sub Form_Activate()
 If inited Then Exit Sub
 If Me.textID.text <> "" Then Me.textPW.SetFocus
End Sub

Private Sub Form_Load()
 labelVer.Caption = VER
 Me.textID.text = lastID
' Me.textPW.text = "0000000000000000"
End Sub

Private Sub Form_Resize()
 If Me.WindowState = 2 Then Me.WindowState = 0
End Sub
Public Sub showState(text As String, Optional color As ColorConstants = ColorConstants.vbBlue)
 Me.labelState.Caption = text
 Me.labelState.ForeColor = color
 DoEvents
End Sub

Private Sub textPW_GotFocus()
 textPW.SelStart = 0
 textPW.SelLength = Len(textPW.text)
End Sub
