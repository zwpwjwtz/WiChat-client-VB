VERSION 5.00
Begin VB.Form FormProfile 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4845
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox textNote 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox textMsg 
      BackColor       =   &H80000009&
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Offline Msg."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label labelID 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "FormProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public readOnly As Boolean


Private Sub Command1_Click()
 If labelID.Caption = nowID Then Exit Sub
 If Len(textNote.text) > 15 Then showMsg "Note cannot be longer than 15 characters.", , vbCritical: Exit Sub
 changeNote labelID.Caption, textNote.text
End Sub

Private Sub Form_Activate()
 textMsg.Locked = Me.readOnly
End Sub

Private Sub Form_Click()
 checkOfflineMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
 checkOfflineMsg
End Sub
Private Sub checkOfflineMsg()
 If Me.readOnly Then Exit Sub
 textMsg.text = Trim(textMsg.text)
 If textMsg.text <> nowOfflineMsg Then
  If Not changeMsg(textMsg.text) Then textMsg.text = nowOfflineMsg
 End If
End Sub
