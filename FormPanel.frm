VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormPanel 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8745
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame frameTop 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   885
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1561
         ButtonWidth     =   2725
         ButtonHeight    =   1455
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Home"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Notification"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Session"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Hide To Tray"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Height          =   390
      Left            =   7080
      TabIndex        =   3
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Friend"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh List"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listFriend 
      Height          =   5415
      Left            =   6240
      TabIndex        =   2
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Icon"
         Object.Width           =   476
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ID_Detail"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8760
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":0FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":1F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":2F16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":3EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":4E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":5E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":6DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":7D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":8D42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9000
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":961C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":9BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":A150
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":A6EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":AC84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":B21E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":B7B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":BD52
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":C2EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":C886
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":CE20
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":D3BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":D954
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":DEEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":E340
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":E792
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormPanel.frx":ED2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   9000
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame frameSession 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   6135
      Begin VB.CommandButton Command4 
         Caption         =   "▽"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   47
         Top             =   5040
         Width           =   255
      End
      Begin VB.Frame frameProcess 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Visible         =   0   'False
         Width           =   5415
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label labelProcessState 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            Height          =   255
            Left            =   600
            TabIndex        =   46
            Top             =   120
            Width           =   4335
         End
      End
      Begin VB.Frame frameFont 
         Caption         =   "Font Effect"
         Height          =   975
         Left            =   0
         TabIndex        =   24
         Top             =   2760
         Visible         =   0   'False
         Width           =   6135
         Begin VB.CheckBox Check2 
            Caption         =   "Italic"
            Height          =   255
            Left            =   840
            TabIndex        =   33
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Right"
            Height          =   255
            Left            =   4320
            TabIndex        =   36
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Center"
            Height          =   255
            Left            =   3360
            TabIndex        =   35
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Left"
            Height          =   255
            Left            =   2520
            TabIndex        =   34
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Bold"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox listFontColor 
            Height          =   300
            ItemData        =   "FormPanel.frx":F2C6
            Left            =   3960
            List            =   "FormPanel.frx":F2E5
            TabIndex        =   30
            Text            =   "Black"
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox listFontSize 
            Height          =   300
            ItemData        =   "FormPanel.frx":F327
            Left            =   2400
            List            =   "FormPanel.frx":F34C
            TabIndex        =   27
            Text            =   "15"
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox listFontFamily 
            Height          =   300
            ItemData        =   "FormPanel.frx":F37C
            Left            =   600
            List            =   "FormPanel.frx":F3AA
            TabIndex        =   26
            Text            =   "宋体"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Align"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   37
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   31
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Font"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView listEmotion 
         Height          =   2055
         Left            =   0
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3625
         View            =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         OLEDragMode     =   1
         FlatScrollBar   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList3"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   1.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   5040
         Width           =   1215
      End
      Begin SHDocVwCtl.WebBrowser textLog 
         Height          =   3420
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   6135
         ExtentX         =   10821
         ExtentY         =   6032
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1085
         MultiRow        =   -1  'True
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SHDocVwCtl.WebBrowser textInput 
         Height          =   975
         Left            =   0
         TabIndex        =   11
         Top             =   4080
         Width           =   6135
         ExtentX         =   10821
         ExtentY         =   1720
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   120
         TabIndex        =   7
         Top             =   3720
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Set Font"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Add Emotion"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert Picture"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert File"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Screen Shot"
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   5760
         Picture         =   "FormPanel.frx":F431
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   6135
      Begin VB.CommandButton buttonRefreshNote 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5160
         Picture         =   "FormPanel.frx":F9BB
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
      Begin MSComctlLib.ListView listNote 
         Height          =   4335
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7646
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   7409
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   3  'Dot
         Height          =   5055
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame frameHome 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   6135
      Begin VB.Frame frameInfo 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   1320
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton buttonCloseInfo 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3960
            Picture         =   "FormPanel.frx":FF45
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton buttonRefreshInfo 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3480
            Picture         =   "FormPanel.frx":104CF
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.Frame frameSetting 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4335
            Left            =   360
            TabIndex        =   17
            Top             =   600
            Width           =   4215
            Begin VB.Frame frameSetting2 
               Height          =   1455
               Left            =   0
               TabIndex        =   48
               Top             =   2760
               Width           =   4215
               Begin VB.Frame framePassword 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame1"
                  Height          =   855
                  Left            =   120
                  TabIndex        =   49
                  Top             =   1680
                  Width           =   3975
                  Begin VB.TextBox textPW 
                     Height          =   375
                     IMEMode         =   3  'DISABLE
                     Left            =   1200
                     MaxLength       =   16
                     PasswordChar    =   "*"
                     TabIndex        =   52
                     Top             =   0
                     Width           =   3015
                  End
                  Begin VB.CommandButton buttonPWOK 
                     Caption         =   "OK"
                     Height          =   375
                     Left            =   1560
                     TabIndex        =   51
                     Top             =   480
                     Width           =   975
                  End
                  Begin VB.CommandButton buttonPWCancel 
                     Caption         =   "Cancel"
                     Height          =   375
                     Left            =   3000
                     TabIndex        =   50
                     Top             =   480
                     Width           =   975
                  End
                  Begin VB.Label labelPWPrompt 
                     Height          =   495
                     Left            =   120
                     TabIndex        =   53
                     Top             =   0
                     Width           =   1095
                  End
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "Change My Password"
                  Height          =   495
                  Left            =   720
                  TabIndex        =   54
                  Top             =   600
                  Width           =   2775
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Account Security"
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
                  Left            =   2160
                  TabIndex        =   55
                  Top             =   120
                  Width           =   1935
               End
            End
            Begin VB.Frame frameSetting1 
               Height          =   2655
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Width           =   4215
               Begin VB.CheckBox Check4 
                  Caption         =   "Flash window or bubble"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   58
                  Top             =   2160
                  Width           =   3375
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "Browse"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   3240
                  TabIndex        =   43
                  Top             =   960
                  Width           =   855
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "Hide window when capture screen"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   56
                  Top             =   1440
                  Width           =   3855
               End
               Begin VB.TextBox textRecordPath 
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   120
                  TabIndex        =   42
                  Text            =   "(Path)"
                  Top             =   960
                  Width           =   3135
               End
               Begin VB.OptionButton Option2 
                  Caption         =   "Custom"
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
                  Left            =   2640
                  TabIndex        =   39
                  Top             =   600
                  Width           =   975
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Along with EXE file"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   480
                  TabIndex        =   40
                  Top             =   600
                  Value           =   -1  'True
                  Width           =   2055
               End
               Begin VB.Label Label8 
                  Caption         =   "When there is new message:"
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
                  Left            =   120
                  TabIndex        =   59
                  Top             =   1800
                  Width           =   2775
               End
               Begin VB.Label Label4 
                  Caption         =   "Chatting record path:"
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
                  Left            =   120
                  TabIndex        =   41
                  Top             =   360
                  Width           =   2775
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Session"
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
                  Left            =   2160
                  TabIndex        =   57
                  Top             =   120
                  Width           =   1935
               End
            End
         End
         Begin VB.Frame frameHelp 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   4215
            Left            =   360
            TabIndex        =   15
            Top             =   600
            Width           =   4215
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   2895
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   16
               Text            =   "FormPanel.frx":10A59
               Top             =   480
               Width           =   3855
            End
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Dot
            Height          =   4935
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   4695
         End
         Begin VB.Label labelInfoCaption 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Width           =   2775
         End
      End
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   825
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1455
         ButtonWidth     =   2302
         ButtonHeight    =   1455
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "My State"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "My Profile"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Settings"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Help"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exit"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label labelFriendList 
      Caption         =   "Friend List(0/0)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu menuHome 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menuHomeProfile 
         Caption         =   "My Profile"
      End
      Begin VB.Menu menuHomeState 
         Caption         =   "My State"
         Begin VB.Menu menuHomeStateOnline 
            Caption         =   "Online"
         End
         Begin VB.Menu menuHomeStateHide 
            Caption         =   "Hide"
         End
         Begin VB.Menu menuHomeStateBusy 
            Caption         =   "Busy"
         End
         Begin VB.Menu menuHomeStatel1 
            Caption         =   "-"
         End
         Begin VB.Menu menuHomeStateOffline 
            Caption         =   "Offline"
         End
      End
      Begin VB.Menu menuHomel1 
         Caption         =   "-"
      End
      Begin VB.Menu menuHomeHide 
         Caption         =   "Hide To Tray"
      End
      Begin VB.Menu menuHomel2 
         Caption         =   "-"
      End
      Begin VB.Menu menuHomeExit 
         Caption         =   "Exit WiChat"
      End
   End
   Begin VB.Menu menuTray 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menuTrayMain 
         Caption         =   "Show Main"
      End
      Begin VB.Menu menuTrayl1 
         Caption         =   "-"
      End
      Begin VB.Menu menuTrayStateOnline 
         Caption         =   "Online"
      End
      Begin VB.Menu menuTrayStateHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu menuTrayStateBusy 
         Caption         =   "Busy"
      End
      Begin VB.Menu menuTrayStateOffline 
         Caption         =   "Offline"
      End
      Begin VB.Menu menuTrayl2 
         Caption         =   "-"
      End
      Begin VB.Menu menuTrayExit 
         Caption         =   "Exit Wichat"
      End
   End
   Begin VB.Menu menuFriend 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menuFriendNewDialog 
         Caption         =   "New Dialog"
      End
      Begin VB.Menu menuFriendl1 
         Caption         =   "-"
      End
      Begin VB.Menu menuFriendShowInfo 
         Caption         =   "Show Information"
      End
      Begin VB.Menu menuFriendChangeNote 
         Caption         =   "Change Note"
      End
      Begin VB.Menu menuFriendl2 
         Caption         =   "-"
      End
      Begin VB.Menu menuFriendDelete 
         Caption         =   "Delete Friend"
      End
   End
   Begin VB.Menu menuSend 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menuSendEnter 
         Caption         =   "ENTER to send"
      End
      Begin VB.Menu menuSendShiftEnter 
         Caption         =   "Shift+ENTER to send"
      End
      Begin VB.Menu menuSendCtrlEnter 
         Caption         =   "Ctrl+ENTER to send"
      End
      Begin VB.Menu menuSendl1 
         Caption         =   "-"
      End
      Begin VB.Menu menuSendNoKey 
         Caption         =   "Disable shortcut key"
      End
   End
   Begin VB.Menu menuLogPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menuLogPopupCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu menuLogPopupCopyText 
         Caption         =   "Copy Text Only"
      End
      Begin VB.Menu menuLogPopupSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu menuLogPopupl1 
         Caption         =   "-"
      End
      Begin VB.Menu menuLogPopupClear 
         Caption         =   "Clear All"
      End
   End
End
Attribute VB_Name = "FormPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum taskType
 taskNone = 0
 taskUpdateState = 1
 taskChangeSession = 2
 taskLogOut = 3
 taskLogIn = 4
 taskUpdateFriList = 5
 taskShowNotification = 6
 taskUpdateAll = 7
 taskGetMsgList = 8
 taskReloadMsg = 9
 taskRebuildConnection = 10
End Enum
Private Const MaxTask As Integer = 16
Private Const restNoteTime As Integer = 60  'In second
Private Const restMsgTime As Integer = 22 'In second

Private taskList(0 To MaxTask) As taskType, taskListPoint As Integer

Public windowTop As Long, windowLeft As Long
Private nowX As Long, nowY As Long, preTab As Integer

Private changePWstage As Integer, oldPW As String, newPW As String
Private humanExit As Boolean, inited As Boolean, clearNoteList As Boolean, isFlash As Boolean

Public WithEvents textArea As HTMLTextAreaElement
Attribute textArea.VB_VarHelpID = -1
Private WithEvents logDoc As HTMLDocument, WithEvents inputDoc As HTMLDocument
Attribute logDoc.VB_VarHelpID = -1
Attribute inputDoc.VB_VarHelpID = -1
Private nowFont As fontStyle

Public Property Get hasInited() As Boolean
 hasInited = inited
End Property
Private Sub init()
 If inited Then Exit Sub
 textLog.Navigate "about:blank": DoEvents
 textInput.Navigate "about:blank": DoEvents
 TabStrip1.Tabs.clear
 loadUserSettings
 loadSessionFile
 If Session_Count < 1 Then
  loadSession nowID
 Else
  loadTab
  loadSession Session_Now
 End If
 applyUserSettings
 inited = True
 clearNoteList = False
 humanExit = False
 isFlash = False
End Sub
Private Sub buttonCloseInfo_Click()
 frameInfo.visible = False
End Sub

Private Sub buttonRefreshInfo_Click()
 applyUserSettings
End Sub

Private Sub buttonRefreshNote_Click()
 showNotification False
End Sub

Private Sub Check1_Click()
 If Check1.value Then nowFont.basic = nowFont.basic Or &H1 Else nowFont.basic = nowFont.basic And &HFFF0
 changeFont
End Sub

Private Sub Check2_Click()
 If Check2.value Then nowFont.basic = nowFont.basic Or &H10 Else nowFont.basic = nowFont.basic And &HFFF0F
 changeFont
End Sub

Private Sub Check3_Click()
 lastUserCaptureHide = Check3.value
End Sub

Private Sub Check4_Click()
 lastUserVisualNotification = Check4.value
End Sub

Private Sub Command1_Click()
 If Trim(textArea.innerText) = "" Then showMsg "Please input the content before sending.", , vbInformation: Exit Sub
 If nowState = State.none Or nowState = State.Offline Then showMsg "Please log in before sending.", , vbInformation: Exit Sub
 If sendMessage(textLog, textInput, nowFont) Then
   clearInput textInput
   changeFont
 Else
   showMsg "Network error. Please try to send it later.", , vbExclamation
 End If
End Sub

Private Sub Command2_Click()
 clearInput textInput
End Sub

Private Sub Command3_Click()
 Dim temp As String
 temp = ModuleAddIn1.BrowseFolders(Me, "Select Directory")
 If temp = "" Then Exit Sub
 textRecordPath.text = temp
End Sub

Private Sub Command4_Click()
 PopupMenu menuSend
End Sub

Private Sub Command5_Click()
 labelPWPrompt.Caption = "Original Password:"
 movePWFrame True
 changePWstage = 1
End Sub
Private Sub buttonPWCancel_Click()
 changePWstage = 0
 movePWFrame False
End Sub
Private Sub buttonPWOK_Click()
 Select Case changePWstage
  Case 1
   oldPW = textPW.text
   labelPWPrompt.Caption = "New Password:"
   textPW.text = vbNullString
   changePWstage = 2
  Case 2
   newPW = textPW.text
   textPW.text = vbNullString
   labelPWPrompt.Caption = "New Password Again:"
   changePWstage = 3
  Case 3
   If newPW <> textPW.text Then
    showMsg "Passwords inconsistent.", "Input error", vbExclamation
   Else
    changePW oldPW, newPW
   End If
   movePWFrame False
   changePWstage = 0
  Case Else
   movePWFrame False
   changePWstage = 0
 End Select
End Sub
Private Sub movePWFrame(show As Boolean)
 setStringSafe oldPW
 setStringSafe newPW
 setStringSafe textPW.text
 framePassword.visible = show
 If show Then framePassword.Top = 500 Else framePassword.Top = 1680
End Sub

Private Sub Form_Activate()
 init
 updateState
End Sub

Private Sub Form_Load()
#If Not IS_DEBUG Then
 monitor True
#End If
 taskListPoint = 0
 inited = False
 humanExit = False
 clearNoteList = True
 'Only operate on value here!
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If humanExit Then Exit Sub
 If showMsg("Sure to exit WiChat?", , vbYesNo + vbInformation) <> vbYes Then Cancel = True
End Sub

Private Sub Form_Resize()
 If Me.WindowState = 2 Then
  Me.WindowState = 0
 ElseIf Me.WindowState = 0 Then
  windowTop = Me.Top
  windowLeft = Me.Left
  updateCaption
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
#If IS_DEBUG Then
 monitor False
#End If
 lastUserFont = nowFont
 saveAll
 changeState Offline
 Unload FormProfile
End Sub

Private Sub Form_Terminate()
 Destroy
End Sub

Private Sub applyUserSettings()
 textRecordPath.text = recordPath(False)
 If textRecordPath.text = vbNullString Then Option1.value = True Else Option2.value = True
 nowFont = lastUserFont
 menuSendEnter.Checked = False
 menuSendShiftEnter.Checked = False
 menuSendCtrlEnter.Checked = False
 menuSendNoKey.Checked = False
 Select Case lastUserSendOperation
  Case 1
   menuSendEnter.Checked = True
  Case 2
   menuSendShiftEnter.Checked = True
  Case 3
   menuSendCtrlEnter.Checked = True
  Case 4
   menuSendNoKey.Checked = True
 End Select
 nowFont = lastUserFont
 listFontFamily.text = nowFont.family
 listFontSize.text = nowFont.size
 listFontColor.text = colorToHex(nowFont.color, True)
 Check1.value = Abs((nowFont.basic And &HFF) > 0)
 Check2.value = Abs((nowFont.basic And &HFF00) > 0)
 Check3.value = lastUserCaptureHide
 Check4.value = lastUserVisualNotification
 Select Case nowFont.align
  Case 2
   Option4.value = True
  Case 3
   Option5.value = True
  Case Else
   Option3.value = True
 End Select
 DoEvents
 changeFont
End Sub

Private Sub frameSession_DragDrop(source As Control, X As Single, Y As Single)
 hideToolBar
End Sub

Private Sub inputDoc_onkeydown()
 Select Case inputDoc.parentWindow.event.keyCode
  Case 27, 112 To 123
   inputDoc.parentWindow.event.keyCode = 0
   inputDoc.parentWindow.event.cancelBubble = True
 End Select
End Sub

Private Sub listEmotion_Click()
 If listEmotion.SelectedItem Is Nothing Then Exit Sub
 addToText "[/em" & Format(listEmotion.SelectedItem.index, "00") & "]"
End Sub

Private Sub listFontColor_Click()
 nowFont.color = colorToHex(listFontColor.text)
 changeFont
End Sub

Private Sub listFontFamily_Click()
 nowFont.family = listFontFamily.text
 changeFont
End Sub
Private Sub listFontSize_Click()
  nowFont.size = listFontSize.text
 changeFont
End Sub

Private Sub listFriend_DblClick()
 If listFriend.HitTest(nowX, nowY) Is Nothing Or listFriend.SelectedItem Is Nothing Then Exit Sub
 frameHome.visible = False
 frameNote.visible = False
 frameSession.visible = True
 If listFriend.SelectedItem.text = TabStrip1.SelectedItem.Caption Then Exit Sub
 Dim id As String
 id = listFriend.SelectedItem.SubItems(1)
 If Not Session_Exist(id) Then
  If Session_Count >= MaxSession Then showMsg "Too many session. Please close one before create new.", , vbExclamation: Exit Sub
  createSession id
 End If
 showSession id
 If Note_count > 0 Then Toolbar1.Buttons(2).Image = 9 Else Toolbar1.Buttons(2).Image = 4
End Sub

Private Sub listFriend_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
 nowX = X: nowY = Y
End Sub

Private Sub listFriend_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
 If Button <> vbRightButton Or listFriend.SelectedItem Is Nothing Then Exit Sub
 If listFriend.HitTest(X, Y) Is Nothing Then Exit Sub
 PopupMenu menuFriend
End Sub

Private Sub listNote_DblClick()
 If listNote.HitTest(nowX, nowY) Is Nothing Then Exit Sub
 Dim tempNote As Notification
 tempNote = Note_peek(Val(listNote.HitTest(nowX, nowY).SubItems(1)))
 With tempNote
  Select Case .type
   Case NoteEvent.FriendAdd
    If showMsg("Sure to add " & .source & " as friend?", "Add as friend", vbYesNo + vbInformation) <> vbYes Then
     delFriend .source
     listNote.HitTest(nowX, nowY).text = "Request from " & .source & " refused."
     listNote.HitTest(nowX, nowY).SmallIcon = 5
     listNote.HitTest(nowX, nowY).SubItems(1) = vbNullString
    Else
     If addFriend(.source) Then
      listNote.HitTest(nowX, nowY).text = "Adding " & .source & " as friend successfully."
      listNote.HitTest(nowX, nowY).SmallIcon = 6
      listNote.HitTest(nowX, nowY).SubItems(1) = vbNullString
      addTask taskUpdateFriList
     Else
      showMsg "Cannot confirm the request from " & .source & ". Please try again."
      listNote.HitTest(nowX, nowY).SmallIcon = 7
      listNote.HitTest(nowX, nowY).SubItems(1) = vbNullString
     End If
    End If
   Case NoteEvent.FriendDelete
   'Anything to show?
   Case NoteEvent.GotMsg
    showSession .source
  End Select
 End With
 If Note_count > 0 Then Toolbar1.Buttons(2).Image = 9 Else Toolbar1.Buttons(2).Image = 4
End Sub


Private Sub listNote_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
 nowX = X: nowY = Y
End Sub


Private Function logDoc_oncontextmenu() As Boolean
 logDoc_oncontextmenu = False
End Function


Private Function logDoc_ondblclick() As Boolean
 If LCase(logDoc.parentWindow.event.srcElement.tagName) = "img" Then
  Dim img As IHTMLImgElement
  Set img = logDoc.parentWindow.event.srcElement
  If Left(img.src, 8) <> "file:///" Then Exit Function
  Shell "cmd /c " & addQuo(URLDecode(Right(img.src, Len(img.src) - 8))), vbHide
 End If
End Function

Private Sub logDoc_onkeydown()
 Select Case logDoc.parentWindow.event.keyCode
  Case 65
   If logDoc.parentWindow.event.ctrlKey Then selectAll
  Case 67
   If logDoc.parentWindow.event.ctrlKey Then copySelected
  Case Else
   logDoc.parentWindow.event.keyCode = 0
   logDoc.parentWindow.event.cancelBubble = True
 End Select
End Sub

Private Sub logDoc_onmouseup()
 If logDoc.parentWindow.event.Button <> 2 Then Exit Sub
 PopupMenu menuLogPopup
End Sub

Private Sub menuFriendChangeNote_Click()
 If listFriend.SelectedItem Is Nothing Then Exit Sub
 Dim temp As String, temp2 As String
 With listFriend.SelectedItem.ListSubItems
  If .Item(1) = .Item(2) Then temp = vbNullString Else temp = Left(.Item(2), InStr(1, .Item(2), .Item(1)) - 2)
  temp2 = InputBox("Please input your note:", "Set note", temp)
  If temp2 <> temp Then changeNote .Item(1), temp2
 End With
End Sub

Private Sub menuFriendDelete_Click()
 If listFriend.SelectedItem Is Nothing Then Exit Sub
 Dim temp As String
 temp = listFriend.SelectedItem.SubItems(1)
 If showMsg("Sure to delete friend " & temp & " ?", "Delete friend", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
 If Not delFriend(temp) Then showMsg "Cannot delete friend " & temp & " . Please try again.", , vbCritical
 If Session_Now = temp Then closeSession textInput
 addTask taskUpdateFriList
End Sub

Private Sub menuFriendNewDialog_Click()
 If listFriend.SelectedItem Is Nothing Then Exit Sub
 showSession listFriend.SelectedItem.SubItems(1)
End Sub

Private Sub menuFriendShowInfo_Click()
 If listFriend.SelectedItem Is Nothing Then Exit Sub
 Dim tempInfo As IDInfo
 tempInfo = getFriendInfo(listFriend.SelectedItem.SubItems(1))
 If tempInfo.id = vbNullString Then showMsg "Cannot get friend info.", , vbCritical: Exit Sub
 FormProfile.show
 setWindowTopMost FormProfile.hwnd
 FormProfile.readOnly = True
 FormProfile.Caption = "Profile of " & tempInfo.id
 FormProfile.labelID.Caption = tempInfo.id
 FormProfile.textMsg.text = tempInfo.offlineMsg
 If listFriend.SelectedItem.SubItems(2) <> listFriend.SelectedItem.SubItems(1) Then
  FormProfile.textNote.text = Left(listFriend.SelectedItem.SubItems(2), InStr(1, listFriend.SelectedItem.SubItems(2), listFriend.SelectedItem.SubItems(1)) - 2)
 End If
End Sub

Private Sub menuHomeExit_Click()
 Unload Me
End Sub

Private Sub menuHomeProfile_Click()
 FormProfile.show
 FormProfile.readOnly = False
 FormProfile.Caption = "Profile of " & nowID
 FormProfile.labelID.Caption = nowID
 FormProfile.textMsg.text = nowOfflineMsg
End Sub

Private Sub menuHomeStateOnline_Click()
 changeState State.onLine
 updateState
End Sub

Private Sub menuHomeStateBusy_Click()
 changeState State.Busy
 updateState
End Sub

Private Sub menuHomeStateHide_Click()
 changeState State.Hide
 updateState
End Sub

Private Sub menuHomeStateOffline_Click()
 changeState State.Offline
 updateState
End Sub

Private Sub menuLogPopupClear_Click()
 If (isNowSessionEmpty And 1) > 0 Then
  If showMsg("Clear all log? The content will be REMOVED PERMANTLY!", , vbYesNo + vbCritical) <> vbYes Then Exit Sub
 End If
 clearSession textLog
End Sub

Private Sub menuLogPopupCopy_Click()
 If Not logSelected Then Exit Sub
 copySelected False
End Sub

Private Sub menuLogPopupCopyText_Click()
 If Not logSelected Then Exit Sub
 copySelected True
End Sub

Private Sub menuLogPopupSelectAll_Click()
 selectAll
End Sub

Private Sub menuSendCtrlEnter_Click()
 lastUserSendOperation = 3
 applyUserSettings
End Sub

Private Sub menuSendEnter_Click()
 lastUserSendOperation = 1
 applyUserSettings
End Sub

Private Sub menuSendNoKey_Click()
 lastUserSendOperation = 4
 applyUserSettings
End Sub

Private Sub menuSendShiftEnter_Click()
 lastUserSendOperation = 2
 applyUserSettings
End Sub

Private Sub menuTrayExit_Click()
 humanExit = True
 doShow
 menuHomeExit_Click
End Sub

Private Sub menuTrayMain_Click()
 doShow
End Sub

Private Sub menuTrayStateBusy_Click()
 menuHomeStateBusy_Click
End Sub

Private Sub menuTrayStateHide_Click()
 menuHomeStateHide_Click
End Sub

Private Sub menuTrayStateOffline_Click()
 menuHomeStateOffline_Click
End Sub

Private Sub Option1_Click()
 textRecordPath.Enabled = False
 Command3.Enabled = False
 textRecordPath.text = vbNullString
End Sub

Private Sub Option2_Click()
 textRecordPath.Enabled = True
 Command3.Enabled = True
End Sub

Private Sub Option3_Click()
 nowFont.align = 1
 changeFont
End Sub

Private Sub Option4_Click()
 nowFont.align = 2
 changeFont
End Sub

Private Sub Option5_Click()
 nowFont.align = 3
 changeFont
End Sub

Private Sub TabStrip1_Click()
 If preTab <> TabStrip1.SelectedItem.index Then
  loadSession TabStrip1.SelectedItem.Caption, False
 Else
  If TabStrip1.SelectedItem.Left - 20 < nowX And nowX < TabStrip1.SelectedItem.Left + 200 And TabStrip1.SelectedItem.Top < nowY And nowY < TabStrip1.SelectedItem.Top + 300 Then
   If TabStrip1.SelectedItem.Caption = nowID Then Exit Sub
   If (isNowSessionEmpty And 1) > 0 Then
    If showMsg("Are you sure to close session with " & TabStrip1.SelectedItem.Caption & " ?" & vbCrLf & "For security reason, all text with be DELETED PERMANETLY!", , vbExclamation + vbYesNo) <> vbYes Then Exit Sub
   ElseIf isNowSessionEmpty = 2 Or Me.textArea.innerHTML <> "" Then
    If showMsg("You have unsent message. Still close this session?", , vbYesNo + vbExclamation) <> vbYes Then Exit Sub
   End If
   closeSession textInput
   TabStrip1.Tabs.Remove TabStrip1.SelectedItem.index
   loadSession TabStrip1.Tabs(1).Caption, False
  End If
 End If
 preTab = TabStrip1.SelectedItem.index
End Sub

Private Sub TabStrip1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
 nowX = X: nowY = Y
 If TabStrip1.SelectedItem.Caption = nowID Then Exit Sub
 If TabStrip1.SelectedItem.Left - 20 < X And X < TabStrip1.SelectedItem.Left + 200 And TabStrip1.SelectedItem.Top < Y And Y < TabStrip1.SelectedItem.Top + 300 Then
  TabStrip1.SelectedItem.Image = 3
 Else
  TabStrip1.SelectedItem.Image = getStateImage(TabStrip1.SelectedItem.Caption)
 End If
End Sub

Private Function textArea_onkeypress() As Boolean
 textArea_onkeypress = True
 Static doc As HTMLDocument, window As HTMLWindow2
 Set doc = textInput.Document
 If doc Is Nothing Then Exit Function
 Set window = doc.parentWindow
 If Not window.event Is Nothing Then
 'Need to forbid F5、F11、F12,etc.
  Select Case window.event.keyCode
   Case 13
    If lastUserSendOperation <> 4 And (window.event.ctrlKey And lastUserSendOperation = 3 Or window.event.shiftKey And lastUserSendOperation = 2 Or Not window.event.ctrlKey And Not window.event.shiftKey And lastUserSendOperation = 1) Then
     Command1_Click
     textArea_onkeypress = False
     textArea.focus
    End If
   Case 27
    textArea_onkeypress = False
  End Select
 End If
End Function

Private Sub textInput_GotFocus()
 hideToolBar
End Sub

Private Sub textLog_GotFocus()
 hideToolBar
End Sub

Private Sub textRecordPath_Change()
 recordPath = textRecordPath.text
End Sub

Private Sub Timer1_Timer()
 Static Count As Long
 Static flashCount As Integer
 
 Count = Count + 1
 If Count >= MaxSessionTime Then
  Count = 0
  addTask taskChangeSession
 End If
 doTask
 If Count Mod restNoteTime = 2 Then addTask taskShowNotification
 If Count Mod restMsgTime = 5 Then addTask taskGetMsgList
 
 If isFlash Then
  FlashWindow Me.hwnd, True
  flashCount = flashCount + 1
  If flashCount > 5 Then
   isFlash = False
   flashCount = 0
   FlashWindow Me.hwnd, False
  End If
 End If
End Sub
Public Function addTask(task As taskType) As Boolean
 addTask = False
 If taskListPoint < MaxTask Then
  taskListPoint = taskListPoint + 1
  taskList(taskListPoint) = task
  addTask = True
 End If
End Function
Private Sub doTask()
 If taskListPoint < 1 Then Exit Sub
 taskListPoint = taskListPoint - 1
 Select Case taskList(taskListPoint + 1)
  Case taskType.taskChangeSession
   changeSession
  Case taskType.taskUpdateState
   updateState
   refreshTab
  Case taskType.taskShowNotification
   updateFriendList listFriend
   refreshTab
   showNotification
  Case taskType.taskUpdateFriList
   updateFriendList listFriend
   refreshTab
  Case taskType.taskUpdateAll
   addTask taskGetMsgList
   addTask taskUpdateFriList
  Case taskType.taskGetMsgList
   Image1.visible = True
   DoEvents
   getMessageList
   Image1.visible = False
   showNotification
  Case taskType.taskReloadMsg
   loadSessionContent Session_Now, textLog, textInput
  Case taskType.taskRebuildConnection
   fixBrokenConnection
 End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.index
  Case 1
   frameHome.visible = True
   frameNote.visible = False
   frameSession.visible = False
  Case 2
   frameHome.visible = False
   frameNote.visible = True
   frameSession.visible = False
  Case 3
   frameHome.visible = False
   frameNote.visible = False
   frameSession.visible = True
  Case 4
   Me.WindowState = 1
   hideToTray
 End Select
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
 hideToolBar
 Select Case Button.index
  Case 2
   frameFont.visible = True
  Case 4
   initEmotionList
   listEmotion.visible = True
  Case 6
   ModuleAddIn1.Title = "添加图片"
   ModuleAddIn1.Filter = "图片文件(*.jpg,*.gif,*.png,*.tiff,*.bmp)|*.jpg;*.gif;*.png;*.tiff;*.bmp"
   ModuleAddIn1.ShowOpen Me
   If ModuleAddIn1.fileName = vbNullString Then Exit Sub
   If FileLen(ModuleAddIn1.fileName) > MaxCapableFileSize And Session_Now <> nowID Then
    If showMsg("Picture too large. Network may fail during sending. Continue?", , vbExclamation + vbYesNo) <> vbYes Then Exit Sub
   End If
   addToText "[/i=" & ModuleAddIn1.fileName & "/]"
  Case 8
   ModuleAddIn1.Title = "发送文件"
   ModuleAddIn1.Filter = "所有文件(*.*)|*.*"
   ModuleAddIn1.ShowOpen Me
   If ModuleAddIn1.fileName = vbNullString Then Exit Sub
   If FileLen(ModuleAddIn1.fileName) > MaxCapableFileSize And Session_Now <> nowID Then
    If showMsg("File too large. Network may fail during sending. Continue?", , vbExclamation + vbYesNo) <> vbYes Then Exit Sub
   End If
   addToText "[/f=" & ModuleAddIn1.fileName & "/]"
  Case 10
   If lastUserCaptureHide > 0 Then Me.Hide
   FormScreenShot.show
  Case Else
 End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.index
  Case 1
   Dim tempID As String
   tempID = Trim(InputBox("Please input the ID you want to add as friend:", "Add Friend"))
   If Not checkID(tempID) Then showMsg "ID invalid.", , vbCritical: Exit Sub
   If addFriend(tempID) Then showMsg "Your request has been send." Else showMsg "Request failed. Maybe he/she does not want to be added as friend.", , vbExclamation
  Case 3
   addTask taskUpdateFriList
 End Select
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.index
  Case 1
   PopupMenu Me.menuHomeState, , 1400, 1500
  Case 3
   menuHomeProfile_Click
  Case 5
   frameInfo.visible = True
   labelInfoCaption.Caption = "Settings"
   buttonRefreshInfo.visible = True
   frameSetting.visible = True
   frameHelp.visible = False
  Case 7
   frameInfo.visible = True
   labelInfoCaption.Caption = "Help"
   buttonRefreshInfo.visible = False
   frameSetting.visible = False
   frameHelp.visible = True
  Case 9
   menuHomeExit_Click
 End Select
End Sub

Private Sub updateState()
 menuHomeStateBusy.Checked = False
 menuTrayStateBusy.Checked = False
 menuHomeStateHide.Checked = False
 menuTrayStateHide.Checked = False
 menuHomeStateOffline.Checked = False
 menuTrayStateOffline.Checked = False
 menuHomeStateOnline.Checked = False
 menuTrayStateOnline.Checked = False
 Select Case nowState
  Case State.Busy
   menuHomeStateBusy.Checked = True
   menuTrayStateBusy.Checked = True
  Case State.Hide
   menuHomeStateHide.Checked = True
   menuTrayStateHide.Checked = True
  Case State.Offline
   menuHomeStateOffline.Checked = True
   menuTrayStateOffline.Checked = True
  Case State.onLine
   menuHomeStateOnline.Checked = True
   menuTrayStateOnline.Checked = True
 End Select
 updateCaption
End Sub
Private Sub updateCaption()
 If Session_Now = nowID Then
  Me.Caption = nowID & "[" & stateToString(nowState) & "]"
  If nowOfflineMsg <> vbNullString Then Me.Caption = Me.Caption & " - " & nowOfflineMsg
 End If
End Sub
Private Sub showNotification(Optional clearList As Boolean = True)
 Dim tempNote() As Notification
 Static i As Integer, j As Integer
 tempNote = Note_get
 
 If clearNoteList Then listNote.ListItems.clear
 For i = 1 To UBound(tempNote)
  If Not clearNoteList Then
   For j = 1 To listNote.ListItems.Count
    If Val(listNote.ListItems(j).SubItems(1)) = tempNote(i).handle Then GoTo con
   Next j
  End If
   Select Case tempNote(i).type
    Case NoteEvent.FriendAdd
     listNote.ListItems.Add , , tempNote(i).source & " want to add you as friend", , 4
     listNote.ListItems.Item(listNote.ListItems.Count).SubItems(1) = str(tempNote(i).handle)
    Case NoteEvent.FriendDelete
     listNote.ListItems.Add , , tempNote(i).source & " remove you from friend list", , 3
     listNote.ListItems.Item(listNote.ListItems.Count).SubItems(1) = str(tempNote(i).handle)
     delFriend tempNote(i).source
    Case NoteEvent.GotMsg
     If Session_Now = tempNote(i).source Then
      addTask taskReloadMsg
     Else
      listNote.ListItems.Add , , "Message from " & tempNote(i).source, , 12
      listNote.ListItems.Item(listNote.ListItems.Count).SubItems(1) = str(tempNote(i).handle)
     End If
     If Me.WindowState = 1 And lastUserVisualNotification > 0 Then  'Show visualized notification
      If Me.visible = True Then
       Me.Caption = "Msg from " & tempNote(i).source & "..."
       isFlash = True
      Else
       TrayBalloon listNote.ListItems(1) & " ...", "WiChat"
      End If
     End If
    Case NoteEvent.none
   End Select
con:
 Next i
 
 If listNote.ListItems.Count > 0 Then Toolbar1.Buttons(2).Image = 9 Else Toolbar1.Buttons(2).Image = 4
 clearNoteList = True
End Sub
Private Function getStateImage(id As String) As Integer
 getStateImage = 0
 If id = nowID Then
  getStateImage = getStateImageIndex(Val(nowState), True)
 Else
  Dim i As Long
  For i = 1 To listFriend.ListItems.Count
   If listFriend.ListItems(i).SubItems(1) = id Then getStateImage = listFriend.ListItems(i).SmallIcon: Exit For
  Next i
  If i > listFriend.ListItems.Count Then getStateImage = getStateImageIndex(State.Offline)
 End If
End Function
Private Function loadSession(id As String, Optional setTabActive As Boolean = True) As Boolean
   loadSession = False
   If id = vbNullString Then Exit Function
   Dim i As Integer
   For i = 1 To TabStrip1.Tabs.Count
    If TabStrip1.Tabs(i).Caption = id Then Exit For
   Next i
   If i > TabStrip1.Tabs.Count Then
    If Not Session_Exist(id) Then
     If Not createSession(id) Then Exit Function
    End If
    TabStrip1.Tabs.Add , , id, getStateImage(id)
   End If
   If setTabActive Then
    For i = 1 To TabStrip1.Tabs.Count
     If TabStrip1.Tabs(i).Caption = id Then TabStrip1.Tabs(i).Selected = True: Exit For
    Next i
   Else
    loadSessionContent id, textLog, textInput
    showNotification
    Me.Caption = "Session with " & id
   End If
   updateCaption
End Function
Private Sub loadTab()
 Dim tempList() As String, i As Integer, j As Integer
 TabStrip1.Tabs.clear
 Session_List tempList
 TabStrip1.Tabs.Add , , nowID, getStateImage(nowID)
 If stringArrayIsDimed(tempList) Then
  For i = 0 To UBound(tempList)
   If tempList(i) <> nowID Then TabStrip1.Tabs.Add , , tempList(i), getStateImage(tempList(i))
  Next i
 End If
End Sub
Private Sub refreshTab()
 Dim i As Integer, j As Integer
 For i = 1 To TabStrip1.Tabs.Count
  TabStrip1.Tabs(i).Image = getStateImage(TabStrip1.Tabs(i).Caption)
 Next i
End Sub
Private Sub showSession(id As String)
 frameHome.visible = False
 frameNote.visible = False
 frameSession.visible = True
 loadSession id, True
End Sub
Private Sub initEmotionList()
 On Error Resume Next
 Dim i As Integer
 If ImageList3.ListImages.Count < 1 Then
  For i = 1 To MaxEmotion
   ImageList3.ListImages.Add i, , LoadPicture(App.path & LOCAL_EMTION_DEFAULT_PATH & "/" & i & ".gif")
  Next i
 End If
 If listEmotion.ListItems.Count < 1 Then
  For i = 1 To ImageList3.ListImages.Count
   listEmotion.ListItems.Add i, , " ", , ImageList3.ListImages(i).index
  Next i
 End If
 Set listEmotion.SelectedItem = Nothing
End Sub
Private Sub hideToolBar()
 listEmotion.visible = False
 frameFont.visible = False
End Sub
Public Sub addToText(str As String)
 If Not bindTextBox Then Exit Sub
 textArea.focus
 Dim doc As IHTMLDocument2
 Dim range1 As IHTMLTxtRange, range2 As IHTMLTxtRange
 Dim pStart As Long ', pEnd As Long
 Dim j As Long
 Set doc = textInput.Document
 Set range1 = doc.selection.createRange
 If range1.parentElement.id <> textArea.id Then Exit Sub
 Set range2 = doc.body.createTextRange
 range2.moveToElementText textArea
 Do While range2.compareEndPoints("StartToStart", range1) < 0
  pStart = pStart + 1
  range2.moveStart "character", 1
 Loop
 If pStart > Len(textArea.innerText) Then pStart = Len(textArea.innerText)
' For j = 1 To pStart
'  If Mid(textArea.value, i, 1) = Chr(13) Then pEnd = pEnd + 1
' Next j
' range2.moveToElementText textArea
' Do While range2.compareEndPoints("StartToStart", range1) < 0
'  pEnd = pEnd + 1
'  range2.moveStart "character", 1
' Loop
' For j = 1 To pEnd
'  If Mid(textArea.value, i, 1) = Chr(13) Then pEnd = pEnd + 1
' Next j
 textArea.innerText = Left(textArea.innerHTML, pStart) & str & Right(textArea.innerHTML, Len(textArea.innerText) - pStart)
 textArea.focus
End Sub
Public Function bindTextBox() As Boolean
 bindTextBox = False
 Dim doc As IHTMLDocument2
 Set doc = textInput.Document
 If doc Is Nothing Then Exit Function
 Set textArea = doc.getElementById(textboxName)
 If textArea Is Nothing Then Exit Function
 doc.body.Scroll = "no"
 Set logDoc = textLog.Document
 Set inputDoc = textInput.Document
 bindTextBox = True
End Function
Public Sub changeFont()
 If Not bindTextBox Then Exit Sub
 Dim obj As IHTMLStyle
 Set obj = textArea.Style
 If nowFont.color <> vbNullString Then obj.color = "#" & nowFont.color
 obj.fontFamily = nowFont.family
 If nowFont.size > 0 Then obj.FontSize = nowFont.size + 2 & "px"
 If (nowFont.basic And &H1) Then obj.fontWeight = "bold" Else obj.fontWeight = vbNullString
 If (nowFont.basic And &H10) Then obj.fontStyle = "italic" Else obj.fontStyle = vbNullString
End Sub
Public Sub showProcess(visible As Boolean, Optional text As String = vbNullString, Optional value As Integer = 0)
 If text = vbNullString Then text = "Processing..."
 If value < 0 Or value > 100 Then value = 100
 If visible Then
  ProgressBar1.value = value
  labelProcessState.Caption = text
  frameProcess.visible = True
  frameSession.Enabled = False
 Else
  ProgressBar1.value = 0
  labelProcessState.Caption = vbNullString
  frameProcess.visible = False
  frameSession.Enabled = True
 End If
 DoEvents
End Sub
Private Function logSelected() As Boolean
 logSelected = False
 If textLog.Document Is Nothing Then Exit Function
 Dim doc As HTMLDocument, range As IHTMLTxtRange
 Set doc = textLog.Document
 Set range = doc.selection.createRange
 If range.htmlText = "" Then Exit Function
 logSelected = True
End Function
Private Sub selectAll()
 Dim doc As HTMLDocument, range As IHTMLTxtRange
 Set doc = textLog.Document
 Set range = doc.selection.createRange
 range.moveStart "character", 0
 range.moveEnd "character", Len(doc.body.innerText)
 range.Select
End Sub
Private Sub copySelected(Optional textOnly As Boolean = False)
 Dim doc As HTMLDocument, range As IHTMLTxtRange
 Set doc = textLog.Document
 Set range = doc.selection.createRange
 If textOnly Then
  Clipboard.clear
  Clipboard.SetText range.text
 Else
  Clipboard.clear
  Clipboard.SetText range.htmlText
 End If
End Sub
