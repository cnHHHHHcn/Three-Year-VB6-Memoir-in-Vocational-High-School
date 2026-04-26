VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Setting"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4275
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4530
   ScaleWidth      =   4275
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "显示热键设置"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3975
      Begin VB.ComboBox ControlKey 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "Form3.frx":0CFA
         Left            =   1440
         List            =   "Form3.frx":0D10
         TabIndex        =   7
         Text            =   "ControlKey"
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox WordKey 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "Form3.frx":0D56
         Left            =   1440
         List            =   "Form3.frx":0DA8
         TabIndex        =   6
         Text            =   "V"
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "字母键："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "控制键："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   405
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "剪贴板存储设置"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.TextBox ListMax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   2
         Text            =   "10"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox OutputIntervalText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2160
         TabIndex        =   1
         Text            =   "1s"
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "内容最大数量："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   435
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "内容粘贴间隔："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2100
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    ClipBoardDispose.SaveDispose
    ClipBoardDispose.ProgramClosed
End Sub

Private Sub Form_Load()
    IsSave.Value = Abs(CInt(Config.Save))
    ListMax.Text = Config.ListMax
    OutputIntervalText = Config.OutputInterval & "s"
    Form3.ControlKey.Text = Form3.ControlKey.List(3)
    ControlKey.Text = Config.HotKey.Control
    WordKey.Text = Config.HotKey.Word
End Sub

Private Sub IsSave_Click()
    Config.Save = CBool(IsSave.Value)
End Sub

Private Sub ListMax_Change()
    Config.ListMax = Val(ListMax.Text)
End Sub

Private Sub OutputIntervalText_Change()
    Config.OutputInterval = Val(OutputIntervalText)
    Form1.PrintClipBoardText.Interval = Config.OutputInterval * 1000
End Sub

Private Sub Look_Click()
    CommonDialog1.ShowOpen
    PathText.Text = CommonDialog1.FileName
End Sub

Private Sub ControlKey_Click()
    Config.HotKey.Control = ControlKey.Text
End Sub

Private Sub WordKey_Click()
    Config.HotKey.Word = WordKey.Text
End Sub
