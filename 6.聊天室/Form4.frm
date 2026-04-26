VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "聊天室 更新日志"
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   9345
   ScaleWidth      =   9645
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8175
      LargeChange     =   5
      Left            =   9240
      Max             =   30
      Min             =   10
      SmallChange     =   5
      TabIndex        =   1
      Top             =   840
      Value           =   10
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   120
      ScaleHeight     =   8145
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   840
      Width           =   8895
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   180
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "更新日志"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form4.Hide
End Sub

Private Sub Timer1_Timer()
'读取公告.txt文件并写入picture控件
Open "更新日志.txt " For Input As #1
Picture1.Cls
Do While Not EOF(1)
    Line Input #1, gh
    a = a & Chr(10) & gh
Loop
Close
Label1.Caption = a
If Picture1.Height < Label1.Height Then VScroll1.Max = Label1.Height - Picture1.Height Else VScroll1.Enabled = False
Timer1.Interval = 0
End Sub

Private Sub Form_Load()
Timer1.Interval = 900
End Sub

Private Sub VScroll1_Change()
Label1.Top = -(VScroll1.Value)
End Sub

Private Sub VScroll1_Scroll()
Label1.Top = -(VScroll1.Value)
End Sub

