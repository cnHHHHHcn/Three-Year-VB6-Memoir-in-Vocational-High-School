VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   21360
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   17400
      Top             =   6360
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "判题"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出题"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   495
      Left            =   13200
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   495
      Left            =   13080
      TabIndex        =   13
      Tag             =   "false"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      TabIndex        =   9
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15120
      TabIndex        =   8
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "错误："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   7
      Top             =   3360
      Width           =   1485
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "正确："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   6
      Top             =   2160
      Width           =   1485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()  '按钮 出题 单击事件
'出0~10的整数
Label2.Caption = Int(Rnd * 11)
Label4.Caption = Int(Rnd * 11)
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command2_Click()   '按钮 判题 单击事件
'判定是否 文本框1 = 标签2和标签4
If Val(Text1.Text) = Val(Label2.Caption) + Val(Label4.Caption) Then
Label6.Caption = "对"
Label9.Caption = Label9.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
Else
Label6.Caption = "错"
Label10.Caption = Label10.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
End If
'如果标签6为 “对”时，跳转到a:后
If Label6.Caption = "对" Then GoTo a
If Val(Label10.Caption) Mod 5 = 0 Then
MsgBox "这都做不对，那在加一道题吧。"
Form1.Text1.Text = Form1.Text1.Text + 1
End If
a:
If Form1.Text1.Text = Label9.Caption Then End
End Sub

Private Sub Form_Load()  '窗体 加载事件
On Error Resume Next
Command2.Enabled = False
Label11.Caption = Form2.Width
Label12.Caption = Form2.Height
End Sub

Private Sub Form_Resize()   '窗体  调节尺寸事件
On Error Resume Next
a = Label11.Caption: b = Label12.Caption
If 0 < Form2.Width And Form2.Width < a Or 0 < Form2.Height And Form2.Height < b Then MsgBox "你想干嘛？", , "系统"
End Sub

Private Sub Timer1_Timer()   '定时器 定时事件
On Error Resume Next
a = Label11.Caption: b = Label12.Caption
Form2.Top = 0: Form2.Width = a
Form2.Left = 0: Form2.Height = b
End Sub
