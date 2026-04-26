VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   12165
   ScaleWidth      =   16980
   StartUpPosition =   3  '窗口缺省
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
      Left            =   5280
      TabIndex        =   3
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "密码"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   495
      Left            =   9120
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "请输入密码："
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
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Width           =   2970
   End
   Begin VB.Label Label1 
      Caption         =   "专属三三"
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
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label3.Caption = Label3.Caption + 1
a = Val(Label3.Caption)
Command2.Visible = True
If a = 1 Then
MsgBox "你笨死了", , ""
Command1.Enabled = False
GoTo g
ElseIf a = 2 Then
MsgBox "哈哈，你好SB" & Chr(10) & "看看密码", , ""
Command1.Enabled = False
Command2.Enabled = True
Text1.Text = ""
End If
If Text1.Text = "333" Then
b:
h = InputBox("you 是否 miss me ?", "")
If h = "是" Then
MsgBox "墓骂墓骂，五一快乐！", , ""
Form1.Hide: Form2.Show
ElseIf h = "否" Then
MsgBox "OK，我懂了没爱了，我走。", , ""
End
Else
GoTo b
End If
Else
MsgBox "你笨死了", , ""
Command1.Enabled = False
Command2.Enabled = True
End If
g:
End Sub

Private Sub Command2_Click()
Label5.Caption = Label5.Caption + 1
a = Val(Label5.Caption)
If a = 1 Then
MsgBox "好了不闹了，密码是sb250", , ""
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
ElseIf a = 2 Then
MsgBox "抱歉SB，密码是250sb", , ""
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
ElseIf a = 3 Then
MsgBox "密码是咱们的信号", , ""
Command1.Enabled = True
Command2.Enabled = False
End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Label4.Caption + 1
a = Val(Label4.Caption)
If a >= 5 Then GoTo c
Command2.Left = Int(Rnd * 10000)
Command2.Top = Int(Rnd * 10000)
If a = 1 Then
MsgBox "按钮：什么玩意摸我！", 32, "按钮"
MsgBox "按钮：嗯~ ~ ~ ！？", 32, "按钮"
MsgBox "按钮：嘿嘿！！！", 64, "按钮"
ElseIf a = 2 Then
MsgBox "按钮：我躲！！！", , "按钮"
ElseIf a = 3 Then
MsgBox "按钮：抓不到气不气。", , "按钮"
ElseIf a = 4 Then
MsgBox "系统：触发技能按钮锁定。", 64, "系统"
MsgBox "按钮：啊啊！！！系统！！！你*$%&", , "按钮"
End If
c:
End Sub


Private Sub Text1_Change()
If Text1.Text <> "" Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub
