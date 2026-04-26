VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   21360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   13680
      Top             =   7560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "随机"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
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
      Left            =   5280
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   15480
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   15600
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "做对     个题，就关闭"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   3240
      Width           =   5235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()   '按钮 确定  单击事件
'截取数值这为整数
q = Text1.Text
Text1.Text = Round(q, 0)
'题数至少40个
If Val(Text1.Text) <= 39 Then
MsgBox "想偷懒，没门！", , "系统"
GoTo g '跳转
End If
a:
'当判定为1时，显示窗体2，判定为2时显示窗体3，否则跳转回来
b = InputBox("请输入编码" & Chr(10) & "编码1，十以内加法" & Chr(10) & "编码2，一百以内加减乘除", "系统")
If b = 1 Then
Form1.Hide: Form2.Show
ElseIf b = 2 Then
Form1.Hide: Form3.Show
Else
GoTo a
End If
g:
End Sub

Private Sub Command2_Click()   '按钮  随机  单击事件
Randomize
'随机出50至100道题
Text1.Text = Int(Rnd * 51 + 50)
a:
'当判定为1时，显示窗体2，判定为2时显示窗体3，否则跳转回来
b = InputBox("请输入编码" & Chr(10) & "编码1，十以内加法" & Chr(10) & "编码2，一百以内加减乘除", "系统")
If b = 1 Then
Form1.Hide: Form2.Show
ElseIf b = 2 Then
Form1.Hide: Form3.Show
Else
GoTo a
End If
End Sub
Private Sub Form_Load()   '窗体 加载事件
On Error Resume Next
Label2.Caption = Form1.Width
Label3.Caption = Form1.Height
End Sub

Private Sub Form_Resize()   '窗体  调节尺寸事件
On Error Resume Next
'如果窗体大小有变化，则输出对话框
a = Label2.Caption: b = Label3.Caption
If 0 < Form1.Width And Form1.Width < a Or 0 < Form1.Height And Form1.Height < b Then MsgBox "你想干嘛？", , "系统"
End Sub

Private Sub Timer1_Timer()   '定时器  定时事件
On Error Resume Next
'固定窗体不能变大或变小
a = Label2.Caption: b = Label3.Caption
Form1.Top = 0: Form1.Width = a
Form1.Left = 0: Form1.Height = b
End Sub


Private Sub Text1_Change()   '文本框  改变事件
'如果文本框为空，则按钮(确定)可用，否则按钮(确定)可用
If Text1.Text <> "" Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
Form2.Label1.Caption = "做对" & Form1.Text1.Text & "个题，就结束"
Form3.Label1.Caption = "做对" & Form1.Text1.Text & "个题，就结束(保留两位小数)"
End Sub

