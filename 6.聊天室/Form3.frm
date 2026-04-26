VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "创建聊天室"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   660
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "是否设置密码"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "聊天室密码："
      Height          =   180
      Left            =   600
      TabIndex        =   5
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "聊天室名称："
      Height          =   180
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()    '是否设置密码
If Check1.Value = 1 Then
    Text2.Enabled = True
    Text2.SetFocus
Else
    Text2.Enabled = False
End If
End Sub

Private Sub Command1_Click()   '返回
Form3.Hide
End Sub

Private Sub Command2_Click()
'必须满足的必要条件
If Text1.Text = "" Then
MsgBox "未设置聊天室名称！", 48: GoTo tiao
ElseIf CBool(Check1.Value) And Text2.Text = "" Then
MsgBox "未设置密码！！！", 48: GoTo tiao
End If
'如果不设置密码，则清空密码框
If Check1.Value = 0 Then Text2.Text = ""
'创建聊天室，初始化聊天室内容
hao = FreeFile
Open Form2.Label4.Caption & Text1.Text & ".txt" For Output As hao
Print #hao, "是否设置密码:" & CBool(Check1.Value) & " 密码:" & Text2.Text
Print #hao, "建立聊天室 用户:" & Form2.Label3.Caption
Close
'创建聊天室人员列表
Open Form2.Label7.Caption & Text1.Text & "名称列表.txt" For Output As hao
Close
'刷新聊天室列表
Form2.Label5.Caption = Text1.Text
Form2.File1.Refresh
Form3.Hide
Check1.Value = 0
Text1.Text = ""
Text2.Text = ""
tiao:
End Sub

Private Sub Text1_Change()
Label2.Caption = "聊天室名不允许重复！！！"
End Sub

