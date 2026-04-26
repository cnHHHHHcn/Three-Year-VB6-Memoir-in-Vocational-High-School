VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "科技"
   ClientHeight    =   6855
   ClientLeft      =   13980
   ClientTop       =   6405
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7500
   Begin VB.Frame Frame1 
      Caption         =   "自定义区"
      Height          =   4695
      Left            =   2880
      TabIndex        =   8
      Top             =   360
      Width           =   4335
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1920
         TabIndex        =   13
         Text            =   "0"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Text            =   "0"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Text            =   "0"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1920
         TabIndex        =   10
         Text            =   "0"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "＜执行"
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "自定义当前分数："
         Height          =   180
         Left            =   480
         TabIndex        =   17
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "自定义复活次数："
         Height          =   180
         Left            =   480
         TabIndex        =   16
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "自定义当前时间："
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   2760
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "自定义无敌时间："
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   3600
         Width           =   1440
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "刷豆"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "无敌+穿墙"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "暂停鬼"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "刷复活"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "刷强化"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "刷时间"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "刷无敌"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()     '无敌+穿墙
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
If Check1.Value = 1 Then
Form1.Command2.Left = Form1.Command2.Left + 1
Form1.Command2.Top = Form1.Command2.Top + 1
Else
Form1.Command2.Left = Form1.Command2.Left - 1
Form1.Command2.Top = Form1.Command2.Top - 1
End If
End Sub

Private Sub Check2_Click()     '暂停鬼
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
If Check2.Value = 1 Then
Form1.Timer2.Enabled = False
End If
If Check2.Value = 0 Then
Form1.Timer2.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Form4.Hide
End Sub

Private Sub Command2_Click()     '刷复活
Form1.Image3.Visible = True
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
End Sub

Private Sub Command3_Click()     '刷强化
Form1.Image5.Visible = True
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
End Sub

Private Sub Command4_Click()     '刷时间定格
Form1.Image4.Visible = True
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
End Sub

Private Sub Command5_Click()     '刷无敌
Form1.Image2.Visible = True
End Sub

Private Sub Command6_Click()     '当前时间运行按钮
Form1.Label4 = Text3.Text
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
End Sub

Private Sub Command7_Click()     '刷豆
For i = 1 To 400
    Form1.Image1(i).Visible = True
Next
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
End Sub

Private Sub Text1_Change()     '当前分数
a = Val(Text1.Text)
Form1.Label1 = a
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
If Text1.Text <= 0 Or Text1.Text = "" Then Text1.Text = 0: MsgBox "数值不能为0或负数"
End Sub

Private Sub Text2_Change()     '自定义复活次数
a = Val(Text2.Text)
Form1.Label6 = a
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
If Text2.Text <= 0 Or Text2.Text = "" Then Text1.Text = 0: MsgBox "数值不能为0或负数"
End Sub

Private Sub Text4_Change()     '自定义无敌时间
a = Val(Text4.Text)
Form1.Label9 = a
If Form1.Command3.Enabled = True Then
Form1.Command3.SetFocus
End If
If Text4.Text <= 0 Or Text4.Text = "" Then Text1.Text = 0: MsgBox "数值不能为0或负数"
End Sub

