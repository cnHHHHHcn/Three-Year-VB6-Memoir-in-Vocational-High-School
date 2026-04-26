VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "设置"
   ClientHeight    =   2085
   ClientLeft      =   9030
   ClientTop       =   3300
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5445
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "应用"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "计算机名称或IPv4地址"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "端口号："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DuanKou = Text1.Text

End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command3_Click()
Name_Ip = Text2.Text
DuanKou = Val(Text1.Text)

End Sub

Private Sub Command4_Click()
Unload Form1
End Sub

Private Sub Form_Load()
Text1.Text = DuanKou
Text2.Text = Name_Ip
End Sub

Private Sub Text1_Change()
If Val(Text1.Text) > 65535 Then MsgBox "不能大于65535！", 48, "错误": Text1.Text = "1001"
If Text1.Text = "" Then MsgBox "不能为空！", 48, "错误": Text1.Text = "1001"
End Sub
