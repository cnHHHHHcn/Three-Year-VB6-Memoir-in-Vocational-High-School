VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "公告"
   ClientHeight    =   9780
   ClientLeft      =   2220
   ClientTop       =   1905
   ClientWidth     =   16980
   LinkTopic       =   "Form2"
   ScaleHeight     =   9780
   ScaleWidth      =   16980
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   15840
      Top             =   8280
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   15240
      Top             =   8280
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   5970
      ItemData        =   "Form2.frx":0000
      Left            =   15240
      List            =   "Form2.frx":0002
      TabIndex        =   19
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "选定联机用户"
      Height          =   1455
      Left            =   15240
      TabIndex        =   16
      Top             =   6720
      Width           =   1455
      Begin VB.CommandButton Command5 
         Caption         =   "发送"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "请输入名字或编码（编码起始位为 1 ）"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "游戏记录"
      Height          =   495
      Left            =   8400
      TabIndex        =   15
      Top             =   8400
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7695
      LargeChange     =   600
      Left            =   240
      SmallChange     =   600
      TabIndex        =   12
      Top             =   360
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10200
      Top             =   6720
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   840
      ScaleHeight     =   7665
      ScaleWidth      =   8145
      TabIndex        =   11
      Top             =   360
      Width           =   8175
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   180
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form2.frx":0004
      Left            =   12840
      List            =   "Form2.frx":0011
      TabIndex        =   10
      Text            =   "简单"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form2.frx":0027
      Left            =   12840
      List            =   "Form2.frx":0034
      TabIndex        =   7
      Text            =   "普通模式"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "进入游戏"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Label10"
      Height          =   180
      Left            =   10920
      TabIndex        =   20
      Top             =   3120
      Width           =   630
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "用户："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11640
      TabIndex        =   14
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "难度："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12000
      TabIndex        =   9
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12000
      TabIndex        =   8
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "强化：让鬼消失(当时间为45秒时出现)"
      Height          =   180
      Left            =   11880
      TabIndex        =   6
      Top             =   7200
      Width           =   3060
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "时间定格：让鬼定住，并增加五秒时间"
      Height          =   180
      Left            =   11880
      TabIndex        =   5
      Top             =   6360
      Width           =   3060
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "生命次数：增加一次复活次数"
      Height          =   180
      Left            =   11880
      TabIndex        =   4
      Top             =   5520
      Width           =   2340
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "无敌：有十五秒时间无视鬼的攻击"
      Height          =   180
      Left            =   11880
      TabIndex        =   3
      Top             =   4680
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   11160
      Picture         =   "Form2.frx":0056
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   405
   End
   Begin VB.Image Image4 
      Height          =   405
      Left            =   11160
      Picture         =   "Form2.frx":6541
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   405
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   11160
      Picture         =   "Form2.frx":E8EB
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   11160
      Picture         =   "Form2.frx":EF3B
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   495
      Left            =   -360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim user

Private Sub Combo1_Change()
'游戏模式与form3的游戏模式互通，combo2控件同理
Select Case Combo1.Text
Case "看脸模式"
Form3.Combo1.Text = "看脸模式"
Case "普通模式"
Form3.Combo1.Text = "普通模式"
Case "无尽模式"
Form3.Combo1.Text = "无尽模式"
Case Else
Combo1.Text = ""
End Select
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "看脸模式" Then
Form1.Label13.Caption = 1
ElseIf Combo1.Text = "普通模式" Then
Form1.Label13.Caption = 0
ElseIf Combo1.Text = "无尽模式" Then
Form1.Label13.Caption = 2
End If
Select Case Combo1.Text
Case "看脸模式"
Form3.Combo1.Text = "看脸模式"
Combo2.Enabled = True
Form3.Combo2.Enabled = True
Case "普通模式"
Form3.Combo1.Text = "普通模式"
Combo2.Enabled = True
Form3.Combo2.Enabled = True
Case "无尽模式"
Form3.Combo1.Text = "无尽模式"
Combo2.Text = ""
Combo2.Enabled = False
Form3.Combo2.Text = ""
Form3.Combo2.Enabled = False
Case Else
Combo1.Text = ""
End Select
End Sub

Private Sub Combo2_Change()
Select Case Combo2.Text
Case "简单"
Form3.Combo2.Text = "简单"
Case "普通"
Form3.Combo2.Text = "普通"
Case "困难"
Form3.Combo2.Text = "困难"
Case Else
Combo2.Text = ""
End Select
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "简单" Then
Form1.Label14.Caption = 0
ElseIf Combo2.Text = "普通" Then
Form1.Label14.Caption = 1
ElseIf Combo2.Text = "困难" Then
Form1.Label14.Caption = 2
End If
Select Case Combo2.Text
Case "简单"
Form3.Combo2.Text = "简单"
Case "普通"
Form3.Combo2.Text = "普通"
Case "困难"
Form3.Combo2.Text = "困难"
Case Else
Combo2.Text = ""
End Select
End Sub

Private Sub Command1_Click()
If Combo1.Text = "无尽模式" And Combo2.Text = "" Then
Form1.Show: Form2.Hide: GoTo gh
End If
If Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox "请输入模式和难度。", 48
Else
Form1.Show: Form2.Hide
End If
gh:
End Sub

Private Sub Command2_Click()
Form2.Hide: Form3.Show
Label2.Caption = Label2.Caption + 1
End Sub

Private Sub Command3_Click()
Form6.Show: Form2.Hide
End Sub

Private Sub Command5_Click()
Open Text3.Text For Output As #1
    Print #1, "p" & Mid(Label9.Caption, 4)
    Print #1, Combo1.Text
    Print #1, Combo2.Text
Close
Timer3.Enabled = True
End Sub

Private Sub Form_Load()
Open "键位.txt " For Append As #1
Close

'读取键位.txt中的键位并写入
Dim a(3)
i = -1
Open "键位.txt " For Input As #1
Do While Not EOF(1)
Line Input #1, a(i + 1)
i = i + 1
Loop
Close
Form3.Text1.Text = a(0)
Form3.Text2.Text = a(1)
Form3.Text3.Text = a(2)
Form3.Text4.Text = a(3)

If Form3.Text1.Text <> "" Or Form3.Text2.Text <> "" Or Form3.Text3.Text <> "" Or Form3.Text4.Text <> "" Then GoTo hjk
'在键位.txt文件中写入“w”“a”“s”“d”
Open "键位.txt" For Output As #1
Print #1, "w"
Print #1, "s"
Print #1, "a"
Print #1, "d"
Close
Dim gh(3): i = 0
Open "键位.txt " For Input As #1
Do While Not EOF(1)
Line Input #1, gh(i)
i = i + 1
Loop
Close
Form3.Text1.Text = "w"
Form3.Text2.Text = "s"
Form3.Text3.Text = "a"
Form3.Text4.Text = "d"
hjk:

End Sub

Private Sub Text3_Change()
On Error Resume Next
'输入标号查找人员
If Not Val(Text3.Text) <= 0 Then Text3.Text = List1.List(Text3.Text)
If Text3.Text = "" Or Text3.Text = " " Then Command5.Enabled = False Else Command5.Enabled = True
Command5.SetFocus
End Sub

Private Sub Timer1_Timer()
'读取公告.txt文件并写入picture控件
Open "公告.txt" For Input As #1
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

Private Sub Timer2_Timer()
Dim a(5)
'当你作为被邀请方
'提取自己的文本文档中的信息
i = 0
Open Form5.Label10.Caption & "\联机\" & Mid(Label9.Caption, 4) & ".txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, a(i)
        i = i + 1
    Loop
Close
'提取出用户名
user = Mid(a(0), 2)
Label10.Caption = user
Select Case a(1)
Case "普通模式"
    Form1.Label13.Caption = 0
Case "看脸模式"
    Form1.Label13.Caption = 1
Case "无尽模式"
    Form1.Label13.Caption = 2
End Select
Select Case a(2)
Case "简单"
    Form1.Label14.Caption = 0
Case "普通"
    Form1.Label14.Caption = 1
Case "困难"
    Form1.Label14.Caption = 2
End Select
'如果提取出字母“p”，则弹出消息框
If Left(a, 1) = "p" Then ma = MsgBox("玩家：" & user & "邀请你玩" & a(1) & " " & a(2), vbYesNo)
If ma = 4 Then
    Form1.Show
    Open Form5.Label10.Caption & "\联机\" & user & ".txt" For Output As #1
        Print #1, "Y"
    Close
End If
End Sub

Private Sub Timer3_Timer()
'当你作为邀请方
'提取对方的文本文档中的信息
Open Form5.Label10.Caption & "\联机\" & user & ".txt" For Input As #2
    Line Input #2, a
Close
'如果提取出字母“Y”，则跳转到form1窗体
If a = "Y" Then Form1.Show
End Sub

Private Sub VScroll1_Change()
Label1.Top = -(VScroll1.Value)
End Sub

Private Sub VScroll1_Scroll()
Label1.Top = -(VScroll1.Value)
End Sub
