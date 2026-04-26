VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "设置"
   ClientHeight    =   7080
   ClientLeft      =   4305
   ClientTop       =   2880
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6390
   Begin VB.CommandButton Command2 
      Caption         =   "注销用户"
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "模式、难度设置"
      Height          =   1575
      Left            =   960
      TabIndex        =   10
      Top             =   4440
      Width           =   3375
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form3.frx":0000
         Left            =   1200
         List            =   "Form3.frx":000D
         TabIndex        =   12
         Text            =   "普通模式"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "Form3.frx":002F
         Left            =   1200
         List            =   "Form3.frx":003C
         TabIndex        =   11
         Text            =   "简单"
         Top             =   960
         Width           =   1215
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
         Left            =   360
         TabIndex        =   14
         Top             =   480
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
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "键位设置"
      Height          =   3735
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   720
      Width           =   3375
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text2 
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text3 
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text4 
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   2
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上："
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
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "下："
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
         Left            =   600
         TabIndex        =   8
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "左："
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
         Left            =   600
         TabIndex        =   7
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "右："
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
         Left            =   600
         TabIndex        =   6
         Top             =   2640
         Width           =   990
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   495
      Left            =   5880
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ai As New ADODB.Connection '声明并创建链接对象
Dim rs As ADODB.Recordset '声明链接对象类型
Dim sql$

Private Sub Combo1_Change()
'游戏模式与form2的游戏模式互通，combo2控件同理
Select Case Combo1.Text
Case "看脸模式"
Form2.Combo1.Text = "看脸模式"
Case "普通模式"
Form2.Combo1.Text = "普通模式"
Case "无尽模式"
Form3.Combo1.Text = "无尽模式"
Case Else
Combo1.Text = ""
End Select
End Sub

Private Sub Combo1_Click()
If Form2.Label2.Caption = 1 Then
Else
Label5.Caption = 1
End If
If Combo1.Text = "看脸模式" Then
Form1.Label13.Caption = 1
ElseIf Combo1.Text = "普通模式" Then
Form1.Label13.Caption = 0
ElseIf Combo1.Text = "无尽模式" Then
Form1.Label13.Caption = 2
End If
Select Case Combo1.Text
Case "看脸模式"
Combo2.Enabled = True
Form2.Combo1.Text = "看脸模式"
Form2.Combo2.Enabled = True
Case "普通模式"
Combo2.Enabled = True
Form2.Combo1.Text = "普通模式"
Form2.Combo2.Enabled = True
Case "无尽模式"
Form2.Combo1.Text = "无尽模式"
Combo2.Text = ""
Combo2.Enabled = False
Form2.Combo2.Text = ""
Form2.Combo2.Enabled = False
Case Else
Combo1.Text = ""
End Select
End Sub

Private Sub Combo2_Change()
Select Case Combo2.Text
Case "简单"
Form2.Combo2.Text = "简单"
Case "普通"
Form2.Combo2.Text = "普通"
Case "困难"
Form2.Combo2.Text = "困难"
Case Else
Combo2.Text = ""
End Select
End Sub

Private Sub Combo2_Click()
If Form2.Label2.Caption = 1 Then
Else
Label5.Caption = 1
End If
If Combo2.Text = "简单" Then
Form1.Label14.Caption = 0
ElseIf Combo2.Text = "普通" Then
Form1.Label14.Caption = 1
ElseIf Combo2.Text = "困难" Then
Form1.Label14.Caption = 2
End If
Select Case Combo2.Text
Case "简单"
Form2.Combo2.Text = "简单"
Case "普通"
Form2.Combo2.Text = "普通"
Case "困难"
Form2.Combo2.Text = "困难"
Case Else
Combo2.Text = ""
End Select
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim a(3)
a(0) = Val(Asc(Text1.Text))
a(1) = Val(Asc(Text2.Text))
a(2) = Val(Asc(Text3.Text))
a(3) = Val(Asc(Text4.Text))
'防止键位重复
For i = 0 To 3
For j = i + 1 To 3
If a(i) = a(j) Or a(i) = a(j) + 32 Or a(i) + 32 = a(j) Then MsgBox "按键重复，请重新输入(不区分大小写)": GoTo G
Next j
Next i
'判定键位是从“a”到“z”，否则报错
For e = 0 To 3
If (65 <= a(e) And a(e) <= 90) Or (97 <= a(e) And a(e) <= 122) Then
Else
MsgBox "无效值，请重新输入。", 48
GoTo G
End If
Next
'写入键位.txt文件
Open "键位.txt" For Output As #1
Print #1, Text1.Text
Print #1, Text2.Text
Print #1, Text3.Text
Print #1, Text4.Text
Close
'判定你在哪个窗体点的“设置”
If Form2.Label2.Caption = 1 Then
    Form2.Show: Form3.Hide
    Form2.Timer1.Interval = 1
    Form2.Label2.Caption = 0
ElseIf Form1.Label7.Caption = 1 Then
    Form1.Show: Form3.Hide:
    Form1.Label7.Caption = 0
    '如果切换模式，则触发
    If Label5.Caption = 1 Then
        MsgBox "请重新开始游戏。", 48
        Form1.Command6.Enabled = False
        Form1.Command4.Enabled = True
        Form1.Command1.Enabled = True
    End If
End If
G:
End Sub

Private Sub Command2_Click()
On Error Resume Next
With ai
    If Form5.Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        .Open
    End If
End With
a = MsgBox("是否注销此账户？", 64 + 4)
If a = 6 Then
    sql = "delete from [用户登录] where 用户名 = ' " & Trim(Form5.Text4.Text) & " '"
    ai.Execute sql
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    sql = "delete from [用户成绩排行] where 用户名 = ' " & Trim(Form5.Text4.Text) & " '"
    rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    sql = "drop table [" & Trim(Form5.Text4.Text) & "]"
    rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
    Set rs = Nothing
    Form5.Show: Form3.Hide
    Form5.Text3 = "": Form5.Text4 = ""
End If
End Sub
