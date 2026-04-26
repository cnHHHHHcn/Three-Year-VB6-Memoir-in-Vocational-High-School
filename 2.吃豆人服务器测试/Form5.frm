VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "登录系统"
   ClientHeight    =   4710
   ClientLeft      =   6240
   ClientTop       =   4650
   ClientWidth     =   6345
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6345
   Begin VB.Frame Frame2 
      Caption         =   "忘记"
      Height          =   4095
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   3
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "确定"
         Height          =   495
         Left            =   1320
         TabIndex        =   19
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   3
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   17
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "取消"
         Height          =   495
         Left            =   3480
         TabIndex        =   16
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "密码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   600
         TabIndex        =   24
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   5400
         Picture         =   "Form5.frx":0000
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "新密码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "用户名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1740
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   5400
         Picture         =   "Form5.frx":467F
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   480
         TabIndex        =   20
         Top             =   1320
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "新用户注册"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton Command4 
         Caption         =   "取消"
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   3
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "确定"
         Height          =   495
         Left            =   1320
         TabIndex        =   1
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "用户名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "密码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   600
         TabIndex        =   4
         Top             =   1680
         Width           =   1305
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "是否下次显示"
      Height          =   495
      Left            =   4680
      TabIndex        =   26
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "注册"
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "忘记"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   9
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   3
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1680
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "确定"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   5760
      Picture         =   "Form5.frx":8CFE
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4440
      Picture         =   "Form5.frx":D37D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   11
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   840
      TabIndex        =   10
      Top             =   1680
      Width           =   1305
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ai As New ADODB.Connection '声明并创建链接对象
Dim rs As ADODB.Recordset '声明链接对象类型
Dim sql$

Private Sub Check1_Click()
'判定是否下次开启更换数据库路径
If Check1.Value = 1 Then
    Open "路径\数据库路径提示.txt" For Output As #1
    Print #1, "如果下一行输入1则不提示共享路径，否则反之"
    Close
Else
    Open "路径\数据库路径提示.txt" For Output As #1
    Print #1, "如果下一行输入1则不提示共享路径，否则反之"
    Print #1, 1
    Close
End If
End Sub

Private Sub Command1_Click() '忘记
Set rs = New ADODB.Recordset
sql = " select * from [用户登录] where 用户名 = '" & Space(1) & Text1.Text & " '"
rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
If rs.EOF = True Then
    MsgBox "无此用户名，请输入新的用户名。", 48
    Frame1.Visible = True
Else
    Frame2.Visible = True
    Text5.Text = Text4.Text
End If
End Sub

Private Sub Command3_Click() '注册用户  登录系统 frame1(新用户注册) 的 确定
'创建数据库集对象的链接
Set rs = New ADODB.Recordset
If Text1.Text <> "" And Text2.Text <> "" Then
    If Not (LenB(Text1.Text) = 12 And LenB(Text2.Text) = 6) Then MsgBox "用户不能少于6位，密码不能少于3位。", 48: GoTo wu
        'sql插入语句
        sql = "insert into [用户登录](用户名,密码) values (' " & Trim(Text1.Text) & " ', ' " & Trim(Text2.Text) & " ')"
              'insert into <表名> ([字段1],[字段2]……) values ('字段值1','字段2'，……)
              'insert into <表名> form array <数组名>
        rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
        Frame1.Visible = False
Else
    MsgBox "用户名和密码不能为空！", 16, "警告"
End If
wu:
End Sub

Private Sub Command3_LostFocus() '用户注册好之后 建<用户>表 登录系统 frame1(新用户注册) 的 确定
On Error Resume Next
Dim ai As New ADODB.Connection
With ai
    If Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        .Open
    End If
End With
ghj = "CREATE  TABLE " & Text1.Text & "(模式 text,难度 text,得分 long,游玩时间 long,结果 text,时间 Datetime)"
      'sql建表语句(不建议用逻辑型)
      'create table <表名> [(字段1 数据类型1,字段2 数据类型2,……)]
ai.Execute ghj
ai.Close
Text1 = "": Text2 = ""
End Sub

Private Sub Command4_Click() 'frame1(新用户注册) 的 取消
Text1.Text = "": Text2.Text = ""
Frame1.Visible = False
End Sub

Private Sub Command5_Click() '检测是否有此用户   登录系统
Static kl
Dim DB As New ADODB.Connection  '链接吃豆人.mdb数据库
With DB     'mdb格式链接
    If Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"    '链接数据库的驱动字符
        .Open
    End If
End With
Dim DB_rs As New ADODB.Recordset '声明链接对象类型
Dim DB_sql$
If Text3.Text <> "" Or Text4.Text <> "" Then
    If LenB(Text3.Text) = 6 And LenB(Text4.Text) = 12 Then
        'sql查询语句
        DB_sql = "select * from [用户登录] where 用户名 = ' " & Trim(Text4.Text) & " 'and 密码 = ' " & Trim(Text3.Text) & " ' "  '查询用户登录表
                 'select [distinct] [*] from <表名> where <条件>
        DB_rs.Open DB_sql, DB, adOpenKeyset, adLockBatchOptimistic
        If DB_rs.EOF = False Then  '查找记录
            Form5.Hide: Form2.Show
            MsgBox "账户登陆成功。", 64, "吃豆人登录系统": kl = 0
            Form2.Label9.Caption = Form2.Label9.Caption + Text4.Text
            Open Text4.Text For Output As #1
            Close
            Else
            kl = kl + 1
            '如果连续错误三次，则播报“账户或密码已错三次。请修改用户密码或注册新用户”
            If kl Mod 3 = 0 Then MsgBox "账户或密码已错三次。请修改用户密码或注册新用户", 64: GoTo er
            MsgBox "账户或密码有错误，请检查。"
er:
        End If
        DB_rs.Close
        Set DB_rs = Nothing
        DB.Close
        Set DB = Nothing
    Else
        MsgBox "用户不能少于6位，密码不能少于3位。", 48
    End If
Else
    MsgBox "用户名和密码不能为空！", 16, "警告"
End If
End Sub

Private Sub Command6_Click() '注册
Frame1.Visible = True
End Sub

Private Sub Command7_Click() 'frame2(忘记) 的 取消
Frame2.Visible = False: Text7.Text = "": Text6.Text = ""
End Sub

Private Sub Command8_Click() '修改用户密码  登录系统 frame2 的 确定
Dim DB As New ADODB.Connection  '链接吃豆人.mdb数据库
With DB     'mdb格式链接
    If Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        .Open
    End If
End With
Dim DB_rs As New ADODB.Recordset '声明链接对象类型
Dim DB_sql$
On Error Resume Next
If Text6.Text <> "" Or Text7.Text <> "" Then
    If Text6.Text = Text7.Text Then
        DB_sql = "update 用户登录 set 密码 = ' " & Text6.Text & " ' where 用户名 = ' " & Text5.Text & " ' "
                 'update <表名> set <字段值1>=<表达式>,<字段值2>=<表达式>,…… where <条件>
        DB_rs.Open DB_sql, DB, adOpenKeyset, adLockBatchOptimistic
        If DB_rs.EOF = False Then '查找记录
            MsgBox "密码修改成功。"
            Frame2.Visible = False
        Else
            MsgBox "无此用户，密码修改失败。", 16
        End If
    Else
        MsgBox "密码值不符", 16
    End If
Else
MsgBox "密码框值为空", 48
End If
End Sub

Private Sub Form_Load() '链接数据库路径
'查看数据库路径文件末尾是否有1
Open "路径\数据库路径提示.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, gh
    Loop
Close
'如果有就跳到 wts 后执行语句
If gh = 1 Then GoTo wts
'如果为“是”显示索要共享的路径，否则为本地路径
a = MsgBox("是否共享数据库？", 64 + 4)
If a = 6 Then
    Label10.Caption = InputBox("请输入共享路径。" & Chr(10) & "请更改括号内容，更改其他内容无法连接！！！", , "\\XSH(机号)\users")
Else
    Label10.Caption = "app.path"
End If
'保存路径
Open "路径\数据库路径.txt" For Output As #2
    Print #2, Label10.Caption
Close
'如果选“是”则下次显示，选“否”则下次不显示
a = MsgBox("是否下次显示？", 32 + 4)
If a = 7 Then
    Open "路径\数据库路径提示.txt" For Append As #1
        Print #1, "1"
    Close
Else
    Check1.Enabled = False
End If
GoTo hh
wts:
Open "路径\数据库路径.txt" For Input As #2
    Line Input #2, ad
    Label10.Caption = ad
Close
hh:
If Label10.Caption = "app.path" Then
    ai.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
    ai.Open
Else
    ai.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
    ai.Open
End If
End Sub

Private Sub Image2_Click() '密码一区
If Text3.PasswordChar = "*" Then Text3.PasswordChar = "" Else Text3.PasswordChar = "*"
End Sub

Private Sub Image3_Click() 'frame2(忘记)新密码一区
If Text6.PasswordChar = "*" Then Text6.PasswordChar = "" Else Text6.PasswordChar = "*"
End Sub

Private Sub Image4_Click() 'frame2(忘记)密码一区
If Text7.PasswordChar = "*" Then Text7.PasswordChar = "" Else Text7.PasswordChar = "*"
End Sub

Private Sub Text1_Change() 'frame1(新用户注册)提示
Label5.Caption = "注：账户和密码不能用汉字，账户名不能用标点符号、不区分大小写，" & Chr(10) & "             和不能用数字，否则无法记录到数据库。"
Label5.ForeColor = vbRed
End Sub

Private Sub Text1_LostFocus() '检测用户输入的用户名与数据库的用户名是否重复 frame1
'创建数据库集对象的链接
On Error Resume Next
Set rs = New ADODB.Recordset
sql = " select * from [用户登录] where 用户名 = '" & Space(1) & Text1.Text & " '"
rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
If rs.EOF = False Then
    MsgBox "用户名已占用，请输入新的用户名。"
    Text1.Text = "": Text1.SetFocus
End If
End Sub

