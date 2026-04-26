VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   8760
   ClientLeft      =   4800
   ClientTop       =   2400
   ClientWidth     =   9600
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   9600
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   3840
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   3240
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   2760
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   2280
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   1800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   1320
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      LargeChange     =   600
      Left            =   1560
      SmallChange     =   600
      TabIndex        =   10
      Top             =   7440
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   600
      Left            =   1560
      SmallChange     =   600
      TabIndex        =   9
      Top             =   7440
      Width           =   6735
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   5895
      LargeChange     =   600
      Left            =   360
      SmallChange     =   600
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5895
      LargeChange     =   600
      Left            =   360
      SmallChange     =   600
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "个人记录"
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查看排行榜"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回游戏"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9120
      Top             =   120
   End
   Begin VB.Frame Frame2 
      Caption         =   "排行榜"
      Height          =   6735
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   8055
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   8055
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Menu ckph 
      Caption         =   "查看排行"
      Visible         =   0   'False
      Begin VB.Menu pt 
         Caption         =   "普通"
         Begin VB.Menu jd0 
            Caption         =   "简单"
         End
         Begin VB.Menu pt0 
            Caption         =   "普通"
         End
         Begin VB.Menu kn0 
            Caption         =   "困难"
         End
      End
      Begin VB.Menu kl 
         Caption         =   "看脸"
         Begin VB.Menu jd1 
            Caption         =   "简单"
         End
         Begin VB.Menu pt1 
            Caption         =   "普通"
         End
         Begin VB.Menu kn1 
            Caption         =   "困难"
         End
      End
      Begin VB.Menu wj 
         Caption         =   "无尽"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ai As New ADODB.Connection '声明并创建链接对象
Dim rs As ADODB.Recordset '声明链接对象类型
Dim sql$, sel$, a, b, c

Private Sub Command1_Click() '返回游戏
Form1.Show: Form6.Hide
End Sub

Private Sub Command2_Click() '查看排行榜
Frame2.Visible = True
VScroll2.Visible = False
VScroll1.Visible = True
End Sub

Private Sub Command3_Click() '个人记录
Frame2.Visible = False
VScroll2.Visible = True
VScroll1.Visible = False
End Sub

Private Sub Form_Click()
Dim a(1 To 6)
Label1.Caption = ""
Label2.Caption = ""
'frame1.label2 的 个人记录
a(1) = "模式" & "              ": a(2) = "难度" & "         ": a(3) = "得分" & "      ": a(4) = "游玩时间" & "   ": a(5) = "结果" & "       ": a(6) = "时间"
For i = 1 To 6
hg = hg & a(i)
Next
Label2.Caption = hg
'frame2.label1 的 排行榜记录
a(1) = "用户名" & "        ": a(2) = "模式" & "            ": a(3) = "难度" & "        ": a(4) = "得分" & "      ": a(5) = "通关时间"
For i = 1 To 5
gh = gh & a(i)
Next
Label1.Caption = gh
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
If ai.State = 0 Then
    If Form5.Label10.Caption = "app.path" Then
        ai.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        ai.Open
    Else
        ai.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Form5.Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        ai.Open
    End If
End If
Dim a(1 To 6)
'frame1.label2 的 个人记录
a(1) = "模式" & "              ": a(2) = "难度" & "         ": a(3) = "得分" & "      ": a(4) = "游玩时间" & "   ": a(5) = "结果" & "       ": a(6) = "时间"
For i = 1 To 6
hg = hg & a(i)
Next
Label2.Caption = hg

Frame1.Caption = Form5.Text4.Text & "用户记录"
VScroll2.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu ckph
End Sub

Private Sub HScroll1_Change()
Label2.Left = -(HScroll1.Value) + 240
End Sub

Private Sub HScroll1_Scroll()
Label2.Left = -(HScroll1.Value) + 240
End Sub

Private Sub HScroll2_Change()
Label1.Left = -(HScroll2.Value) + 240
End Sub

Private Sub HScroll2_LostFocus()
Label1.Left = -(HScroll2.Value) + 240
End Sub

Private Sub jd0_Click()
Timer2.Enabled = True
End Sub

Private Sub jd1_Click()
Timer5.Enabled = True
End Sub

Private Sub wj_Click()
Timer8.Enabled = True
End Sub

Private Sub kn0_Click()
Timer4.Enabled = True
End Sub

Private Sub kn1_Click()
Timer7.Enabled = True
End Sub

Private Sub pt0_Click()
Timer3.Enabled = True
End Sub

Private Sub pt1_Click()
Timer6.Enabled = True
End Sub

Private Sub Timer1_Timer() '提取个人记录并输出到 frame1.label2
On Error Resume Next
Set rs = New ADODB.Recordset
sql = "select * from [" & Form5.Text4.Text & "] "
rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
Do While Not rs.EOF
    i = 0
    Do While i <= 5
        hg = hg & rs(i).Value & "        "
        i = i + 1
    Loop
    hg = hg & Chr(10)
    rs.MoveNext
Loop
Set rs = Nothing
rs.Close

Label2.Caption = Label2.Caption & Chr(10) & hg
'当记录超出frame1的宽度时启用vscroll2(垂直滚动条)控件进行浏览
If Frame1.Height < Label2.Height Then VScroll2.Max = Label2.Height - Frame1.Height Else VScroll2.Enabled = False
If Frame1.Width < Label2.Width Then HScroll2.Max = Label2.Width - Frame1.Width Else HScroll2.Enabled = False
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer() '提取排行榜记录并输出到 frame2.label1
Dim l(1 To 6): Label1.Caption = "": gh = "": hg = "": hg = ""
'frame2.label1 的 排行榜记录
l(1) = "用户名" & "        ": l(2) = "模式" & "            ": l(3) = "难度" & "        ": l(4) = "得分" & "      ": l(5) = "通关时间"
For i = 1 To 5
gh = gh & l(i)
Next
Label1.Caption = gh:  gh = ""
a = "普通模式":  b = "简单": c = 100
'调用过程PHB
Call PHB

'当记录超出frame2的宽度时启用vscroll1(垂直滚动条)控件进行浏览
If Frame2.Height < Label1.Height Then VScroll1.Max = Label1.Height - Frame2.Height Else VScroll1.Enabled = False
If Frame2.Width < Label1.Width Then HScroll1.Max = Label1.Width - Frame2.Width Else HScroll1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Dim l(1 To 6): Label1.Caption = "": gh = "": hg = ""
'frame2.label1 的 排行榜记录
l(1) = "用户名" & "        ": l(2) = "模式" & "            ": l(3) = "难度" & "        ": l(4) = "得分" & "      ": l(5) = "通关时间"
For i = 1 To 5
gh = gh & l(i)
Next
Label1.Caption = gh:  gh = ""
a = "普通模式": b = "普通": c = 100
'调用过程PHB
Call PHB

'当记录超出frame2的宽度时启用vscroll1(垂直滚动条)控件进行浏览
If Frame2.Height < Label1.Height Then VScroll1.Max = Label1.Height - Frame2.Height Else VScroll1.Enabled = False
If Frame2.Width < Label1.Width Then HScroll1.Max = Label1.Width - Frame2.Width Else HScroll1.Enabled = False
Timer3.Enabled = False
End Sub

Sub PHB()
On Error Resume Next
Dim OPL As New ADODB.Connection  '链接吃豆人.mdb数据库
With OPL     'mdb格式链接
    If Form5.Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Form5.Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"    '链接数据库的驱动字符
        .Open
    End If
End With
Dim OPL_rs As New ADODB.Recordset '声明链接对象类型
Dim OPL_sql$

'用户表，提取用户最好的记录
OPL_sel = "select min(游玩时间) as 通关时间 from [" & Form5.Text4.Text & "]" & " where 模式 = '" & a & "' and 难度 = '" & b & "'" & " and 得分 =  " & c
OPL_rs.Open OPL_sel, OPL, adOpenKeyset, adLockBatchOptimistic

'提取各字段值，并记录到gl变量中
Do While Not OPL_rs.EOF
    gl = OPL_rs(0).Value
    Exit Do
Loop
Set OPL_rs = Nothing

'用户表→排行表，如果排行榜中有此用户的记录，就更新记录，否则插入记录
Dim DB As New ADODB.Connection  '链接吃豆人.mdb数据库
With DB     'mdb格式链接
    If Form5.Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Form5.Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        .Open
    End If
End With
Dim DB_rs As New ADODB.Recordset '声明链接对象类型
Dim DB_into$, DB_up$, DB_sel$

Set DB_rs = New ADODB.Recordset
DB_sel = "select 用户名  from [用户成绩排行] where 用户名 ='" & Space(1) & Form5.Text4.Text & "' and  模式 = '" & Space(1) & a & "' and 难度 = '" & Space(1) & b & "'" & " and 得分 =  " & c
DB_rs.Open DB_sel, DB, adOpenKeyset, adLockBatchOptimistic

If DB_rs.EOF = True Then
    Set DB_rs = New ADODB.Recordset
    DB_into = "insert into [用户成绩排行](用户名,模式,难度,得分,通关时间) values (' " & Form5.Text4.Text & " ',' " & a & " ',' " & b & " ', " & c & " , " & gl & " ) " '插入记录
    DB_rs.Open DB_into, DB, adOpenKeyset, adLockBatchOptimistic
    
    Set DB_rs = Nothing
Else
    Set DB_rs = New ADODB.Recordset
    DB_up = "update 用户成绩排行 set 通关时间 = " & gl & " where 模式 = ' " & a & " 'and  难度 = ' " & b & " 'and  得分 =  " & c & "  and 用户名 = ' " & Form5.Text4.Text & " '"  '更新记录
    DB_rs.Open DB_up, DB, adOpenKeyset, adLockBatchOptimistic
    
    Set DB_rs = Nothing

End If

'排行表，对记录进行排序
Dim D As New ADODB.Connection  '链接吃豆人.mdb数据库
With D     'mdb格式链接
    If Form5.Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Form5.Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        .Open
    End If
End With
Dim D_rs As New ADODB.Recordset '声明链接对象类型
Dim D_sel
D_sel = "select * from [用户成绩排行] where 模式 = '" & Space(1) & a & "' and 难度 = '" & Space(1) & b & "'" & " and 得分 =  " & c & "  order by 通关时间 asc "
D_rs.Open D_sel, D, adOpenKeyset, adLockBatchOptimistic

'提取记录，输出到label1
Do While Not D_rs.EOF
    i = 0
    Do While i <= 4
        hg = hg & D_rs(i).Value & "        "
        i = i + 1
    Loop
    hg = hg & Chr(10)
    D_rs.MoveNext
Loop
Set D_rs = Nothing
Label1.Caption = Label1.Caption & Chr(10) & hg
End Sub

Private Sub Timer4_Timer()
Dim l(1 To 6): Label1.Caption = "": gh = "": hg = ""
'frame2.label1 的 排行榜记录
l(1) = "用户名" & "        ": l(2) = "模式" & "            ": l(3) = "难度" & "        ": l(4) = "得分" & "      ": l(5) = "通关时间"
For i = 1 To 5
gh = gh & l(i)
Next
Label1.Caption = gh:  gh = ""
a = "普通模式": b = "困难": c = 100
'调用过程PHB
Call PHB

'当记录超出frame2的宽度时启用vscroll1(垂直滚动条)控件进行浏览
If Frame2.Height < Label1.Height Then VScroll1.Max = Label1.Height - Frame2.Height Else VScroll1.Enabled = False
If Frame2.Width < Label1.Width Then HScroll1.Max = Label1.Width - Frame2.Width Else HScroll1.Enabled = False
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Dim l(1 To 6): Label1.Caption = "": gh = "": hg = ""
'frame2.label1 的 排行榜记录
l(1) = "用户名" & "        ": l(2) = "模式" & "            ": l(3) = "难度" & "        ": l(4) = "得分" & "      ": l(5) = "通关时间"
For i = 1 To 5
gh = gh & l(i)
Next
Label1.Caption = gh:  gh = ""
a = "看脸模式": b = "简单": c = 50
'调用过程PHB
Call PHB

'当记录超出frame2的宽度时启用vscroll1(垂直滚动条)控件进行浏览
If Frame2.Height < Label1.Height Then VScroll1.Max = Label1.Height - Frame2.Height Else VScroll1.Enabled = False
If Frame2.Width < Label1.Width Then HScroll1.Max = Label1.Width - Frame2.Width Else HScroll1.Enabled = False
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
Dim l(1 To 6): Label1.Caption = "": gh = "": hg = ""
'frame2.label1 的 排行榜记录
l(1) = "用户名" & "        ": l(2) = "模式" & "            ": l(3) = "难度" & "        ": l(4) = "得分" & "      ": l(5) = "通关时间"
For i = 1 To 5
gh = gh & l(i)
Next
Label1.Caption = gh:  gh = ""
a = "看脸模式": b = "普通": c = 50

'调用过程PHB
Call PHB

'当记录超出frame2的宽度时启用vscroll1(垂直滚动条)控件进行浏览
If Frame2.Height < Label1.Height Then VScroll1.Max = Label1.Height - Frame2.Height Else VScroll1.Enabled = False
If Frame2.Width < Label1.Width Then HScroll1.Max = Label1.Width - Frame2.Width Else HScroll1.Enabled = False
Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
Dim l(1 To 6): Label1.Caption = "": gh = "": hg = ""
'frame2.label1 的 排行榜记录
l(1) = "用户名" & "        ": l(2) = "模式" & "            ": l(3) = "难度" & "        ": l(4) = "得分" & "      ": l(5) = "通关时间"
For i = 1 To 5
gh = gh & l(i)
Next
Label1.Caption = gh:  gh = ""
a = "看脸模式": b = "困难": c = 50
'调用过程PHB
Call PHB

'当记录超出frame2的宽度时启用vscroll1(垂直滚动条)控件进行浏览
If Frame2.Height < Label1.Height Then VScroll1.Max = Label1.Height - Frame2.Height Else VScroll1.Enabled = False
If Frame2.Width < Label1.Width Then HScroll1.Max = Label1.Width - Frame2.Width Else HScroll1.Enabled = False
Timer7.Enabled = False
End Sub

Private Sub Timer8_Timer()
On Error Resume Next
Dim l(1 To 6), gl(0 To 1): Label1.Caption = "": gh = "": hg = ""
'frame2.label1 的 排行榜记录
l(1) = "用户名" & "        ": l(2) = "模式" & "            ": l(3) = "难度" & "        ": l(4) = "得分" & "      ": l(5) = "通关时间"
For i = 1 To 5
gh = gh & l(i)
Next
Label1.Caption = gh:  gh = ""
a = "无尽模式": b = "无"

'On Error Resume Next
Dim OPL As New ADODB.Connection  '链接吃豆人.mdb数据库
With OPL
    If Form5.Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Form5.Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        .Open
    End If
End With
Dim OPL_rs As New ADODB.Recordset '声明链接对象类型
Dim OPL_sel$
'用户表，提取用户最好的记录
OPL_sel = "select max(得分) as 无尽得分,min(游玩时间) as 通关时间 from [" & Form5.Text4.Text & "]" & " where 模式 = '" & a & "' and 难度 = '" & b & "'"
OPL_rs.Open OPL_sel, OPL, adOpenKeyset, adLockBatchOptimistic

'提取各字段值，并记录到gl变量中
Do While Not OPL_rs.EOF
    i = 0
    Do While i <= 1
        gl(i) = OPL_rs(i).Value
        i = i + 1
    Loop
    Exit Do
Loop
Set OPL_rs = Nothing


'用户表→排行表，如果排行榜中有此用户的记录，就更新记录，否则插入记录
Dim DB As New ADODB.Connection  '链接吃豆人.mdb数据库
With DB     'mdb格式链接
    If Form5.Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Form5.Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        .Open
    End If
End With
Dim DB_rs As New ADODB.Recordset '声明链接对象类型
Dim DB_into$, DB_up$, DB_sel$

Set DB_rs = New ADODB.Recordset
DB_sel = "select 用户名  from [用户成绩排行] where 用户名 ='" & Space(1) & Form5.Text4.Text & "' and  模式 = '" & Space(1) & a & "' and 难度 = '" & Space(1) & b & "'"
DB_rs.Open DB_sel, DB, adOpenKeyset, adLockBatchOptimistic

If DB_rs.EOF = True Then
    Set DB_rs = New ADODB.Recordset
    DB_into = "insert into [用户成绩排行](用户名,模式,难度,得分,通关时间) values (' " & Form5.Text4.Text & " ',' " & a & " ',' " & b & " ', " & gl(0) & " , " & gl(1) & " ) " '插入记录
    DB_rs.Open DB_into, DB, adOpenKeyset, adLockBatchOptimistic
    
    Set DB_rs = Nothing
Else
    Set DB_rs = New ADODB.Recordset
    DB_up = "update 用户成绩排行 set 通关时间 = " & gl(1) & " , 得分 = " & gl(0) & " where 模式 = ' " & a & " 'and  难度 = ' " & b & " ' and 用户名 = ' " & Form5.Text4.Text & " '"  '更新记录
    DB_rs.Open DB_up, DB, adOpenKeyset, adLockBatchOptimistic
    
    Set DB_rs = Nothing
End If

'排行表，对记录进行排序
Dim D As New ADODB.Connection  '链接吃豆人.mdb数据库
With D     'mdb格式链接
    If Form5.Label10.Caption = "app.path" Then
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
        .Open
    Else
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Form5.Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
        .Open
    End If
End With
Dim D_rs As New ADODB.Recordset '声明链接对象类型
Dim D_sel
D_sel = "select * from [用户成绩排行] where 模式 = '" & Space(1) & a & "' and 难度 = '" & Space(1) & b & "' order by 得分 desc ,通关时间 asc "
D_rs.Open D_sel, D, adOpenKeyset, adLockBatchOptimistic

'提取记录，输出到label1
Do While Not D_rs.EOF
    i = 0
    Do While i <= 4
        hg = hg & D_rs(i).Value & "        "
        i = i + 1
    Loop
    hg = hg & Chr(10)
    D_rs.MoveNext
Loop
Set D_rs = Nothing
Label1.Caption = Label1.Caption & Chr(10) & hg


'当记录超出frame2的宽度时启用vscroll1(垂直滚动条)控件进行浏览
If Frame2.Height < Label1.Height Then VScroll1.Max = Label1.Height - Frame2.Height Else VScroll1.Enabled = False
If Frame2.Width < Label1.Width Then HScroll1.Max = Label1.Width - Frame2.Width Else HScroll1.Enabled = False
Timer8.Enabled = False
End Sub

Private Sub VScroll1_Change()
Label1.Top = -(VScroll1.Value) + 120
End Sub

Private Sub VScroll1_Scroll()
Label1.Top = -(VScroll1.Value) + 120
End Sub

Private Sub VScroll2_Change()
Label2.Top = -(VScroll2.Value) + 120
End Sub

Private Sub VScroll2_Scroll()
Label2.Top = -(VScroll2.Value) + 120
End Sub
