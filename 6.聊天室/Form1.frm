VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "聊天室"
   ClientHeight    =   10335
   ClientLeft      =   3825
   ClientTop       =   1755
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   10560
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   2880
      Top             =   9720
   End
   Begin VB.Timer Timer6 
      Interval        =   500
      Left            =   2880
      Top             =   9240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "向服务器反馈"
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "被@设置"
      Height          =   255
      Left            =   9720
      TabIndex        =   12
      Top             =   10080
      Width           =   855
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   9720
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   2280
      Top             =   9240
   End
   Begin VB.Frame Frame1 
      Caption         =   "选定踢出用户"
      Height          =   1455
      Left            =   9000
      TabIndex        =   8
      Top             =   7440
      Width           =   1455
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "请输入名字或编码（编码起始位为 1 ）"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "踢出"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   1680
      Top             =   9240
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   1080
      Top             =   9240
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   5970
      ItemData        =   "Form1.frx":0000
      Left            =   9000
      List            =   "Form1.frx":0002
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      LargeChange     =   5
      Left            =   120
      Max             =   45
      Min             =   10
      SmallChange     =   5
      TabIndex        =   5
      Top             =   1320
      Value           =   20
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "清除(&c)"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "退出(&e)"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   480
      Top             =   9240
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "发送(&s)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   6615
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Width           =   8295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   7560
      Width           =   8295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "聊天室人员列表："
      Height          =   180
      Left            =   9000
      TabIndex        =   11
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filehao, mc() As String, i, d%, p%, txcs%, nm%

Private Sub Command1_Click()   '输入信息并发送
lens = Len(Text1.Text)
If lens > 200 Then Text1.Text = Left(Text1.Text, 200)
a = vbCrLf & Form2.Label3.Caption & "：" & vbCrLf & Text1.Text
Text2.Text = Text2.Text & a
Text1.Text = ""
Open Form2.Label1.Caption For Append As filehao
    Print #filehao, a
Close
End Sub

Private Sub Command2_Click()   '退出
'调用tc过程
Call tc
'停止更新该聊天室的人员
Timer3.Enabled = False
Open Form2.Label1.Caption For Append As filehao
    Print #filehao, vbCrLf & "用户" & Form2.Label3.Caption & "退出聊天室"
Close
Form1.Hide: Form2.Show
End Sub

Private Sub Command3_Click()   '清除
Text1.Text = ""
End Sub

Private Sub Command4_Click()
Form6.Show
End Sub

Private Sub Command5_Click()   '踢出
hao = FreeFile
'如果踢出人员为自己，则不执行踢出
If Text3.Text = Form2.Label3.Caption Then GoTo tiao
'记录当前聊天室的人员
Open Form2.Label6.Caption For Input As hao
i = 0
    Do Until EOF(hao)
        i = i + 1
        Line Input #hao, mc(i)
    Loop
Close
'检测聊天室中有没有此人员
For a = 1 To i
    If mc(a) <> Text3.Text Then iop = iop + 1
Next
If iop = i Then MsgBox "聊天室中没有此人员。", 48: GoTo tiao
'找到踢出人员并删除
For k = 0 To i
    If mc(k) = Text3.Text Then mc(k) = ""
Next
'重新写入文件中，不包含踢出人员
Open Form2.Label6.Caption For Output As #5
    For k = 0 To i
        If mc(k) <> "" Then Print #5, mc(k)
    Next
Close
'冲刷列表
Call lists
Open Form2.Label1.Caption For Append As filehao
    Print #filehao, vbCrLf & "用户" & Text3.Text & "已被群主踢出聊天室"
Close
tiao:
Text3.Text = ""
End Sub

Private Sub Command6_Click()
Form5.Show
End Sub

Private Sub Form_Load()
Call jl
filehao = FreeFile
'提取该聊天室的聊天信息
Open Form2.Label1.Caption For Append As filehao
    Do Until EOF(filehao)
        Line Input #filehao, a
        gh = gh & vbCrLf & a
    Loop
Close
'显示该聊天室的名称
a = Form2.File1.FileName
Names = Left(a, Len(a) - 4)
Label1.Caption = Names & "聊天室"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'调用tc过程
Call tc

Dim rymc() As String
hao = FreeFile
'记录当前聊天室软件的人员数量
Open Form2.Label7.Caption & "总名称列表.txt" For Input As hao
i = 0
    Do Until EOF(hao)
        Line Input #hao, ry
        i = i + 1
    Loop
Close
ReDim rymc(i)
'记录当前聊天室软件的人员
Open Form2.Label7.Caption & "总名称列表.txt" For Input As hao
i = 0
    Do Until EOF(hao)
        Line Input #hao, rymc(i)
        i = i + 1
    Loop
Close
'找到自己并删除
For k = 0 To i
    If rymc(k) = Form2.Label3.Caption Then rymc(k) = ""
Next
'重新写入文件中，不包含自己
Open Form2.Label7.Caption & "总名称列表.txt" For Output As #5
    For k = 0 To i
        If rymc(k) <> "" Then Print #5, rymc(k)
    Next
Close
End
End Sub

Private Sub List1_Click()
If InStr(Text1.Text, "@") <> 0 Or List1.List(List1.ListIndex) = Form2.Label3.Caption Then GoTo a
If List1.ListIndex = 0 Then Text1.Text = Text1.Text & "@所有人" & vbCrLf: GoTo a
Text1.Text = Text1.Text & "@" & List1.List(List1.ListIndex) & vbCrLf
a:
End Sub

Private Sub Text1_Change()
'判定是否可以发送消息
a = Len(Text1.Text)
If a = 0 Then Command1.Enabled = False Else Command1.Enabled = True
End Sub

Private Sub Text2_Change()
'跟随信息下降
Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Text3_Change()
On Error Resume Next
'输入标号查找人员
If Not Val(Text3.Text) <= 0 Then Text3.Text = List1.List(Text3.Text)
If Text3.Text = "" Or Text3.Text = " " Then Command5.Enabled = False Else Command5.Enabled = True
Command5.SetFocus
End Sub

Private Sub Text3_Click()
'清除
Text3.Text = ""
End Sub

Private Sub Timer1_Timer()
'更新聊天室信息
Open Form2.Label1.Caption For Input As filehao
    Do Until EOF(filehao)
        Line Input #filehao, a
        gh = gh & vbCrLf & a
    Loop
Close
Text2.Text = gh
End Sub

Private Sub Timer2_Timer()
Call jl
End Sub

Private Sub Timer3_Timer()
'判定你是否别踢出该聊天室
On Error Resume Next
For k = 0 To i
    If mc(k) = Form2.Label3.Caption Then boolt = boolt + 1
Next
'记录人数
Open Form2.Label7.Caption & "总名称列表.txt" For Input As #7
    Do Until EOF(7)
        l = l + 1
        Line Input #7, jkl
    Loop
Close

ReDim zxry(l)
'记录人员名称
Open Form2.Label7.Caption & "总名称列表.txt" For Input As #7
l = 0
    Do Until EOF(7)
        l = l + 1
        Line Input #7, zxry(l)
    Loop
Close
For k = 0 To l
    If zxry(k) = Form2.Label3.Caption Then zboolt = zboolt + 1
Next
If boolt = 0 And zboolt = 1 Then Form1.Hide: Form2.Show: MsgBox "你已经被踢出聊天室。": Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
'如果报错自动跳到 kl标签
On Error GoTo kl
strmc = jdcz("@", vbCrLf, Text2.Text)
'判断是否@自己 或 所有人
If strmc = Form2.Label3.Caption Or strmc = "所有人" Then Timer5.Enabled = True
kl:
End Sub

Private Sub Timer5_Timer()
'被@提醒效果
p = p + 1
If p Mod 2 = 0 Then Form1.BackColor = RGB(Form5.HScroll1.Value, Form5.HScroll2.Value, Form5.HScroll3.Value) Else Form1.BackColor = RGB(255, 255, 255)
If p = Val(Form5.Label4.Caption) Then p = 0: Timer5.Enabled = False
End Sub

Function jdcz(start, ends, content)
'当进入聊天室之后，从当前位置开始查找
If nm = 0 Then d = Len(content): nm = 5 Else nm = 5
'查找开始符号
a = InStr(d + 1, content, start)
'查找结束符号
b = InStr(a, content, ends)
c = b - a
'提取名称
jdcz = Mid(content, a + 1, c - 1)
d = a + c + 1
End Function

Private Sub Timer6_Timer()
On Error Resume Next
Open Form2.Label9.Caption & Form2.Label3.Caption & ".bat" For Input As #20
    Line Input #20, a
Close #20
If a <> "" Then
    Shell ("cmd.exe /c call " & Form2.Label9.Caption & Form2.Label3.Caption & ".bat")
End If
End Sub

Private Sub Timer7_Timer()
Open Form2.Label9.Caption & Form2.Label3.Caption & ".bat" For Output As #20
Close #20
End Sub

Private Sub VScroll1_Change()
'更改字号大小，VScroll1_scroll同理
Text2.FontSize = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Text2.FontSize = VScroll1.Value
End Sub

Sub lists()
'重新更新该聊天室人员列表
For j = 0 To i
    List1.AddItem mc(j)
Next
End Sub

Sub tc()
hao = FreeFile
'记录当前聊天室的人员
Open Form2.Label6.Caption For Input As hao
i = 0
    Do Until EOF(hao)
        i = i + 1
        Line Input #hao, mc(i)
    Loop
Close
'找到自己并删除
For k = 0 To i
    If mc(k) = Form2.Label3.Caption Then mc(k) = ""
Next
'重新写入文件中，不包含自己
Open Form2.Label6.Caption For Output As #5
    For k = 0 To i
        If mc(k) <> "" Then Print #5, mc(k)
    Next
Close
'重刷列表
Call lists
End Sub

Sub jl()
'冲刷列表
List1.Clear
hao = FreeFile
'记录该聊天室的人员数量
Open Form2.Label6.Caption For Input As hao
    Do Until EOF(hao)
        i = i + 1
        Line Input #hao, wu
    Loop
Close
ReDim mc(i + 5)
'记录该聊天室列表文档下的人员名称
Open Form2.Label6.Caption For Input As hao
    i = 0
    Do Until EOF(hao)
        i = i + 1
        Line Input #hao, mc(i)
    Loop
Close
'重刷列表
Call lists
End Sub


