VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4005
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4005
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "更新日志"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   4320
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1920
      Top             =   3960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "创建聊天室"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4(&r)"
      Height          =   180
      Left            =   10000
      TabIndex        =   6
      Top             =   10000
      Width           =   90
   End
   Begin VB.CommandButton Command3 
      Caption         =   "刷新"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "进入聊天室"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2250
      Left            =   600
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Label8"
      Height          =   180
      Left            =   360
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label8"
      Height          =   180
      Left            =   360
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   180
      Left            =   360
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   180
      Left            =   360
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   180
      Left            =   360
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label3 
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
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户："
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b, mc() As String

Private Sub Command1_Click()    '创建聊天室
Form3.Show
End Sub

Private Sub Command2_Click()
'判定是否选中聊天室
If Label1.Caption = "Label1" Then MsgBox "未选择聊天室。": GoTo tiao
'进入聊天室
Dim strings As String
hao = FreeFile
Open Label1.Caption For Input As hao
    Seek #hao, 14
    Line Input #hao, strings1
    Line Input #hao, strings2
Close
'记录创建该聊天室的用户
yh = Mid(strings2, 10)
'提取是否设置密码为 TRUE 还是 FALSE
setm = Mid(strings1, 1, 5)
'记录该聊天室密码
mima = Mid(strings1, 9)
'判定该用户是否为该聊天室的群主
If yh = Label3.Caption Then Form1.Frame1.Visible = True Else Form1.Frame1.Visible = False: Form1.Timer3.Enabled = False
'如果为 TRUE 则输入密码，否则进入
If setm = "True " Then
    ma = InputBox("聊天室密码为：")
    If ma = mima Then GoTo mc
Else
    GoTo mc
End If
GoTo tiao
mc:

Form1.Show: Form2.Hide
hao = FreeFile

'记录进出聊天室的人员
Open Label6.Caption For Append As hao
Print #hao, Label3.Caption
Close
'对该聊天室通报 “用户 *** 进入”
filehao = FreeFile
Open Form2.Label1.Caption For Append As filehao
    Print #filehao, vbCrLf & "用户" & Label3.Caption & "进入聊天室"
Close
Form1.Timer3.Enabled = True
tiao:

End Sub

Private Sub Command3_Click()
File1.Refresh
End Sub

Private Sub Command4_Click()
'重新设定路径
h = FreeFile
a = InputBox("请输入路径：", , File1.Path)
File1.Path = a
Open "聊天室路径.txt" For Output As h
    Print #h, a
Close
End Sub

Private Sub Command5_Click()
Form4.Show
End Sub

Private Sub File1_Click()
'选中聊天室刷新
Label1.Caption = File1.Path & "\" & File1.FileName
Label5.Caption = Left(File1.FileName, Len(File1.FileName) - 4)
Label6.Caption = b & Label5.Caption & "名称列表.txt"
End Sub

Private Sub Form_Load()
'On Error GoTo ab
'设定聊天室路径
h = FreeFile
Open App.Path & "\" & "聊天室路径.txt" For Input As h
    Line Input #h, a
Close
'a = "c:\users"
c = a & "\聊天室\"
b = a & "\聊天室名称列表\"
e = a & "\人员文档\"
Label4.Caption = c
Label6.Caption = b
Label7.Caption = b
Label8.Caption = b & "服务器反馈.txt"
Label9.Caption = e
File1.Path = c
'设置名称
gh:
Label3.Caption = InputBox("请输入名字。")
Select Case Label3.Caption
Case "", "王丹", "崔敬爱", "赵兰英", "董铭", "康建晓", "李海静", "郭晓坤", "张江华", "刘红莲", "张爱学", "贺林涛", "秦文龙", "习近平", "毛泽东", "周恩来", "邓小平"
    GoTo gh
End Select
'用户名第一位不能为数字
a = Asc(Label3.Caption)
If Asc("0") <= a And a <= Asc("9") Then GoTo gh
GoTo bc
ab:
Open b & "总名称列表.txt" For Output As #1
Close
bc:
'查找是否有相同名称的用户
hao = FreeFile
Open Label6.Caption & "总名称列表.txt" For Input As hao
    Do Until EOF(hao)
        Line Input #hao, ry
        If ry = Label3.Caption Then GoTo gh
    Loop
Close

'记录名称
hao = FreeFile
Open b & "总名称列表.txt" For Append As hao
    Print #hao, Label3.Caption
Close
Open e & Label3.Caption & ".bat" For Output As hao
Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rymc() As String
hao = FreeFile
Open Label7.Caption & "总名称列表.txt" For Input As hao
i = 0
    Do Until EOF(hao)
        Line Input #hao, ry
        i = i + 1
    Loop
Close
ReDim rymc(i)
Open Label7.Caption & "总名称列表.txt" For Input As hao
i = 0
    Do Until EOF(hao)
        Line Input #hao, rymc(i)
        i = i + 1
    Loop
Close
For k = 0 To i
    If rymc(k) = Form2.Label3.Caption Then rymc(k) = ""
Next
Open Label7.Caption & "总名称列表.txt" For Output As #5
    For k = 0 To i
        If rymc(k) <> "" Then Print #5, rymc(k)
    Next
Close
End
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'记录人数
Open b & "总名称列表.txt" For Input As #6
    Do Until EOF(6)
        i = i + 1
        Line Input #6, jkl
    Loop
Close

ReDim mc(i)
'记录人员名称
Open b & "总名称列表.txt" For Input As #6
i = 0
    Do Until EOF(6)
        i = i + 1
        Line Input #6, mc(i)
    Loop
Close

'判定你是否别踢出该聊天室
For k = 0 To i
    If mc(k) = Label3.Caption Then boolt = boolt + 1
Next
If boolt = 0 Then MsgBox "你已经被服务器管理员注销。": End
End Sub
