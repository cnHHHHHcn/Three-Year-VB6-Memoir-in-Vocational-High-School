VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8400
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "确定"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text2 
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
      Left            =   5400
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "检测"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "提取“输入”文件内容 "
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "启动"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   960
      Width           =   7335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "位置锁定:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   7
      Top             =   500
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "去除HTML标签"
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
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   2640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim checkposition() As String

Private Sub Command1_Click()
a = LabelTrim(ex_Label(ex_Label(ex_Label(Alltrim(Text1.Text), "<br>"), "<p>"), "</p>"))
Open "结果.txt" For Output As #1
    Print #1, a
Close
End Sub

Function LabelTrim(str As String) As String
'变量列表        作用
'LElabel         输出 "<" 符号的位置
'RIlabel         输出 ">" 符号的位置
'labelstr        截取的标签文本
'LENlabel        截取的标签文本的长度
'str             文本框中的文本
'labelwz         查找截取标签文本的位置
'LEstr           截取标签文本左边的文本
'RIstr           截取标签文本右边的文本


ks:
'查找标签的 "<"，">"
LElabel = InStr(str, "<")
RIlabel = InStr(str, ">")
If LElabel = 0 And RIlabel = 0 Then GoTo end2
'计算标签的长度
LENlabel = RIlabel - LElabel + 1
'截取标签文本
For i = LElabel To RIlabel
    labelstr = labelstr & Mid(str, i, 1)
Next
'-------------------------------------
strks:
'防止报错
On Error Resume Next
'查找标签的位置
labelwz = InStr(str, labelstr)
'如果未查找到标签位置，则不执行此语句块
If labelwz = 0 Then GoTo end1
'截取标签左边文本
LEstr = Left(str, labelwz - 1)
'截取标签右边文本
RIstr = Mid(str, labelwz + LENlabel)
'连接左边和右边文本
str = LEstr & RIstr
If labelwz <> 0 Then GoTo strks
'-------------------------------------
end1:
If RIlabel <> 0 Then labelstr = "": GoTo ks
end2:
LabelTrim = str
End Function

Function ex_Label(str As String, str_Label As String) As String
labelLEN = Len(str_Label)
strks:
'防止报错
On Error Resume Next
'查找标签的位置
labelwz = InStr(str, str_Label)
'如果未查找到标签位置，则不执行此语句块
If labelwz = 0 Then GoTo end1
'截取标签左边文本
LEstr = Left(str, labelwz - 1)
'截取标签右边文本
RIstr = Mid(str, labelwz + labelLEN)
'连接左边和右边文本
str = LEstr & vbCrLf & RIstr
If labelwz <> 0 Then GoTo strks
end1:
ex_Label = str
End Function

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Function Alltrim(str As String) As String
strks:
'防止报错
On Error Resume Next
'查找 换行符 的位置
codewz = InStr(str, Chr(13))
'如果未查找到 换行符 位置，则不执行此语句块
If codewz = 0 Then GoTo end1
'截取 换行符 左边文本
LEstr = Left(str, codewz - 1)
'截取 换行符 右边文本
RIstr = Mid(str, codewz + 1)
'连接左边和右边文本
str = LEstr & RIstr
If codewz <> 0 Then GoTo strks
end1:
Alltrim = str
End Function

Sub HTMLcheck()
On Error Resume Next
start = 1
start:
labelstr = ""
LEposition = InStr(start, Text1.Text, "<")
RIposition = InStr(start, Text1.Text, ">")
LENnum = RIposition - LEposition + 1
labelstr = Mid(Text1.Text, LEposition, LENnum)
textstr = LEposition & "     " & RIposition & "     " & LENnum & "   标签:" & labelstr
Debug.Print textstr
If LEposition <> 0 Then
    check = check + 1: start = RIposition + 1
    ReDim Preserve checkposition(check)
    checkposition(check) = textstr
    GoTo start
End If
End Sub

Function strcount(checkstr As String)
start = 1
start:
position = InStr(start, Text1.Text, checkstr)
If CBool(position) Then start = position + 1: strcount = strcount + 1: Debug.Print checkstr, strcount: GoTo start

End Function

Private Sub Command3_Click()
On Error Resume Next
Open "输入.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, str1
        strs = strs & vbCrLf & str1
    Loop
Close
Text1.Text = Right(strs, Len(strs) - 1)
End Sub

Private Sub Command4_Click()
HTMLcheck
Lnum = strcount("<"): Rnum = strcount(">")
Open App.Path & "\检查结果.txt" For Output As #1
    Print #1, "共计" & UBound(checkposition) & "个标签"
    Print #1, "“<”共计" & Lnum & "个     " & "“>”共计" & Rnum & "个"
    Print #1,
    For i = 1 To UBound(checkposition)
        Print #1, checkposition(i)
    Next
Close
Set a = CreateObject("wscript.shell")
a.run "检查结果.txt"
If Lnum = Rnum Then Command1.Enabled = True Else Command1.Enabled = False
End Sub

Private Sub Command5_Click()
Text1.SelStart = Val(Text2.Text)
Text1.SetFocus
End Sub

Private Sub Form_Load()
Command3_Click
End Sub

