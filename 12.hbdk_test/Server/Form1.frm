VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   11685
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4920
      TabIndex        =   18
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Write_Client_ini 
      Caption         =   "编写客户端配置"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame SaveFrame 
      Caption         =   "保存"
      Enabled         =   0   'False
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   5535
      Begin VB.CommandButton Save_DNS 
         Caption         =   "保存域名"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox List_Jion_File 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   11
         ToolTipText     =   "将把以上列表信息与本文本框内指定的文件制作成域名表"
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton ChoiceFile 
         Caption         =   "选择"
         Height          =   495
         Left            =   4800
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "文件："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Frame RequestFrame 
      Caption         =   "申请区"
      Height          =   5295
      Left            =   5880
      TabIndex        =   5
      Top             =   0
      Width           =   5655
      Begin VB.ListBox Request_List 
         Appearance      =   0  'Flat
         Height          =   4350
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0010
         TabIndex        =   8
         Top             =   240
         Width           =   5295
      End
      Begin VB.CommandButton Send 
         Caption         =   "发送"
         Height          =   495
         Left            =   3960
         TabIndex        =   7
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CheckBox Extent 
         Caption         =   "是否同意"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   4725
         Width           =   1815
      End
      Begin VB.Label Count_Num 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         ToolTipText     =   "左键 + 1；右键 - 1"
         Top             =   4750
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "次数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   4800
         Width           =   945
      End
   End
   Begin VB.Timer ReBootPort 
      Interval        =   3000
      Left            =   3720
      Top             =   6000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "端口检测报告"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   6000
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   10440
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin VB.TextBox Path 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   2
      Top             =   5400
      Width           =   6255
   End
   Begin VB.FileListBox File 
      Appearance      =   0  'Flat
      Height          =   1830
      Left            =   120
      Pattern         =   "*.txt"
      TabIndex        =   1
      ToolTipText     =   "右键刷新"
      Top             =   120
      Width           =   5535
   End
   Begin VB.ListBox DNS_list 
      Appearance      =   0  'Flat
      Height          =   2190
      ItemData        =   "Form1.frx":0023
      Left            =   120
      List            =   "Form1.frx":0025
      TabIndex        =   0
      ToolTipText     =   "右键清空"
      Top             =   2040
      Width           =   5535
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   1
      Left            =   10440
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   2
      Left            =   10440
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   3
      Left            =   10440
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   4
      Left            =   10440
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   5
      Left            =   10440
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   6
      Left            =   10440
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   7
      Left            =   10440
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   8
      Left            =   10440
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   9
      Left            =   10440
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   10
      Left            =   10800
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   11
      Left            =   10800
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   12
      Left            =   10800
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   13
      Left            =   10800
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   14
      Left            =   10800
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   15
      Left            =   10800
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   16
      Left            =   10800
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   17
      Left            =   10800
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   18
      Left            =   10800
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   19
      Left            =   10800
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   20
      Left            =   11160
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   21
      Left            =   11160
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   22
      Left            =   11160
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   23
      Left            =   11160
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   24
      Left            =   11160
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1010
   End
   Begin VB.CheckBox ChangeModel 
      Caption         =   "侦听模式"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "路径："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   5475
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'服务器模式切换
Private Sub ChangeModel_Click()
If ChangeModel.Value = 0 Then
    ChangeModel.Caption = "侦听模式"
    SaveFrame.Enabled = False
    File.Enabled = True
Else
    ChangeModel.Caption = "接收模式"
    DNS_list.Clear
    SaveFrame.Enabled = True
    File.Enabled = False
End If
For i = 0 To 24
    If Server(i).State = 7 Then Server(i).SendData "State:" & IIf(ChangeModel.Value = 0, "侦听中……", "接收中……")
Next
End Sub

'选择文件目录下的域名表
Private Sub ChoiceFile_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
    Open "D:\Program\工程\Visual-basic-6.0\12.hbdk_test\test.txt" For Input As #1
        Do Until Not EOF(1)
            Line Input #1, Temp
            DNS_list.AddItem Temp
        Loop
    Close
End Sub

'允许客户端查询其他账户的次数
Private Sub Count_Num_Change()
Dim a As Integer
a = Val(Count_Num.Caption)
If a <= 0 Then Count_Num.Caption = "1"
If a >= 4 Then Count_Num.Caption = "3"
End Sub

'次数加减
Private Sub Count_Num_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Count_Num.Caption = Count_Num.Caption + 1
If Button = 2 Then Count_Num.Caption = Count_Num.Caption - 1
End Sub

'清理 DNS_List 的内容
Private Sub DNS_list_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And ChangeModel.Value = 0 Then DNS_list.Clear
End Sub

'是否允许客户端查询
Private Sub Extent_Click()
If Extent.Value = 1 Then Count_Num.Enabled = True Else Count_Num.Enabled = False: Count_Num.Caption = "1"
End Sub

Private Sub ReBootPort_Timer()
'端口重启
For i = 0 To Server.Count - 1
    If Not (Server(i).State = 7 Or Server(i).State = 2) Then Server_COde.ReBoot Server(i)
Next
End Sub

'客户端申请列表
Private Sub Request_List_Click()
For i = 0 To Request_List.ListCount - 1
    If Request_List.Selected(i) Then
        Extent.Enabled = True
        Count_Num.Enabled = ture
        Exit For
    Else
        Extent.Enabled = False
        Extent.Value = 0
        Count_Num.Enabled = False
        Count_Num.Caption = "1"
    End If
Next
End Sub

'设置值

'State 属性的设置值是:

'常数                       值 描述
'sckClosed                  0  缺省的。关闭
'sckOpen                    1  打开
'sckListening               2  侦听
'sckConnectionPending       3  连接挂起
'sckResolvingHost           4  识别主机
'sckHostResolved            5  已识别主机
'sckConnecting              6  正在连接
'sckConnected               7  已连接
'sckClosing                 8  同级人员正在关闭连接
'sckError                   9  错误


Private Sub server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，在接受新的连接之前先关闭此连接。
If Server(Index).State <> sckClosed Then Server(Index).Close
'接受具有 requestID 参数的 连接。
Server(Index).Accept requestID
End Sub

Private Sub server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'为进入的数据声明一个变量。
Dim strData As String, Label As String
Dim StrLen As Integer, LabelLen As Integer

'清除原始数据
strData = ""

'获取新的数据
Server(Index).GetData strData

'去空格
strData = Trim(strData)

'获取 strData 变量的长度
StrLen = Len(strData)

'获取 strData 变量中的标签给 Label 变量
Label = Server_COde.GetServerLabel(strData, "<", ">")

'获取 LabelLen 变量的长度
LabelLen = Len(Label)

'获取 strData 变量中的内容（除标签外）
strText = Right(strData, StrLen - LabelLen)

Select Case Label
Case "<State>"
    Server(Index).SendData "State:" & IIf(ChangeModel.Value = 0, "侦听中……", "接收中……")
    GoTo None
Case "<Put>"
    If ChangeModel.Value = 1 Then
        'Put
        DNS_list.AddItem strText
        GoTo Send
    End If
Case "<Get>"
     If ChangeModel.Value = 0 Then
        'Get
        For i = 0 To DNS_list.ListCount - 1
            If Left(Trim(DNS_list.List(i)), 8) = strText Then Server(Index).SendData DNS_list.List(i)
        Next
        GoTo Send
    End If
Case "<Request>"
    'Request
    Request_List.AddItem Server(Index).LocalPort & "端口：" & strText
    GoTo None
End Select
Send:
Call Server_COde.Send_Close(Server(Index))
None:
End Sub

'端口检测报告
Private Sub Command1_Click()
For i = 0 To Server.UBound
    RePort_Data = Public_Code.RePort(Server(i), True, IIf(i = 0, True, False), Server.Count)
Next
MsgBox RePort_Data
End Sub

Private Sub File_Click()
On Error Resume Next
'列表清空
DNS_list.Clear
'读取文件内容放入 DNS_List
Path.Text = File.Path & "\" & File.FileName
Open File.Path & "\" & File.FileName For Input As #1
    Do Until EOF(1)
        Line Input #1, strDNS
        DNS_list.AddItem strDNS
    Loop
Close
End Sub

'文件区刷新
Private Sub File_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then File.Refresh
End Sub

'配置 端口环境 与 软件环境
Private Sub Form_Load()
'端口 1011 - 1035
Dim Port As Integer
Port = 1010
For i = 0 To Server.UBound
    Port = Port + 1
    With Server(i)
        .LocalPort = Port
        .RemotePort = 1010
       ' .LocalHostName = HostName
    End With
    Server_COde.ReBoot Server(i)
Next
File.Path = App.Path & "\DNS-form"
Path.Text = File.Path
End Sub

'盘符路径
Private Sub Path_Change()
File.Path = GetDOS.GetPath(Path.Text)
End Sub


Private Sub GetNameList(ByVal File As String, NameList() As String)
    
    Dim FirstName As String, FirstNameList() As String, List() As String
    Dim FirstNameCount As Integer, Line As Integer, _
        NameAdd As Integer, FirstNameAdd As Integer
    Dim IndexAdd As Long, Columns As Integer
    
    '报错跳到 Add 标签
    On Error GoTo Add
    '获取空闲文件号
    FileNum = FreeFile()
    '打开文件
    Open File For Input As #FileNum
        '当 指针 指向文件尾部时 结束循环
        Do Until EOF(FileNum)
        '提取姓氏
            '给 List数组 拓展空间，并保存 List数组 中的值
            ReDim Preserve List(NameAdd)
            '把文件中的姓名存入 List数组 中
            Line Input #FileNum, List(NameAdd)
            '给 FirstNameList数组 拓展空间，并保存 FirstNameList数组 中的值
            ReDim Preserve FirstNameList(FirstNameAdd)
            '获取姓氏
            FirstName = Left(List(NameAdd), 1)
            '历遍 FirstNameList数组 中的姓氏
            For i = LBound(FirstNameList) To UBound(FirstNameList)
                '如果 FirstName 等于 FirstNameList数组 中的其中一个，则退出循环，否则 FirstNameCount 变量值 + 1，并且存入新的姓氏
                If FirstName = FirstNameList(i) Then Exit For Else FirstNameCount = FirstNameCount + 1: FirstNameList(UBound(FirstNameList)) = FirstName
            Next
            '如果 FirstNameCount 变量 等于 FirstNameList数组 个数，则 FirstNameAdd 变量值 + 1
            If FirstNameCount = UBound(FirstNameList) Then FirstNameAdd = FirstNameAdd + 1
            '把 FirstNameCount 变量值清除为 0，并为 NameAdd 变量值 + 1
            FirstNameCount = 0: NameAdd = NameAdd + 1
        Loop
    '关闭文件
    Close #FileNum
    
    '获取 FirstNameList数组 元素个数 赋入 Line 变量
    Line = UBound(FirstNameList) - LBound(FirstNameList) + 1
    
    '给 NameList数组 拓展空间 成为二维数组，并保存 NameList数组 中的值
    ReDim Preserve NameList(Line, IndexAdd)
    Columns = IndexAdd + 1
    
    '把 FirstNameList数组 所有的元素传到 NameList数组 最左列
    For i = LBound(FirstNameList) To UBound(FirstNameList)
        NameList(i, 0) = FirstNameList(i)
    Next
    
    '历遍 List数组 中的 姓名
    For i = LBound(List) To UBound(List)
        '历遍 NameList数组 中的 姓氏
        For j = LBound(NameList) To UBound(NameList)
            '如果 List数组元素 中的姓氏 等于 NameList数组(j,0) 中的其中一个，则执行下列 for语句
            If Left(List(i), 1) = NameList(j, 0) Then
back:
                '执行中报错 NameList数组 下标越界，则 执行 报错拓展法 (转到 add 标签)
                For k = 1 To Columns
                    '如果有空位，则插入 NameList数组 中，并且退出 for循环 (第三层)
                    If NameList(j, k) = "" Then NameList(j, k) = List(i): Exit For
                Next k
                '插入完成 跳出第二层循环
                Exit For
            End If
        Next j
    Next i
    
    GoTo Go:
    '报错拓展法
Add:
    '让列加一
    IndexAdd = IndexAdd + 1
    ReDim Preserve NameList(Line, IndexAdd)
    '变量 Columns 用于判定 列空间 是否够用
    '若 列空间 不足则执行 报错拓排法
    Columns = IndexAdd + 1
    '恢复 报错处理机制
    Resume
    GoTo back
Go:

' debug 输出
For i = LBound(NameList) To UBound(NameList)
    For j = LBound(NameList, 2) To UBound(NameList, 2)
        Debug.Print NameList(i, j) & "|";
    Next
    Debug.Print
Next
End Sub

'制作域名表
Private Sub Save_DNS_Click()
Dim NameList() As String
Dim Name As String
If Not List_Jion_File.Text = "" Then
    If MsgBox("系统将列表与" & List_Jion_File.Text & "文件合并成域名表，是否保存？", vbYesNo, "保存") = 6 Then
        DNS_name = InputBox("请输入文件名。" & Chr(10) & "不能与" & File.Path & "下内文件重名，否则覆盖！！！", "保存")
        If DNS_name <> "" Then
            '打开用户指定的文件
            Open File.Path & "\" & DNS_name & ".txt" For Output As #1
                '调用 GetNamelist 子程序
                Call GetNameList(List_Jion_File.Text, NameList())
                For k = 0 To DNS_list.ListCount - 1
                    '提取名称
                    Name = Right(DNS_list.List(k), 4)
                    For i = LBound(NameList) To UBound(NameList)
                        '进行姓氏判断
                        If NameList(i, 0) = Left(Trim(Name), 1) Then
                            '查找名称，并写入文件
                            For j = LBound(NameList, 2) To UBound(NameList, 2)
                                If Trim(Name) = Trim(Left(NameList(i, j), 3)) Then Print #1, DNS_list.List(k) & Right(NameList(i, j), Len(NameList(i, j)) - 3): Exit For
                            Next
                            Exit For
                        End If
                    Next
                Next
            Close
        Else
            MsgBox "文件名为空！"
        End If
    End If
Else
    MsgBox "未指定文件！"
End If
End Sub

'申请查询结果与次数返回客户端
Private Sub Send_Click()
Dim Index As Integer
For i = 0 To Request_List.ListCount - 1
    If Request_List.Selected(i) Then
        RequestText = Request_List.List(i)
        a = InStr(RequestText, "端口") - 1
        Index = Val(Mid(RequestText, 1, a)) - 1011
        If Server(Index).State = 7 Then Server(Index).SendData "Return:" & Extent.Value & "Count:" & Count_Num.Caption
        Request_List.RemoveItem i
        Call Server_COde.Send_Close(Server(Index))
        Exit For
    End If
Next
Extent.Enabled = False
Extent.Value = 0
Count_Num.Enabled = False
Count_Num.Caption = "1"
End Sub

Private Sub Write_Client_ini_Click()

'获取 Server 名称
Shell ("cmd.exe /c hostname > " & App.Path & "\Name.txt")
ServerHost = Public_Code.GetHostName

'获取 Server IPv4
Shell ("cmd.exe /c ipconfig > " & App.Path & "\Name.txt")
Open App.Path & "\IP.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, ServerIP
        If CBool(InStr(1, ServerIP, "IPv4")) Then ServerIP = Mid(ServerIP, InStr(1, ServerIP, ": ") + 2): Exit Do
    Loop
Close #1

'写入
Open "配置.csini" For Output As #1
    Print #1, "[+Server-]"
    Print #1, "   <Name>" & ServerHost
    Print #1, "   <IP>" & ServerIP
    Print #1, "[-Server+]"
Close #1
End Sub
