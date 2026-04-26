VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   8790
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Append_Text 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8160
      Top             =   2280
   End
   Begin VB.Timer Server_Listening 
      Interval        =   100
      Left            =   8280
      Top             =   120
   End
   Begin VB.CommandButton Send 
      Caption         =   "send"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "测试链接"
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   10920
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   2895
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   7920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "admin008"
   End
   Begin VB.Label Label2 
      Caption         =   "接收端"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "发送端"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
'设置值

'State 属性的设置值是:

'常数 值 描述
'sckClosed                  0 缺省的。关闭
'sckOpen                    1 打开
'sckListening               2 侦听
'sckConnectionPending       3 连接挂起
'sckResolvingHost           4 识别主机
'sckHostResolved            5 已识别主机
'sckConnecting              6 正在连接
'sckConnected               7 已连接
'sckClosing                 8 同级人员正在关闭连接
'sckError                   9 错误


If Server.State = 0 Then MsgBox "关闭状态"
If Server.State = 1 Then MsgBox "打开"
If Server.State = 2 Then MsgBox "正在侦听"
If Server.State = 3 Then MsgBox "连接挂起"
If Server.State = 4 Then MsgBox "识别主机"
If Server.State = 5 Then MsgBox "已识别主机"
If Server.State = 6 Then MsgBox "正在连接"
If Server.State = 7 Then MsgBox "链接成功"
If Server.State = 8 Then MsgBox "同级人员正在关闭连接"
If Server.State = 9 Then MsgBox "未知错误"


End Sub

Private Sub ReBoot()
Server.Close
Server.Listen
End Sub

Private Sub Send_Click()
If Server.State = 7 Then
    Server.SendData Text1.Text & " "
End If
End Sub

Private Sub Form_Load()
Server.LocalPort = 1010
Server.Listen
End Sub

Private Sub server_ConnectionRequest(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If Server.State <> sckClosed Then _
Server.Close
'接受具有 requestID 参数的
'连接。
Server.Accept requestID
End Sub

Private Sub server_DataArrival(ByVal bytesTotal As Long)
'为进入的数据声明一个变量。
'调用 GetData 方法，并将数据赋予名为 txtOutput
'的 TextBox 的 Text 属性。
Dim strData As String
Server.GetData strData
Text2.Text = ""
Text2.Text = strData
End Sub

Private Sub Server_Listening_Timer()
If Server.State <> 7 Then Call ReBoot
End Sub

Private Sub Text2_Change()
Name = Text2.Text
Open App.Path & "\房间列表.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, a
        If a = Name Then Exit Sub
    Loop
Close


Open App.Path & "\房间列表.txt" For Append As #1
    Print #1, Text2.Text
Close
End Sub
