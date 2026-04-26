VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   10605
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer GameServerStateCheck 
      Interval        =   10
      Left            =   1920
      Top             =   240
   End
   Begin VB.Timer ConnectGameServer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   720
   End
   Begin MSWinsockLib.Winsock GameClient 
      Left            =   1920
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer ConnectServer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   7920
   End
   Begin VB.Timer StateCheck 
      Interval        =   10
      Left            =   1800
      Top             =   7440
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   1800
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Create_Server 
      Caption         =   "创建房间"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox PlayerList 
      Appearance      =   0  'Flat
      Height          =   5250
      ItemData        =   "Form1.frx":0000
      Left            =   360
      List            =   "Form1.frx":0002
      TabIndex        =   7
      ToolTipText     =   "右键刷新"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "选定房间"
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton Join 
         Caption         =   "申请"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "服务器连接状态："
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label SendGameServerData 
      Caption         =   "Label6"
      Height          =   495
      Left            =   3480
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label GameClientState 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "房间连接状态："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   8520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label State 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      TabIndex        =   13
      Top             =   8760
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "房间："
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
      Left            =   480
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2355
      _cy             =   661
   End
   Begin VB.Label User 
      AutoSize        =   -1  'True
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
      Left            =   8280
      TabIndex        =   9
      Top             =   360
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户："
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
      Left            =   7080
      TabIndex        =   8
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公告"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   3
      Left            =   8280
      TabIndex        =   3
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   2
      Left            =   8280
      TabIndex        =   2
      Top             =   5760
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "连机模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   1
      Left            =   7440
      TabIndex        =   1
      Top             =   2880
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单机模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   0
      Left            =   7440
      TabIndex        =   0
      Top             =   1560
      Width           =   2100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim indexs, Port As Integer

'重启
Sub ReBoot()
Client.Close
Client.Connect Text3.Text, 1021
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Data = ""
Client.GetData Data
'房主退出
If Len(Data) = 4 And Data = "Quit" Then
    Form4.Player1_Name.Caption = ""
    Form4.Player1_State.Value = 0
    Form4.Player2_Name.Caption = ""
    Form4.Player2_State.Value = 0
    Form4.Hide: Form1.Show
    MsgBox "房主已解散房间。", 64
Else
    '游戏准备阶段玩家1状态
    If Data = "1" Then
        Form4.Player1_State.Value = 1
    ElseIf Data = "0" Then
        Form4.Player1_State.Value = 0
    Else
        Form4.Player1_Name = Data
    End If
End If

End Sub

Private Sub ConnectGameServer_Timer()
'链接游戏服务器端口（1011 - 1020）
If Not (1010 < Port And Port < 1020) Then Port = 1010
If GameClient.State = 9 Or GameClient.State = 0 Then Port = Port + 1: GameClient_ReBoot

If GameClient.State = 7 Then
    Select Case SendGameServerData.Caption
    '向服务器发送本机名
    Case "Give"
        GameClient.SendData Label5.Caption
    '向服务器发送删除本机名命令
    Case "Del"
        GameClient.SendData "<DEL>" & Trim(Label5.Caption)
    End Select
    '恢复原始状态
    SendGameServerData.Caption = "None"
    ConnectGameServer.Enabled = False
End If
End Sub

'链接服务器
Sub GameClient_ReBoot()
GameClient.Close
GameClient.RemotePort = Port
GameClient.Connect "admin015"
End Sub

'链接房间
Private Sub ConnectServer_Timer()
If Client.State = 7 Then
    Form4.Show: Form1.Hide
    Form4.Player2_Name.Caption = User.Caption
    Form4.StateCheck.Enabled = False
    Form4.Server.Close
    Client.SendData User.Caption
    ConnectServer.Enabled = False
Else
    ReBoot
End If
End Sub


Private Sub Create_Server_Click()
'向游戏服务器发送房间名
SendGameServerData.Caption = "Give"
ConnectGameServer.Enabled = True
GameServerStateCheck.Enabled = True
StateCheck.Enabled = False
Form4.StateCheck.Enabled = True
'标识为房主
Form4.House_Owner_Check.Caption = "1"
'房间初始化
Form4.Server.Close
Form4.Server.LocalPort = 1021
Form4.Server.Listen
Form4.Show: Form1.Hide
Form4.Player1_Name.Caption = User.Caption
End Sub

Private Sub Form_Load()
'获取游戏名
Start:
Namestr = InputBox("请输入名称：" & Chr(10) & "按 1 退出", "Game")

If Namestr = "1" Then End
If Namestr = "" Then GoTo Start Else User.Caption = Namestr

Shell ("cmd.exe /c hostname > " & App.Path & "\Name.txt")
Users = GetHostName()


'!!!!!!!!!!!!编写完成后，要删掉!!!!!!!!!!!!
Client.RemoteHost = User
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'游戏端口 1021
Client.RemotePort = 1021
End Sub

'获取本机名称函数
Function GetHostName() As String
Open App.Path & "\Name.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, HostName
    Loop
Close
Label5.Caption = HostName
GetHostName = HostName
End Function

'UI特效
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(indexs).FontSize = 28
Label1(indexs).ForeColor = vbBlack
End Sub

Private Sub GameClient_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
strData = ""
GameClient.GetData strData
'若游戏服务器发送 "Close" ，则关闭与服务器链接
If strData = "Close" Then GoTo end_sub
'去空格
strData = Trim(strData)
'把房间数据存入 strArray数组 中
strarray = Split(strData, "|")
'刷新房间列表
For i = LBound(strarray) To UBound(strarray)
    If strarray(i) <> "" And strarray(i) <> "Close" Then PlayerList.AddItem strarray(i)
Next
end_sub:
GameClient.Close
End Sub

Private Sub GameServerStateCheck_Timer()
'GameClientState
If GameClient.State = 0 Then GameClientState.Caption = "关闭状态"
If GameClient.State = 1 Then GameClientState.Caption = "打开"
If GameClient.State = 2 Then GameClientState.Caption = "正在侦听"
If GameClient.State = 3 Then GameClientState.Caption = "连接挂起"
If GameClient.State = 4 Then GameClientState.Caption = "识别主机"
If GameClient.State = 5 Then GameClientState.Caption = "已识别主机"
If GameClient.State = 6 Then GameClientState.Caption = "正在连接"
If GameClient.State = 7 Then GameClientState.Caption = "链接成功"
If GameClient.State = 8 Then GameClientState.Caption = "同级人员正在关闭连接"
If GameClient.State = 9 Then GameClientState.Caption = "链接错误"
If GameClient.State = 7 And SendGameServerData.Caption = "Get" Then PlayerListRefresh: PlayerList.Enabled = True: SendGameServerData.Caption = "None"
End Sub

Private Sub Join_Click()
'加入房间
Form4.House_Owner_Name.Caption = Text3.Text
Client.Close
Client.Connect Text3.Text, 1021
ConnectServer.Enabled = True
'标识为房间成员
Form4.House_Owner_Check.Caption = "0"
End Sub

'模式选择
Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 0
    Form1.Hide: Form2.Show
    Frame1.Visible = False: PlayerList.Visible = False: Create_Server.Visible = False: Label3.Visible = False
    For i = 0 To 1: Label4(i).Visible = False: Next
    GameClientState.Visible = False: State.Visible = False
Case 1
    Frame1.Visible = True: PlayerList.Visible = True: Create_Server.Visible = True: Label3.Visible = True
    For i = 0 To 1: Label4(i).Visible = True: Next
    GameClientState.Visible = True: State.Visible = True
Case 2
    End
Case 3
    Form1.Hide: Form5.Show
End Select
End Sub

'UI特效
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
indexs = Index
Label1(indexs).FontSize = 30
Label1(indexs).ForeColor = vbRed
End Sub

'选择房间
Private Sub PlayerList_Click()
Text3.Text = PlayerList.List(PlayerList.ListIndex)
End Sub

'右键刷新房间列表
Private Sub PlayerList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    SendGameServerData.Caption = "Get"
    PlayerList.Enabled = False
    PlayerListRefresh
End If
End Sub

'房间刷新
Sub PlayerListRefresh()
PlayerList.Clear
ConnectGameServer.Enabled = True
If GameClient.State = 7 Then GameClient.SendData "GetRoom": ConnectGameServer.Enabled = False
End Sub

Private Sub StateCheck_Timer()
'form1
If Client.State = 0 Then State.Caption = "1关闭状态"
If Client.State = 1 Then State.Caption = "1打开"
If Client.State = 2 Then State.Caption = "1正在侦听"
If Client.State = 3 Then State.Caption = "1连接挂起"
If Client.State = 4 Then State.Caption = "1识别主机"
If Client.State = 5 Then State.Caption = "1已识别主机"
If Client.State = 6 Then State.Caption = "1正在连接"
If Client.State = 7 Then State.Caption = "1链接成功"
If Client.State = 8 Then State.Caption = "1同级人员正在关闭连接"
If Client.State = 9 Then State.Caption = "1链接错误"

'form4
If Client.State = 0 Then Form4.State.Caption = "1关闭状态"
If Client.State = 1 Then Form4.State.Caption = "1打开"
If Client.State = 2 Then Form4.State.Caption = "1正在侦听"
If Client.State = 3 Then Form4.State.Caption = "1连接挂起"
If Client.State = 4 Then Form4.State.Caption = "1识别主机"
If Client.State = 5 Then Form4.State.Caption = "1已识别主机"
If Client.State = 6 Then Form4.State.Caption = "1正在连接"
If Client.State = 7 Then Form4.State.Caption = "1链接成功"
If Client.State = 8 Then Form4.State.Caption = "1同级人员正在关闭连接"
If Client.State = 9 Then Form4.State.Caption = "1链接错误"
End Sub

Private Sub Text3_Change()
Join.Enabled = True
Join.SetFocus
End Sub

