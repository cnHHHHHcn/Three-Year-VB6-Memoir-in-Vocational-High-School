VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3765
   ClientLeft      =   7260
   ClientTop       =   7020
   ClientWidth     =   6630
   LinkTopic       =   "Form4"
   ScaleHeight     =   3765
   ScaleWidth      =   6630
   Begin VB.Timer GameReady 
      Interval        =   1000
      Left            =   6120
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Timer StateCheck 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   3360
   End
   Begin VB.Frame Frame2 
      Caption         =   "Player2"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
      Begin VB.CheckBox Player2_State 
         Height          =   495
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   200
      End
      Begin VB.Label Player2_Name 
         AutoSize        =   -1  'True
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
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player1"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      Begin VB.CheckBox Player1_State 
         Height          =   495
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   200
      End
      Begin VB.Label Player1_Name 
         AutoSize        =   -1  'True
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   165
      End
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   5400
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label House_Owner_Name 
      Caption         =   "Label3"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label House_Owner_Check 
      Caption         =   "Label3"
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "将在5秒后，进入游戏"
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
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label State 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1200
      TabIndex        =   7
      Top             =   3360
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "状态："
      Height          =   180
      Left            =   600
      TabIndex        =   6
      Top             =   3360
      Width           =   540
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim House_Owner As Boolean

Sub Del_Room()
'向游戏服务器请求删除当前房间名
Form1.SendGameServerData.Caption = "Del"
Form1.GameServerStateCheck.Enabled = True
Form1.ConnectGameServer.Enabled = True
End Sub

Sub GamePlay_House_Owner_T()
Form3.Send_GameData.Close
Form3.Send_GameData.LocalPort = 1021
Form3.Send_GameData.Listen
End Sub

Sub GamePlay_House_Owner_F(House_Owner_HostName As String)
Form3.Send_GameData.Close
Form3.Send_GameData.Connect House_Owner_HostName, 1021
End Sub

'重启
Private Sub ReBoot()
Server.Close
Server.Listen
End Sub

Private Sub Command1_Click()
On Error Resume Next
Form1.Show: Form4.Hide

Del_Room

Player1_Name.Caption = ""
Player1_State.Value = 0
Player2_Name.Caption = ""
Player2_State.Value = 0

'向对方发送离开 标识Quit
If House_Owner Then Server.SendData "Quit" Else Form1.Client.SendData "Quit"
End Sub

Private Sub Form_Activate()
'房间成员标识
House_Owner = CBool(Val(House_Owner_Check.Caption))
Command1.Enabled = True
End Sub

Private Sub GameReady_Timer()
On Error Resume Next
Static Sec As Integer
'如果双方准备就绪，则进入就绪阶段
If CBool(Player1_State.Value) And CBool(Player2_State.Value) Then
    Frame1.Enabled = False
    Frame2.Enabled = False
    Command1.Enabled = False
    StateCheck.Enabled = False
    Del_Room
    Label1.Visible = True
    Sec = Sec + 1
    Label1.Caption = "将在" & (5 - Sec) & "秒后，进入游戏"
    If House_Owner Then
        GamePlay_House_Owner_T
        Form3.Player(1).Caption = Form1.User.Caption
    Else
        GamePlay_House_Owner_F (House_Owner_Name.Caption)
    End If
    
    
    '进入游戏界面
    If (5 - Sec) = 0 Then Form4.Hide: Form3.Show: GameReady.Enabled = False
    Server.Close: Form1.Client.Close
End If
End Sub

'------------------------游戏玩家当前状态---------------------------
'*******************************************************************
Private Sub Player1_State_Click()
If Server.State = 7 Then
    If Player1_State.Value = 1 Then
        Server.SendData "1"
    Else
        Server.SendData "0"
    End If
End If
End Sub

Private Sub Player2_State_Click()
If Form1.Client.State = 7 Then
    If Player2_State.Value = 1 Then
        Form1.Client.SendData "1"
    Else
        Form1.Client.SendData "0"
    End If
End If
End Sub
'*******************************************************************

Private Sub State_Change()
Select Case State.Caption
Case "4链接成功"
    Server.SendData Player1_Name.Caption
    Form4.Frame1.Enabled = True
    Form4.Frame2.Enabled = False
Case "1链接成功"
    Form4.Frame1.Enabled = False
    Form4.Frame2.Enabled = True
End Select
End Sub

Private Sub StateCheck_Timer()
If Server.State = 0 Then State.Caption = "4关闭状态"
If Server.State = 1 Then State.Caption = "4打开"
If Server.State = 2 Then State.Caption = "4正在侦听"
If Server.State = 3 Then State.Caption = "4连接挂起"
If Server.State = 4 Then State.Caption = "4识别主机"
If Server.State = 5 Then State.Caption = "4已识别主机"
If Server.State = 6 Then State.Caption = "4正在连接"
If Server.State = 7 Then State.Caption = "4链接成功"
If Server.State = 8 Then State.Caption = "4同级人员正在关闭连接"
If Server.State = 9 Then State.Caption = "4链接错误"
If Server.State <> 7 Then ReBoot
End Sub

Private Sub Server_ConnectionRequest(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If Server.State <> sckClosed Then _
Server.Close
'接受具有 requestID 参数的
'连接。
Server.Accept requestID
End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
'为进入的数据声明一个变量。
Dim Data As String
Data = ""
Server.GetData Data
'玩家2退出
If Len(Data) = 4 And Data = "Quit" Then
    Player2_Name.Caption = ""
    Player2_State.Value = 0
Else
    '游戏准备阶段玩家2状态
    If Data = "1" Then
        Player2_State.Value = 1
    ElseIf Data = "0" Then
        Player2_State.Value = 0
    Else
        Player2_Name.Caption = Data
    End If
End If
End Sub

