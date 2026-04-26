VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   8100
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Send_Order 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   4560
   End
   Begin VB.CommandButton Boot 
      Caption         =   "连接 Server"
      Height          =   495
      Left            =   1920
      TabIndex        =   19
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Request_Block 
      Caption         =   "申请结果"
      Height          =   735
      Left            =   2160
      TabIndex        =   15
      Top             =   480
      Width           =   1455
      Begin VB.Label Request_Return 
         AutoSize        =   -1  'True
         Caption         =   "查询：None"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Number 
         AutoSize        =   -1  'True
         Caption         =   "None"
         Height          =   180
         Left            =   660
         TabIndex        =   16
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "次数："
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.Frame Get_And_Put 
      Caption         =   "获取与申请"
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   3495
      Begin VB.CommandButton GetU_P 
         Caption         =   "获取"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Request 
         Caption         =   "申请"
         Height          =   495
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Request_Text 
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
         Height          =   435
         Left            =   960
         TabIndex        =   12
         ToolTipText     =   "输入主机名"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Copy_Text 
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Port_Report 
      Caption         =   "端口检测报告"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Return_Text 
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Send 
      Caption         =   "发送"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   3495
      Begin VB.TextBox Text 
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
         Left            =   960
         TabIndex        =   2
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton Send_text 
         Caption         =   "发送"
         Height          =   495
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Timer StateCheck 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   4320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Order 
      AutoSize        =   -1  'True
      Caption         =   "None"
      Height          =   180
      Left            =   720
      TabIndex        =   21
      Top             =   4440
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "指令："
      Height          =   180
      Left            =   240
      TabIndex        =   20
      Top             =   4440
      Width           =   540
   End
   Begin VB.Label ServerState 
      AutoSize        =   -1  'True
      Caption         =   "服务器状态：None"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Width           =   1440
   End
   Begin VB.Label ServerIP 
      AutoSize        =   -1  'True
      Caption         =   "<IP>"
      Height          =   180
      Left            =   1080
      TabIndex        =   7
      Top             =   3960
      Width           =   360
   End
   Begin VB.Label CopyIP 
      AutoSize        =   -1  'True
      Caption         =   "服务器IP："
      Height          =   180
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "点击复制服务器IP"
      Top             =   3960
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Port As Integer
Dim UserHost As String, ServerHost As String, strData As String

Private Sub Boot_Click()
StateCheck.Enabled = True
Boot.Enabled = False
End Sub

Private Sub Command1_Click()
MsgBox OperaKey.ClassList(App.Path & "\配置.csini")
End Sub

'复制 Server IP
Private Sub CopyIP_Click()
Set a = CreateObject("wscript.shell")
Copy_Text.Text = ServerIP.Caption
Copy_Text.SetFocus
Copy_Text.SelLength = Len(Copy_Text.Text)
a.SendKeys "^c"
End Sub

'配置 端口环境 与 程序环境
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_Activate()
Copy_Text.Top = -1000
Copy_Text.Left = -1000
UserHost = Public_Code.GetHostName
Client.RemoteHost = UserHost
Client.LocalPort = 1010
Client.RemotePort = 1011
Client_Code.ReBoot Client, "None", 1011

ServerHost = OperaKey.GetKey(App.Path & "\配置.csini", "Server", "Name")
ServerIP.Caption = OperaKey.GetKey(App.Path & "\配置.csini", "Server", "IP")
End Sub

Private Sub Form_Load()
Shell ("cmd.exe /c hostname > " & App.Path & "\Name.txt")
Open App.Path & "\配置.csini" For Input As #1
    Do Until EOF(1)
        Line Input #1, StrText
        Position = InStr(1, StrText, ">")
        If Position <> 0 Then ServerHost = Right(StrText, Len(StrText) - Position)
    Loop
Close
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------

'端口检测报告
Private Sub Port_Report_Click()
MsgBox Public_Code.RePort(Client, False, True, 1)
End Sub

'申请查看其它账户
Private Sub Request_Click()
Order.Caption = "Request"
StateCheck.Enabled = True
Send_Order.Enabled = True
End Sub

Private Sub Send_Order_Timer()

If Client.State = 7 Then

    Select Case Order.Caption
    Case "Get"
        strData = "<Get>" & Request_Text.Text
        If Request_Text.Enabled Then
            If Number.Caption > 0 And Request_Text <> "" Then
                Client.SendData strData: Number.Caption = Number.Caption - 1: Request_Text.Text = ""
                If Number.Caption = "0" Then Number.Caption = "None": Request_Return.Caption = "查询：None"
            End If
        Else
            Client.SendData "<Get>" & UserHost
        End If
        
    Case "Request"
        RequestName = InputBox("请输入姓名:", "Client")
        strData = "<Request>  " & UserHost & " " & RequestName & "申请查看其它电脑的河北对口账户"
        Client.SendData strData: Request_Block.Caption = "申请中……"
        
    Case "Put"
        strData = "<Put>" & UserHost & "    " & Text.Text
        Client.SendData strData: Text.Text = ""
        
    End Select
    Order.Caption = "None"
    Send_Order.Enabled = False
End If

End Sub

'连接 Server
Private Sub StateCheck_Timer()
If Client.State = 7 Then StateCheck.Enabled = False: GoTo ends
If Not (1010 < Port And Port < 1035) Then Port = 1010
If Not (Client.State = 7 Or Client.State = 2) Then Port = Port + 1: ServerState.Caption = "服务器状态：None": Client_Code.ReBoot Client, ServerHost, Port
ends:
End Sub

'当连接成功时，获取 Server 状态
Private Sub Client_Connect()
Client.SendData "<State>"
End Sub

'接收 Server
Private Sub Client_DataArrival(ByVal bytesTotal As Long)


Dim strData As String

Client.GetData strData
strData = Trim(strData)

'含有 Close ，则与 Server 断开
If CBool(InStr(strData, "Close")) Then Client.Close: ServerState.Caption = "服务器状态：None"

On Error Resume Next
'获取Server标签
Position = InStr(1, strData, ":")
Label = Left(strData, Position - 1)

'进行匹配
Select Case Label
Case "State"
    '显示现在 Server 状态
    State = Mid(strData, Position + 1)
    ServerState.Caption = "服务器状态：" & State
    If State = "侦听中……" Then
        Get_And_Put.Enabled = True
        Send.Enabled = False
    ElseIf State = "接收中……" Then
        Get_And_Put.Enabled = False
        Send.Enabled = True
    End If
    GoTo ends
Case "Return"
    '显示申请结果
    ReturnStr = Mid(strData, Position + 1)
    Request_Text.Enabled = CBool(Val(Left(ReturnStr, 1)))
    Request_Return.Caption = "查询：" & IIf(Request_Text.Enabled, "True", "False")
    Request_Block.Caption = "申请结果"
    Number.Caption = Mid(ReturnStr, 8, 1)
    GoTo ends
End Select

'返回账户
Return_Text.Text = Left(strData, InStr(1, strData, "Close") - 1)
ends:
End Sub

'向 Server 获取账户信息
Private Sub GetU_P_Click()
Order.Caption = "Get"
StateCheck.Enabled = True
Send_Order.Enabled = True
Get_And_Put.Enabled = False
Boot.Enabled = True
End Sub

'向 Server 发送 名称与主机名
Private Sub Send_text_Click()
Order.Caption = "Put"
StateCheck.Enabled = True
Send_Order.Enabled = True
Send.Enabled = False
Boot.Enabled = True
End Sub

