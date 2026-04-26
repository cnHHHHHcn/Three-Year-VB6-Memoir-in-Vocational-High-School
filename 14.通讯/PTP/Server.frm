VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Server 
   Caption         =   "Server"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8535
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Dispose 
      Caption         =   "配置"
      Height          =   2415
      Left            =   5520
      TabIndex        =   5
      Top             =   600
      Width           =   2655
      Begin VB.CommandButton SockOpen 
         Caption         =   "启动"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton SockClose 
         Caption         =   "关闭"
         Height          =   495
         Left            =   1440
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox RemotePortText 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   13
         Text            =   "10001"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox LocalHostName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox RemoteHostName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton SetPort 
         Caption         =   "确定"
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox LocalPortText 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   6
         Text            =   "10001"
         Top             =   200
         Width           =   855
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   2640
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "远程端口："
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   525
         Width           =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         X1              =   1320
         X2              =   1320
         Y1              =   840
         Y2              =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "本地电脑名："
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   2640
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "远程电脑名："
         Height          =   180
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "本地端口："
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   900
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6600
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Save 
      Caption         =   "保存"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6120
      Top             =   3360
   End
   Begin VB.TextBox DataText 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
   Begin VB.TextBox SendText 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3960
      Width           =   5055
   End
   Begin VB.CommandButton Send 
      Caption         =   "发送"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   5520
      TabIndex        =   3
      Top             =   3480
      Width           =   540
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Send_Click()
    If Winsock1.State = 7 Or SendText.Text = "" Then _
        Winsock1.SendData SendText.Text: _
        DataText.Text = DataText.Text & vbCrLf & LocalHostName.Text & ":" & vbCrLf & SendText.Text: _
        SendText.Text = ""
End Sub

Private Sub Save_Click()
    PublicMethod.SaveData DataText.Text
End Sub

Private Sub Form_Load()
Query:
    Flag = InputBox("请选择模式：" & Chr(13) & "C 为 Client(客户端)" & Chr(13) & "S 为 Server(服务端)")
    If LCase(Flag) = "s" Then
        Server.Show: Client.Hide
        Winsock1.LocalPort = Val(LocalPortText.Text)
        Winsock1.Close
        Winsock1.Listen
        PublicMethod.ServerDefaultDispose
    ElseIf LCase(Flag) = "c" Then
        Server.Hide: Client.Show
        PublicMethod.ClientDefaultDispose
    Else
        GoTo Query
    End If
End Sub

Private Sub SetPort_Click()
    Winsock1.Close
    Port = Val(LocalPortText.Text)
    If 1 <= Port And Port <= 2 ^ 16 - 1 Then Winsock1.LocalPort = Port: MsgBox "本地端口设置成功!", 64 Else MsgBox "本地端口未设置成功。", 48
End Sub

Private Sub SockOpen_Click()
    Winsock1.Close
    Winsock1.Listen
End Sub

Private Sub SockClose_Click()
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    Winsock1.SendData "<Host>" & LocalHostName.Text
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If Winsock1.State <> sckClosed Then Winsock1.Close
'接受具有 requestID 参数的
'连接。
Winsock1.Accept requestID
End Sub

Private Sub Label1_Click()
    Winsock1.Close
    Winsock1.Listen
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = "Server状态:" & PublicMethod.GetState(Winsock1.State)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    Winsock1.GetData Data
    If Left(Data, 6) = "<Host>" Then RemoteHostName.Text = Mid(Data, 7): Exit Sub
    DataText.Text = DataText.Text & vbCrLf & RemoteHostName.Text & ":" & vbCrLf & Data
    DataText.SelStart = Len(DataText.Text)
End Sub



