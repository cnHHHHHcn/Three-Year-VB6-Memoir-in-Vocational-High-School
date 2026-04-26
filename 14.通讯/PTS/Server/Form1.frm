VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   12210
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox ComputerIP 
      Height          =   495
      Left            =   3960
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Timer ServerClosed 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5400
      Top             =   4800
   End
   Begin VB.CommandButton Save 
      Caption         =   "保存"
      Height          =   495
      Left            =   3960
      TabIndex        =   39
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Timer PortReboot 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7920
      Top             =   240
   End
   Begin VB.Timer GetPortState 
      Interval        =   500
      Left            =   7440
      Top             =   240
   End
   Begin VB.Frame Dispose 
      Caption         =   "配置"
      Height          =   1335
      Left            =   240
      TabIndex        =   23
      Top             =   4680
      Width           =   3375
      Begin VB.CommandButton SockClose 
         Caption         =   "关闭"
         Height          =   375
         Left            =   2040
         TabIndex        =   41
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton SockOpen 
         Caption         =   "启动"
         Height          =   375
         Left            =   2040
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox ComputerName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   600
         TabIndex        =   27
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton SetPort 
         Caption         =   "确定"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox PortText 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   600
         TabIndex        =   25
         Text            =   "10001"
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   1800
         X2              =   1800
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "名称："
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   645
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "端口："
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   290
         Width           =   540
      End
   End
   Begin VB.Frame UserCollection 
      Caption         =   "服务器在线人员："
      Height          =   4215
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   6255
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   9
         Left            =   1440
         TabIndex        =   12
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   8
         Left            =   1440
         TabIndex        =   11
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   7
         Left            =   1440
         TabIndex        =   10
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   6
         Left            =   1440
         TabIndex        =   9
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   5
         Left            =   1440
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   1440
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   1440
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox UserText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   9
         Left            =   4320
         TabIndex        =   52
         Top             =   3645
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   8
         Left            =   4320
         TabIndex        =   51
         Top             =   3285
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   7
         Left            =   4320
         TabIndex        =   50
         Top             =   2925
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   6
         Left            =   4320
         TabIndex        =   49
         Top             =   2565
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   5
         Left            =   4320
         TabIndex        =   48
         Top             =   2205
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   4
         Left            =   4320
         TabIndex        =   47
         Top             =   1845
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   3
         Left            =   4320
         TabIndex        =   46
         Top             =   1485
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   2
         Left            =   4320
         TabIndex        =   45
         Top             =   1125
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   1
         Left            =   4320
         TabIndex        =   44
         Top             =   765
         Width           =   270
      End
      Begin VB.Label IPList 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Index           =   0
         Left            =   4320
         TabIndex        =   43
         Top             =   405
         Width           =   270
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   9
         Left            =   2520
         TabIndex        =   38
         Top             =   3650
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   8
         Left            =   2520
         TabIndex        =   37
         Top             =   3290
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   7
         Left            =   2520
         TabIndex        =   36
         Top             =   2930
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   6
         Left            =   2520
         TabIndex        =   35
         Top             =   2570
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   5
         Left            =   2520
         TabIndex        =   34
         Top             =   2210
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   33
         Top             =   1850
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   3
         Left            =   2520
         TabIndex        =   32
         Top             =   1490
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   2
         Left            =   2520
         TabIndex        =   31
         Top             =   1130
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   1
         Left            =   2520
         TabIndex        =   30
         Top             =   770
         Width           =   450
      End
      Begin VB.Label PortStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Index           =   0
         Left            =   2520
         TabIndex        =   29
         Top             =   400
         Width           =   450
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   22
         Top             =   3650
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   3290
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   20
         Top             =   2930
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   19
         Top             =   2570
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   2210
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   1850
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1490
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1130
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   770
         Width           =   900
      End
      Begin VB.Label PortList 
         AutoSize        =   -1  'True
         Caption         =   "端口:*****"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   400
         Width           =   900
      End
   End
   Begin VB.TextBox LogText 
      Appearance      =   0  'Flat
      Height          =   4095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   11400
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   1
      Left            =   11400
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   2
      Left            =   11400
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   3
      Left            =   11400
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   4
      Left            =   11400
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   5
      Left            =   11400
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   6
      Left            =   11400
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   7
      Left            =   11400
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   8
      Left            =   11400
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   9
      Left            =   11400
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server Log:"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Connect As New WinsockConnect
Dim ConnectedUser As String

Private Sub Form_Load()
    ComputerName.Text = Connect.GetHostName
End Sub

Private Sub GetPortState_Timer()
    For i = 0 To 9
        PortStateList(i).Caption = "状态:" & Connect.StateReturn(Server(i))
    Next
End Sub

Private Sub PortReboot_Timer()
    For i = Server.LBound To Server.UBound
        If Not (Server(i).State = 7 Or Server(i).State = 2) Then
            Server(i).Close
            Server(i).Listen
            UserText(i).Text = ""
            ConnectedUser = SendConnectedUserCollection
            For j = 0 To 9
                If Server(j).State = 7 Then Server(j).SendData ConnectedUser
            Next
        End If
    Next
End Sub

Private Sub Save_Click()
    SockPackDispose.SaveData LogText.Text
End Sub

Private Sub Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If Server(Index).State <> sckClosed Then _
Server(Index).Close
'接受具有 requestID 参数的
'连接。
Server(Index).Accept requestID
End Sub

Private Sub Server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Data As String
    Dim DataFrame As SockPack
    Data = ""
    Server(Index).GetData Data
    Select Case SockPackDispose.GetLabelText(Data, "<+", "->", False)
        Case "Host"
            UserText(Index).Text = SockPackDispose.GetLabelText(Data, "<+Host->", "<-Host+>", False)
            LogText.Text = LogText.Text & vbCrLf & _
                           UserText(Index).Text & " Connected to Server" & vbCrLf & _
                           "Port:" & Server(Index).LocalPort
            ConnectedUser = SendConnectedUserCollection
            For i = 0 To 9
                If Server(i).State = 7 Then Server(i).SendData ConnectedUser
            Next
            Exit Sub
        Case "Package"
            SockPackDispose.ReadPackage Data, DataFrame
            SockPackDispose.ServerLog LogText, Server(i), DataFrame
            If DataFrame.SendMethod = GroupCost Then
                For i = LBound(DataFrame.AcceptUser) To UBound(DataFrame.AcceptUser)
                    For j = 0 To 9
                        If Server(j).State = 7 Then
                            If DataFrame.AcceptUser(i) = UserText(j).Text Then Server(j).SendData Data
                        End If
                    Next
                Next
            Else
                For j = 0 To 9
                    Debug.Print Server(j).State; UserText(j).Text
                    Debug.Print Server(j).State = 7 And UserText(j).Text <> DataFrame.SendUser
                    If Server(j).State = 7 And UserText(j).Text <> DataFrame.SendUser Then Server(j).SendData Data
                Next
            End If
    End Select
    LogText.SelStart = Len(LogText.Text)
End Sub

Private Function SendConnectedUserCollection() As String
    Dim UserCollection As String
    On Error Resume Next
    For i = 0 To 9
        If Server(i).State = 7 Then UserCollection = UserCollection & UserText(i).Text & ";"
    Next
    SendConnectedUserCollection = "<+ConnectedUser->" & Left(UserCollection, Len(UserCollection) - 1) & "<-ConnectedUser+>"
End Function

Private Sub SetPort_Click()

    Port = Val(PortText.Text)
    If 1 <= Port And Port <= 2 ^ 16 - 11 Then
        For i = 0 To 9
            PortList(i).Caption = "端口:" & Trim(Str(Port + i))
            Server(i).LocalPort = Port + i
            IPList(i).Caption = "IP:" & Server(i).LocalIP
        Next
    Else
        MsgBox "端口设置无效！", 48
    End If
End Sub

Private Sub SockClose_Click()
    For i = 0 To 9
        If Server(i).State = 7 Then Server(i).SendData "<+Server->Server Closed"
        UserText(i).Text = ""
    Next
    PortReboot.Enabled = False
    ServerClosed.Enabled = True
End Sub

Private Sub SockOpen_Click()
    For i = 0 To 9
        Server(i).Listen
    Next
    PortReboot.Enabled = True
End Sub

Private Sub ServerClosed_Timer()
    For i = 0 To 9
        Server(i).Close
    Next
    ServerClosed.Enabled = False
End Sub
