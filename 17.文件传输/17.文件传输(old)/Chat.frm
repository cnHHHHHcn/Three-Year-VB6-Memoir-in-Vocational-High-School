VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Chat 
   Caption         =   "Chat"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9195
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7680
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox SendUserList 
      Appearance      =   0  'Flat
      Height          =   750
      ItemData        =   "Chat.frx":0000
      Left            =   7320
      List            =   "Chat.frx":0002
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton RemoveUserItem 
      Caption         =   "移去"
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton UserItemClear 
      Caption         =   "清除"
      Height          =   375
      Left            =   8280
      TabIndex        =   18
      Top             =   5520
      Width           =   855
   End
   Begin VB.Frame Dispose 
      Caption         =   "配置"
      Height          =   4095
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   2055
      Begin VB.Timer GetStatePort 
         Interval        =   500
         Left            =   1440
         Top             =   3000
      End
      Begin VB.CommandButton SockClose 
         Caption         =   "关闭"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton SockOpen 
         Caption         =   "启动"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "Chat.frx":0004
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox RemoteComputerName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox RemotePortText 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   10
         Text            =   "10001"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox LocalPortText 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Text            =   "9001"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox LocalComputerName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label ErrorDisplay 
         AutoSize        =   -1  'True
         Caption         =   "Error："
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label StateList 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "远程 IP :"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1005
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "远程端口:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1365
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "本地端口:"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "本地 IP :"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   285
         Width           =   810
      End
   End
   Begin VB.ListBox ConnectedUserList 
      Appearance      =   0  'Flat
      Height          =   3090
      ItemData        =   "Chat.frx":0035
      Left            =   7320
      List            =   "Chat.frx":003C
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Save 
      Caption         =   "保存"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton RequireTransfer 
      Caption         =   "文件传输请求"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox SendText 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   4560
      Width           =   4695
   End
   Begin VB.TextBox DataText 
      Appearance      =   0  'Flat
      Height          =   4215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   7200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "服务器在线人员列表:"
      Height          =   180
      Left            =   7320
      TabIndex        =   22
      Top             =   720
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "指定发送人员区:"
      Height          =   180
      Left            =   7320
      TabIndex        =   21
      Top             =   4320
      Width           =   1350
   End
End
Attribute VB_Name = "Chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Public TransferFlag As Boolean
Private ConnectingPort As Long

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo GetDataErr
    Dim Data As String, Head As String, ConnectedUser() As String
    Client.GetData Data
    If SockPackDispose.GetPackageHead(Data, Head) Then
        Select Case Head
            Case "ConnectedUser"
                ConnectedUser = Split(Data, ";")
                ConnectedUserList.Clear
                For i = LBound(ConnectedUser) To UBound(ConnectedUser)
                    ConnectedUserList.AddItem ConnectedUser(i)
                Next
            Case "Server"
                 If CBool(InStr(1, Data, "Close")) Then GoTo GetDataErr
                 If CBool(InStr(1, Data, "IsOnline")) Then
                    Dim Host As HostAddress
                    Host.IP = Client.LocalIP
                    Host.Port = Client.LocalPort
                    Host.IsOnline = True
                    Client.SendData SockPackDispose.CreateHostPackage(Host)
                End If
            Case "Package"
                Dim DataInfo As SockPack
                SockPackDispose.ReadInfoPackage Data, DataInfo, True
                SockPackDispose.GetPackageHead DataInfo.Data, Head
                Select Case Head
                    Case "Request"
                        Dim RequestInfo As Request
                        SockPackDispose.ReadRequestPackage DataInfo.Data, RequestInfo, True
                        If (CBool(InStr(1, RequestInfo.Info, "TransferRequest")) And RequestInfo.IsRequest) Then
                            RequestInfo.IsRequest = False
                            If TransferFlag Then RequestInfo.IsAccept = False: GoTo Send
                            If MsgBox("用户:" & DataInfo.SendUser & "请求与你进行文件传输。" & Chr(13) & "是否同意？", 32 + 4, "UDPClient") = vbYes Then
                                LoadDiskTransfer RequestInfo.Info, DataInfo.SendUser
                                DiskTransfer.TransferDataFile.Listen: DiskTransfer.TransferID = True
                                For i = 0 To DiskTransfer.LocalDrive.ListCount - 1
                                    Drive = Drive & Left(DiskTransfer.LocalDrive.List(i), 2) & "|"
                                Next
                                Drive = "{" & Left(Drive, Len(Drive) - 1) & "}"
                                RequestInfo.Info = "TransferRequest" & Drive
                                RequestInfo.IsAccept = True
                            Else
                                RequestInfo.IsAccept = False
                            End If
Send:                       DataInfo.Data = SockPackDispose.CreateRequestPackage(RequestInfo)
                            DataInfo.AcceptUser(0) = DataInfo.SendUser: DataInfo.SendUser = Client.LocalIP: DataInfo.SendMethod = GroupCost
                            Client.SendData SockPackDispose.CreateInfoPackage(DataInfo): Exit Sub
                        End If
                        If RequestInfo.IsAccept Then LoadDiskTransfer RequestInfo.Info, DataInfo.SendUser: DiskTransfer.TransferDataFile.Connect: DiskTransfer.TransferID = False Else MsgBox "请求失败。", 64, "DiskTransferError"
                    Case Else
                        If DataInfo.SendMethod = GroupCost Then
                            For i = LBound(DataInfo.AcceptUser) To UBound(DataInfo.AcceptUser)
                                If DataInfo.AcceptUser(i) = LocalComputerName.Text Then SockPackDispose.ClientPirnt DataText, DataInfo
                            Next
                        Else
                            SockPackDispose.ClientPirnt DataText, DataInfo
                        End If
                        FlashWindow Chat.hwnd, 2: DataText.SelStart = Len(DataText.Text)
                End Select
        End Select
    End If
    GoTo GO
GetDataErr:
    SockOpen.Enabled = True
    SockClose.Enabled = False
    ErrorDisplay.Caption = "Error：" & Chr(13) & "服务器关闭。"
    Client.Close
GO:
End Sub

Private Sub LoadDiskTransfer(DriveInfo As String, RemoteUser As String)
    DriveList = Split(SockPackDispose.GetLabelText(DriveInfo, "{", "}", False), "|")
    With DiskTransfer
        .Show
        .DirList.Close
        .DirList.LocalPort = 28256
        .DirList.RemoteHost = RemoteUser
        .DirList.RemotePort = 28256
        .DirList.Bind
        .TransferDataFile.Close
        .TransferDataFile.LocalPort = 28257
        .TransferDataFile.RemoteHost = RemoteUser
        .TransferDataFile.RemotePort = 28257
        .RemoteTreeView.Nodes.Item("RemoteRoot").Text = RemoteUser
        .Status.Panels.Item(1).Text = RemoteUser & "状态:"
        For i = LBound(DriveList) To UBound(DriveList)
            .RemoteTreeView.Nodes.Add "RemoteRoot", tvwChild, "R" & Left(DriveList(i), 1), DriveList(i), "drive"
        Next
    End With
    TransferFlag = True
End Sub

Private Sub Command1_Click()
    Server.Show
End Sub

Private Sub ConnectedUserList_Click()
    Dim User As String, Flag As Boolean
    User = ConnectedUserList.Text: Flag = True
    For i = 0 To SendUserList.ListCount - 1
        If SendUserList.List(i) = User Then Flag = False
    Next
    If Flag Then SendUserList.AddItem User
    For i = 0 To SendUserList.ListCount - 1
        If SendUserList.List(i) = LocalComputerName.Text Then SendUserList.RemoveItem i
    Next
End Sub

Private Sub Form_Load()
    LocalComputerName.Text = Client.LocalIP
    ErrorDisplay.Caption = "Error：" & Chr(13) & "NONE"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SockClose_Click
End Sub

Private Sub RequireTransfer_Click()
    Dim DataInfo As SockPack, RequestInfo As Request
    DataInfo.SendUser = Client.LocalIP
    DataInfo.SendMethod = GroupCost
    If SendUserList.ListCount = 1 Then
        ReDim DataInfo.AcceptUser(0)
        DataInfo.AcceptUser(0) = SendUserList.List(0)
        For i = 0 To DiskTransfer.LocalDrive.ListCount - 1
            Drive = Drive & Left(DiskTransfer.LocalDrive.List(i), 2) & "|"
        Next
        Drive = "{" & Left(Drive, Len(Drive) - 1) & "}"
        RequestInfo.Info = "TransferRequest" & Drive
        RequestInfo.IsRequest = True
        DataInfo.Data = SockPackDispose.CreateRequestPackage(RequestInfo)
        Client.SendData SockPackDispose.CreateInfoPackage(DataInfo)
    Else
        MsgBox "请选择或指定一个用户！", 48
    End If
End Sub

Private Sub Save_Click()
    SockPackDispose.SaveData DataText.Text
End Sub

Private Function SetDispose() As Boolean
    On Error Resume Next
    
    Dim LFlag As Boolean, RFlag As Boolean, LHFlag As Boolean, RHFlag As Boolean
    Dim ErrorText As String
    Dim Count As Byte
    
    Client.Close
    LocalHostPort = Val(LocalPortText.Text)
    RemoteHostPort = Val(RemotePortText.Text)
    If LocalComputerName.Text <> "" Then LHFlag = True Else ErrorText = ErrorText & "本地名称，": Count = Count + 1
    If (1 <= LocalHostPort And LocalHostPort <= 2 ^ 16 - 11) And (LocalHostPort <> 28256 And LocalHostPort <> 28257) Then LFlag = True Else Count = Count + 1: ErrorText = ErrorText & IIf(Count Mod 2 = 0, "本地端口" & vbCrLf, "本地端口，")
    If RemoteComputerName.Text <> "" Then RHFlag = True: Client.RemoteHost = RemoteComputerName.Text Else Count = Count + 1: ErrorText = ErrorText & IIf(Count Mod 2 = 0, "远程名称" & vbCrLf, "远程名称，")
    If 1 <= RemoteHostPort And RemoteHostPort <= 2 ^ 16 - 11 Then RFlag = True: Client.RemotePort = RemoteHostPort Else Count = Count + 1: ErrorText = ErrorText & IIf(Count Mod 2 = 0, "远程端口" & vbCrLf, "远程端口，")
    ErrorText = Left(ErrorText, Len(ErrorText) - 1)
    Flag = RFlag And RHFlag And LFlag
    If Not Flag Then ErrorDisplay.Caption = "Error：" & Chr(13) & ErrorText & "未设置成功！"
    SetDispose = Flag
End Function

Private Sub SendText_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Frame As SockPack
    If KeyCode = 13 Then
        If Shift = 1 Then Exit Sub
        Frame.SendUser = Client.LocalIP
        TextLen = Len(SendText.Text)
        If Not (1 <= TextLen And TextLen <= 1000) Then SendText.Text = Left(SendText.Text, 1000)
        If Client.State <> 0 And SendText.Text <> "" Then Frame.Data = SendText.Text
        If SendUserList.ListCount > 0 Then
            Frame.SendMethod = GroupCost
            For i = 0 To SendUserList.ListCount - 1
                ReDim Preserve Frame.AcceptUser(i)
                Frame.AcceptUser(i) = SendUserList.List(i)
            Next
            SockPackDispose.ClientPirnt DataText, Frame
        Else
            Frame.SendMethod = AllCost
        End If
        If Client.State <> 0 And SendText.Text <> "" Then
            Client.SendData SockPackDispose.CreateInfoPackage(Frame)
            SendText.Text = "": DataText.SelStart = Len(DataText.Text)
        End If
    End If
End Sub

Private Sub SockClose_Click()
    Dim Host As HostAddress
    Host.IP = Client.LocalIP
    Host.IsDel = True
    Client.SendData SockPackDispose.CreateHostPackage(Host)
    Client.Close
    ConnectedUserList.Clear
    SockClose.Enabled = False
    SockOpen.Enabled = True
End Sub

Private Sub SockOpen_Click()
    On Error GoTo Mind
    If SetDispose Then
        Client.Close
        Client.Bind CLng(LocalPortText.Text), LocalComputerName.Text
        GoTo GO
Mind:
        Shell "cmd.exe /c ipconfig /all & pause", vbMaximizedFocus
        MsgBox "指定本地或远程 IP 不正确，请在网络适配器中指定 本地IP 。" & Chr(10) & "请选择 IPv4 地址 首选IP", 16, "UDPClientError"
        SockClose.Enabled = False: SockOpen.Enabled = True
        Exit Sub
GO:
        SockClose.Enabled = True
        SockOpen.Enabled = False
        ErrorDisplay.Caption = "Error：" & Chr(13) & "NONE"
        Dim Host As HostAddress
        Host.IP = Client.LocalIP
        Host.Port = Client.LocalPort
        Host.IsAdd = True
        Client.SendData SockPackDispose.CreateHostPackage(Host)
    End If
End Sub

Private Sub GetStatePort_Timer()
    StateList.Caption = "Client状态：" & Chr(13) & SockPackDispose.StateReturn(Client)
    If Client.State = 0 Then ConnectedUserList.Clear
End Sub

Private Sub UserItemClear_Click()
    SendUserList.Clear
End Sub

Private Sub RemoveUserItem_Click()
    For i = SendUserList.ListCount - 1 To 0 Step -1
        If SendUserList.Selected(i) Then SendUserList.RemoveItem i
    Next
End Sub
