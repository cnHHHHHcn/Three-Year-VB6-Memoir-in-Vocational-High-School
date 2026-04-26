VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   Caption         =   "Server"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   5565
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton SaveLog 
      Caption         =   "保存Server日志"
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock UDPServer 
      Left            =   3960
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer IsOnline 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   1680
   End
   Begin VB.TextBox Log 
      Appearance      =   0  'Flat
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   2880
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   "配置"
      Height          =   1935
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1320
         Top             =   1440
      End
      Begin VB.CommandButton SockOpen 
         Caption         =   "启动"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton SockClose 
         Caption         =   "关闭"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox LocalComputerName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox LocalPortText 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   2
         Text            =   "10001"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label State 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "本地 IP :"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "本地端口:"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   645
         Width           =   810
      End
   End
   Begin VB.ListBox UserList 
      Appearance      =   0  'Flat
      Height          =   2550
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.ListBox OnlineIP 
      Appearance      =   0  'Flat
      Height          =   2010
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Function SetDispose() As Boolean
    On Error Resume Next
    Dim LFlag As Boolean
    LocalHostPort = Val(LocalPortText.Text)
    If (1 <= LocalHostPort And LocalHostPort <= 2 ^ 16 - 11) And (LocalHostPort <> 28256 And LocalHostPort <> 28257) Then LFlag = True Else ErrorText = ErrorText & "本地端口，"
    If Not LFlag Then MsgBox ErrorText & "未设置成功！", 48, "UDPServerError"
    SetDispose = LFlag
End Function

Private Sub Form_Load()
    LocalComputerName.Text = UDPServer.LocalIP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SockClose_Click
End Sub

Private Sub SaveLog_Click()
    SockPackDispose.SaveData Log.Text
End Sub

Private Sub SockClose_Click()
    On Error Resume Next
    For i = 0 To UserList.ListCount - 1
        SendDataInfo "<+Server->Close<-Server+>", UserList.List(i)
    Next
    UDPServer.Close
    UserList.Clear: OnlineIP.Clear
    SockClose.Enabled = False
    SockOpen.Enabled = True
    Log.Text = Log.Text & vbCrLf & "UDPServer Closed"
End Sub

Private Sub SockOpen_Click()
    On Error GoTo Mind
    If SetDispose Then
        UDPServer.Close
        UDPServer.Bind CLng(LocalPortText.Text), LocalComputerName.Text
        GoTo GO
Mind:
        Shell "cmd.exe /c ipconfig /all & pause", vbMaximizedFocus
        MsgBox "指定 IP 不正确，请在网络适配器中指定 IP 。" & Chr(10) & "请选择 IPv4 地址 首选IP", 16, "UDPClientError"
        SockClose.Enabled = False
        Exit Sub
GO:
        SockClose.Enabled = True
        SockOpen.Enabled = False
    End If
    Log.Text = Log.Text & vbCrLf & "UDPServer Opened"
End Sub

Private Sub Timer1_Timer()
    State.Caption = "Server状态：" & Chr(13) & SockPackDispose.StateReturn(UDPServer)
    Static Add As Long
    If Add = 20 Then
        For i = 0 To UserList.ListCount - 1
            SendDataInfo "<+Server->IsOnline<-Server+>", UserList.List(i)
        Next
        IsOnline.Enabled = True
        Add = 0
    End If
    Add = Add + 1
End Sub

Private Sub IsOnline_Timer()
    If OnlineIP.ListCount > 0 Then
        UserList.Clear
        For i = 0 To OnlineIP.ListCount - 1
            UserList.AddItem OnlineIP.List(i)
        Next
        OnlineIP.Clear
        For i = 0 To UserList.ListCount - 1
            SendDataInfo GetConnectedUser, UserList.List(i)
        Next
    End If
    IsOnline.Enabled = False
End Sub

Private Sub UDPServer_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim Data As String, TempData As String, Head As String, _
        OperaFlag As String, UserRange As String
    Dim Host As HostAddress, DataInfo As SockPack
    UDPServer.GetData Data
    Log.Text = Log.Text & vbCrLf
    TempData = Data
    If SockPackDispose.GetPackageHead(TempData, Head) Then
        Select Case Head
            Case "Host"
                SockPackDispose.ReadHostPackage TempData, Host, True
                If Host.IsAdd Then UserList.AddItem Host.IP & ":" & Host.Port: OperaFlag = "Add"
                If Host.IsDel Then
                    For i = UserList.ListCount - 1 To 0
                        If CBool(InStr(1, UserList.List(i), Host.IP)) Then UserList.RemoveItem i
                    Next
                    OperaFlag = "Del"
                End If
                If Host.IsOnline Then
                    Log.Text = Log.Text & Head & "   " & Host.IP & ":" & Host.Port & " Online"
                    OnlineIP.AddItem Host.IP & ":" & Host.Port: Exit Sub
                End If
                For i = 0 To UserList.ListCount - 1
                    SendDataInfo GetConnectedUser, UserList.List(i)
                Next
                Log.Text = Log.Text & Head & "   " & OperaFlag & " " & Host.IP & ":" & Host.Port & " " & "Send AllHost Refreshed"
            Case "Package"
                SockPackDispose.ReadInfoPackage TempData, DataInfo, True
                If DataInfo.SendMethod = AllCost Then
                    OperaFlag = "AllCost": UserRange = "All Host"
                    For i = 0 To UserList.ListCount - 1
                        SendDataInfo Data, UserList.List(i)
                    Next
                Else
                    OperaFlag = "GroupCost"
                    For i = LBound(DataInfo.AcceptUser) To UBound(DataInfo.AcceptUser)
                        UserRange = UserRange & "," & DataInfo.AcceptUser(i)
                        For j = 0 To UserList.ListCount - 1
                            If CBool(InStr(1, UserList.List(j), DataInfo.AcceptUser(i))) Then SendDataInfo Data, UserList.List(j)
                        Next
                    Next
                    UserRange = Mid(UserRange, 2)
                End If
                Log.Text = Log.Text & Head & "   " & OperaFlag & "   " & UserRange & vbCrLf & "Data:" & vbCrLf & DataInfo.Data
        End Select
    End If
End Sub

Private Sub SendDataInfo(Data As String, UserInfo As String)
    Pos = InStr(1, UserInfo, ":")
    UDPServer.RemoteHost = Left(UserInfo, Pos - 1)
    UDPServer.RemotePort = Mid(UserInfo, Pos + 1)
    UDPServer.SendData Data
End Sub

Private Function GetConnectedUser() As String
    On Error Resume Next
    Dim User As String
    For i = 0 To UserList.ListCount - 1
        User = User & Left(UserList.List(i), InStr(1, UserList.List(i), ":") - 1) & ";"
    Next
    User = Left(User, Len(User) - 1)
    GetConnectedUser = "<+ConnectedUser->" & User & "<-ConnectedUser+>"
End Function

