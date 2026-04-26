VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9120
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton OpenFileTransferForm 
      Caption         =   "文件传输"
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ListBox SendUserList 
      Appearance      =   0  'Flat
      Height          =   750
      ItemData        =   "Form1.frx":0000
      Left            =   6840
      List            =   "Form1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton RemoveUserItem 
      Caption         =   "移去"
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton UserItemClear 
      Caption         =   "清除"
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   5520
      Width           =   855
   End
   Begin VB.Frame Dispose 
      Caption         =   "配置"
      Height          =   3735
      Left            =   5160
      TabIndex        =   5
      Top             =   480
      Width           =   2055
      Begin VB.Timer Connecting 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   1560
         Top             =   1440
      End
      Begin VB.Timer GetStatePort 
         Interval        =   500
         Left            =   1440
         Top             =   3000
      End
      Begin VB.CommandButton SockClose 
         Caption         =   "关闭"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton SockOpen 
         Caption         =   "启动"
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton SetDispose 
         Caption         =   "设置"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "Form1.frx":0004
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox RemoteComputerName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox RemotePortText 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   10
         Text            =   "10001"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox LocalPortText 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Text            =   "9001"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox LocalComputerName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label StateList 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "远程名称:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1125
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "远程端口:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1485
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "本地端口:"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   765
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "本地名称:"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.ListBox ConnectedUserList 
      Appearance      =   0  'Flat
      Height          =   3090
      ItemData        =   "Form1.frx":0033
      Left            =   7440
      List            =   "Form1.frx":0035
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Save 
      Caption         =   "保存"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Send 
      Caption         =   "发送"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox SendText 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox DataText 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   7320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "服务器在线人员列表:"
      Height          =   180
      Left            =   7320
      TabIndex        =   23
      Top             =   480
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "指定发送人员区:"
      Height          =   180
      Left            =   6840
      TabIndex        =   22
      Top             =   4320
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Dim Connect As New WinsockConnect
Dim StartPort As Long, ConnectingPort As Long

Private Sub Client_Connect()
    Client.SendData "<+Host->" & LocalComputerName.Text & "<-Host+>"
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String, ConnectedUser() As String
    Dim DataFrame As SockPack
    Client.GetData Data
    Select Case SockPackDispose.GetLabelText(Data, "<+", "->", False)
        Case "ConnectedUser"
            ConnectedUser = Split(SockPackDispose.GetLabelText(Data, "<+ConnectedUser->", "<-ConnectedUser+>", False), ";")
            ConnectedUserList.Clear
            For i = LBound(ConnectedUser) To UBound(ConnectedUser)
                ConnectedUserList.AddItem ConnectedUser(i)
            Next
        Case "Package"
            SockPackDispose.ReadPackage Data, DataFrame
            If DataFrame.SendMethod = GroupCost Then
                For i = LBound(DataFrame.AcceptUser) To UBound(DataFrame.AcceptUser)
                    If DataFrame.AcceptUser(i) = LocalComputerName.Text Then SockPackDispose.ClientPirnt DataText, DataFrame
                Next
            Else
                SockPackDispose.ClientPirnt DataText, DataFrame
            End If
        Case "Server"
             DataText.Text = DataText.Text & vbCrLf & "Server:" & vbCrLf & Mid(Data, 11)
    End Select
    FlashWindow Form1.hwnd, 1
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
End Sub

Private Sub GetStatePort_Timer()
    StateList.Caption = "Client状态：" & Chr(13) & Connect.StateReturn(Client)
    If Client.State <> 7 Then ConnectedUserList.Clear
End Sub

Private Sub OpenFileTransferForm_Click()
    Form2.Show
End Sub

Private Sub Save_Click()
    SockPackDispose.SaveData DataText.Text
End Sub

Private Sub Send_Click()
    Dim Frame As SockPack
    Frame.SendUser = LocalComputerName.Text
    TextLen = Len(SendText.Text)
    If Not (1 <= TextLen And TextLen <= 200) Then SendText.Text = Left(SendText.Text, 200)
    If Client.State = 7 And SendText.Text <> "" Then Frame.Data = SendText.Text
    If SendUserList.ListCount <> 0 Then
        Frame.SendMethod = GroupCost
        For i = 0 To SendUserList.ListCount - 1
            ReDim Preserve Frame.AcceptUser(i)
            Frame.AcceptUser(i) = SendUserList.List(i)
        Next
    Else
        Frame.SendMethod = AllCost
    End If
    If Client.State = 7 And SendText.Text <> "" Then
        Client.SendData SockPackDispose.CreatePackage(Frame)
        SendText.Text = ""
        SockPackDispose.ClientPirnt DataText, Frame
    End If
End Sub

Private Sub SetDispose_Click()
    Dim LHFlag As Boolean, LFlag As Boolean, RFlag As Boolean, RHFlag As Boolean
    Dim ErrorText As String
    
    On Error Resume Next
    
    Client.Close
    StartPort = Val(RemotePortText.Text)
    LocalHostPort = Val(LocalPortText.Text)
    
    If LocalComputerName.Text <> "" Then LHFlag = True: Client.LocalHostName = LocalComputerName.Text Else ErrorText = ErrorText & "本地名称，"
    If 1 <= StartPort And StartPort <= 2 ^ 16 - 11 Then LFlag = True: Client.LocalPort = LocalHostPort Else ErrorText = ErrorText & "本地端口，"
    If RemoteComputerName.Text <> "" Then RHFlag = True: Client.RemoteHost = RemoteComputerName.Text Else ErrorText = ErrorText & "远程名称，"
    If 1 <= StartPort And StartPort <= 2 ^ 16 - 11 Then RFlag = True: Client.RemotePort = StartPort Else ErrorText = ErrorText & "远程端口，"
    
    ErrorText = Left(ErrorText, Len(ErrorText) - 1)
    
    If RFlag And RHFlag And LFlag Then
        MsgBox "设置成功！", 64
        SockOpen.Enabled = True
        SetDispose.Enabled = False
    Else
        MsgBox ErrorText & "未设置成功！", 48
    End If
End Sub

Private Sub Connecting_Timer()
    If Not (StartPort <= ConnectingPort And ConnectingPort < StartPort + 9) Then ConnectingPort = StartPort - 1
    If Not (Client.State = 7 Or Client.State = 2) Then
        ConnectingPort = ConnectingPort + 1
        Connect.ClientReBoot ConnectingPort, Client, Client.RemoteHost
        RemotePortText.Text = ConnectingPort
    Else
        Connecting.Enabled = False
    End If
End Sub

Private Sub SockClose_Click()
    Connecting.Enabled = False
    SockClose.Enabled = False
    SetDispose.Enabled = True
    Client.Close
    ConnectedUserList.Clear
End Sub

Private Sub SockOpen_Click()
    Client.Close
    SockOpen.Enabled = False
    SockClose.Enabled = True
    Connecting.Enabled = True
End Sub

Private Sub UserItemClear_Click()
    SendUserList.Clear
End Sub

Private Sub RemoveUserItem_Click()
    For i = 0 To SendUserList.ListCount - 1
        If SendUserList.Selected(i) = True Then SendUserList.RemoveItem i: Exit For
    Next
End Sub
