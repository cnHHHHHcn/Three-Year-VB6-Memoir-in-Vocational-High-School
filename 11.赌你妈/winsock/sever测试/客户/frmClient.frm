VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "TCP Client"
   ClientHeight    =   5070
   ClientLeft      =   6945
   ClientTop       =   3300
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   9405
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   6840
      Top             =   3960
   End
   Begin VB.CommandButton Send 
      Caption         =   "send"
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8760
      Top             =   3240
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   7200
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新启动"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton set 
      Caption         =   "设置"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "链接"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtOutput 
      Height          =   2535
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtSend 
      Height          =   2535
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   4680
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "当前状态是："
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "当前端口号是："
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "接收端"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "发送端"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload tcpClient
End Sub

Private Sub Command2_Click()
tcpClient.Close
tcpClient.RemoteHost = Form1.Text2.Text
tcpClient.RemotePort = Form1.Text1.Text
Label4.Caption = tcpClient.RemotePort
tcpClient.Connect
Call ZhuangTai
End Sub

Private Sub Command3_Click()
MsgBox tcpClient.RemotePort
End Sub

Private Sub Form_Load()
Name_Ip = "Admin008"
DuanKou = 1001
'Winsock 控件的名字为 tcpClient。
'注意：要指定远程主机，可以使用
' IP 地址（例如："121.111.1.1"），也可以使用
'计算机的“好听的名字”如下所示。
tcpClient.RemoteHost = Name_Ip
tcpClient.RemotePort = DuanKou
End Sub

Private Sub cmdConnect_Click()
On Error Resume Next
Label4.Caption = tcpClient.RemotePort
   '调用 Connect 方法，初始化连接。
tcpClient.Connect
Call ZhuangTai
End Sub

Private Sub txtSendData_Change()
   tcpClient.SendData txtSend.Text
End Sub

Private Sub Send_Click()
If tcpClient.State = 7 Then
    tcpClient.SendData txtSend.Text & " "
End If
End Sub

Private Sub set_Click()
tcpClient.Close
Form1.Show
End Sub

Private Sub tcpClient_DataArrival _
(ByVal bytesTotal As Long)
    Dim strData As String
    tcpClient.GetData strData
    If Left(strData, 9) = "Text3Data" Then
        Text3.Text = strData
    Else
        txtOutput.Text = ""
        txtOutput.Text = strData
    End If
Call ZhuangTai
End Sub
Private Sub ZhuangTai()

If tcpClient.State = 2 Then Label6.Caption = "正在侦听"
If tcpClient.State = 7 Then Label6.Caption = "链接成功"
If tcpClient.State = 0 Then Label6.Caption = "关闭状态"
If tcpClient.State = 9 Then Label6.Caption = "链接错误"
If tcpClient.State = 5 Then Label6.Caption = "已识别主机"
If tcpClient.State = 3 Then Label6.Caption = "连接挂起"
If tcpClient.State = 4 Then Label6.Caption = "识别主机"
If tcpClient.State = 6 Then Label6.Caption = "正在连接"
If tcpClient.State = 8 Then Label6.Caption = "同级人员正在关闭连接"
If tcpClient.State = 1 Then Label6.Caption = "打开"

End Sub




Private Sub Timer1_Timer()
Call ZhuangTai
End Sub

Private Sub Timer2_Timer()
If Label6.Caption = "链接错误" Then DuanKou = DuanKou + 1
End Sub

Private Sub txtOutput_Change()

End Sub
