VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5775
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List 
      Height          =   1680
      ItemData        =   "Form1.frx":0000
      Left            =   600
      List            =   "Form1.frx":0002
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3840
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "启动"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3120
      Top             =   3480
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   4320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Send 
      Caption         =   "send"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "当前端口号是："
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "当前状态是："
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "发送端"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Port As Integer

Private Sub ZhuangTai()

If Client.State = 2 Then Label6.Caption = "正在侦听"
If Client.State = 7 Then Label6.Caption = "链接成功"
If Client.State = 0 Then Label6.Caption = "关闭状态"
If Client.State = 9 Then Label6.Caption = "链接错误"
If Client.State = 5 Then Label6.Caption = "已识别主机"
If Client.State = 3 Then Label6.Caption = "连接挂起"
If Client.State = 4 Then Label6.Caption = "识别主机"
If Client.State = 6 Then Label6.Caption = "正在连接"
If Client.State = 8 Then Label6.Caption = "同级人员正在关闭连接"
If Client.State = 1 Then Label6.Caption = "打开"

End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Client.GetData strData
strData = Trim(strData)
strArray = Split(strData, "|")
For i = LBound(strArray) To UBound(strArray)
    If strArray(i) <> "" Then List.AddItem strArray(i)
Next
End Sub

Private Sub Command2_Click()
Client.Close
Client.RemotePort = Port
Label4.Caption = Client.RemotePort
Client.Connect
Call ZhuangTai
End Sub

Private Sub Form_Load()
Shell ("cmd.exe /c hostname > " & App.Path & "\client.txt")
Open App.Path & "\client.txt" For Input As #1
    Do Until EOF(1)
    Line Input #1, HostName
    Loop
Close
Client.RemoteHost = HostName
Client.RemotePort = 1011
End Sub

Sub Client_ReBoot()
Client.Close
Client.RemotePort = Port
Label4.Caption = Port
Client.Connect
End Sub

Sub Server_ReBoot()
Client.Close
Client.Listen
End Sub

Private Sub Send_Click()
If Client.State = 7 Then
    Client.SendData txtSend.Text & " "
    txtSend.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
Call ZhuangTai
End Sub

Private Sub Timer2_Timer()
If Not (1010 < Port And Port < 1020) Then Port = 1010
If Label6.Caption = "链接错误" Then Port = Port + 1: Client_ReBoot
End Sub

Sub SendData(UserName As String)
Open App.Path & "\client.txt" For Input As #1
    Do While EOF(1)
        Line Input #1, HostName
    Loop
Close
If Client.State = 7 Then
    Client.SendData HostName & " " & UserName
End If
Server_ReBoot
End Sub

