VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8490
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "send"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock udpPeerA 
      Left            =   2760
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtSend 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
udpPeerA.SendData "sb"
End Sub

Private Sub Form_Load() '控件的名字为 udpPeerA
With udpPeerA
'重点:必须将 RemoteHost 的值
'修改为计算机的名字。
.RemoteHost = "PeerB"
.RemotePort = 1002  '连接的端口号。
.Bind 1001           '绑定到本地的端口。
End With
End Sub

Private Sub udpPeerA_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
udpPeerA.GetData strData
txtOutput.Text = strData
End Sub
