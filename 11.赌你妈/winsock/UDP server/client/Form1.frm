VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7860
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "send"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtOutPut 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2880
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock udpPeerB 
      Left            =   3240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtSend 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'控件的名字为 udpPeerB。
With udpPeerB
'重点:必须将 RemoteHost 的值改为
'计算机的名字。
.RemoteHost = "PeerA"
.RemotePort = 1001    '要连接的端口。
.Bind 1002                '绑定到本地的端口上。
End With
End Sub

Private Sub txtSend_Change()
'在键入后立即发送文本。
udpPeerB.SendData txtSend.Text
End Sub

Private Sub udpPeerB_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
udpPeerB.GetData strData
txtOutPut.Text = strData
End Sub
