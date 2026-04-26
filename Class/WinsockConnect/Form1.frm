VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer ServerReboot 
      Interval        =   1000
      Left            =   3360
      Top             =   840
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   1200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   1
      Left            =   1560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   2
      Left            =   1920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   3
      Left            =   2280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   4
      Left            =   2640
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Win As New WinsockConnect
Dim pack As New NetWorkPackage

Private Sub Command2_Click()
    Dim i As Package
    Call pack.a(i)
    Print i.AcceptUser(0)
End Sub

Private Sub Form_Load()
    For i = 0 To 4
        Server(i).LocalPort = 1011 + i
        Server(i).Listen
    Next
End Sub

Private Sub server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
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
    Server(Index).GetData Data
    Win.StarNetwork Server, Index, Data
End Sub

Private Sub ServerReboot_Timer()
    '端口重启
    For i = 0 To Server.Count - 1
        Win.ServerReboot Server(i)
    Next
End Sub
