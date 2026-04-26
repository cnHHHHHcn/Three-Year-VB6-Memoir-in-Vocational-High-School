VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3840
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim a() As Byte

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If Winsock1.State <> sckClosed Then _
Winsock1.Close
'接受具有 requestID 参数的
'连接。
Winsock1.Accept requestID
End Sub

Private Sub Command1_Click()
    Dim FilePoint As Long
    Dim Index As Long
    Dim SizeCount As Long
    SizeCount = 1
    Open "E:\红蜘蛛.zip" For Binary As #1
        Do Until EOF(1)
            ReDim Preserve a(Index)
            FilePoint = FilePoint + 1
            Get #1, FilePoint, a(Index)
            Index = Index + 1
            If Index Mod 8192 = 0 Or EOF(1) Then
                i = i + 1
                Debug.Print i
                DoEvents
                Winsock1.SendData a
                Index = 0
            End If
        Loop
    Close
End Sub

Private Sub Command2_Click()
'    Winsock1.Close
'    Winsock1.LocalPort = 1001
'    Winsock1.Listen
    Winsock1.RemoteHost = "admin031"
    Winsock1.RemotePort = 1001
    Winsock1.Connect "admin031", 1001
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Winsock1.State
End Sub
