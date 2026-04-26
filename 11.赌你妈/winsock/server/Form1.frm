VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "GameServer"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6840
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox RoomList 
      Height          =   4200
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.Timer Clear 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   0
   End
   Begin VB.CommandButton Server_Looking 
      Caption         =   "查看端口状态"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Timer Server_Array_Listening 
      Interval        =   100
      Left            =   4440
      Top             =   3720
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   6000
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   1
      Left            =   6000
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   2
      Left            =   6000
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   3
      Left            =   6000
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   4
      Left            =   6000
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   5
      Left            =   6000
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   6
      Left            =   6000
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   7
      Left            =   6000
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   8
      Left            =   6000
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   9
      Left            =   6000
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'设置值

'State 属性的设置值是:

'常数                       值 描述
'sckClosed                  0  缺省的。关闭
'sckOpen                    1  打开
'sckListening               2  侦听
'sckConnectionPending       3  连接挂起
'sckResolvingHost           4  识别主机
'sckHostResolved            5  已识别主机
'sckConnecting              6  正在连接
'sckConnected               7  已连接
'sckClosing                 8  同级人员正在关闭连接
'sckError                   9  错误


Static Function RePort(Equipment As Winsock, Array_Bool As Boolean, Clear_Bool As Boolean, Count As Integer) As String
On Error Resume Next
Dim Port() As Integer, State() As String
Static i As Integer, RePortData As String
If Clear_Bool Then i = 0: RePortData = ""
RePort = "端口     " & "状态" & Chr(10)
If Array_Bool Then
    i = i + 1
    ReDim Preserve Port(1 To i), State(1 To i)
    Port(i) = Equipment.LocalPort
    State(i) = StateReturn(Equipment)
Else
    ReDim Preserve Port(1), State(1)
    Port(1) = Equipment.LocalPort
    State(1) = StateReturn(Equipment)
End If

For j = i To UBound(Port)
    RePortData = RePortData & Port(j) & String(9 - Len(Str(Port(j))), " ") & State(j) & Chr(10)
Next

If Count = i Then RePort = RePort & RePortData

End Function

Function StateReturn(Equipment As Winsock) As String
    If Equipment.State = 0 Then StateReturn = "关闭"
    If Equipment.State = 1 Then StateReturn = "打开"
    If Equipment.State = 2 Then StateReturn = "侦听"
    If Equipment.State = 3 Then StateReturn = "连接挂起"
    If Equipment.State = 4 Then StateReturn = "识别主机"
    If Equipment.State = 5 Then StateReturn = "已识别主机"
    If Equipment.State = 6 Then StateReturn = "正在连接"
    If Equipment.State = 7 Then StateReturn = "连接成功"
    If Equipment.State = 8 Then StateReturn = "同级人员正在关闭连接"
    If Equipment.State = 9 Then StateReturn = "连接错误"
End Function

'共 10 个端口 1011 - 1020
Private Sub Form_Load()
Port = 1010
For i = 0 To Server.Count - 1
    Port = Port + 1
    Server(i).LocalPort = Port
    Debug.Print Port
    Server(i).Listen
Next
End Sub

'重启
Private Sub ReBoot(ByVal Index As Integer)
Server(Index).Close
Server(Index).Listen
End Sub

'端口监测
Private Sub Server_Array_Listening_Timer()
For i = 0 To Server.Count - 1
    If Server(i).State <> 7 Then Call ReBoot(i)
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

Private Sub server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'为进入的数据声明一个变量。
Dim strData As String

'清除原始数据
strData = ""

'获取新的数据
Server(Index).GetData strData

'去空格
strData = Trim(strData)

'添加房间
If strData <> "GetRoom" And GetServerLabel(strData, "<", ">") <> "<DEL>" Then RoomList.AddItem strData

'删除房间
If GetServerLabel(strData, "<", ">") = "<DEL>" Then
    RoomName = GetClientName(strData)
    For i = 0 To RoomList.ListCount - 1
        If RoomList.List(i) = RoomName Then RoomList.RemoveItem i
    Next
End If

'发送房间列表数据
If strData = "GetRoom" Then
    For i = 0 To RoomList.ListCount - 1
        Server(Index).SendData RoomList.List(i) & "|"
    Next
End If
'向客户端发送 标签 Close
ClientClose (Index)
End Sub

'向客户端发送关闭命令
Sub ClientClose(Index As Integer)
Server(Index).SendData "Close"
End Sub

'检测状态
Private Sub Server_Looking_Click()
For i = 0 To 9
    RePort_Data = RePort(Server(i), True, IIf(i = 0, True, False), 10)
Next
MsgBox RePort_Data
End Sub

'获取服务器标签
Function GetServerLabel(Data As String, StartString As String, EndString As String) As String
On Error Resume Next
StartNum = InStr(1, Data, StartString)
EndNum = InStr(1, Data, EndString)
For i = StartNum To EndNum
    GetServerLabel = GetServerLabel & Mid(Data, i, 1)
Next
End Function

'获取客户端名称
Function GetClientName(Data As String) As String
On Error Resume Next
Num = InStr(1, Data, ">")
Num = Num + 1
GetClientName = Mid(Data, Num)
End Function



'**********************************************************没用***************************************************
'-----------------------------------------------------------------------------------------------------------------
Private Sub Save_Data()

On Error Resume Next

Names = Left(Text1.Text, Len(Text1.Text) - 1)
Open App.Path & "\房间列表.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, a
        If a = Names Then Exit Sub
    Loop
Close

Open App.Path & "\房间列表.txt" For Append As #1
    Print #1, Names
Close

End Sub

Private Sub Clear_Timer()
Open App.Path & "\房间列表.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, Name1
        Line Input #1, Name2
        If Name1 = Name2 Then a = 0
    Loop
Close
End Sub
'-----------------------------------------------------------------------------------------------------------------
