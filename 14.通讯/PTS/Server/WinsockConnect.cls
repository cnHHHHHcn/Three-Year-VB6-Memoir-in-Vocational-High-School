VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinsockConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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

Public Equipment As Winsock

'类模块方法中 形参值的优先级 要高于 属性值 _
 优先级 形参值 > 属性值
'--------------------------------------------------------------------------------------------------------------
Private Function GetPropertyEquipment(Equipment As Winsock) As Winsock
    Set GetPropertyEquipment = IIf(Equipment Is Nothing, Me.Equipment, Equipment)
End Function
'--------------------------------------------------------------------------------------------------------------

Public Sub ClientReBoot(ByVal Port As Integer, Optional Equipment As Winsock = Nothing, Optional RemoteHostName As String = "None")
    
    Set Equipment = GetPropertyEquipment(Equipment)
    
    With Equipment
        .Close
        .RemotePort = Port
        If RemoteHostName = "None" Or RemoteHostName = "" Then
            .Connect
        Else
            .Connect RemoteHostName
        End If
    End With
End Sub

Public Sub ServerReBoot(Optional Equipment As Winsock = Nothing)
    
    Set Equipment = GetPropertyEquipment(Equipment)
    
    With Equipment
        .Close
        .Listen
    End With
End Sub

'向客户端发送关闭
Public Function Server_Send_Close(Optional Equipment As Winsock = Nothing)
    
    Set Equipment = GetPropertyEquipment(Equipment)
    
    With Equipment
        .SendData "Close"
    End With
End Function

Public Static Function ServerRePort(Array_Bool As Boolean, Clear As Boolean, Optional Equipment As Winsock = Nothing, Optional Count As Integer = 1) As String
    On Error Resume Next
    Dim Port() As Integer, State() As String
    
    Set Equipment = GetPropertyEquipment(Equipment)
    
    Static i As Integer, ReportData As String
    If Clear Then i = 0: ReportData = ""
    ServerRePort = "端口     " & "状态" & vbCrLf
    If Array_Bool Then
        i = i + 1
        ReDim Preserve Port(1 To i), State(1 To i)
        Port(i) = Equipment.LocalPort
        State(i) = StateReturn(Equipment)
    Else
        i = 1
        ReDim Preserve Port(1), State(1)
        Port(1) = Equipment.LocalPort
        State(1) = StateReturn(Equipment)
    End If
    
    ReportData = ReportData & Port(i) & String(9 - Len(Str(Port(i))), " ") & State(i) & vbCrLf
    
    If Count = i Then ServerRePort = ServerRePort & ReportData

End Function

Public Function ClientRePort(Optional Equipment As Winsock = Nothing) As String
    
    Set Equipment = GetPropertyEquipment(Equipment)
    
    ReportData = "本地端口     远程端口     状态" & Chr(10)
    LocalPort = Equipment.LocalPort
    RemotePort = Equipment.RemotePort
    
    ClientRePort = ReportData & LocalPort & String(14 - Len(Str(LocalPort)), " ") & RemotePort & String(14 - Len(Str(RemotePort)), " ") & StateReturn(Equipment)
    
    Debug.Print ClientRePort
    
End Function

Public Function StateReturn(Optional Equipment As Winsock = Nothing) As String
    
    Set Equipment = GetPropertyEquipment(Equipment)

    If Equipment.State = 0 Then StateReturn = "设备关闭": Exit Function
    If Equipment.State = 1 Then StateReturn = "设备打开": Exit Function
    If Equipment.State = 2 Then StateReturn = "正在侦听": Exit Function
    If Equipment.State = 3 Then StateReturn = "连接挂起": Exit Function
    If Equipment.State = 4 Then StateReturn = "识别主机": Exit Function
    If Equipment.State = 5 Then StateReturn = "已识别主机": Exit Function
    If Equipment.State = 6 Then StateReturn = "正在连接": Exit Function
    If Equipment.State = 7 Then StateReturn = "连接成功": Exit Function
    If Equipment.State = 8 Then StateReturn = "同级人员正在关闭连接": Exit Function
    If Equipment.State = 9 Then StateReturn = "连接错误": Exit Function
End Function

Public Function GetHostName() As String
    Dim NetName As String
    NetName = Space(256)
    GetComputerName NetName, 256
    GetHostName = Trim(NetName)
End Function

