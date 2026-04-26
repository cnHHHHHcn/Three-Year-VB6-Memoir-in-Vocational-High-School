Attribute VB_Name = "Public_Code"

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

Sub EquipmentState(Equipment As Winsock, StateShow As Control)
If Equipment.State = 0 Then StateShow.Caption = "关闭状态"
If Equipment.State = 1 Then StateShow.Caption = "打开"
If Equipment.State = 2 Then StateShow.Caption = "正在侦听"
If Equipment.State = 3 Then StateShow.Caption = "连接挂起"
If Equipment.State = 4 Then StateShow.Caption = "识别主机"
If Equipment.State = 5 Then StateShow.Caption = "已识别主机"
If Equipment.State = 6 Then StateShow.Caption = "正在连接"
If Equipment.State = 7 Then StateShow.Caption = "链接成功"
If Equipment.State = 8 Then StateShow.Caption = "同级人员正在关闭连接"
If Equipment.State = 9 Then StateShow.Caption = "链接错误"
End Sub

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

Function GetHostName() As String
Shell ("cmd.exe /c hostname > " & App.Path & "\Name.txt")
Open App.Path & "\Name.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, HostName
    Loop
Close
GetHostName = HostName
End Function
