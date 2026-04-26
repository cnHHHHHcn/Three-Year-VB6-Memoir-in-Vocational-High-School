Attribute VB_Name = "Server_COde"
'重启
Function ReBoot(Equipment As Winsock)
With Equipment
    .Close
    .Listen
End With
End Function

'向客户端发送关闭
Function Send_Close(Equipment As Winsock)
With Equipment
    .SendData "Close"
End With
End Function

'获取服务器标签
Function GetServerLabel(ByVal Data As String, StartString As String, EndString As String) As String
On Error Resume Next
StartNum = InStr(1, Data, StartString)
EndNum = InStr(1, Data, EndString)
For i = StartNum To EndNum
    GetServerLabel = GetServerLabel & Mid(Data, i, 1)
Next
End Function
