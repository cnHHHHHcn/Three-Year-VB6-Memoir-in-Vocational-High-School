Attribute VB_Name = "PublicMethod"
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Sub ServerDefaultDispose()
    Server.LocalPortText.Enabled = True
    Server.RemotePortText.Text = ""
    Server.RemotePortText.Enabled = False
    Server.LocalHostName = GetHostName
End Sub

Public Sub ClientDefaultDispose()
    Client.RemoteHostName.Enabled = True
    Client.Winsock1.RemotePort = Val(Client.RemotePortText.Text)
    Client.LocalHostName = GetHostName
End Sub

Public Function GetState(ByVal State As Integer) As String
    If State = 2 Then GetState = "正在侦听": Exit Function
    If State = 7 Then GetState = "链接成功": Exit Function
    If State = 0 Then GetState = "关闭状态": Exit Function
    If State = 9 Then GetState = "链接错误": Exit Function
    If State = 5 Then GetState = "已识别主机": Exit Function
    If State = 3 Then GetState = "连接挂起": Exit Function
    If State = 4 Then GetState = "识别主机": Exit Function
    If State = 6 Then GetState = "正在连接": Exit Function
    If State = 8 Then GetState = "同级人员正在关闭连接": Exit Function
    If State = 1 Then GetState = "打开": Exit Function
End Function

Public Function GetHostName() As String
    Dim Name As String
    Name = Space(255)
    GetComputerName Name, 255
    GetHostName = Name
End Function

Public Sub SaveData(Data As String)
    Num = FreeFile
    DateStr = Trim(Str(Year(Date))) & "-" & Trim(Str(Month(Date))) & "-" & Trim(Str(Day(Date)))
    Open App.Path & "\" & DateStr & ".txt" For Append As #Num
        Print #Num, Data
    Close
    
End Sub
