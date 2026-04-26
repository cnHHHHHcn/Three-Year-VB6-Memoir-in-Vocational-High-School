Attribute VB_Name = "SockPackDispose"
'定义结构体 SockPack
Public Type SockPack
    SendUser As String
    AcceptUser() As String
    SendMethod As Cost
    Data As String
End Type

'定义枚举集合 Cost
Public Enum Cost
    AllCost = 1
    GroupCost = 2
End Enum

'创建字符串包 CreatePackage
Public Function CreatePackage(InformationPackage As SockPack) As String
    Dim UserFrame As String, DataFrame As String, Group As String, _
        Is_Group As Boolean
    '组装 SendUser 内容 并赋入 UserFrame 变量
    UserFrame = "<SendUser:" & InformationPackage.SendUser & ">"
    '组装 SendMethod 内容 并赋入 UserFrame 变量
    UserFrame = UserFrame & "{" & InformationPackage.SendMethod & "}"
    '如果 发送方法为 AllCost(广播) 则 Is_Group 为 False 否则 Is_Group 为 True
    If InformationPackage.SendMethod = AllCost Then
        Is_Group = False
    Else
        '组装 AcceptUser 内容 并赋入 Group 变量
        For i = LBound(InformationPackage.AcceptUser) To UBound(InformationPackage.AcceptUser)
            Group = Group & InformationPackage.AcceptUser(i) & ";"
        Next
        Is_Group = True: Group = Left(Group, Len(Group) - 1)
    End If
    '把 Group 变量 赋入 UserFrame 变量
    UserFrame = UserFrame & "[AcceptUser:" & IIf(Is_Group, Group, "None") & "]"
    '组装 Data 内容 并赋入 DataFrame 变量
    DataFrame = "(Data:" & InformationPackage.Data & ")"
    '组装 UserFrame 和 DataFrame 内容 并添加 Package 标签
    CreatePackage = "<+Package->" & UserFrame & DataFrame & "<-Package+>"
End Function

Public Function ReadPackage(ByVal SockPackage As String, Result As SockPack) As Boolean
    If GetLabelText(SockPackage, "<+", "->", False) = "Package" And GetLabelText(SockPackage, "<-", "+>", False) = "Package" Then
        SockPackage = GetLabelText(SockPackage, "<+Package->", "<-Package+>", False)
        Result.SendUser = GetLabelText(SockPackage, "<SendUser:", ">", False)
        Result.SendMethod = Val(GetLabelText(SockPackage, "{", "}", False))
        If Result.SendMethod = GroupCost Then
            Result.AcceptUser = Split(GetLabelText(SockPackage, "[AcceptUser:", "]", False), ";")
        Else
            ReDim Result.AcceptUser(0)
            Result.AcceptUser(0) = "AllHost"
        End If
        Result.Data = GetLabelText(SockPackage, "(Data:", ")", False)
        ReadPackage = True
    Else
        ReadPackage = False
    End If
End Function

'获取标签内容 GetLabelText
Public Function GetLabelText(ByVal Data As String, LBorder As String, RBorder As String, Is_Border As Boolean) As String
    On Error Resume Next
    Dim LBorderPosition As Long, RBorderPosition As Long, _
        LabelTextLen As Long
    '查找 左边框 与 右边框 的位置
    LBorderPosition = InStr(1, Data, LBorder)
    RBorderPosition = InStr(1, Data, RBorder)
    
    '如果 Is_Border 为 True 把右边框的长度添加 否则 把左边框的长度添加
    If Is_Border Then
        RBorderPosition = RBorderPosition + Len(RBorder)
    Else
        LBorderPosition = LBorderPosition + Len(LBorder)
    End If
    
    '获取标签长度
    LabelTextLen = RBorderPosition - LBorderPosition
    '截取标签内容
    GetLabelText = Mid(Data, LBorderPosition, LabelTextLen)
End Function

Public Sub ClientPirnt(Text As TextBox, Frame As SockPack)
    Dim AcceptUserList As String
    On Error Resume Next
    If Frame.SendMethod = GroupCost Then
        For i = LBound(Frame.AcceptUser) To UBound(Frame.AcceptUser)
            AcceptUserList = AcceptUserList & Frame.AcceptUser(i) & ";"
        Next
        AcceptUserList = Left(AcceptUserList, Len(AcceptUserList) - 1)
    Else
        AcceptUserList = "AllHost"
    End If
    Text.Text = Text.Text & vbCrLf & _
                Frame.SendUser & " Send to " & AcceptUserList & ":" & vbCrLf & _
                Frame.Data
End Sub

Public Sub ServerLog(ByVal Text As TextBox, SockEquipment As Winsock, Frame As SockPack)
    For i = LBound(Frame.AcceptUser) To UBound(Frame.AcceptUser)
        AcceptUserBuffer = AcceptUserBuffer & Frame.AcceptUser(i) & ";"
    Next
    AcceptUserBuffer = Left(AcceptUserBuffer, Len(AcceptUserBuffer) - 1)
    Text.Text = Text.Text & vbCrLf & _
                Frame.SendUser & " send to " & AcceptUserBuffer & vbCrLf & _
                "SendMethod:" & Frame.SendMethod & "  Port:" & SockEquipment.LocalPort & vbCrLf & _
                "Data:" & Frame.Data
End Sub

Public Sub SaveData(Data As String)
    Num = FreeFile
    DateStr = Trim(Str(Year(Date))) & "-" & Trim(Str(Month(Date))) & "-" & Trim(Str(Day(Date)))
    Open App.Path & "\" & DateStr & ".txt" For Append As #Num
        Print #Num, Data
    Close
End Sub
