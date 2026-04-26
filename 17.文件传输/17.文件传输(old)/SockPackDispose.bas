Attribute VB_Name = "SockPackDispose"
Private Head As String

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

Public Type HostAddress
    IP As String
    Port As Long
    IsAdd As Boolean
    IsDel As Boolean
    IsOnline As Boolean
End Type

Public Type FileTransfer
    File As String
    FileTotal As Long
    IsRequest As Boolean
    IsAccept As AcceptCode
    IsACK As Boolean
    IsFIN As Boolean
    ErrorInfo As String
End Type

Public Enum AcceptCode
    No = 1
    Yes = 2
    Inquire = 3
End Enum

Public Type DriveInfo
    Path As String
    IsRequire As Boolean
    IsACK As Boolean
    Folder() As String
    File() As String
End Type

Public Type Request
    Info As String
    IsRequest As Boolean
    IsAccept As Boolean
End Type

Public Function GetPackageHead(Package As String, Head As String) As Boolean
    FirstStr = GetLabelText(Package, "<+", "->", False)
    StrLength = Len(FirstStr) + 4
    SecondStr = GetLabelText(Right(Package, StrLength), "<-", "+>", False)
    GetPackageHead = (FirstStr = SecondStr) And Not (FirstStr = "" Or SecondStr = "")
    If GetPackageHead Then Package = GetLabelText(Package, "<+" & FirstStr & "->", "<-" & SecondStr & "+>", False): Head = FirstStr
End Function

Public Function CreateRequestPackage(RequestInfo As Request) As String
    DataInfo = "(" & RequestInfo.Info & ")"
    If RequestInfo.IsRequest Then Flag = "?Request!": GoTo NextLine
    Flag = IIf(RequestInfo.IsAccept, "?YES!", "?NO!")
NextLine:
    CreateRequestPackage = "<+Request->" & Flag & DataInfo & "<-Request+>"
End Function

Public Function ReadRequestPackage(ByVal RequestPackage As String, Result As Request, Optional IsNoHead As Boolean) As Boolean
    If IsNoHead Then GoTo Code
    GetPackageHead RequestPackage, Head
    If Head = "Request" Then
Code:
        Result.Info = GetLabelText(RequestPackage, "(", ")", False)
        Result.IsRequest = CBool(InStr(1, RequestPackage, "?Request!"))
        Result.IsAccept = IIf(GetLabelText(RequestPackage, "?", "!", False) = "YES", True, False)
        ReadRequestPackage = True
    Else
        ReadRequestPackage = False
    End If
End Function

Public Function CreateDrivePackage(DriveInfo As DriveInfo) As String
    On Error Resume Next
    Dim Folder As String, File As String
    DataPath = "/?" & DriveInfo.Path & "?\"
    If DriveInfo.IsRequire Then Flag = "?Require!"
    If DriveInfo.IsACK Then
        Flag = "?ACK!"
        For i = LBound(DriveInfo.Folder) To UBound(DriveInfo.Folder)
            Folder = Folder & DriveInfo.Folder(i) & "|"
        Next
        Folder = Left(Folder, Len(Folder) - 1)
        For i = LBound(DriveInfo.File) To UBound(DriveInfo.File)
            File = File & DriveInfo.File(i) & "|"
        Next
        File = Left(File, Len(File) - 1)
        DataFolder = "/*" & Folder & "*\"
        DataFile = "/*:" & File & ":*\"
    End If
    CreateDrivePackage = "<+Dir->" & Flag & DataPath & DataFolder & DataFile & "<-Dir+>"
End Function

Public Function ReadDrivePackage(ByVal DrivePackage As String, DriveInfo As DriveInfo, Optional IsNoHead As Boolean) As Boolean
    If IsNoHead Then GoTo Code
    GetPackageHead DrivePackage, Head
    If Head = "Dir" Then
Code:
        DriveInfo.Path = GetLabelText(DrivePackage, "/?", "?\", False)
        DriveInfo.IsRequire = CBool(InStr(1, DrivePackage, "?Require!"))
        If InStr(1, DrivePackage, "?ACK!") <> 0 Then
            DriveInfo.IsACK = True
            DriveInfo.Folder = Split(GetLabelText(DrivePackage, "/*", "*\", False), "|")
            DriveInfo.File = Split(GetLabelText(DrivePackage, "/*:", ":*\", False), "|")
        End If
        ReadDrivePackage = True
    Else
        ReadDrivePackage = False
    End If
End Function

Public Function CreateFileTransferPackage(FileTransferInfo As FileTransfer) As String
    Data = "/*:" & FileTransferInfo.File & ":*\{" & FileTransferInfo.FileTotal & "}"
    Data = Data & IIf(FileTransferInfo.ErrorInfo <> "", "{-(" & FileTransferInfo.ErrorInfo & ")-}", "")
    If FileTransferInfo.IsRequest Then TempInfo = "?Request!": GoTo NextLine
    If FileTransferInfo.IsACK Then TempInfo = "?ACK!": GoTo NextLine
    If FileTransferInfo.IsFIN Then TempInfo = "?FIN!": GoTo NextLine
    If FileTransferInfo.IsAccept <> 0 Then TempInfo = "?+" & FileTransferInfo.IsAccept & "-!"
NextLine:
    CreateFileTransferPackage = "<+FileTransfer->" & TempInfo & Data & "<-FileTransfer+>"
End Function

Public Function ReadFileTransferPackage(ByVal FileTransferPackage As String, Result As FileTransfer, Optional IsNoHead As Boolean) As Boolean
    If IsNoHead Then GoTo Code
    GetPackageHead FileTransferPackage, Head
    If (Head = "FileTransfer") Then
Code:
        ReadFileTransferPackage = True
        Result.File = GetLabelText(FileTransferPackage, "/*:", ":*\", False)
        Result.FileTotal = GetLabelText(FileTransferPackage, "{", "}", False)
        Result.IsRequest = CBool(InStr(1, FileTransferPackage, "?Request!"))
        Result.IsACK = CBool(InStr(1, FileTransferPackage, "?ACK!"))
        Result.IsFIN = CBool(InStr(1, FileTransferPackage, "?FIN!"))
        Result.ErrorInfo = GetLabelText(FileTransferPackage, "{-(", ")-}", False)
        Result.IsAccept = Val(GetLabelText(FileTransferPackage, "?+", "-!", False))
    Else
        ReadFileTransferPackage = False
    End If
End Function

Public Function CreateHostPackage(HostInfo As HostAddress) As String
    DataIP = "(+" & HostInfo.IP & "-)"
    If HostInfo.IsAdd Then Flag = "?ADD!": DataPort = "[+" & HostInfo.Port & "-]"
    If HostInfo.IsOnline Then Flag = "?Online!": DataPort = "[+" & HostInfo.Port & "-]"
    If HostInfo.IsDel Then Flag = "?DEL!"
    CreateHostPackage = "<+Host->" & Flag & DataIP & DataPort & "<-Host+>"
End Function

Public Function ReadHostPackage(ByVal HostPackage As String, Result As HostAddress, Optional IsNoHead As Boolean) As Boolean
    If IsNoHead Then GoTo Code
    GetPackageHead HostPackage, Head
    If Head = "Host" Then
Code:
        ReadHostPackage = True
        Result.IP = GetLabelText(HostPackage, "(+", "-)", False)
        If CBool(InStr(1, HostPackage, "?ADD!")) Then Result.IsAdd = True: Result.Port = GetLabelText(HostPackage, "[+", "-]", False)
        If CBool(InStr(1, HostPackage, "?Online!")) Then Result.IsOnline = True: Result.Port = GetLabelText(HostPackage, "[+", "-]", False)
        If CBool(InStr(1, HostPackage, "?DEL!")) Then Result.IsDel = True
    Else
        ReadHostPackage = False
    End If
End Function

'创建字符串包 CreateInfoPackage
Public Function CreateInfoPackage(InformationPackage As SockPack) As String
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
    DataFrame = "(Data-" & InformationPackage.Data & "-Data)"
    '组装 UserFrame 和 DataFrame 内容 并添加 Package 标签
    CreateInfoPackage = "<+Package->" & UserFrame & DataFrame & "<-Package+>"
End Function

Public Function ReadInfoPackage(ByVal SockPackage As String, Result As SockPack, Optional IsNoHead As Boolean) As Boolean
    If IsNoHead Then GoTo Code
    GetPackageHead SockPackage, Head
    If Head = "Package" Then
Code:
        Result.SendUser = GetLabelText(SockPackage, "<SendUser:", ">", False)
        Result.SendMethod = Val(GetLabelText(SockPackage, "{", "}", False))
        If Result.SendMethod = GroupCost Then
            Result.AcceptUser = Split(GetLabelText(SockPackage, "[AcceptUser:", "]", False), ";")
        Else
            ReDim Result.AcceptUser(0)
            Result.AcceptUser(0) = "AllHost"
        End If
        Result.Data = GetLabelText(SockPackage, "(Data-", "-Data)", False)
        ReadInfoPackage = True
    Else
        ReadInfoPackage = False
    End If
End Function

'获取标签内容 GetLabelText
Public Function GetLabelText(ByVal Data As String, LBorder As String, RBorder As String, Is_Border As Boolean) As String
    'On Error Resume Next
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
    If LabelTextLen < 0 Then Exit Function
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

Public Function StateReturn(Equipment As Winsock) As String
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

