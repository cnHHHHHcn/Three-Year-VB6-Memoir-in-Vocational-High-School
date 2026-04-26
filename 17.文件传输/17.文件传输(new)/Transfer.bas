Attribute VB_Name = "Transfer"
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
 
Private Type MEMORYSTATUS
    dwLength As Long           '没用
    dwMemoryLoad As Long       '正在使用的物理内存的近似百分比
    dwTotalPhys As Long        '实际物理内存量（以字节为单位）
    dwAvailPhys As Long        '当前可用的物理内存量（以字节为单位）
    dwTotalPageFile As Long    '系统或当前进程的当前已提交内存限制，以字节为单位，以较小者为准
    dwAvailPageFile As Long    '当前进程可以提交的最大内存量（以字节为单位）
    dwTotalVirtual As Long     '当前位于调用进程的虚拟地址空间的用户模式部分中的未保留和未提交的内存量（以字节为单位）。
    dwAvailVirtual As Long     '保留。 此值始终为 0。
End Type

Private Const iMax As Long = 65535

Public Enum DOS
    Folder = 1
    File = 2
End Enum

Public Function BitList(ByVal BitTotal As Long) As String
    Dim Total As Single
    Unit = Array("B", "KB", "MB", "GB")
    If BitTotal >= 0 Then Total = BitTotal Else BitList = "-1": Exit Function
    Do Until Total < 1024
        Total = Round(Total / 1024, 2)
        i = i + 1
    Loop
    BitList = Total & Unit(i)
End Function

Public Function FileMemoryCheck(File As String) As Boolean
    Dim MemoryInfo As MEMORYSTATUS, _
        SuperiorlimitMemory As Long, CurrentMemory As Long, FileTotal As Long
    FileTotal = FileLen(File)
    If FileTotal <= 0 Then Exit Function
    GlobalMemoryStatus MemoryInfo
    SuperiorlimitMemory = MemoryInfo.dwTotalPhys * 0.8
    CurrentMemory = MemoryInfo.dwAvailPhys
    FileMemoryCheck = (FileTotal + CurrentMemory <= SuperiorlimitMemory)
End Function

Public Sub GetFolderPath(ByVal Path As String, SubPath() As String)
    Dim Pos As Integer, Index As Integer
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Debug.Print Path
    Do Until Pos >= Len(Path)
        ReDim Preserve SubPath(Index)
        Pos = InStr(Pos + 1, Path, "\")
        SubPath(Index) = Mid(Path, 1, Pos)
        Debug.Print SubPath(Index)
        Index = Index + 1
    Loop
End Sub

Public Function GetPathName(ByVal Path_String As String, Flag As DOS) As String
    Do While Path_String <> GetFolderName
        GetFolderName = Path_String
        Path_String = Right(Path_String, Len(Path_String) - InStr(Path_String, "\"))
    Loop
    If Flag = 1 Then
        If Path_String <> "" Then GetPathName = Path_String Else GetPathName = "None"
    Else
        If CBool(InStr(1, Path_String, ".")) Then GetPathName = Path_String Else GetPathName = "None"
    End If
End Function

Public Function FileTransfer(File As String, Transfersock As Winsock, Progress As ProgressBar) As Boolean
    If File = "" Or Transfersock.State <> 7 Then Exit Function
    FileTransfer = True
    Dim FreeF As Integer
    Dim bytData() As Byte
    FreeF = FreeFile '获得空闲的文件号
    Open File For Binary As #FreeF '打开文件
        DoEvents
        LenFile = LOF(FreeF) '获得文件长度
        Progress.Max = LenFile
        
        If LenFile <= iMax Then '如果要发送的文件小于数据块大小，直接发送
            ReDim bytData(1 To LenFile) '根据文件长度重新定义数组大小
            Get #FreeF, , bytData '把文件读入到数组里
            Transfersock.SendData bytData '发送数据
            GoTo CloseFile
        End If
        Do Until (iPos >= (LenFile - iMax)) '发送整块数据的循环
            ReDim bytData(1 To iMax)
            Get #FreeF, iPos + 1, bytData
            DoEvents
            Transfersock.SendData bytData
            iPos = iPos + iMax '移动iPos，使它指向下来要读的数据
            Progress.Value = iPos
        Loop
        '这里要注意的是，必须检查文件有没有剩下的数据，如果文件大小正好等于数据块大小的
        '整数倍，那么就没有剩下的数据了，发送剩下的不够一个数据块的数据
        ReDim bytData(1 To LenFile - iPos)
        Get #FreeF, iPos + 1, bytData
        Transfersock.SendData bytData
CloseFile:
    Close #FreeF
End Function
