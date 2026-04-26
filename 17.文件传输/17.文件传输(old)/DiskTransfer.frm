VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form DiskTransfer 
   Caption         =   "DiskTransfer"
   ClientHeight    =   9735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   12150
   StartUpPosition =   3  '窗口缺省
   Begin ComctlLib.ListView DownloadListView 
      Height          =   1695
      Left            =   240
      TabIndex        =   21
      Top             =   6240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin ComctlLib.ProgressBar FileProgress 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   9360
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   19
      Top             =   8955
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TreeView LocalTreeView 
      Height          =   5175
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9128
      _Version        =   327682
      Indentation     =   706
      LineStyle       =   1
      Style           =   7
      Appearance      =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6960
      TabIndex        =   14
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Timer DataConnect 
      Interval        =   500
      Left            =   11640
      Top             =   6840
   End
   Begin VB.Timer SockState 
      Interval        =   1000
      Left            =   11640
      Top             =   7200
   End
   Begin VB.Timer CloseMsg 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   11640
      Top             =   7560
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      _Version        =   327681
      FullWidth       =   233
      FullHeight      =   65
   End
   Begin VB.Frame Frame1 
      Caption         =   "对方请求"
      Height          =   1455
      Left            =   10440
      TabIndex        =   7
      Top             =   6360
      Width           =   1095
      Begin VB.OptionButton Option1 
         Caption         =   "询问"
         Height          =   200
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "同意"
         Height          =   200
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "拒绝"
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock DirList 
      Left            =   10680
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock TransferDataFile 
      Left            =   10200
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame RemoteArea 
      Caption         =   "远程端"
      Height          =   6015
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
      Begin VB.DirListBox RemoteDir 
         Height          =   510
         Left            =   4200
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.FileListBox RemoteFile 
         Height          =   450
         Left            =   4200
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.DriveListBox RemoteDrive 
         Height          =   300
         Left            =   4200
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox RemotePath 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   5520
         Width           =   5535
      End
      Begin ComctlLib.TreeView RemoteTreeView 
         Height          =   5175
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   9128
         _Version        =   327682
         Indentation     =   706
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
   End
   Begin VB.Frame LocalArea 
      Caption         =   "本地端"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox LocalPath 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   5520
         Width           =   5535
      End
      Begin VB.DriveListBox LocalDrive 
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.FileListBox LocalFile 
         Height          =   450
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.DirListBox LocalDir 
         Height          =   510
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskTransfer.frx":0000
            Key             =   "computer"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskTransfer.frx":01DA
            Key             =   "link"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskTransfer.frx":03B4
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskTransfer.frx":061A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskTransfer.frx":0880
            Key             =   "open"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskTransfer.frx":0AE6
            Key             =   "file"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   10320
      TabIndex        =   13
      Top             =   8160
      Width           =   1215
   End
End
Attribute VB_Name = "DiskTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10

Public TransferID As Boolean
Private RNode As Node, _
        MsgMindFlag As Boolean, CurrentFlag As Boolean, TransferDataFlag As Boolean, TransferError As Boolean
        

Private Sub TransferMind()
    Dim FileInfo As FileTransfer
    FileInfo.File = DownloadListView.ListItems.Item(1).SubItems(2)
    Select Case DownloadListView.ListItems.Item(1).SubItems(4)
        Case "询问"
            If DownloadListView.ListItems.Item(1).SubItems(1) = "上传" Then
                CloseMsg.Enabled = True
                If MessageBox(DiskTransfer.hwnd, "同意" & DownloadListView.ListItems.Item(1).SubItems(2) & "文件进行传输ma~？" & vbCrLf & "(注:不交互自动拒绝。)", "Mind", 64) = 1 Then
                    CloseMsg.Enabled = False
                    If Not CurrentFlag Then
                        DownloadListView.ListItems.Item(1).SubItems(4) = "同意"
                        FileInfo.IsAccept = Yes: FileInfo.FileTotal = FileLen(FileInfo.File)
                        DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo)
                    End If
                    CurrentFlag = False
                End If
            End If
        Case "拒绝"
            DownloadListView.ListItems.Remove 1
        Case "同意"
            If DownloadListView.ListItems.Item(1).SubItems(1) = "下载" And (Not TransferDataFlag) Then
                Status.Panels.Item(4).Text = FileInfo.File
                FileInfo.IsACK = True: FileInfo.ErrorInfo = "接受文件中"
                DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo): TransferDataFlag = True
            End If
    End Select
End Sub

Private Sub CloseMsg_Timer()
    Dim MindHWND As Long, FileInfo As FileTransfer
    FileInfo.File = DownloadListView.ListItems.Item(1).SubItems(2)
    FileInfo.IsAccept = No
    DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo)
    DownloadListView.ListItems.Remove 1: CurrentFlag = True
    MindHWND = FindWindow(vbNullString, "Mind")
    SendMessage MindHWND, WM_CLOSE, 0, 0
    CloseMsg.Enabled = False: MsgMindFlag = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Chat.TransferFlag = False
End Sub

Private Sub SockState_Timer()
    Static Add As Long, FilePoint As Long
    If DownloadListView.ListItems.Count >= 1 Then
        TransferMind
        If MsgMindFlag Then MsgMindFlag = False
    End If
    If TransferDataFlag Then
        FilePoint = FileProgress.Value
        Status.Panels.Item(8).Text = Transfer.BitList(FilePoint) & "/" & DownloadListView.ListItems.Item(1).SubItems(3)
        Status.Panels.Item(6).Text = Transfer.BitList(CLng(FileProgress.Value) - FilePoint)
        If Add <> 0 Then Exit Sub
        Add = Add + 1
        Animation1.Open App.Path & "\cartoon-FileTransfer.avi"
        Animation1.Play
    Else
        Add = 0
        Animation1.Close
    End If
    If TransferError Then
        Dim FileInfo As FileTransfer
        FileInfo.File = DownloadListView.ListItems.Item(1).Text
        TCP_Transfer FileInfo
    End If
    Label1.Caption = TransferDataFile.State & DirList.State & CloseMsg.Enabled
End Sub

Private Sub Command1_Click()
    Static a As Integer
    If a Mod 2 = 0 Then
        Animation1.Open App.Path & "\cartoon-FileTransfer.avi"
        Animation1.Play
    Else
        Animation1.Close
    End If
    a = a + 1
    Status.Panels.Item(2).Text = "文件最大传输2GB以内。"
    Status.Panels.Item(4).Text = "C:\1.txt"
    Status.Panels.Item(6).Text = "1023MB/s"
    Status.Panels.Item(8).Text = "1023MB/1032MB"
End Sub

Private Sub Form_Load()
    LocalTreeView.ImageList = ImageList1
    LocalTreeView.LabelEdit = tvwManual
    RemoteTreeView.ImageList = ImageList1
    RemoteTreeView.LabelEdit = tvwManual
    LocalTreeView.Nodes.Add , , "LocalRoot", "计算机", "computer"
    RemoteTreeView.Nodes.Add , , "RemoteRoot", "NONE", "link"
    For i = 0 To LocalDrive.ListCount - 1
        LocalTreeView.Nodes.Add "LocalRoot", tvwChild, "L" & Left(LocalDrive.List(i), 1), Left(LocalDrive.List(i), 2), "drive"
    Next
    With DownloadListView.ColumnHeaders
        .Add 1, , "请求", 300
        .Add 2, , "传输方式", 700
        .Add 3, , "文件名", 5000
        .Add 4, , "文件大小"
        .Add 5, , "请求"
    End With
    With Status.Panels
        .Add(1, "ls", DirList.RemoteHostIP & "状态:").Width = 1900
        .Add(2, "remotestate", "NONE").Width = 2700
        .Add(3, "fn", "文件:").Width = 600
        .Add(4, "fileinfo").Width = 2800
        .Add(5, "fs", "传输速度:").Width = 1000
        .Add(6, "transferspeed", "0B/s").Width = 900
        .Add(7, "fp", "进度:").Width = 600
        .Add(8, "fileprogress").Width = 1500
    End With
End Sub

Private Function RequestState(Data As FileTransfer) As String
    For i = 0 To 2
        If Option1(i).Value Then RequestState = Option1(i).Caption: Data.IsAccept = i + 1: Exit For
    Next
    DirList.SendData SockPackDispose.CreateFileTransferPackage(Data)
End Function

Private Sub TCP_Transfer(FileInfo As FileTransfer)
    If Transfer.FileMemoryCheck(FileInfo.File) Then
        FileInfo.ErrorInfo = "文件读取中。": TransferDataFlag = True
        DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo)
        Transfer.FileTransfer FileInfo.File, TransferDataFile, FileProgress
        DownloadListView.ListItems.Remove 1
        FileInfo.ErrorInfo = "文件读取完成。"
        DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo)
    Else
        FileInfo.ErrorInfo = "发送失败。Error:内存空间不足"
        DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo): TransferError = True
    End If
End Sub

Private Sub DirList_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim Data As String, Head As String
    DirList.GetData Data
    SockPackDispose.GetPackageHead Data, Head
    Select Case Head
        Case "FileTransfer"
            Dim FileInfo As FileTransfer
            SockPackDispose.ReadFileTransferPackage Data, FileInfo, True
            Status.Panels.Item(2).Text = FileInfo.ErrorInfo
            Select Case True
                Case FileInfo.IsRequest
                    FileInfo.IsRequest = False
                    '添加至队列
                    Dim Item As ListItem
                    Index = DownloadListView.ListItems.Count
                    For i = 1 To Index
                        If DownloadListView.ListItems.Item(i).SubItems(2) = FileInfo.File Then Exit For
                    Next
                    If i > Index Then
                        Set Item = DownloadListView.ListItems.Add(i)
                        With Item
                            .SubItems(1) = "上传"
                            .SubItems(2) = FileInfo.File
                            .SubItems(3) = Transfer.BitList(FileLen(FileInfo.File))
                            .SubItems(4) = RequestState(FileInfo)
                        End With
                    End If
                    If DownloadListView.ListItems.Item(1).SubItems(4) = "询问" Then MsgMindFlag = True
                Case (FileInfo.IsAccept = No Or FileInfo.IsAccept = Yes Or FileInfo.IsAccept = Inquire)
                    For i = 1 To DownloadListView.ListItems.Count
                        If FileInfo.File = DownloadListView.ListItems.Item(i).SubItems(2) Then
                            With DownloadListView.ListItems.Item(i)
                                .SubItems(4) = Switch(FileInfo.IsAccept = No, "拒绝", FileInfo.IsAccept = Yes, "同意", FileInfo.IsAccept = Inquire, "询问")
                                .SubItems(3) = Transfer.BitList(FileInfo.FileTotal)
                                If i = 1 Then
                                    If FileInfo.FileTotal < 0 Or DownloadListView.ListItems.Item(1).Text = "-1" Then Status.Panels.Item(2).Text = "文件最大传输2GB以内。": Exit Sub
                                    FileProgress.Max = FileInfo.FileTotal
                                End If
                            End With
                        End If
                    Next
                Case FileInfo.IsACK
                    FileInfo.IsACK = False
                    TCP_Transfer FileInfo
                Case FileInfo.IsFIN
                    TransferDataFlag = False
                    DownloadListView.ListItems.Remove 1
            End Select
        Case "Dir"
            Dim DirInfo As DriveInfo
            SockPackDispose.ReadDrivePackage Data, DirInfo, True
            Select Case True
                Case DirInfo.IsRequire
                    '发送本地磁盘数据
                    DirInfo.IsACK = True
                    RemoteDir.Refresh: RemoteFile.Refresh
                    RemoteDir.Path = DirInfo.Path
                    RemoteFile.Path = DirInfo.Path
                    ReDim DirInfo.Folder(RemoteDir.ListCount - 1)
                    For i = 0 To RemoteDir.ListCount - 1
                        DirInfo.Folder(i) = Transfer.GetPathName(RemoteDir.List(i), Folder)
                    Next
                    ReDim DirInfo.File(RemoteFile.ListCount - 1)
                    For i = 0 To RemoteFile.ListCount - 1
                        DirInfo.File(i) = RemoteFile.List(i)
                    Next
                    DirList.SendData SockPackDispose.CreateDrivePackage(DirInfo)
                Case DirInfo.IsACK
                    For i = LBound(DirInfo.Folder) To UBound(DirInfo.Folder)
                        RemoteTreeView.Nodes.Add RNode.Key, tvwChild, RNode.Key & "-" & i, DirInfo.Folder(i), "close"
                    Next
                    For j = LBound(DirInfo.File) To UBound(DirInfo.File)
                        RemoteTreeView.Nodes.Add RNode.Key, tvwChild, RNode.Key & "-" & i + j, DirInfo.File(j), "file"
                    Next
            End Select
    End Select
End Sub

Private Sub DownloadListView_ItemClick(ByVal Item As ComctlLib.ListItem)
    If Item.SubItems(1) = "下载" Then Exit Sub
    If Item.SubItems(1) = "上传" Then
        Dim FileInfo As FileTransfer
        FileInfo.File = Item.SubItems(2)
        FileInfo.FileTotal = FileLen(Item.SubItems(2))
        If FileInfo.FileTotal < 0 Then GoTo NoCode
        If (Item.SubItems(4) = "询问" Or Item.SubItems(4) = "拒绝") Then
            Item.SubItems(4) = "同意": FileInfo.IsAccept = Yes
        Else
NoCode:     Item.SubItems(4) = "拒绝": FileInfo.IsAccept = No
        End If
    End If
    DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo)
End Sub

Private Sub RefreshPath(Path As String, Key As String)
    LocalDir.Path = Path: LocalFile.Path = Path
    LocalDir.Refresh: LocalFile.Refresh
    For i = 0 To LocalDir.ListCount - 1
        LocalTreeView.Nodes.Add Key, tvwChild, Key & "-" & i, Transfer.GetPathName(LocalDir.List(i), Folder), "close"
    Next
    For j = 0 To LocalFile.ListCount - 1
        LocalTreeView.Nodes.Add Key, tvwChild, Key & "-" & i + j, LocalFile.List(j), "file"
    Next
End Sub

Private Sub LocalPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error GoTo AddKey
        Const PathHead = "计算机\"
        Dim tmpPath As String, a() As String
        tmpPath = IIf(InStr(1, LocalPath.Text, PathHead) <> 0, Mid(LocalPath.Text, 5), LocalPath.Text)
        Transfer.GetFolderPath tmpPath, a
        Dim i As Integer, Key As String
        Key = "L" & LCase(Left(a(0), 1))
        Do
            LocalTreeView.Nodes.Remove Key & "-" & i
            i = i + 1
        Loop
AddKey: For i = LBound(a) To UBound(a)
            RefreshPath a(i), Key
            '磁盘刷新
        Next
    End If
End Sub

Private Sub LocalTreeView_Expand(ByVal Node As ComctlLib.Node)
    If Node.Image = "close" Then Node.Image = "open"
End Sub

Private Sub LocalTreeView_Collapse(ByVal Node As ComctlLib.Node)
    If Node.Image = "open" Then Node.Image = "close"
End Sub

Private Sub LocalTreeView_NodeClick(ByVal Node As Node)
    On Error GoTo ExitProc
    Dim i As Integer
    LocalPath.Text = Node.FullPath
    If Node.Image = "open" Then Node.Image = "close"
    If Not Node.FullPath = "计算机" Then
        Do
            LocalTreeView.Nodes.Remove Node.Key & "-" & i
            i = i + 1
        Loop
ExitProc:
        If Node.Image = "close" Or Node.Image = "drive" Then
            RefreshPath Mid(Node.FullPath, 5) & "\", Node.Key
        End If
    End If
    If Node.Image = "file" Then
        Dim List As ListItem
        Index = DownloadListView.ListItems.Count + 1
        Set List = DownloadListView.ListItems.Add(Index)
        List.SubItems(1) = "上传"
        List.SubItems(2) = Mid(Node.FullPath, InStr(1, Node.FullPath, "\") + 1)
        List.SubItems(3) = Transfer.BitList(FileLen(List.SubItems(2)))
        List.SubItems(4) = "同意"
    End If
End Sub

Private Sub RemoteTreeView_Expand(ByVal Node As ComctlLib.Node)
    If Node.Image = "close" Then Node.Image = "open"
End Sub

Private Sub RemoteTreeView_Collapse(ByVal Node As ComctlLib.Node)
    If Node.Image = "open" Then Node.Image = "close"
End Sub

Private Sub RemoteTreeView_NodeClick(ByVal Node As ComctlLib.Node)
    RemotePath.Text = Node.FullPath
    If Node.Image = "file" Then
        Dim FileInfo As FileTransfer, Item As ListItem, _
            TempFile As String
        Index = DownloadListView.ListItems.Count
        TempFile = Mid(Node.FullPath, InStr(1, Node.FullPath, "\") + 1)
        For j = 1 To Index
            If TempFile = DownloadListView.ListItems.Item(j).SubItems(2) Then Exit For
        Next
        If j > Index Then
            FileInfo.File = TempFile
            FileInfo.IsRequest = True
            DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo)
            Set Item = DownloadListView.ListItems.Add(j)
            Item.SubItems(1) = "下载"
            Item.SubItems(2) = FileInfo.File
        End If
    End If
    If Node.Image = "drive" Or Node.Image = "close" Then
        On Error GoTo ExitProc
        Dim DriveInfo As DriveInfo, i As Integer
        DriveInfo.Path = Mid(Node.FullPath, InStr(1, Node.FullPath, "\") + 1) & "\"
            Do
                RemoteTreeView.Nodes.Remove Node.Key & "-" & i
                i = i + 1
            Loop
ExitProc:
        DriveInfo.IsRequire = True
        Set RNode = Node
        DirList.SendData SockPackDispose.CreateDrivePackage(DriveInfo)
    End If
End Sub

Private Sub DataConnect_Timer()
    If TransferDataFile.State <> 7 Then
        TransferDataFile.Close
        If TransferID Then TransferDataFile.Listen Else TransferDataFile.Connect
    Else
        DataConnect.Enabled = False
    End If
End Sub

Private Sub TransferDataFile_ConnectionRequest(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If TransferDataFile.State <> sckClosed Then TransferDataFile.Close
'接受具有 requestID 参数的
'连接。
TransferDataFile.Accept requestID
End Sub

Private Sub TransferDataFile_DataArrival(ByVal bytesTotal As Long)
    Dim bytData() As Byte
    Dim f
    f = FreeFile
    If DownloadListView.ListItems.Count > 0 Then
        Open "C:\" & Transfer.GetPathName(DownloadListView.ListItems.Item(1).SubItems(2), File) For Binary As #f ''strFileName是文件名
            lLenFile = LOF(f)
            ReDim bytData(1 To bytesTotal)
            TransferDataFile.GetData bytData
            'lLenFile=0表示是第一次打开文件，这里有个问题，就是
            '如果如果该文件存在的话，就会出错，应该在打开前检查文件是否存在。（这里我省略了）
            If lLenFile = 0 Then
                Put #f, 1, bytData
            Else
                Put #f, lLenFile + 1, bytData
            End If
            Status.Panels.Item(4).Text = DownloadListView.ListItems.Item(1).SubItems(2)
            If FileProgress.Max = FileProgress.Value Then
                Dim FileInfo As FileTransfer
                FileInfo.File = DownloadListView.ListItems.Item(1).SubItems(2)
                FileInfo.IsFIN = True: TransferDataFlag = False
                DirList.SendData SockPackDispose.CreateFileTransferPackage(FileInfo)
                DownloadListView.ListItems.Remove 1
            Else
                FileProgress.Value = lLenFile
            End If
        Close #f
    End If
End Sub

