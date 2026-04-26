VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form unControlForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "解控端"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   10980
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5820
   Begin VB.CommandButton WriteScanList 
      Caption         =   "写入文件"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Scan 
      Caption         =   "扫描(关闭)"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ListBox ScanList 
      Appearance      =   0  'Flat
      Height          =   1470
      Left            =   3240
      TabIndex        =   10
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Chat 
      Caption         =   "同意"
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Chat 
      Caption         =   "拒绝"
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.ListBox ChatRequestList 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Close 
      Caption         =   "关闭"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Open 
      Caption         =   "开启"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox TeacherIP 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Text            =   "192.168.1.60"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox ControlPort 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Text            =   "1689"
      Top             =   360
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ScanIP 
      Left            =   5280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "扫描结果："
      Height          =   180
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "聊天请求 IP："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "教师 IP："
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   760
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "控制端口："
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "unControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1



Private CloseData() As Byte, OpenData() As Byte, ChatData() As Byte

Private Sub Chat_Click(Index As Integer)
    For i = ChatRequestList.ListCount - 1 To 0 Step -1
        Pos = InStr(1, ChatRequestList.List(i), ":")
        Winsock.RemoteHost = Left(ChatRequestList.List(i), Pos - 1)
        Winsock.RemotePort = Mid(ChatRequestList.List(i), Pos + 1)
        If ChatRequestList.Selected(i) Then
            Winsock.SendData "Chat:" & IIf(CBool(Index), "Yes", "No")
            If CBool(Index) Then
                Winsock.RemotePort = Val(ControlPort.Text)
                IPbytes = Split(Winsock.RemoteHostIP, ".")
                TeacherIPBytes = Split(TeacherIP.Text, ".")
                For j = 0 To 3
                    ChatData(UBound(ChatData) - 7 + j) = Val(TeacherIPBytes(j))
                    ChatData(UBound(ChatData) - 3 + j) = Val(IPbytes(j))
                Next
                Winsock.SendData ChatData
            End If
            ChatRequestList.RemoveItem i
        End If
    Next
End Sub

Private Sub WriteScanList_Click()
    Client.Config.Class = "ScanIP"
    Client.Config.DeleteClass
    Client.Config.CreateClass
    For i = 0 To ScanList.ListCount - 1
        Client.Config.CreateKey "", "", ScanList.List(i)
    Next
End Sub

Private Sub Open_Click()
    SetWinsock
End Sub

Private Sub Close_Click()
    Winsock.Close
End Sub

Private Sub Form_Load()
    ReadPackage App.Path & "\Package\Close.txt", CloseData
    ReadPackage App.Path & "\Package\Open.txt", OpenData
    ReadPackage App.Path & "\Package\Chat.txt", ChatData
    TeacherIP.Text = ConfigMod.TeacherInfo.IP: SetWinsock
    ControlPort.Text = ConfigMod.TeacherInfo.Port
End Sub

Private Sub SetWinsock()
    Winsock.Close
    Winsock.Protocol = sckUDPProtocol
    Winsock.Bind 45000, Winsock.LocalIP
End Sub

Public Function ReadPackage(File As String, Bytes() As Byte) As Boolean
    On Error Resume Next
    Dim Temp As String, TempArray() As String
    Erase Bytes: freenum = FreeFile
    Open File For Input As #freenum
        Do Until EOF(freenum)
            Line Input #1, Temp
            TempArray = Split(Temp, ":")
            ReDim Bytes(UBound(TempArray) - LBound(TempArray) + 1)
            For i = LBound(TempArray) To UBound(TempArray)
                Bytes(i + 1) = CLng("&H" & TempArray(i))
            Next
        Loop
    Close #freenum
End Function

Private Sub Scan_Click()
    If Mid(Scan.Caption, 4, 2) = "关闭" Then
        ScanIP.Close
        ScanIP.Bind Val(ControlPort.Text)
        Scan.Caption = "扫描(开启中)"
    Else
        ScanIP.Close
        Scan.Caption = "扫描(关闭)"
    End If
End Sub

Private Sub ScanIP_DataArrival(ByVal bytesTotal As Long)
    Dim PackageBytes() As Byte, HostName As String
    ScanIP.GetData PackageBytes
    If (PackageBytes(0) = 1 And PackageBytes(1) = 1 And PackageBytes(2) = 0 And PackageBytes(3) = 0 And PackageBytes(4) = 160) Then
        For i = ScanList.ListCount - 1 To 0 Step -1
            If CBool(InStr(1, ScanList.List(i), ScanIP.RemoteHostIP)) Then Exit For
        Next
        If i = -1 Then
            For i = 14 To 31
                HostName = HostName & Chr(PackageBytes(i))
            Next
            HostName = Left(HostName, InStr(1, HostName, Chr(0)) - 1)
            ScanList.AddItem (HostName & ";" & ScanIP.RemoteHostIP)
        End If
    End If
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrorArea
    Dim Temp As String, TempIPBytes() As String, TeacherIPBytes() As String
    Winsock.GetData Temp
    If Temp = "ChatRequest" Then ChatRequestList.AddItem Winsock.RemoteHostIP & ":" & Winsock.RemotePort: Exit Sub
    If Temp = "unControl" Then
        TempIPBytes = Split(Winsock.RemoteHostIP, ".")
        TeacherIPBytes = Split(TeacherIP.Text, ".")
        For i = 0 To 3
            OpenData(13 + i) = TeacherIPBytes(i)
            CloseData(UBound(CloseData) - 7 + i) = Val(TeacherIPBytes(i))
            OpenData(UBound(OpenData) - 7 + i) = Val(TeacherIPBytes(i))
            CloseData(UBound(CloseData) - 3 + i) = Val(TempIPBytes(i))
            OpenData(UBound(OpenData) - 3 + i) = Val(TempIPBytes(i))
        Next
        For i = LBound(OpenData) To UBound(OpenData)
            Debug.Print OpenData(i)
        Next
        Winsock.Close
        Randomize
        Winsock.Bind Int(Rnd * 15536) + 50000, Winsock.LocalIP
        Winsock.RemoteHost = Winsock.RemoteHostIP
        Winsock.RemotePort = Val(ControlPort.Text)
        Winsock.SendData CloseData
        Winsock.SendData OpenData
        SetWinsock
    End If
ErrorArea:
    Exit Sub
End Sub
