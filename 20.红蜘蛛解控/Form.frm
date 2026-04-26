VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Client 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3915
   ClientLeft      =   75
   ClientTop       =   11070
   ClientWidth     =   4875
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4875
   Begin VB.Frame Frame2 
      Caption         =   "联网"
      Height          =   1455
      Left            =   2280
      TabIndex        =   13
      Top             =   1920
      Width           =   2295
      Begin VB.TextBox SubNet 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   19
         Text            =   "255.255.255.0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox ControlIP 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton OpenOnline 
         Caption         =   "打开联网"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CloseOnline 
         Caption         =   "取消联网"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "子网掩码:"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "控制 IP:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "教师信息"
      Height          =   1695
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton GetTeacherInfo 
         Caption         =   "抓取(关闭)"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox TeacherHost 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   760
         TabIndex        =   11
         Text            =   "teacher"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TeacherIP 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   400
         TabIndex        =   9
         Text            =   "192.168.1.60"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "IP:"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   880
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "主机名:"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   380
         Width           =   630
      End
   End
   Begin VB.CommandButton WindowAuto 
      Caption         =   "窗口调节"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton ChatRequest 
      Caption         =   "聊天请求(关闭)"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Timer OpenServer 
      Interval        =   100
      Left            =   1800
      Top             =   600
   End
   Begin VB.CommandButton CheckControl 
      Caption         =   "自动解控(关闭)"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox RemoteServerIP 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton KillREDA 
      Caption         =   "杀死红蜘蛛进程"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   1800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton ShowREDAWindow 
      Caption         =   "窗口化"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "远程服务器 IP:"
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Option Base 1



Private GetTeacherFlag As Boolean, ChatFlag As Boolean, CheckControlFlag As Boolean
Public Config As New OperateKey

Private Sub SetWinsock()
    Winsock.Close
    Winsock.Bind Int(15536 * Rnd) + 50000, Winsock.LocalIP
    Winsock.RemoteHost = RemoteServerIP.Text
    Winsock.RemotePort = 45000
End Sub

Private Sub ChatRequest_Click()
    SetWinsock
    Winsock.SendData "ChatRequest"
    ChatRequest.Caption = "聊天请求(请求中)": ChatFlag = True
    CheckControl.Enabled = False: ShowREDAWindow.Enabled = False
End Sub

Private Sub CheckControl_Click()
    Winsock.Close
    Winsock.Bind ConfigMod.TeacherInfo.Port: CheckControlFlag = True
    CheckControl.Caption = "自动解控(开启中)"
End Sub


Private Sub GetTeacherInfo_Click()
    Winsock.Close
    Winsock.Bind ConfigMod.TeacherInfo.Port
    GetTeacherFlag = True: GetTeacherInfo.Caption = "抓取(开启中)"
End Sub

Private Sub OpenOnline_Click()
    Shell "cmd.exe /c netsh interface ip set address name=" & "本地连接" & " source=DHCP"
    Shell "cmd.exe /c netsh interface set interface " & "本地连接" & " admin=disable"
    Shell "cmd.exe /c netsh interface set interface " & "本地连接" & " admin=enabled"
End Sub

Private Sub CloseOnline_Click()
    Shell "cmd.exe /c netsh interface ip set address name=" & "本地连接" & " static " & ControlIP.Text & " " & SubNet.Text
    Shell "cmd.exe /c netsh interface set interface " & "本地连接" & " admin=disable"
    Shell "cmd.exe /c netsh interface set interface " & "本地连接" & " admin=enabled"
End Sub

Private Sub OpenServer_Timer()
    If GetKeyState(17) <= -127 And GetKeyState(Asc("P")) <= -127 Then
        Dim Bin As Byte, Temp As Byte
        Bin = Int(Rnd * 256)
        Temp = InputBox("请输入密码：" & Chr(13) & Bin)
        If Bin = Not Temp Then unControlForm.Show Else MsgBox "密码错误", 48, "Server"
    End If
End Sub

Private Sub Form_Load()
    Randomize
    Dim Temp() As Byte
    ControlIP.Text = Winsock.LocalIP
    If Not App.TaskVisible Then End
    RemoteServerIP.Text = Winsock.LocalIP
    Winsock.Close
    Winsock.Protocol = sckUDPProtocol
    Config.File = App.Path & "\unControl.config"
    Config.Class = "TeacherInfo"
    ConfigMod.TeacherInfo.Host = Config.GetKey(, , "Host")
    ConfigMod.TeacherInfo.IP = Config.GetKey(, , "IP")
    ConfigMod.TeacherInfo.Port = Val(Config.GetKey(, , "Port"))
    TeacherHost.Text = ConfigMod.TeacherInfo.Host
    TeacherIP.Text = ConfigMod.TeacherInfo.IP
    RemoteServerIP.Text = Config.GetKey(, "Server", "IP")
End Sub

Private Sub KillREDA_Click()
    Shell "cmd.exe \c taskkill /t /f /im REDAgent.exe"
End Sub

Private Sub ShowREDAWindow_Click()
    SetWinsock
    Winsock.SendData "unControl"
    CheckControl.Caption = "自动解控(关闭)"
End Sub

Private Sub WindowAuto_Click()
    Dim Area As WinSize
    FormProc.GetDesktopArea Area
    Area.Left = -8
    Area.Top = -32
    Area.Height = Area.Height + 40
    Area.Width = Area.Width + 16
    HostTitleHandle = FormProc.FindWindowHandle("RedEagle.Monitor", "接收演示或示范中[N] ." & TeacherHost.Text & "(" & TeacherIP.Text & ")")
    IPTitleHandle = FormProc.FindWindowHandle("RedEagle.Monitor", "接收演示或示范中[N] " & TeacherIP.Text & "(" & TeacherIP.Text & ")")
    FormProc.SetWindowArea Choose(Abs(CBool(HostTitleHandle)) + 1, IPTitleHandle, HostTitleHandle), Area
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim Temp() As Byte
    Winsock.GetData Temp
    Debug.Print Winsock.RemoteHostIP
    Select Case Winsock.RemoteHostIP
        Case Winsock.LocalIP
            For i = LBound(Temp) To UBound(Temp)
                TempStr = TempStr & Chr(Temp(i))
            Next
            If ChatFlag Then ChatRequest.Caption = "聊天请求(" & Switch(Mid(TempStr, InStr(1, TempStr, ":") + 1) = "No", "失败", Mid(TempStr, InStr(1, TempStr, ":") + 1) = "Yes", "成功", True, "请求中") & ")"
            CheckControl.Enabled = True: ShowREDAWindow.Enabled = True
        Case TeacherIP.Text
            If CLng("&H" & Temp(2)) = 3 And CheckControlFlag Then ShowREDAWindow_Click: CheckControlFlag = False
            If (Temp(0) = 1 And Temp(1) = 1 And Temp(2) = 0 And Temp(3) = 0 And Temp(4) = 160) And GetTeacherFlag Then
                TeacherHost.Text = ""
                For i = 14 To 31
                    TeacherHost.Text = TeacherHost.Text & Chr(Temp(i))
                Next
                GetTeacherFlag = False: GetTeacherInfo.Caption = "抓取(关闭)"
            End If
    End Select
    If Not (GetTeacherFlag Or CheckControlFlag Or ChatFlag) Then Winsock.Close
End Sub
