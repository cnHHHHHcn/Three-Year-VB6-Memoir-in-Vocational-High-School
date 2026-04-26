VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8805
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "解除黑屏"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "关闭广播"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "黑屏"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "自动解控"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2640
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "全屏"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   900
      Width           =   855
   End
   Begin VB.TextBox RemoteIP 
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Text            =   "225.2.2.1"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox RemotePort 
      Height          =   270
      Left            =   1080
      TabIndex        =   1
      Text            =   "5512"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "手动解控"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1320
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Private Type ConfigInfo
    Host As String
    IP As String
    Port As Long
End Type

Private ControlScreen() As Byte
Private ControlRelease() As Byte
Private VideoPlay() As Byte
Private VideoStop() As Byte

Private Control2() As Byte

Private ChatPackage() As Byte
Private VirtualOnline() As Byte
Private TeacherInfo As ConfigInfo

Private Sub Command2_Click()
    Winsock2.Close
    Winsock2.Protocol = sckUDPProtocol
    Winsock2.Bind 5542, "26.170.212.13"
End Sub

Private Sub Command3_Click()
     SetTeacherInfo
     Winsock1.SendData ControlScreen
End Sub

Private Sub Command4_Click()
     SetTeacherInfo
     Winsock1.SendData VideoStop
End Sub

Private Sub Command5_Click()
    SetTeacherInfo
    Winsock1.SendData ControlRelease
    Winsock1.RemotePort = 4988
    Winsock1.SendData ControlRelease
End Sub

Private Sub Form_Load()
    Winsock1.Close
    Winsock1.Protocol = sckUDPProtocol
    Winsock1.Bind 61300, "26.170.212.13"
    'Des IP:13 - 16; SubNet Mask:25 -28
    ReadPackage "ControlScreen.txt", ControlScreen
    'Des IP:13 - 16
    ReadPackage "ControlRelease.txt", ControlRelease
    
    'Des IP:21 - 24
    'ReadPackage "Control2.txt", Control2
    
    'Des IP:13 - 16; Scr IP:105 - 108 _
     Play Style: 29 ,0 屏幕 , 1 全屏
    ReadPackage "VideoPlay.txt", VideoPlay
    ReadPackage "VideoStop.txt", VideoStop
    'For i = LBound(Control2) To UBound(Control2)
    '    Debug.Print Control2(i), i
    'Next
End Sub

Private Sub Command1_Click()
    SetTeacherInfo
    Winsock1.SendData VideoStop
    VideoPlay(29) = Check1.Value
    Winsock1.SendData VideoPlay
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Dim Temp() As Byte
    Winsock.GetData Temp
    '屏幕播放特征码：227，64，0，0
    If Winsock2.RemoteHostIP = "26.170.212.13" Then
        If Temp(0) = 227 And Temp(1) = 64 And Temp(2) = 0 And Temp(3) = 0 Then Command1_Click
    End If
End Sub

Public Sub ReadPackage(File As String, Bytes() As Byte)
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
End Sub

Public Sub SetTeacherInfo()
    Winsock1.RemoteHost = RemoteIP.Text
    Winsock1.RemotePort = Val(RemotePort.Text)
End Sub
