VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4485
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3960
      Top             =   960
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   318
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   600
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "admin015"
      RemotePort      =   10002
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "admin015"
      RemotePort      =   10001
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   327681
      FullWidth       =   121
      FullHeight      =   9
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lenFile As Long, FreeName As String, Secnd As Long, Secnd2 As Long

Private Sub Command1_Click()
    Winsock1.Connect "192.168.1.16"
    Winsock2.Connect "192.168.1.16"
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Label1.Caption = "已接收数据" & FileLen("d:\" & FreeName) & "\" & lenFile & "(" & Int((FileLen("d:\" & FreeName) / lenFile) * 100) & "%)" & _
    Chr(13) & Int(lenFile / 1024 / 1024) & "Mb"
    ProgressBar1.Value = FileLen("d:\" & FreeName)
    Label3.Caption = Winsock1.State & Winsock2.State
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    Secnd = FileLen("d:\" & FreeName)
    Label2.Caption = Int(((Secnd - Secnd2) / 1024) / 1024) & " mb/s"
    Secnd2 = FileLen("d:\" & FreeName)
    If Left(Label2.Caption, 1) = 0 Then Timer2.Enabled = False
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Timer2.Enabled = True
    Dim bytData() As Byte
    Dim lLenFile As Long
    Dim f
    f = FreeFile
    Open "d:\" & FreeName For Binary As #f ''strFileName是文件名
    lLenFile = LOF(f)
    ReDim bytData(1 To bytesTotal)
    Winsock1.GetData bytData
    If lLenFile = 0 Then ''lLenFile=0表示是第一次打开文件，这里有个问题，就是''如果如果该文件存在的话，就会出错，应该在打开前检查文件是否存在。（这里我省略了）
        Put #f, 1, bytData
    Else
        Put #f, lLenFile + 1, bytData

    End If
Close #f
    
    
    
End Sub
Private Sub Form_Load()
    Winsock1.Connect "192.168.1.16"
    Winsock2.Connect "192.168.1.16"
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Dim data_ As String
    Winsock2.GetData data_
    Animation1.Open App.Path & "\1.avi"
    Animation1.Play
    If Left(data_, 1) = "-" Then
        Do Until InStr(1, data_, "\") = 0
            data_ = Right(data_, Len(data_) - InStr(1, data_, "\"))
            FreeName = data_
        Loop
    Else
        CLng_ = CLng(data_)
        ProgressBar1.Max = CLng_
        lenFile = CLng_
    End If
End Sub
