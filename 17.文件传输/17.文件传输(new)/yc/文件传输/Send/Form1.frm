VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   4515
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1320
      Top             =   1560
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   10002
   End
   Begin MSWinsockLib.Winsock wins 
      Left            =   600
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   10001
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   327681
      FullWidth       =   257
      FullHeight      =   33
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   180
      Left            =   840
      TabIndex        =   2
      Top             =   2160
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Const iMax = 65535 '是每个数据块的大小。
    Dim FreeF As Integer ''空闲的文件号
    Dim LenFile As Long ''文件的长度
    Dim bytData() As Byte ''存放数据的数组
    Dim Pos As Long

Private Sub Command1_Click()

    Animation1.Open App.Path & "\1.avi"
    Animation1.Play
    On Error GoTo Error
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then Exit Sub
    
    Winsock1.SendData "-" & CommonDialog1.FileName
    
    FreeF = FreeFile ''获得空闲的文件号
    Open CommonDialog1.FileName For Binary As #FreeF ''打开文件
    DoEvents
    LenFile = LOF(FreeF) ''获得文件长度
    
    Winsock1.SendData Str(LenFile)
    
    ProgressBar1.Max = LenFile
    

    If LenFile <= iMax Then ''如果要发送的文件小于数据块大小，直接发送
        ReDim bytData(1 To LenFile) ''根据文件长度重新定义数组大小
        Get #FreeF, , bytData ''把文件读入到数组里
        Close #FreeF ''关闭文件
        
        wins.SendData bytData ''发送数据
        Exit Sub
    End If

    Do Until (iPos >= (LenFile - iMax)) ''发送整块数据的循环
        ReDim bytData(1 To iMax)
        Get #FreeF, iPos + 1, bytData
        DoEvents
        wins.SendData bytData
        iPos = iPos + iMax ''移动iPos，使它指向下来要读的数据
        ProgressBar1.Value = iPos
        Label1.Caption = "已发送数据:" & CLng(ProgressBar1.Value) & "\" & LenFile 'Int((iPos / LenFile) * 100)
    Loop
''这里要注意的是，必须检查文件有没有剩下的数据，如果文件大小正好等于数据块大小的
'' 整数倍，那么就没有剩下的数据了'
    ReDim bytData(1 To LenFile - iPos) ''发送剩下的不够一个数据块的数据
    Get #FreeF, iPos + 1, bytData
    wins.SendData bytData
    Close #FreeF
    
    Exit Sub
Error:
    If Err.Number = 40006 Then MsgBox "链接错误！", , "Error"
    

End Sub

Private Sub Command2_Click()
    Animation1.Open App.Path & "\1.avi"
    Animation1.Play
End Sub

Private Sub Command3_Click()
    CommonDialog1.ShowOpen
    Transfer.FileTransfer CommonDialog1.FileName, wins, ProgressBar1
End Sub

Private Sub Form_Load()
    wins.Listen
    Winsock1.Listen
End Sub

Private Sub Timer1_Timer()
Label2.Caption = wins.State & Winsock1.State
End Sub

Private Sub wins_ConnectionRequest _
(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If wins.State <> sckClosed Then _
wins.Close
'接受具有 requestID 参数的
'连接。
wins.Accept requestID
End Sub

Private Sub winsock1_ConnectionRequest _
(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If Winsock1.State <> sckClosed Then _
Winsock1.Close
'接受具有 requestID 参数的
'连接。
Winsock1.Accept requestID
End Sub


