VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10920
   StartUpPosition =   3  '窗口缺省
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3255
      Left            =   240
      TabIndex        =   21
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.TextBox StateText 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   8640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Text            =   "Form1.frx":009D
      Top             =   1680
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   6000
      ScaleHeight     =   435
      ScaleWidth      =   2940
      TabIndex        =   14
      Top             =   6360
      Width           =   3000
      Begin VB.Label downloadnum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1200
         TabIndex        =   15
         Top             =   120
         Width           =   180
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFF80&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "执行"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   2880
      TabIndex        =   11
      Top             =   840
      Width           =   5655
   End
   Begin VB.ComboBox Combo1 
      Height          =   435
      ItemData        =   "Form1.frx":00E3
      Left            =   1200
      List            =   "Form1.frx":0117
      TabIndex        =   9
      Text            =   "cd"
      Top             =   840
      Width           =   1575
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "添加"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "下载"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.ListBox downloadlist 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      ItemData        =   "Form1.frx":017B
      Left            =   240
      List            =   "Form1.frx":017D
      TabIndex        =   6
      Top             =   5400
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "连接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox SavePath 
      Height          =   435
      Left            =   1680
      TabIndex        =   1
      Top             =   7560
      Width           =   7215
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "下载任务："
      Height          =   315
      Left            =   360
      TabIndex        =   19
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "状态："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8640
      TabIndex        =   18
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label downloadflna 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6600
      TabIndex        =   17
      Top             =   6000
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "文件:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6000
      TabIndex        =   16
      Top             =   6000
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "命令："
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   900
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "保存路径："
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "URL："
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fileim() As String, statestr As String

Private Sub Combo1_Click()
If Combo1.Text = "help" Then
    Set a = CreateObject("wscript.shell")
    a.run "ftphelp.txt"
    Combo1.Text = ""
End If
End Sub

Private Sub Command1_Click()
If Not SavePath.Text = "" Then
    If Right(SavePath.Text, 1) <> "\" Then SavePath.Text = SavePath.Text & "\"
    Shell ("cmd.exe /c md " & SavePath.Text)
    Open "information.txt" For Output As #1
        Print #1, SavePath.Text
    Close
Else
    MsgBox "未输入保存路径！", 48
End If
End Sub

Private Sub Command2_Click()
Text4.Text = Inet1.OpenURL(Text1.Text)
End Sub

Private Sub Command3_Click()
Dim download As Boolean
For i = 1 To downloadlist.ListCount
    downloadflna.Caption = filename(fileim(i))
    download = False
    fileallsize = FileLen(fileim(i))
    Do Until download
        Inet1.AccessType = icUseDefault
        Dim data() As Byte
        '把该文件作为字节数组来检索。
        data() = Inet1.OpenURL(fileim(i), icByteArray)
        Open SavePath.Text & downloadflna.Caption For Binary Access Write As #1
            Put #1, , data()
            filesize = LOF(1)
        Close #1
        Number = Val(Left(downloadnum.Caption, Len(downloadnum) - 1))
        downloadnum.Caption = Fix((filesize / fileallsize) * 100) & "％"
        Shape1.Width = Number * 30
        If Number = 100 Then download = True: downloadlist.RemoveItem (0)
    Loop
Next

End Sub

Private Sub Command5_Click()
Inet1.Execute Text1.Text & " " & Combo1.Text & " " & Text3.Text
End Sub

Private Sub Form_Load()
Open "information.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, Strs
    Loop
Close
SavePath.Text = Strs
Shell ("cmd.exe /c md " & SavePath.Text)
End Sub

'常数                   值  描述
'icNone                 0   无状态可报告。
'icHostResolvingHost    1   该控件正在查询所指定的主机的 IP 地址。
'icHostResolved         2   该控件已成功地找到所指定的主机的 IP 地址。
'icConnecting           3   该控件正在与主机连接。
'icConnected            4   该控件已与主机连接成功。
'icRequesting           5   该控件正在向主机发送请求。
'icRequestSent          6   该控件发送请求已成功。
'icReceivingResponse    7   该控件正在接收主机的响应。
'icResponseReceived     8   该控件已成功地接收到主机的响应。
'icDisconnecting        9   该控件正在解除与主机的连接。
'icDisconnected         10  该控件已成功地与主机解除了连接。
'icError                11  与主机通讯时出现了错误。
'icResponseCompleted    12  该请求已经完成，并且所有数据均已接收到。




Private Sub Inet1_StateChanged(ByVal State As Integer)

   Dim vtData As Variant ' Data variable.
   Select Case State
   Case icNone '0
     StateText.Text = "无状态可报告"
   Case icHostResolvingHost '1
     StateText.Text = "正在查询所指定的主机的 IP 地址"
   Case icHostResolved '2
     StateText.Text = "已成功地找到所指定的主机的 IP 地址"
   Case icConnecting '3
     StateText.Text = "正在与主机连接"
   Case icConnected '4
     StateText.Text = "已与主机连接成功"
   Case icRequesting '5
     StateText.Text = "正在向主机发送请求"
   Case icRequestSent '6
     StateText.Text = "发送请求已成功"
   Case icReceivingResponse '7
     StateText.Text = "正在接收主机的响应"
   Case icResponseReceived '8
     StateText.Text = "已成功地接收到主机的响应"
   Case icDisconnecting '9
     StateText.Text = "正在解除与主机的连接"
   Case icDisconnected '10
     StateText.Text = "已成功地与主机解除了连接"
   Case icError '11
      '出现错误时，返回 ResponseCode 和 ResponseInfo。
      vtData = Inet1.ResponseCode & ":" & vbCrLf & Inet1.ResponseInfo
      StateText.Text = vtData
   Case icResponseCompleted ' 12
      'State = 12 时，用 GetChunk 方法检索服务器的响应。
      Dim strData As String
      Dim bDone As Boolean: bDone = False

      '取得第一个块。
      vtData = Inet1.GetChunk(1024, icString)
      DoEvents

      Do While Not bDone
         strData = strData & vtData
         '取得下一个块。
         vtData = Inet1.GetChunk(1024, icString)
         DoEvents

         If Len(vtData) = 0 Then bDone = True
      Loop
      StateText.Text = strData
   End Select
   
End Sub

Private Sub Command4_Click()
ReDim fileim(List1.ListCount)
For i = 1 To List1.ListCount
    fileim(i) = List1.List(i - 1)
Next
End Sub

Function filename(file As String) As String
Do While file <> filestr
    filestr = file
    file = Right(file, Len(file) - InStr(file, "\"))
Loop
filename = file
End Function

Private Sub StateText_Change()
StateText.Text = statestr
End Sub

Private Sub StateText_Click()
statestr = StateText.Text
End Sub
