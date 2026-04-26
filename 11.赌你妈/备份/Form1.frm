VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10605
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5160
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Create_Server 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer PlayerWaiting 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   840
      Top             =   1080
   End
   Begin VB.Timer PlayerCheck_Send 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   6840
   End
   Begin VB.Timer PlayerCheck 
      Interval        =   1000
      Left            =   1200
      Top             =   480
   End
   Begin VB.Timer PlayerRefresh 
      Interval        =   1000
      Left            =   480
      Top             =   480
   End
   Begin VB.ListBox PlayerList 
      Appearance      =   0  'Flat
      Height          =   5970
      ItemData        =   "Form1.frx":0000
      Left            =   360
      List            =   "Form1.frx":0002
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "选定房间"
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton Command5 
         Caption         =   "申请"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "请输入名字或编码（编码起始位为 1 ）"
         Top             =   360
         Width           =   1215
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2355
      _cy             =   1931
   End
   Begin VB.Label Player2 
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label User 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8280
      TabIndex        =   9
      Top             =   360
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7080
      TabIndex        =   8
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   3
      Left            =   8280
      TabIndex        =   3
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   2
      Left            =   8280
      TabIndex        =   2
      Top             =   5760
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "连机模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   1
      Left            =   7440
      TabIndex        =   1
      Top             =   2880
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单机模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Index           =   0
      Left            =   7440
      TabIndex        =   0
      Top             =   1560
      Width           =   2100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indexs
Dim Play1 As Boolean, Play2 As Boolean

Private Sub Command5_Click()
Player2.Caption = Text3.Text
'发送1
PlayerCheck.Enabled = False: PlayerCheck_Send.Enabled = True
Call games_connect.Request(App.Path & "\联机\" & Player2.Caption & ".txt")
Text3.Text = ""
End Sub

Private Sub Create_Server_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1(indexs).FontSize = 26
Label1(indexs).ForeColor = vbBlack
End Sub

Private Sub Form_Load()
On Error Resume Next
gh:
namestr = InputBox("请输入名称")
User.Caption = namestr

'用户名第一位不能为数字
a = Asc(namestr)
If Asc("0") <= a And a <= Asc("9") Then GoTo gh
Select Case namestr
Case "", "王丹", "崔敬爱", "赵兰英", "董铭", "康建晓", "李海静", "郭晓坤", "张江华", "刘红莲", "张爱学", "贺林涛", "秦文龙", "习近平", "毛泽东", "周恩来", "邓小平"
    GoTo gh
End Select
'查找是否有相同名称的用户
hao = FreeFile
Open App.Path & "\联机\玩家列表.txt" For Input As hao
    Do Until EOF(hao)
        Line Input #hao, ry
        If ry = namestr Then GoTo gh
    Loop
Close

Open App.Path & "\联机\玩家列表.txt" For Append As #1
    Print #1, namestr
Close

Open App.Path & "\联机\" & namestr & ".txt" For Output As #1
Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rymc() As String
hao = FreeFile
Open App.Path & "\联机\玩家列表.txt" For Input As hao
i = 0
    Do Until EOF(hao)
        Line Input #hao, ry
        i = i + 1
    Loop
Close
ReDim rymc(i)
Open App.Path & "\联机\玩家列表.txt" For Input As hao
i = 0
    Do Until EOF(hao)
        Line Input #hao, rymc(i)
        i = i + 1
    Loop
Close
For k = 0 To i
    If rymc(k) = User.Caption Then rymc(k) = ""
Next
Open App.Path & "\联机\玩家列表.txt" For Output As #5
    For k = 0 To i
        If rymc(k) <> "" Then Print #5, rymc(k)
    Next
Close
Shell ("cmd.exe /c del /f /q " & Chr(34) & App.Path & "\联机\" & User.Caption & ".txt" & Chr(34))
End
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 0
    Form1.Hide: Form2.Show
    Frame1.Visible = False: PlayerList.Visible = False
Case 1
    Frame1.Visible = True: PlayerList.Visible = True
Case 2
    Form_Unload (0)
    End
End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
indexs = Index
Label1(indexs).FontSize = 30
Label1(indexs).ForeColor = vbRed
End Sub

Private Sub PlayerCheck_Send_Timer()
Play1 = games_connect.Response(App.Path & "\联机\" & User.Caption & ".txt", App.Path & "\联机\" & Player2.Caption & ".txt")
End Sub

Private Sub PlayerCheck_Timer()   '接收1
Static emm
If Not emm Then
    emm = games_connect.Indication(App.Path & "\联机\" & User.Caption & ".txt", App.Path & "\联机\" & Player2.Caption & ".txt")
Else
    Play2 = games_connect.Confirmation(App.Path & "\联机\" & User.Caption & ".txt", App.Path & "\联机\" & Player2.Caption & ".txt")
End If
End Sub

Private Sub PlayerRefresh_Timer()
PlayerList.Clear
PlayerList.AddItem "-----NONE-----"
Open App.Path & "\联机\玩家列表.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, hhh
        PlayerList.AddItem hhh
    Loop
Close
End Sub

Private Sub Text3_Change()
On Error Resume Next
'输入标号查找玩家
If Not Val(Text3.Text) <= 0 Then Text3.Text = PlayerList.List(Val(Text3.Text))
If Text3.Text = "" Or Text3.Text = " " Then Command5.Enabled = False Else Command5.Enabled = True
Command5.SetFocus
End Sub
