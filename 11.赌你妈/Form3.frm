VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form3"
   ScaleHeight     =   8880
   ScaleWidth      =   11640
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer RoomStateChecking 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   9360
      Top             =   240
   End
   Begin VB.Timer StateCheck 
      Interval        =   1000
      Left            =   7800
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7920
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Timer GameStateCheck 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5400
      Top             =   4200
   End
   Begin VB.Timer reload_ammo_nosice 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10440
      Top             =   840
   End
   Begin MSWinsockLib.Winsock Send_GameData 
      Left            =   9720
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer AmmoColor 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5400
      Top             =   1200
   End
   Begin VB.Timer Start 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5400
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "射击"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   4200
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "选择子弹数"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Information 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   7320
      Width           =   11655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   6960
      Width           =   9855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "发送"
      Height          =   375
      Left            =   10320
      TabIndex        =   1
      Top             =   6960
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "射击"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Timer AI 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7680
      Top             =   960
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "玩家2："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6120
      TabIndex        =   21
      Top             =   1440
      Width           =   1110
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "玩家1："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   1110
   End
   Begin VB.Label Player 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   19
      Top             =   1440
      Width           =   165
   End
   Begin VB.Label Player 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   7200
      TabIndex        =   18
      Top             =   1440
      Width           =   165
   End
   Begin VB.Label StateShow 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   7200
      TabIndex        =   17
      Top             =   240
      Width           =   90
   End
   Begin WMPLibCtl.WindowsMediaPlayer nosice 
      Height          =   615
      Left            =   9840
      TabIndex        =   15
      Top             =   600
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
      _cy             =   1085
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   4400
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10320
      TabIndex        =   14
      Top             =   240
      Width           =   720
   End
   Begin VB.Shape Shape3 
      Height          =   35
      Left            =   5280
      Top             =   1800
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5400
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   5400
      Top             =   960
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   7040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "生命值:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   960
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "生命值:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   11
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   10
      Top             =   960
      Width           =   165
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11640
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "技能："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   9
      Top             =   4920
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "技能："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   4920
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   3
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   4
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   5
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   7
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   0
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   1
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   2
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   3
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   4
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   5
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   6
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   7
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11640
      Y1              =   6920
      Y2              =   6920
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "说："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   7000
      Width           =   480
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   165
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim ammo() As Boolean '弹夹
Dim t As Integer '实弹数
Dim tb As Integer  'AI计算实弹数
Dim f As Integer '虚弹数'
Dim fb As Integer 'AI计算虚弹数
Dim u As Integer '当前第几发子弹'
Dim tool() As Integer '我方技能'
Dim dtool() As Integer '敌方技能'
Dim H_A As Boolean '伤害增加判定'
Dim ammostatc  'AI查看子弹射击判定
Dim AmmoData As String   '发送子弹数据


'                           过程 和 函数
'-------------------------------------------------------------------------
Sub fire(Player As Control, num As Integer)
On Error Resume Next
'子弹射击子程序

num = num + 1

Debug.Print "此发子弹为", num
Debug.Print ammo(num)

If ammo(num) Then
    If H_A Then Player = Player - 2 Else Player = Player - 1
    tb = tb - 1
Else
    fb = fb - 1: nosice.URL = App.Path & "\素材\gunfirenull.mp3"
End If
H_A = False

End Sub

Sub End_Game()
For i = 0 To 1
    Command2(i).Enabled = False
Next
For i = 0 To 7
    Image1(i).Enabled = False
Next
ReDim ammo(10 To 10): u = 0
Command3.Enabled = True
GameStateCheck.Enabled = False
End Sub

'拿去技能子程序
Sub GiveTool(PictureObject1 As Control, PictureObject2 As Control, ToolCode As Integer)
PictureObject2.Picture = LoadPicture
Select Case ToolCode
Case 1
    PictureObject1.Picture = LoadPicture(App.Path & "\素材\攻击.bmp")
Case 2
    PictureObject1.Picture = LoadPicture(App.Path & "\素材\生命治疗.bmp")
Case 3
    PictureObject1.Picture = LoadPicture(App.Path & "\素材\检查子弹.bmp")
Case 4
    PictureObject1.Picture = LoadPicture(App.Path & "\素材\退弹.bmp")
End Select
End Sub

Sub GamePlay_House_Owner_T()
Send_GameData.Close
Send_GameData.LocalPort = 1021
Send_GameData.Listen
End Sub

Sub GamePlay_House_Owner_F(House_Owner_HostName As String)
Send_GameData.Close
Send_GameData.Connect House_Owner_HostName, 1021
End Sub

Sub Start_Game()
On Error Resume Next
ReDim tool(8), dtool(8)
Label2.Caption = "4"
Label4.Caption = "4"
Command3.Enabled = True
Command2(0).Enabled = False
Command2(1).Enabled = False
For i = 0 To 7
    Image1(i).Picture = LoadPicture
    Image2(i).Picture = LoadPicture
    tool(i + 1) = 0: dtool(i + 1) = 0
Next
GameStateCheck.Enabled = True
ReDim ammo(10 To 10)
End Sub

'获取技能子过程
Sub GetTool(PictureObject As Control, ToolCode As Integer)
ToolCode = Int(Rnd * 4) + 1
Select Case ToolCode
Case 1
    PictureObject.Picture = LoadPicture(App.Path & "\素材\攻击.bmp")
Case 2
    PictureObject.Picture = LoadPicture(App.Path & "\素材\生命治疗.bmp")
Case 3
    PictureObject.Picture = LoadPicture(App.Path & "\素材\检查子弹.bmp")
Case 4
    PictureObject.Picture = LoadPicture(App.Path & "\素材\退弹.bmp")
End Select
End Sub

Function GetGameLabel(Data As String, StartString As String, EndString As String) As String
StartNum = InStr(7, Data, StartString)
EndNum = InStr(7, Data, EndString)
For i = StartNum To EndNum
    GetGameLabel = GetGameLabel & Mid(Data, i, 1)
Next
End Function

'-------------------------------------------------------------------------


'*************************************************************************
'   标签         作用

'<Data>[Ammo]    数据 子弹

'<Data>[AmmoNum] 数据 子弹标

'<Data>[Blood]   数据 血量

'<Data>[Text]    数据 游戏信息

'<Data>[Tool]    数据 工具

'<Text>          玩家互相发送的信息

'*************************************************************************


'-------------------------------------------------------------------------------------------------------------------------
Private Sub AI_Timer()    'AI人机
On Error GoTo errorammo
Randomize
Dim o As Integer, l As Integer

PlayerCheck.Caption = "1"
'记录我第一次的生命值
FirstNum = Label2.Caption

'使用技能
Static countnum
countnum = countnum + 1
If countnum <= 4 Then
Starts:
    If c >= 5 Then c = 0: GoTo endss
    l = Int(Rnd * 8)
    If dtool(l + 1) = 0 Then c = c + 1: GoTo Starts
    Call Image2_MouseDown(l, 1, 0, 10, 0)
endss:
    countnum = 4
    
Else
    tbs = tb: fbs = fb
    If ammo(u + 1) Then butler = "实": tb = tb - 1 Else butler = "虚": fb = fb - 1
    If Not ammo(u + 1) Then
        For i = 1 To 8
            If dtool(i) = 0 Then Call GetTool(Image2(i - 1), dtool(i)): Exit For
        Next
    End If
errorammo:

    '子弹查看技能
    If ammostatc <> "" Then
        If ammostatc Then
            Image3.Picture = LoadPicture(App.Path & "\素材\大喷左.bmp"): Call Command2_Click(0)
            If ammostatc Then nosice.URL = App.Path & "\素材\gunfirekill.mp3"
            Text1.Text = Text1.Text & vbCrLf & "AI向你发射子弹，子弹为" & butler & "弹"
        Else
            Image3.Picture = LoadPicture(App.Path & "\素材\大喷右.bmp"): Call Command2_Click(1)
            Text1.Text = Text1.Text & vbCrLf & "AI向自己发射子弹，子弹为" & butler & "弹"
        End If
        ammostatc = "": GoTo ends
    End If
    
    '子弹发射
    If tbs < fbs Then
        Image3.Picture = LoadPicture(App.Path & "\素材\大喷右.bmp"): Call Command2_Click(1)
        If ammo(u) Then nosice.URL = App.Path & "\素材\gunfire.mp3"
        Text1.Text = Text1.Text & vbCrLf & "AI向自己发射子弹，子弹为" & butler & "弹"
    ElseIf fbs < tbs Then
        Image3.Picture = LoadPicture(App.Path & "\素材\大喷左.bmp"): Call Command2_Click(0)
        If ammo(u) Then nosice.URL = App.Path & "\素材\gunfirekill.mp3"
        Text1.Text = Text1.Text & vbCrLf & "AI向你发射子弹，子弹为" & butler & "弹"
    ElseIf tbs = fbs And (tbs <> 0 And fbs <> 0) Then
        o = Int(Rnd * 2)
        Select Case o
        Case 0
            Image3.Picture = LoadPicture(App.Path & "\素材\大喷左.bmp")
            If ammo(u + 1) Then nosice.URL = App.Path & "\素材\gunfirekill.mp3"
        Case 1
            Image3.Picture = LoadPicture(App.Path & "\素材\大喷右.bmp")
            If ammo(u + 1) Then nosice.URL = App.Path & "\素材\gunfire.mp3"
        End Select
        Call Command2_Click(o)
        If o = 0 Then Text1.Text = Text1.Text & vbCrLf & "AI向你发射子弹，子弹为" & butler & "弹" Else Text1.Text = Text1.Text & vbCrLf & "AI向自己发射子弹，子弹为" & butler & "弹"
    End If
    
    '记录我第二次的生命值
    SecondNum = Label2.Caption
    '拿取我的技能
    Dim Number As Integer
    If PlayerCheck = "1" And KillObject.Caption = "me" Then
        If FirstNum > SecondNum Then
            For i = 1 To 8
                Number = Int(Rnd * 8)
                If tool(Number + 1) <> 0 Then Call Image1_MouseDown(Number, 2, 0, 0, 0): Exit For
            Next
        End If
    End If
ends:
    '控制权交给对方
    If u = UBound(ammo) Then GoTo ghj
    If tbs > 0 Or fbs > 0 Then
            For i = 0 To 1
                Command2(i).Enabled = True
            Next
            For i = 0 To 7
                Image1(i).Enabled = True
            Next
    Else
ghj:
        Call End_Game
    End If
    AI.Enabled = False
End If
End Sub
'-------------------------------------------------------------------------------------------------------------------------

Private Sub AmmoColor_Timer()
Shape1.FillColor = &HFFFFFF
AmmoColor.Enabled = False
End Sub

Private Sub reload_ammo()   '装填子弹
Starts_Sub:
t = 0: f = 0
Randomize

For i = 1 To UBound(ammo)
    ammo(i) = CBool(Int(Rnd * 1.8))
    If ammo(i) Then
        t = t + 1
    Else
        f = f + 1
    End If
Next
If t = UBound(ammo) Or f = UBound(ammo) Then GoTo Starts_Sub

tb = t: fb = f
Debug.Print
Debug.Print "------------------"
For i = 1 To UBound(ammo)
Debug.Print ammo(i)
Next
Debug.Print "------------------"
For i = 1 To UBound(ammo)
    AmmoData = AmmoData & IIf(ammo(i), "1", "0")
Next
Send_GameData.SendData "<Data>[Ammo]" & AmmoData
ReDim Preserve tool(8), dtool(8)
Label8.Caption = "实弹" & t & "发，虚弹" & f & "发"
For i = 0 To 7
    Image1(i).Enabled = True
Next
reload_ammo_nosice.Enabled = True
Timer1.Enabled = True
End Sub

Private Sub Command1_Click()
MsgBox AmmoData
End Sub

Private Sub Command2_Click(Index As Integer)     '射击
FirstNum = Label4.Caption
Select Case Index
Case 0
    KillObject.Caption = "me"
    fire Label2, u
Case 1
    KillObject.Caption = "AI"
    fire Label4, u
End Select
'开启拿取技能
SecondNum = Label4.Caption
If PlayerCheck = "0" And KillObject.Caption = "AI" Then
    If FirstNum > SecondNum Then
        For i = 1 To 8
            Image2(i - 1).Enabled = True
            Image1(i - 1).Enabled = False
        Next
        For i = 0 To 1
            Command2(i).Enabled = False
        Next
    End If
Else
    Command4.SetFocus
End If
End Sub

Private Sub Command2_GotFocus(Index As Integer)
On Error Resume Next
PlayerCheck.Caption = "0"
If ammo(u + 1) Then butler = "实" Else butler = "虚"
Select Case Index
Case 0
    Text1.Text = Text1.Text & vbCrLf & "你向自己发射子弹，子弹为" & butler & "弹"
    If ammo(u + 1) Then nosice.URL = App.Path & "\素材\gunfirekill.mp3"
    '开启系统送技能
    If Not ammo(u + 1) Then
        For i = 1 To 8
            If tool(i) = 0 Then Call GetTool(Image1(i - 1), tool(i)): Exit For
        Next
    End If
Case 1
    Text1.Text = Text1.Text & vbCrLf & "你向AI发射子弹，子弹为" & butler & "弹"
    If ammo(u + 1) Then nosice.URL = App.Path & "\素材\gunfire.mp3"
End Select
End Sub

Private Sub Command2_LostFocus(Index As Integer)     '射击
For i = 0 To 1
    Command2(i).Enabled = False
Next
If Not u = UBound(ammo) Then AI.Enabled = True
End Sub

Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
    Image3.Picture = LoadPicture(App.Path & "\素材\大喷左.bmp")
Case 1
    Image3.Picture = LoadPicture(App.Path & "\素材\大喷右.bmp")
End Select
End Sub

Private Sub Command3_Click()   '选择子弹数
Start:
num = InputBox("请选择子弹数（4 - 8）", "子弹数")
If 1 < num And num < 9 Then ReDim ammo(num): Command3.Enabled = False: reload_ammo: Label7.Enabled = False Else GoTo Start
End Sub

Private Sub Command5_Click()
If Send_GameData.State = 7 Then Send_GameData.SendData "<Text>" & Text2.Text: Information.Text = Information.Text & "<Text>" & Text2.Text
End Sub

Private Sub Form_Activate()
Start_Game
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbBlack
End Sub

Private Sub GameStateCheck_Timer()
If u = UBound(ammo) Then Call End_Game
If Val(Label2.Caption) <= 0 Then
    MsgBox "you dead!"
    Call Start_Game
    Label7.Enabled = True
ElseIf Val(Label4.Caption) <= 0 Then
    MsgBox "you win!"
    Call Start_Game
    Label7.Enabled = True
End If
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)   '技能
If Button = 2 Then
    For i = 1 To 8
        If dtool(i) = 0 Then dtool(i) = tool(Index + 1): Exit For
    Next
    tool(Index + 1) = 0
    Call GiveTool(Image2(i - 1), Image1(Index), dtool(i))
    For i = 1 To 8
        Image1(i - 1).Enabled = False
    Next
    For i = 0 To 1
        Command2(i).Enabled = False
    Next
Else
    If tool(Index + 1) <> 0 Then
        Image1(Index).Picture = LoadPicture("")
        Select Case tool(Index + 1)
        Case 1
            H_A = du_gameing.Hurt_Add(Label4)
            jineng = "伤害增加"
        Case 2
            Call du_gameing.Blood_Add(Label2)
            jineng = "血量增加"
        Case 3
            Call du_gameing.Look_Ammo(ammo(u + 1), Shape1)
            jineng = "子弹查看"
            AmmoColor.Enabled = True
        Case 4
            Call du_gameing.Exit_Ammo(ammo(u + 1), Shape1)
            u = u + 1
            jineng = "退弹": nosice.URL = App.Path & "\素材\reload.mp3"
            AmmoColor.Enabled = True
        End Select
        tool(Index + 1) = 0
        Text1.Text = Text1.Text & vbCrLf & "你使用了" & jineng
    End If
End If
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)  '技能
If Button = 2 Then
    For i = 1 To 8
        If tool(i) = 0 Then tool(i) = dtool(Index + 1): Exit For
    Next
    dtool(Index + 1) = 0
    Call GiveTool(Image1(i - 1), Image2(Index), tool(i))
    For i = 1 To 8
        Image2(i - 1).Enabled = False
    Next
    For i = 0 To 1
        Command2(i).Enabled = False
    Next
    AI.Enabled = True
Else
    '防止混淆 [ X = 10 ] 验证
    If X = 10 Then
        If dtool(Index + 1) <> 0 Then
            Image2(Index).Picture = LoadPicture("")
            Debug.Print Index + 1, dtool(Index + 1)
            Select Case dtool(Index + 1)
            Case 1
                H_A = du_gameing.Hurt_Add(Label2)
                djineng = "伤害增加"
            Case 2
                Call du_gameing.Blood_Add(Label4)
                djineng = "血量增加"
            Case 3
                ammostatc = CBool(IIf(ammo(u + 1), 1, 0))
                djineng = "子弹查看"
            Case 4
                Call du_gameing.Exit_Ammo(ammo(u + 1), Shape1)
                If ammo(u + 1) Then tb = tb - 1 Else fb = fb - 1
                u = u + 1
                djineng = "退弹": nosice.URL = App.Path & "\素材\reload.mp3"
                AmmoColor.Enabled = True
            End Select
            dtool(Index + 1) = 0
            Text1.Text = Text1.Text & vbCrLf & "AI使用了" & djineng
        End If
    End If
End If
End Sub

Private Sub Label7_Click()
Form1.Show: Form2.Hide
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbRed
End Sub

Private Sub reload_ammo_nosice_Timer()
Static num
num = num + 1
If num >= UBound(ammo) Then Start.Enabled = True: reload_ammo_nosice.Enabled = False
nosice.URL = App.Path & "\素材\reload.mp3"
End Sub

Private Sub RoomStateChecking_Timer()
If Send_GameData.State <> 7 Then Form1.Show: Form3.Hide: MsgBox "房间中断，已返回界面。": RoomStateChecking.Enabled = False
End Sub

Private Sub Send_GameData_ConnectionRequest(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If Send_GameData.State <> sckClosed Then _
Send_GameData.Close
'接受具有 requestID 参数的
'连接。
Send_GameData.Accept requestID
End Sub

Private Sub Send_GameData_DataArrival(ByVal bytesTotal As Long)
'为进入的数据声明一个变量。
Dim Data As String
Data = ""
Send_GameData.GetData Data
If Left(Data, 6) = "<Data>" Then
    Select Case GetGameLabel(Data, "[", "]")
    Case "[Ammo]"
        AmmoData = Mid(Data, InStr(1, Data, "]") + 1)
        ReDim ammo(Len(AmmoData))
        For i = 1 To Len(AmmoData)
            ammo(i) = CBool(Val(Mid(AmmoData, i, 1)))
        Next
    Case "[AmmoNum]"
    
    Case "[Blood]"
    
    Case "[Text]"
    
    Case "[Tool]"
    
    End Select
Else
    Information.Text = Information.Text & Data
End If
End Sub

Private Sub Start_Timer()
Randomize
Static i, countnum
i = i + 1: countnum = countnum + 1
If i = 9 Then i = 0: GoTo ends
    If tool(i) = 0 Then
        tool(i) = Int(Rnd * 4) + 1
        Select Case tool(i)
        Case 1
            Image1(i - 1).Picture = LoadPicture(App.Path & "\素材\攻击.bmp")
        Case 2
            Image1(i - 1).Picture = LoadPicture(App.Path & "\素材\生命治疗.bmp")
        Case 3
            Image1(i - 1).Picture = LoadPicture(App.Path & "\素材\检查子弹.bmp")
        Case 4
            Image1(i - 1).Picture = LoadPicture(App.Path & "\素材\退弹.bmp")
        End Select
    End If
    
    If dtool(i) = 0 Then
        dtool(i) = Int(Rnd * 4) + 1
        Debug.Print i, dtool(i)
        dtool(i) = Int(Rnd * 4) + 1
        Select Case dtool(i)
        Case 1
            Image2(i - 1).Picture = LoadPicture(App.Path & "\素材\攻击.bmp")
        Case 2
            Image2(i - 1).Picture = LoadPicture(App.Path & "\素材\生命治疗.bmp")
        Case 3
            Image2(i - 1).Picture = LoadPicture(App.Path & "\素材\检查子弹.bmp")
        Case 4
            Image2(i - 1).Picture = LoadPicture(App.Path & "\素材\退弹.bmp")
        End Select
    End If
ends:
If countnum = 2 Then Start.Enabled = False: countnum = 0
Command2(0).Enabled = True: Command2(1).Enabled = True
GameStateCheck.Enabled = True
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub StateCheck_Timer()
If Send_GameData.State = 0 Then StateShow.Caption = "关闭状态"
If Send_GameData.State = 1 Then StateShow.Caption = "打开"
If Send_GameData.State = 2 Then StateShow.Caption = "正在侦听"
If Send_GameData.State = 3 Then StateShow.Caption = "连接挂起"
If Send_GameData.State = 4 Then StateShow.Caption = "识别主机"
If Send_GameData.State = 5 Then StateShow.Caption = "已识别主机"
If Send_GameData.State = 6 Then StateShow.Caption = "正在连接"
If Send_GameData.State = 7 Then StateShow.Caption = "链接成功"
If Send_GameData.State = 8 Then StateShow.Caption = "同级人员正在关闭连接"
If Send_GameData.State = 9 Then StateShow.Caption = "链接错误"
If Send_GameData.State <> 7 Then
    If CBool(Val(Form4.House_Owner_Check.Caption)) Then
        GamePlay_House_Owner_T
    Else
        GamePlay_House_Owner_F (Form4.House_Owner_Name.Caption)
    End If
Else
    RoomStateChecking.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
Static Sec
Sec = Sec + 1
If Sec Mod 2 = 0 Then Label8.ForeColor = vbRed Else Label8.ForeColor = vbBlack
If Sec = 10 Then Timer1.Enabled = False
End Sub
