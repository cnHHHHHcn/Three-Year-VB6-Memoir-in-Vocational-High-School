VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8625
   ClientLeft      =   3285
   ClientTop       =   1095
   ClientWidth     =   11610
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11610
   Begin VB.Timer AmmoColor 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5440
      Top             =   1080
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   11400
      TabIndex        =   12
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Timer Start 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5440
      Top             =   520
   End
   Begin VB.Timer AI 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "射击"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   5880
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   6840
      Width           =   11655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "选择子弹数"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "射击"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   4200
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "装填子弹"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   1095
      Left            =   10080
      TabIndex        =   14
      Top             =   240
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
   Begin VB.Image Image3 
      Height          =   615
      Left            =   4440
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label7 
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
      TabIndex        =   13
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape Shape3 
      Height          =   35
      Left            =   5280
      Top             =   1680
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   5400
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   5400
      Top             =   840
      Width           =   495
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
      TabIndex        =   10
      Top             =   240
      Width           =   165
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   11640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11640
      Y1              =   6800
      Y2              =   6800
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   7
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   6
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   5
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   4
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   7
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   5
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   4
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   3
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
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
      Top             =   4800
      Width           =   945
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
      TabIndex        =   7
      Top             =   4800
      Width           =   945
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11640
      Y1              =   4680
      Y2              =   4680
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
      TabIndex        =   6
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "生命值："
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
      TabIndex        =   5
      Top             =   840
      Width           =   1260
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
      TabIndex        =   4
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "生命值："
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
      TabIndex        =   3
      Top             =   840
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   5640
      Y1              =   -120
      Y2              =   6920
   End
End
Attribute VB_Name = "Form2"
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

Private Sub AI_Timer()    'AI人机
On Error GoTo errorammo
Randomize
Dim o As Integer, l As Integer
Static countnum
countnum = countnum + 1

If countnum <= 4 Then
Starts:
    If c >= 5 Then c = 0: GoTo endss
    l = Int(Rnd * 8)
    If dtool(l + 1) = 0 Then c = c + 1: GoTo Starts
    Call Image2_Click(l)
endss:
    countnum = 4
Else
    tbs = tb: fbs = fb
    If ammo(u + 1) Then butler = "实": tb = tb - 1 Else butler = "虚": fb = fb - 1
errorammo:
    
    If ammostatc <> "" Then
        If ammostatc Then
            Image3.Picture = LoadPicture(App.Path & "\素材\大喷左.bmp"): Call Command2_Click(0)
            Text1.Text = Text1.Text & vbCrLf & "AI向你发射子弹，子弹为" & butler & "弹"
        Else
            Image3.Picture = LoadPicture(App.Path & "\素材\大喷右.bmp"): Call Command2_Click(1)
            Text1.Text = Text1.Text & vbCrLf & "AI向自己发射子弹，子弹为" & butler & "弹"
        End If
        ammostatc = "": GoTo ends
    End If
    If tbs < fbs Then
        Image3.Picture = LoadPicture(App.Path & "\素材\大喷右.bmp"): Call Command2_Click(1)
        Text1.Text = Text1.Text & vbCrLf & "AI向自己发射子弹，子弹为" & butler & "弹"
    ElseIf fbs < tbs Then
        Image3.Picture = LoadPicture(App.Path & "\素材\大喷左.bmp"): Call Command2_Click(0)
        Text1.Text = Text1.Text & vbCrLf & "AI向你发射子弹，子弹为" & butler & "弹"
    ElseIf tbs = fbs And (tbs <> 0 And fbs <> 0) Then
        o = Int(Rnd * 2)
        Select Case o
        Case 0
            Image3.Picture = LoadPicture(App.Path & "\素材\大喷左.bmp")
        Case 1
            Image3.Picture = LoadPicture(App.Path & "\素材\大喷右.bmp")
        End Select
        Call Command2_Click(o)
        If o = 0 Then Text1.Text = Text1.Text & vbCrLf & "AI向你发射子弹，子弹为" & butler & "弹" Else Text1.Text = Text1.Text & vbCrLf & "AI向自己发射子弹，子弹为" & butler & "弹"
    End If
    
ends:
    
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

Sub End_Game()
For i = 0 To 1
    Command2(i).Enabled = False
Next
For i = 0 To 7
    Image1(i).Enabled = False
Next
ReDim ammo(10 To 10): u = 0
Command3.Enabled = True
End Sub

Private Sub AmmoColor_Timer()
Shape1.FillColor = &HFFFFFF
AmmoColor.Enabled = False
End Sub

Private Sub Command1_Click()   '装填子弹
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
If f = UBound(ammo) Then
    rndnum = Int(Rnd * UBound(ammo)) + 1
    For i = 1 To rndnum
        ammo(Int(Rnd * UBound(ammo)) + 1) = True
        f = f - rndnum: t = rndnum
    Next
ElseIf t = UBound(ammo) Then
    rndnum = Int(Rnd * UBound(ammo)) + 1
    For i = 1 To rndnum
        ammo(Int(Rnd * UBound(ammo)) + 1) = False
        t = t - rndnum: f = rndnum
    Next
End If
tb = t: fb = f
Debug.Print
Debug.Print "------------------"
For i = 1 To UBound(ammo)
Debug.Print ammo(i)
Next
Debug.Print "------------------"

ReDim Preserve tool(8), dtool(8)
Label8.Caption = "实弹" & t & "发，虚弹" & f & "发"
Command1.Enabled = False
Command2(0).Enabled = True
Command2(1).Enabled = True
For i = 0 To 7
    Image1(i).Enabled = True
Next
Timer1.Enabled = True
Start.Enabled = True
End Sub

Private Sub Command2_Click(Index As Integer)     '射击
Select Case Index
Case 0
    fire Label2, u
Case 1
    fire Label4, u
End Select
Command4.SetFocus
End Sub

Sub Fire_()

End Sub


Private Sub Command2_GotFocus(Index As Integer)
On Error Resume Next
If ammo(u + 1) Then butler = "实" Else butler = "虚"
Select Case Index
Case 0
    Text1.Text = Text1.Text & vbCrLf & "你向自己发射子弹，子弹为" & butler & "弹"
Case 1
    Text1.Text = Text1.Text & vbCrLf & "你向AI发射子弹，子弹为" & butler & "弹"
End Select
End Sub

Private Sub Command2_LostFocus(Index As Integer)     '射击
For i = 0 To 1
    Command2(i).Enabled = False
Next
For i = 0 To 7
    Image1(i).Enabled = False
Next
If u = UBound(ammo) Then Call End_Game Else AI.Enabled = True
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
If 3 < num And num < 9 Then ReDim ammo(num): Command3.Enabled = False: Command1.Enabled = True Else GoTo Start
End Sub

Sub fire(player As Control, num As Integer)
On Error Resume Next
'子弹射击子程序

num = num + 1

Debug.Print "此发子弹为", num
Debug.Print ammo(num)

If ammo(num) Then
    If H_A Then player = player - 2 Else player = player - 1
    tb = tb - 1
Else
    fb = fb - 1
End If
H_A = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbBlack
End Sub

Private Sub Image1_Click(Index As Integer)   '技能
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
        jineng = "退弹"
        AmmoColor.Enabled = True
    End Select
    tool(Index + 1) = 0
    Text1.Text = Text1.Text & vbCrLf & "你使用了" & jineng
End If
End Sub

Private Sub Image2_Click(Index As Integer)   '技能
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
    djineng = "退弹"
    AmmoColor.Enabled = True
End Select
dtool(Index + 1) = 0
Text1.Text = Text1.Text & vbCrLf & "AI使用了" & djineng
End Sub

Private Sub Label7_Click()
Form1.Show: Form2.Hide
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbRed
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
If countnum = 4 Then Start.Enabled = False: countnum = 0
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
If Val(Label2.Caption) <= 0 Then
    MsgBox "you dead!"
    Call Start_Game
ElseIf Val(Label4.Caption) <= 0 Then
    MsgBox "you win!"
    Call Start_Game
End If
End Sub

Sub Start_Game()
Label2.Caption = "4"
Label4.Caption = "4"
Command3.Enabled = True
Command2(0).Enabled = False
Command2(1).Enabled = False
For i = 0 To 7
    Image1(i).Picture = LoadPicture
    Image2(i).Picture = LoadPicture
    tool(i + 1) = 0: dtool(i + 1) = 0: ReDim ammo(10 To 10)
Next
End Sub

Private Sub Timer1_Timer()
Static sec
sec = sec + 1
If sec Mod 2 = 0 Then Label8.ForeColor = vbRed Else Label8.ForeColor = vbBlack
If sec = 10 Then Timer1.Enabled = False
End Sub
