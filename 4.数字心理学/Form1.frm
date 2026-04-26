VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12915
   ClientLeft      =   4170
   ClientTop       =   1260
   ClientWidth     =   20190
   LinkTopic       =   "Form1"
   ScaleHeight     =   12915
   ScaleWidth      =   20190
   Begin VB.Timer Timer3 
      Left            =   11520
      Top             =   6240
   End
   Begin VB.Timer Timer2 
      Left            =   10560
      Top             =   6240
   End
   Begin VB.Timer Timer1 
      Left            =   9480
      Top             =   6240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "结果"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   23
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   9120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   2
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   1
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   0
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   22
      Top             =   7800
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   21
      Top             =   7800
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "年份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   20
      Top             =   7800
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   4800
      TabIndex        =   19
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   8040
      TabIndex        =   18
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   1800
      TabIndex        =   17
      Top             =   4200
      Width           =   255
   End
   Begin VB.Line Line6 
      X1              =   4920
      X2              =   2400
      Y1              =   3840
      Y2              =   6960
   End
   Begin VB.Line Line5 
      X1              =   4920
      X2              =   7560
      Y1              =   3840
      Y2              =   6960
   End
   Begin VB.Line Line4 
      X1              =   4920
      X2              =   4920
      Y1              =   5160
      Y2              =   6960
   End
   Begin VB.Line Line3 
      X1              =   3840
      X2              =   6000
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   6720
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   7560
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   4320
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   2640
      TabIndex        =   14
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3480
      TabIndex        =   13
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   5280
      TabIndex        =   12
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   7080
      TabIndex        =   11
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   6120
      TabIndex        =   10
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   4440
      TabIndex        =   9
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4440
      TabIndex        =   8
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3240
      TabIndex        =   7
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   5520
      TabIndex        =   6
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5520
      TabIndex        =   5
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6480
      TabIndex        =   4
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   3
      Top             =   4440
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Dim a(1 To 30), b(1 To 3), c(1 To 3)
gh1 = Val(Text3.Text): gh2 = Val(Text2.Text): gh3 = Val(Text1.Text)
If (1 <= gh1 And gh1 <= 12) And (1 <= gh2 And gh2 <= 31) And (1900 <= gh3 And gh3 <= 3000) Then
    
    '年份输出label1(1)后两位 , label1(2)前两位
    For i = 1 To Len(Text1.Text)
        a(i) = Val(Mid(Text1.Text, i, 1))
    Next
    If a(3) + a(4) > 10 Then
        a(5) = a(3) + a(4)
        For i = 1 To Len(Str(a(5))) - 1
            a(i + 5) = Val(Mid(a(5), i, 1))
        Next
        a(8) = a(6) + a(7)
        Label1(1).Caption = a(8)
    Else
        a(8) = a(3) + a(4)
        If a(8) = 10 Then a(8) = 1
        Label1(1).Caption = a(8)
    End If
    
    If a(1) + a(2) > 10 Then
        a(9) = a(1) + a(2)
        For i = 1 To Len(Str(a(9))) - 1
            a(i + 9) = Val(Mid(a(9), i, 1))
        Next
        a(12) = a(10) + a(11)
        Label1(1).Caption = a(12)
    Else
        a(12) = a(1) + a(2)
        If a(12) = 10 Then a(12) = 1
        Label1(2).Caption = a(12)
    End If
    '日期输出label1(4)
    For i = 1 To Len(Text2.Text)
        a(i + 12) = Val(Mid(Text2.Text, i, 1))
    Next
    If a(13) + a(14) > 10 Then
        a(15) = a(13) + a(14)
        For i = 1 To Len(Str(a(15))) - 1
            a(i + 15) = Val(Mid(a(15), i, 1))
        Next
        a(18) = a(17) + a(16)
        Label1(4).Caption = a(18)
    Else
        a(18) = a(13) + a(14)
        If a(18) = 10 Then a(18) = 1
        Label1(4).Caption = a(18)
    End If
    '月份输出label1(5)
    For i = 1 To Len(Text3.Text)
        a(i + 18) = Val(Mid(Text3.Text, i, 1))
    Next
    a(21) = a(19) + a(20)
    If a(21) = 10 Then a(21) = 1
    Label1(5).Caption = a(21)
    
     'label1(3)= label1(1) + label1(2)
    a(22) = Val(Label1(1).Caption) + Val(Label1(2).Caption)
    If a(22) >= 10 Then
        For i = 1 To Len(Str(a(22))) - 1
            a(i + 22) = Val(Mid(a(22), i, 1))
        Next
        a(25) = a(23) + a(24)
        Label1(3).Caption = a(25)
    Else
        a(25) = Val(Label1(1).Caption) + Val(Label1(2).Caption)
        Label1(3).Caption = a(25)
    End If
    
    'label1(6)= label1(4) + label1(5)
    a(22) = Val(Label1(4).Caption) + Val(Label1(5).Caption)
    If a(22) >= 10 Then
        For i = 1 To Len(Str(a(22))) - 1
            a(i + 22) = Val(Mid(a(22), i, 1))
        Next
        a(25) = a(23) + a(24)
        Label1(6).Caption = a(25)
    Else
        a(25) = Val(Label1(4).Caption) + Val(Label1(5).Caption)
        Label1(6).Caption = a(25)
    End If
     'label1(0)= label1(3) + label1(6)
    a(22) = Val(Label1(3).Caption) + Val(Label1(6).Caption)
    If a(22) >= 10 Then
        For i = 1 To Len(Str(a(22))) - 1
            a(i + 22) = Val(Mid(a(22), i, 1))
        Next
        a(25) = a(23) + a(24)
        Label1(0).Caption = a(25)
    Else
        a(25) = Val(Label1(3).Caption) + Val(Label1(6).Caption)
        Label1(0).Caption = a(25)
    End If
    
    b(1) = Val(Label1(0).Caption)
    b(2) = Val(Label1(1).Caption)
    b(3) = Val(Label1(2).Caption)
    For i = 1 To 3
        c(i) = Val(Label1(3).Caption) + b(i)
    Next
    For k = 1 To 3
        If c(k) >= 10 Then
            For i = 1 To Len(Str(c(k))) - 1
                a(i + 22) = Val(Mid(c(k), i, 1))
            Next
            a(25) = a(23) + a(24)
            c(k) = a(25)
        End If
    Next
    Label1(7).Caption = c(3)
    Label1(8).Caption = c(2)
    Label1(9).Caption = c(1)
    
    b(1) = Val(Label1(0).Caption)
    b(2) = Val(Label1(4).Caption)
    b(3) = Val(Label1(5).Caption)
    For i = 1 To 3
        c(i) = Val(Label1(6).Caption) + b(i)
    Next
    For k = 1 To 3
        If c(k) >= 10 Then
            For i = 1 To Len(Str(c(k))) - 1
                a(i + 22) = Val(Mid(c(k), i, 1))
            Next
            a(25) = a(23) + a(24)
            c(k) = a(25)
        End If
    Next
    Label1(10).Caption = c(3)
    Label1(11).Caption = c(2)
    Label1(12).Caption = c(1)
    
    a(1) = Val(Label1(10).Caption): a(2) = Val(Label1(11).Caption)
    If a(1) + a(2) >= 10 Then
        a(9) = a(1) + a(2)
        For i = 1 To Len(Str(a(9))) - 1
            a(i + 9) = Val(Mid(a(9), i, 1))
        Next
        a(12) = a(10) + a(11)
        Label1(13).Caption = a(12)
    Else
        Label1(13).Caption = a(1) + a(2)
    End If
    
    a(1) = Val(Label1(7).Caption): a(2) = Val(Label1(8).Caption)
    If a(1) + a(2) >= 10 Then
        a(9) = a(1) + a(2)
        For i = 1 To Len(Str(a(9))) - 1
            a(i + 9) = Val(Mid(a(9), i, 1))
        Next
        a(12) = a(10) + a(11)
        Label1(14).Caption = a(12)
    Else
        Label1(14).Caption = a(1) + a(2)
    End If
    
    a(1) = Val(Label1(9).Caption): a(2) = Val(Label1(12).Caption)
    If a(1) + a(2) >= 10 Then
        a(9) = a(1) + a(2)
        For i = 1 To Len(Str(a(9))) - 1
            a(i + 9) = Val(Mid(a(9), i, 1))
        Next
        a(12) = a(10) + a(11)
        Label1(15).Caption = a(12)
    Else
        Label1(15).Caption = a(1) + a(2)
    End If
    
    Open "结果.txt" For Output As #1
    Print #1, "一切以自己为准！！！！"
    gh = Label1(0).Caption
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
    gh = Val(Label1(4).Caption)
    Select Case gh
    Case 1
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
    gh = Val(Label1(1).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
    Close
    Timer1.Interval = 1
Else
MsgBox "超出范围！", 48, "警告"
End If
End Sub

Private Sub Command2_Click()
Set a = CreateObject("wscript.shell")
a.run "结果.txt"
Command2.Enabled = False
End Sub

Private Sub Timer1_Timer()
Open "结果.TXT" For Append As #1
Print #1, "20-40岁"
    gh = Val(Label1(10).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
        gh = Val(Label1(11).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
        gh = Val(Label1(13).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
Close
Timer1.Enabled = False
Timer2.Interval = 1
End Sub

Private Sub Timer2_Timer()
Open "结果.TXT" For Append As #1
Print #1, "40-60岁"
    gh = Val(Label1(9).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
        gh = Val(Label1(12).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
        gh = Val(Label1(15).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
Close
Timer2.Enabled = False
Timer3.Interval = 1
End Sub

Private Sub Timer3_Timer()
Open "结果.TXT" For Append As #1
Print #1, "60岁之后"
    gh = Val(Label1(7).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
        gh = Val(Label1(8).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
        gh = Val(Label1(14).Caption)
    Select Case gh
    Case 1
        Print #1, "1号人格"
        Print #1, "优点：创造、有领导能力、自信、有分析能力、独有组织系统，天性聪明.多才多艺，最有原则性的艺术家，是个很得人心的朋友，许多人被吸引，而且为人仁德，获得朋友的喜爱。"
        Print #1,
        Print #1, "缺点：傲慢固执、孤独、个人主义、沉迷者.冷漠苛刻，自我为中心，容易沉迷于某种不良嗜好。"
        Print #1,
        Print #1, "特质：感情独立，在交际应酬和钱财管理方面都不依赖对方，向往各自的私人空间，活动自由，孤独的。"
        Print #1,
        Print #1, "突破点：试着放下自我，少发表个人意见，多时间来倾听，并给予他人讲话的机会。"
    Case 2
        Print #1, "2号人格"
        Print #1, "优点：很会沟通有耐性宽容温和有礼知足有同情心很温和有爱心受欢迎，在集体生活中能展示自己的才华，适合做外交官；善沟通、有耐心、配合、读心术、和平主义"
        Print #1,
        Print #1, "缺点：犹豫不决易被影响依赖性强得过且过没条理易意志消沉，优柔寡断的性格会错失很多机遇，缺少主张；敏感、胡思乱想、冷暴力"
        Print #1,
        Print #1, "特质：重感情、温和有礼、也有较强的依赖心理金少主张，很少去争取自己的机会和权利。"
        Print #1,
        Print #1, "突破点：学会独立，积极的抓住机会，展现自我的才华会获得成功，需要突破胆量"
    Case 3
        Print #1, "3号人格"
        Print #1, "优点：行动善于言辞乐观精力充沛有创意多才多艺为人乐观，是欢乐使者，有极强的行动力，有演说者的才能。行动力强、善表达、善销售、热情仗义、有感染力"
        Print #1,
        Print #1, "缺点：任性冲动爱批评人过度紧张情绪化我行我素，常常抱怨与批评，怀疑自己，容易愤怒，有时候说话不经过大脑，为人处事不成熟；爱面子、任性、冲动、渴望被关注"
        Print #1,
        Print #1, "特质：感情冲动敏感，夸张处理事情，容易被激怒弃于感性与愉悦的沟通，迅速与人建立起关系。"
        Print #1,
        Print #1, "突破点：说话技巧一点，管住自己的嘴，先思考后开口"
    Case 4
        Print #1, "4号人格"
        Print #1, "优点：策划毅力聪明善于组织稳定追求知识，求知欲强烈有耐心，喜欢按步骤执行，为人忠诚值得信赖，有责任感，积极分析者。逻辑感强、重数据分析、善于组织策划、重承诺"
        Print #1,
        Print #1, "缺点：固执，步调缓慢，死板，犹豫不决，单刀直入，心直口快 缺乏安金感 自我设限 死板 不愿改变"
        Print #1,
        Print #1, "特质：很率直的感情，说话时要说重点和事实不能夸张，喜欢单直入，容易得罪人。"
        Print #1,
        Print #1, "突破点：放松放宽一些不是原则则的问题，可以灵活外理，大胆尝试新鲜事物物，让生命有更多的选择"
    Case 5
        Print #1, "5号人格"
        Print #1, "优点：方向强，随性思考，别出心裁，喜欢自由自在，易被启发，喜欢冒险，尝试名种新鲜事物，对人热情豪爽，朝气蓬勃有远见，自由自在、重情义、冷幽默、趋势感强、喜欢新鲜感"
        Print #1,
        Print #1, "缺点：固执，猜疑，易怒，过度表现，放纵自己不愿意被束缚会错失很多机会。"
        Print #1,
        Print #1, "特质：喜欢无拘无束的自由生活，方向感强也相对固执，在依赖独立两个极端中摇摆不定，经常不是人上人就就是下人。"
        Print #1,
        Print #1, "突破点：放下自己的固执己见，适当听取别人的建议对自己的言行加以约束，会获得更多成功的机会"
    Case 6
        Print #1, "6号人格"
        Print #1, "优点：忠诚，有同情心，和平主义者，有智慧，顾家，良好的策划能力，财富值特高、特别是在钱财管理方向，有社会责任、善于处理问题"
        Print #1,
        Print #1, "缺点：浪费，爱挑剔，自负，有压力感，喋喋不休，敏感和忧郁，容易迷失在琐事里，会因为害怕暴露缺点而避开参加一些活动。挑别、完美主义、胆小谨慎、内心隐蔽、易看不上他人"
        Print #1,
        Print #1, "特质：第六感强，尤其是女性在情感方面敏感，很注重家庭观念，性子急，轻过程只重视结果。"
        Print #1,
        Print #1, "突破点：活跃一些，主动一些，爱别人的同时，也应该会照顾自己"
    Case 7
        Print #1, "7号人格"
        Print #1, "优点：人缘好，精于分析，博学，别出心裁，有正义感，精明能干，内省力强，善于交际，博爱与包容，懂得提出问题和分析问题。幸运、善于分析、逻辑性强"
        Print #1,
        Print #1, "缺点：行动力弱，不体贴，迟缓冷淡，行动力遇到变故时易郁郁寡欢，通常因为恐惧害怕而选选择逃避。疑心重、享乐型、爱质疑、迟缓、易计较"
        Print #1,
        Print #1, "特质：故事型性格的人，在人际关系上能左右举源，但由于他们答应了太多人，却无法满足到每因此别人会认为他们爱放鸽子个人"
        Print #1,
        Print #1, "突破点：欣赏别别人的长处，即使他们们的喜好、性格和你不大相同，当发生重大问题时，用于接受事实，不要惧怕改变"
    Case 8
        Print #1, "8号人格"
        Print #1, "优点：影响力、责任感重、企业家潜质乐于帮助别人，注重名誉，可称为大慈善家。喜好权力、情感丰富、善销售抗压力强"
        Print #1,
        Print #1, "缺点：压力、独裁、消极、野心、专制、追求虚名，一旦少年得志、会执迷于金钱和权力，会因为要获得好处而放弃自己的良好，爱面子、记仇易钻牛角尖、纠结、掌控欲强"
        Print #1,
        Print #1, "特质：一个责任性的人，答应了人就一一定去完成，不然就会坐立不安，大多时大器晚成"
        Print #1,
        Print #1, "突破点：若始终保持善良、真诚、正直的品格以及不懈的努力，会活的平稳安逸，真正属于自己的幸福。"
    Case 9
        Print #1, "9号人格"
        Print #1, "优点：成功、爱冒险、奉献、情感丰富有趣的、洞察力很强，随机应变，说话做事都会被认可而且机会很多，机会多、有远见、有谈资、机灵活泼、易被同"
        Print #1,
        Print #1, "缺点：奢侈、无原则、寂寞、情感泛滥、自我，比较不知足，欲望没有封顶，兴趣爱好广泛，经常三分钟热度章志不坚定,专注度差、持久性弱、不易落地、原则性不强"
        Print #1,
        Print #1, "特质：一个全面型的人，性格开朗，机灵活泼、善于交际，在情感上多情浪漫爱做梦"
        Print #1,
        Print #1, "突破点：机会多，但是要经得住考验与诱惑，脚踏实地，梦想和目标不要脱离现实"
    End Select
Close
    Command2.Enabled = True
    Command2.SetFocus
    MsgBox "请点击结果按钮，查看结果。"
Timer3.Enabled = False
End Sub
