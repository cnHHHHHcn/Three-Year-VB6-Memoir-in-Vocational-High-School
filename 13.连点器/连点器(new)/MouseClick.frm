VERSION 5.00
Begin VB.Form MouseClick 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MouseClick"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   LinkTopic       =   "MouseClick"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Left            =   960
      TabIndex        =   32
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Timer GetKeyCheck 
      Interval        =   10
      Left            =   6240
      Top             =   0
   End
   Begin VB.CheckBox MouseLock 
      Caption         =   "鼠标锁定"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Timer MousePos 
      Interval        =   10
      Left            =   3840
      Top             =   2160
   End
   Begin VB.Timer TranscribeRun 
      Enabled         =   0   'False
      Left            =   4800
      Top             =   360
   End
   Begin VB.Frame Frame1 
      Caption         =   "鼠标设置"
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "按回车键保存，按右键刷新鼠标键选项"
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton KeyBoardSwitch 
         Caption         =   "键盘检测:开"
         Height          =   375
         Left            =   1200
         TabIndex        =   31
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton OpenExchange 
         Caption         =   "录制文件编写"
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Timer Run 
         Enabled         =   0   'False
         Left            =   4200
         Top             =   240
      End
      Begin VB.TextBox ChangeMousePosY 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   24
         Top             =   2325
         Width           =   1575
      End
      Begin VB.TextBox ChangeMousePosX 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         TabIndex        =   23
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Timer Ready 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4200
         Top             =   1320
      End
      Begin VB.CommandButton SaveMouseSetting 
         Caption         =   "保存"
         Height          =   375
         Left            =   3960
         TabIndex        =   15
         Top             =   2880
         Width           =   735
      End
      Begin VB.ComboBox Mouse 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "MouseClick.frx":0000
         Left            =   2040
         List            =   "MouseClick.frx":0002
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox IntervalTime 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox ReadyTime 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label TranscribeList 
         AutoSize        =   -1  'True
         Caption         =   "按键:0  状态:0"
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
         Left            =   2400
         TabIndex        =   29
         ToolTipText     =   "在录制时才有变化"
         Top             =   4440
         Width           =   1680
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   4800
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label ProgramStateList 
         AutoSize        =   -1  'True
         Caption         =   "状态:停止中"
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
         Left            =   120
         TabIndex        =   28
         Top             =   4440
         Width           =   1320
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000A&
         X1              =   2280
         X2              =   2280
         Y1              =   3360
         Y2              =   3840
      End
      Begin VB.Label MousePosYList 
         AutoSize        =   -1  'True
         Caption         =   "Y："
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
         Left            =   2400
         TabIndex        =   26
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label MousePosXList 
         AutoSize        =   -1  'True
         Caption         =   "X："
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
         Left            =   120
         TabIndex        =   25
         Top             =   3480
         Width           =   360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   4800
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "鼠标Y位置："
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
         TabIndex        =   22
         Top             =   2400
         Width           =   1740
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "鼠标X位置："
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
         TabIndex        =   21
         Top             =   1920
         Width           =   1740
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000A&
         X1              =   2280
         X2              =   2280
         Y1              =   4320
         Y2              =   4800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   2280
         X2              =   2280
         Y1              =   3840
         Y2              =   4320
      End
      Begin VB.Label ReadyList 
         AutoSize        =   -1  'True
         Caption         =   "预备时间：0s"
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
         Left            =   2400
         TabIndex        =   14
         Top             =   3960
         Width           =   1440
      End
      Begin VB.Label IntervalList 
         AutoSize        =   -1  'True
         Caption         =   "间隔：0ms"
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
         Left            =   120
         TabIndex        =   13
         Top             =   3960
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   4800
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "选择鼠标键："
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
         TabIndex        =   7
         Top             =   300
         Width           =   1890
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "时间间隔(单位:ms)："
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
         Top             =   840
         Width           =   3030
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "预备时间(单位:s)："
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
         TabIndex        =   5
         Top             =   1320
         Width           =   2865
      End
   End
   Begin VB.ListBox MouseSettingList 
      Appearance      =   0  'Flat
      Height          =   1830
      ItemData        =   "MouseClick.frx":0004
      Left            =   5280
      List            =   "MouseClick.frx":0006
      TabIndex        =   0
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "按键设置"
      Height          =   2535
      Left            =   5280
      TabIndex        =   8
      ToolTipText     =   "按回车键保存"
      Top             =   120
      Width           =   2895
      Begin VB.TextBox TranscribePauseText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1200
         TabIndex        =   19
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox TranscribeText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1200
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox PauseText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1200
         TabIndex        =   12
         Top             =   800
         Width           =   1455
      End
      Begin VB.TextBox BootText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1200
         TabIndex        =   10
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "暂停键:"
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
         TabIndex        =   20
         Top             =   1965
         Width           =   1110
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   3600
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "录制键:"
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
         TabIndex        =   17
         Top             =   1485
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "暂停键:"
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
         TabIndex        =   11
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "启动键:"
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
         TabIndex        =   9
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "鼠标项显示："
      Height          =   180
      Left            =   5280
      TabIndex        =   16
      Top             =   2760
      Width           =   1080
   End
End
Attribute VB_Name = "MouseClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'引用 类 OperateKey 操作键值 分别赋入 ProjectDispose 和 MouseSetting 变量
Public ProjectDispose As New OperateKey   '作用域：公用
Private MouseSetting As New OperateKey    '作用域：私用

'引用 类 MouseButton 鼠标按键 赋入 SetMouse 变量
Public WithEvents SetMouse As MouseButton     '作用域：公用
Attribute SetMouse.VB_VarHelpID = -1

'定义 模块级变量

'MouseKeyReturn    鼠标按键  类型：字符
'IntervalReturn    时间间隔  类型：数值整形
'ReadyTimeReturn   预备时间  类型：数值整形
'IndexMax    索引最大值 与 TranscribeResultData 连用  类型：数值整形
'MousePositionX    鼠标屏幕位置X  类型：数值长整形
'MousePositionY    鼠标屏幕位置Y  类型：数值长整形
'MouseLocked    鼠标屏幕位置锁定标识符  类型：布尔
'BootFlag    启动标识符  类型：布尔
Dim MouseKeyReturn As String, IntervalReturn As Integer, ReadyTimeReturn As Integer, IndexMax As Integer, _
    MousePositionX As Long, MousePositionY As Long, _
    MouseLocked As Boolean, BootFlag As Boolean

'FileName   文件名称  类型：字符
'Comment    解释说明  类型：字符
'Index   记录号索引  类型：数值整形
'WriteFlag    文件写入标识符  类型：布尔
Dim FileName As String, Comment As String, _
    Index As Integer, _
    WriteFlag As Boolean

'定义 模块级数组
'TranscribeResultData   录制数据  类型：MouseButtonData  动态数组
Dim TranscribeResultData() As MouseButtonData

'按键检测决定是否启动
'--------------------------------------------------------------------------------------------------------------------------------------
'按键接收子过程
Private Sub AcceptKey(KeyAscii As Integer)
    '让接收按键(UCase(Chr(KeyAscii))) 与 下面进行匹配
    Select Case UCase(Chr(KeyAscii))
        '启动
        Case UCase(BootText.Text)
            If Not BootFlag Then
                '调用 启动自检函数 BootCheck
                If BootCheck() Then Ready.Enabled = True: BootFlag = True: ProgramStateList.Caption = "状态:运行中" Else BootFlag = False
            End If
        '暂停
        Case UCase(PauseText.Text)
            Run.Enabled = False
            TranscribeRun.Enabled = False: BootFlag = False
            SetMouse.MouseLock MousePositionX, MousePositionY, False
            ProgramStateList.Caption = "状态:停止中"
        '录制启动
        Case UCase(TranscribeText.Text)
            '如果 文件写入标识符(WriteFlag) 为 True ，则执行下列代码
            If Not WriteFlag Then
                Dim Width As Long, Height As Long
                SetMouse.GetDeskTopArea Width, Height
GetName:
                '向用户询问 鼠标录制名称 赋入 FileName 变量
                FileName = InputBox("请输入录制文件名称：" & Chr(10) & "输入 NO 退出录制", "录制", "NO")
                '如果 FileName 为 空串 ，则接着向用户询问
                If UCase(FileName) = "NO" Then Exit Sub
                If Len(Trim(FileName)) = 0 Then GoTo GetName
                '如果 PublicMethod模块 的 检查重复函数 NameRepeatCheck 返回 True ，则接着向用户询问
                If PublicMethod.NameRepeatCheck(FileName, ProjectDispose.KeyList(, "MouseTranscribe"), ",") Then MsgBox "鼠标录制名称重复！", 48: GoTo GetName
GetComment:
                '向用户询问 鼠标录制说明 赋入 Comment 变量
                Comment = InputBox("说明：", "录制")
                '如果 Comment 为 空串 ，则接着向用户询问
                If Len(Trim(Comment)) = 0 Then GoTo GetComment
                '以随机文件的方式打开 以 FileName变量的制 为名称的文件，并标识 文件记录号 为 #1
                FileName = FileName & "(" & Width & "-" & Height & ")" & ".MDS"
                Open App.Path & "\Transcribe-Data\" & FileName For Random As #1
                
                Hook.SetHook
                '文件写入标识符(WriteFlag) 标记为 True
                WriteFlag = True
                ProgramStateList.Caption = "状态:录制中"
            End If
        '录制暂停
        Case UCase(TranscribePauseText.Text)
            Hook.UnHook
            '文件写入标识符(WriteFlag) 标记为 False
            WriteFlag = False
            '关闭 文件记录号 为 #1 的文件
            Close #1
            
            ProgramStateList.Caption = "状态:已录制完毕"
            
            TranscribeList.Caption = "按键:0  状态:0"
            '如果 PublicMethod模块 的 检查重复函数 NameRepeatCheck 返回 False ，则执行下面代码块
            If Not PublicMethod.NameRepeatCheck(FileName, ProjectDispose.KeyList(, "MouseTranscribe"), ",") Then
                '将此 鼠标录制数据文件 记录到 Mouse.Dispose 文件中
                ProjectDispose.CreateKey "", "MouseTranscribe", FileName & ";" & Comment
            End If
    End Select
End Sub

Private Sub Command1_Click()
    Command1.Caption = Val(Command1.Caption) + 1
End Sub

Private Sub GetKeyCheck_Timer()
    '调用子过程 CheckKeyState 检测按键
    CheckKeyState BootText.Text, PauseText.Text, TranscribeText.Text, TranscribePauseText.Text
End Sub

'检测按键
Private Sub CheckKeyState(ParamArray Key())
    '报错继续执行
    On Error Resume Next
    
    Dim Code As Integer
    '历遍循环 Key数组
    For i = LBound(Key) To UBound(Key)
        '如果 当前Key元素 为字符串是转化成大写ASCII码，否则为ASCII码
        Code = IIf(VarType(Key(i)) = 8, Asc(UCase(Key(i))), Key(i))
        '调用按键全局检测
        If PublicMethod.GetKeyState(Code) <> 1 And PublicMethod.GetKeyState(Code) <> 0 Then AcceptKey Code
    Next
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------

'启动自检 函数
Private Function BootCheck() As Boolean
    Const Enter = 13
    Dim ErrorText As String
    Dim MFlag As Boolean, IFlag As Boolean, RFlag As Boolean
    
    
    If IntervalTime.Text <> "" Then IntervalTime_KeyPress Enter
    If ReadyTime.Text <> "" Then ReadyTime_KeyPress Enter
    If ChangeMousePosX.Text <> "" Then ChangeMousePosX_KeyPress Enter
    If ChangeMousePosY.Text <> "" Then ChangeMousePosY_KeyPress Enter
    
    '默认为 False
    BootCheck = False
    If Mouse.Text <> "" Then MFlag = True Else ErrorText = ErrorText & "鼠标键，"
    If Val(Mid(IntervalList.Caption, 4)) <> 0 Then IFlag = True Else ErrorText = ErrorText & "预备时间，"
    If Val(Mid(ReadyList.Caption, 6)) <> 0 Then RFlag = True Else ErrorText = ErrorText & "时间间隔，"
    
    '如果 ErrorText 没有文本 ，则继续执行
    On Error Resume Next '报错继续运行
    ErrorText = Left(ErrorText, Len(ErrorText) - 1)
    
    '三个选项的值都为 True 则，返回值 为 True
    If MFlag And IFlag And RFlag Then BootCheck = True Else MsgBox "BootError：程序自检错误！" & Chr(10) & ErrorText & "未设置值！", 48
End Function

'当窗体释放时 释放钩子
Private Sub Form_Unload(Cancel As Integer)
    Hook.UnHook
End Sub

'保存按键设置
'--------------------------------------------------------------------------------------------------------------------------------------
'按键接收通用子过程 KeyPressGeneral
Private Sub KeyPressGeneral(Text As TextBox, KeyName As String)
    '调用 按键重复检查函数 KeyPressRepeatCheck
    If KeyPressRepeatCheck(BootText.Text, PauseText.Text, TranscribeText.Text, TranscribePauseText.Text) Then
        Text.Text = ""
    Else
        '调用 按键检查函数 KeyPressCheck
        Result = KeyPressCheck(Text.Text)
        '如果返回结果为 False，则 报错 ，否则 更改项目配置
        If Result = False Then
            MsgBox "按键值无效！", 48
        Else
            Text.Text = Result
            ProjectDispose.SetKey Result, , , KeyName     '写入
            GetKeyCheck.Enabled = True
        End If
    End If
End Sub

'按键检查函数
Private Function KeyPressCheck(CheckKey As String)
    '如果字符为大写，则返回；如果字符为小写，则转大写
    If (65 <= Asc(CheckKey) And Asc(CheckKey) <= 90) Or (97 <= Asc(CheckKey) And Asc(CheckKey) <= 122) Then
        KeyPressCheck = UCase(CheckKey)
    '否则返回 False
    Else
        KeyPressCheck = False
    End If
End Function

'按键重复检查函数
Private Function KeyPressRepeatCheck(ParamArray KeyText()) As Boolean
    '默认 KeyPressRepeatCheck 为 False
    KeyPressRepeatCheck = False
    '历遍 KeyText数组的总元素个数 - 1 次
    For i = LBound(KeyText) To UBound(KeyText) - 1
        '历遍 KeyText数组 总元素个数 - i 次
        For j = i + 1 To UBound(KeyText)
            '如果按键重复，则 KeyPressRepeatCheck 返回 True
            If KeyText(i) = KeyText(j) Then MsgBox "按键有重复！", 48, "连点器": KeyPressRepeatCheck = True
        Next
    Next
End Function

Private Sub BootText_KeyPress(KeyAscii As Integer)  '启动按键
    GetKeyCheck.Enabled = False
    '如果按下按键为 Enter，则 调用 KeyPressGeneral 子程序 ，以下同理
    If KeyAscii = 13 Then KeyPressGeneral BootText, "Boot"
End Sub

Private Sub KeyBoardSwitch_Click()
    GetKeyCheck.Enabled = Not GetKeyCheck.Enabled
    KeyBoardSwitch.Caption = "键盘检测:" & IIf(GetKeyCheck.Enabled, "开", "关")
End Sub

Private Sub OpenExchange_Click()
    GetKeyCheck.Enabled = False
    ScreenExchange.Show
End Sub

Private Sub PauseText_KeyPress(KeyAscii As Integer) '暂停按键
    GetKeyCheck.Enabled = False
    If KeyAscii = 13 Then KeyPressGeneral PauseText, "Pause"
End Sub

Private Sub TranscribeText_KeyPress(KeyAscii As Integer) '录制按键
    GetKeyCheck.Enabled = False
    If KeyAscii = 13 Then KeyPressGeneral TranscribeText, "Transcribe"
End Sub

Private Sub TranscribePauseText_KeyPress(KeyAscii As Integer)  '录制暂停按键
    GetKeyCheck.Enabled = False
    If KeyAscii = 13 Then KeyPressGeneral TranscribePauseText, "TranscribePause"
End Sub

'按键自动保存(专为懒人设计)
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    KeyPressGeneral BootText, "Boot"
    KeyPressGeneral PauseText, "Pause"
    KeyPressGeneral TranscribeText, "Transcribe"
    KeyPressGeneral TranscribePauseText, "TranscribePause"
    GetKeyCheck.Enabled = True
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '调用 子过程 LoadProjectDispose
    '右键刷新 加载连点器配置
    If Button = 2 Then LoadProjectDispose
End Sub

Private Sub MouseLock_Click()
    'Dim MousePositionX As Long, MousePositionY As Long
    '根据 名为 MouseLock复选框的值 来决定 MousePos定时器 是否可用
    MousePos.Enabled = Not CBool(MouseLock.Value)
    '把 名为 MouseLock复选框的值 赋入 模块级 MouseLocked 变量
    MouseLocked = CBool(MouseLock.Value)
    '把鼠标在屏幕的位置 赋入 模块级 MousePositionX、MousePositionY 变量
    MousePositionX = Val(Mid(MousePosXList.Caption, 3))
    MousePositionY = Val(Mid(MousePosYList.Caption, 3))
End Sub

Private Sub SaveMouseSetting_Click()
    Dim MouseSettingName As String
    '调用 启动自检函数 BootCheck，查看是否有问题
    If BootCheck() Then
GetName:
        '向用户询问 鼠标设置名称 赋入 MouseSettingName 变量
        MouseSettingName = InputBox("名称：", "连点器")
        '如果 MouseSettingName 为 空串 ，则接着向用户询问
        If MouseSettingName = "" Then GoTo GetName
        
        '如果 PublicMethod模块 的 检查重复函数 NameRepeatCheck 返回 True，则接着向用户询问
        If PublicMethod.NameRepeatCheck(MouseSettingName, MouseSetting.ClassList(), ",") Then MsgBox "鼠标设置项名称重复！", 48: GoTo GetName
        
        '创建 鼠标设置项
        MouseSetting.CreateClass MouseSettingName
        '在已创建的鼠标设置项中，添加 键名 键值(模块级变量)
        MouseSetting.CreateKey "", MouseSettingName, "MouseKey;" & MouseKeyReturn, "Interval;" & IntervalReturn, "ReadyTime;" & ReadyTimeReturn, _
                                                     "Lock;" & MouseLocked, "MouseX;" & MousePositionX, "MouseY;" & MousePositionY
    End If
End Sub

'录制鼠标动作
Private Sub SetMouse_MouseChange(Button As Integer, State As Integer, X As Long, Y As Long)
    Dim Data As MouseButtonData
    TranscribeList.Caption = "按键:" & Button & "  状态:" & State
    With Data
        .Button = Button
        .State = State
        .MouseX = X
        .MouseY = Y
    End With
    If WriteFlag Then
        Index = Index + 1
        Put #1, Index, Data
    Else
        Index = 0
    End If
End Sub

'鼠标设置项
'------------------------------------------------------------------------------------------------
'加载连点器配置 子程序
Private Sub LoadProjectDispose()
    Dim i As Integer
    '加载按键
    BootText.Text = ProjectDispose.GetKey(, , "Boot")
    PauseText.Text = ProjectDispose.GetKey(, , "Pause")
    TranscribeText.Text = ProjectDispose.GetKey(, , "Transcribe")
    TranscribePauseText.Text = ProjectDispose.GetKey(, , "TranscribePause")
    
    '清除名为 Mouse 的组合框中的所有元素
    Mouse.Clear
    
    '加载鼠标录制项
    '获取 鼠标录制项 以数组的形式 赋入 MouseTranscribe 变量
    MouseTranscribe = Split(ProjectDispose.KeyList(, "MouseTranscribe"), ",")
    '历遍 MouseTranscribe 数组
    For i = LBound(MouseTranscribe) To UBound(MouseTranscribe)
        '向 名称为 Mouse 的组合框 添加 鼠标录制项
        Mouse.AddItem MouseTranscribe(i)
    Next
End Sub

'子过程 鼠标设置项 刷新
Private Sub MouseSettingRefresh()    '清除名为 MouseSettingList 的列表框中的元素
    MouseSettingList.Clear

    '获取 新鼠标设置项 存入 MouseSettingNameArray
    MouseSettingNameArray = Split(MouseSetting.ClassList(), ",")
    For i = 0 To UBound(MouseSettingNameArray)
        '放入 名为 MouseSettingList 列表框 中
        MouseSettingList.AddItem MouseSettingNameArray(i)
    Next
End Sub

Private Sub Form_Load()
        
    Set SetMouse = New MouseButton
    
    '设置 类 ProjectDispose 和 MouseSetting 的属性
    ProjectDispose.File = App.Path & "\Mouse.dispose"
    ProjectDispose.Class = "Mouse"
    MouseSetting.File = App.Path & "\MouseSetting.Mouse"
    
    '调用 子过程 LoadProjectDispose 来加载连点器配置
    LoadProjectDispose

    '调用 子过程 MouseSettingRefresh 来加载鼠标设置项
    MouseSettingRefresh
End Sub

Private Sub MouseSettingList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '如果鼠标按键为 右键 ，则调用 鼠标设置项 刷新
    If Button = 2 Then MouseSettingRefresh
    
    '判定 选择的 鼠标录制项 是否为空值 _
     如果不是，则执行下列语句块 ，否则不执行
    If MouseSettingList.List(MouseSettingList.ListIndex) <> "" Then
        '如果鼠标按键为 左键 ，则执行下列语句
        If Button = 1 Then
            '设置 MouseSetting 的 Class 属性为 当前 名为 MouseSettingList 列表框 的选项
            MouseSetting.Class = MouseSettingList.List(MouseSettingList.ListIndex)
            
            '获取当前 鼠标设置项 中的 键名 赋入 MouseSettingKeyNameArray
            'MouseSettingKeyNameArray = Split(MouseSetting.KeyList(), ",")   暂定
            
            '加载 鼠标设置项 中的 MouseKey , Interval , ReadyTime 的键值
            Mouse.Text = MouseSetting.GetKey(, , "MouseKey")
            Run.Interval = MouseSetting.GetKey(, , "Interval")
            IntervalList.Caption = "间隔：" & Run.Interval & "ms"
            ReadyList.Caption = "预备时间：" & MouseSetting.GetKey(, , "ReadyTime") & "s"
            
            '如果选择 录制文件 ，则使 鼠标锁定不可用
            MouseLock.Enabled = Not CBool(InStr(1, Mouse.Text, ".MDS"))
            If Not MouseLock.Enabled Then MouseLock.Value = 0
            
            '返回 鼠标设置项 的 Lock 值
            LockResult = MouseSetting.GetKey(, , "Lock")
            If LockResult = "True" Then
                '如果返回结果为 "True" ，则 鼠标锁定，并获取鼠标位置 X、Y
                MouseLock.Value = 1
                MousePosXList.Caption = "X：" & MouseSetting.GetKey(, , "MouseX")
                MousePosYList.Caption = "Y：" & MouseSetting.GetKey(, , "MouseY")
            Else
                '如果返回结果为 "False" ，则 鼠标不锁定，并不获取鼠标位置 X、Y
                MouseLock.Value = 0
            End If
        End If
    End If
End Sub
'------------------------------------------------------------------------------------------------

'鼠标按键值设置
Private Sub Mouse_Click()
    '把 名为 Mouse 的文本框的值 赋入 模块级 MouseKeyReturn 变量
    MouseKeyReturn = Mouse.Text
    Mouse.ToolTipText = ProjectDispose.GetKey(, "MouseTranscribe", MouseKeyReturn)
    IntervalTime.Enabled = True
    '如果选择 录制文件 ，则使 鼠标锁定不可用
    MouseLock.Enabled = Not CBool(InStr(1, Mouse.Text, ".MDS"))
    If Not MouseLock.Enabled Then
        '强制 时间间隔 默认位 1ms
        IntervalTime.Text = 1
        MouseLock.Value = 0
        IntervalTime.Text = 1
        IntervalTime_KeyPress 13
        IntervalTime.Enabled = False
    End If
End Sub

'预备时间设置
Private Sub ReadyTime_KeyPress(KeyAscii As Integer)
    Dim RT As Integer
    '如果按下按键为 Enter，则保存
    If KeyAscii = 13 Then
        '把 名称为 ReadyTime的文本框 的值赋入 RT 变量
        RT = Val(ReadyTime.Text)
        '判定值 是否在 1 - 10 之间
        If 0 < RT And RT < 11 Then
            '显示 RT 变量的值 并赋入 模块级 ReadyTimeReturn 变量
            ReadyList.Caption = "预备时间：" & RT & "s"
            ReadyTimeReturn = RT
        Else
            '报错
            MsgBox "预备时间值无效！", 48
        End If
        ReadyTime.Text = ""
    End If
End Sub

'时间间隔设置
Private Sub IntervalTime_KeyPress(KeyAscii As Integer)
    Dim IT As Integer
    '如果按下按键为 Enter，则保存
    If KeyAscii = 13 Then
        '把 名称为 IntervalTime的文本框 的值赋入 IT 变量
        IT = Val(IntervalTime.Text)
        '判定 值 是否在 1 - 65534 之间
        If 0 < IT And IT < 65535 Then
            '把 IT 变量的值 赋入 运行定时器的 Interval 属性 与 模块级 IntervalReturn 变量，并显示
            Run.Interval = IT
            IntervalReturn = IT
            IntervalList.Caption = "间隔：" & IT & "ms"
        Else
            '报错
            MsgBox "时间间隔值无效！", 48
        End If
        IntervalTime.Text = ""
    End If
End Sub

'鼠标X位置设置
Private Sub ChangeMousePosX_KeyPress(KeyAscii As Integer)
    '如果按下按键为 Enter，则保存
    If KeyAscii = 13 Then
        '把 名为 ChangeMousePosX文本框 中的值 赋入 模块级 MousePositionX 变量
        MousePositionX = Val(ChangeMousePosX.Text)
        '设置鼠标X位置
        MousePosXList.Caption = "X：" & MousePositionX
        '清空 名为 ChangeMousePosX文本框 中的值
        ChangeMousePosX.Text = ""
        '使 名为 MouseLock复选框 的值为 1
        MouseLock.Value = 1
        '调用 MouseLock_Click子过程
        Call MouseLock_Click
    End If
End Sub

'鼠标Y位置设置
Private Sub ChangeMousePosY_KeyPress(KeyAscii As Integer)
    '如果按下按键为 Enter，则保存
    If KeyAscii = 13 Then
        '把 名为 ChangeMousePosY文本框 中的值 赋入 模块级 MousePositionY 变量
        MousePositionY = Val(ChangeMousePosY.Text)
        '设置鼠标Y位置
        MousePosYList.Caption = "Y：" & MousePositionY
        '清空 名为 ChangeMousePosY文本框 中的值
        ChangeMousePosY.Text = ""
        '使 名为 MouseLock复选框 的值为 1
        MouseLock.Value = 1
        '调用 MouseLock_Click子过程
        Call MouseLock_Click
    End If
End Sub

'预备时间定时器
Private Sub Ready_Timer()
    Dim Second As Integer
    '截取 ReadyList标签 中的秒数转换成数值型，并赋入 Second 变量
    Second = Val(Mid(ReadyList.Caption, 6, 2))
    '每过一秒 Second - 1
    Second = Second - 1
    '当 Second 的值为 0 时，启用 运行定时器 关闭预备时间定时器
    If Second = 0 Then MouseLock_Click: Run.Enabled = True: Ready.Enabled = False
    '显示运算后的值
    ReadyList.Caption = "预备时间：" & Second & "s"
End Sub

'每隔 10ms 获取鼠标在屏幕的位置
Private Sub MousePos_Timer()
    Dim MousePosX As Long, MousePosY As Long
    '调用 SetMouse 类模块 的 API GetCursorPos 获取鼠标位置
    SetMouse.GetMousePosition MousePosX, MousePosY
    '显示
    MousePosXList.Caption = "X：" & MousePosX
    MousePosYList.Caption = "Y：" & MousePosY
End Sub

'运行定时器
Private Sub Run_Timer()
    SetMouse.MouseLock MousePositionX, MousePositionY, MouseLocked
    '让名称为 Mouse的组合框 与下列结果进行匹配 _
     匹配成功，则 调用 SetMouse 类模块 SendMouseKey 方法
    Select Case Mouse.Text
        
        Case "LeftClick"
            SetMouse.SendMouseKey MouseLeftKey
            
        Case "LeftDown"
            SetMouse.SendMouseKey MouseLeftKey_Down
            
        Case "LeftUp"
            SetMouse.SendMouseKey MouseLeftKey_Up
            
        Case "RightClick"
            SetMouse.SendMouseKey MouseRightKey
            
        Case "RightDown"
            SetMouse.SendMouseKey MouseRightKey_Down
            
        Case "RightUp"
            SetMouse.SendMouseKey MouseRightKey_Up
            
        Case "MiddleClick"
            SetMouse.SendMouseKey MouseMiddleKey
            
        Case "MiddleDown"
            SetMouse.SendMouseKey MouseMiddleKey_Down
            
        Case "MiddleUp"
            SetMouse.SendMouseKey MouseMiddleKey_Up
            
        Case Else
            '如果都不匹配，则打开 Transcribe-Data 文件夹下的 录制文件
            Num = FreeFile
            FileName = Mouse.Text
            '清除 TranscribeResultData 数组元素，并把索引最大值设为 0
            ReDim TranscribeResultData(0)
            IndexMax = 0
            '以 随机文件 的方式 打开 鼠标录制文件
            Open App.Path & "\Transcribe-Data\" & FileName For Random As #Num
                '当 文件指针 指到 文件尾部 则结束循环
                Do Until EOF(Num)
                    '使索引最大值 + 1
                    IndexMax = IndexMax + 1
                    '定义数组 TranscribeResultData ，使元素 + 1
                    ReDim Preserve TranscribeResultData(IndexMax)
                    '获取 鼠标录制文件 数据 赋入 新增 元素空间
                    Get #Num, IndexMax, TranscribeResultData(IndexMax)
                Loop
            Close #Num
            '把 Run定时器 的赋入 Transcribe定时器 的 时间间隔属性值 中
            TranscribeRun.Interval = Run.Interval
            '关闭 Run 定时器，打开 TranscribeRun 定时器
            Run.Enabled = False: TranscribeRun.Enabled = True
    End Select
End Sub

'录制运行定时器
Private Sub TranscribeRun_Timer()
    On Error Resume Next
    '过程级变量
    
    Static Index As Long, MouseDowningAdd As Integer
    Dim MouseKey As Integer, _
        Button As Integer, State As Integer
    Dim MouseClickFlag As Boolean, MouseDowningFlag As Boolean, MouseUpFlag As Boolean
    
    If Not BootFlag Then Index = 0
    
    '设置常量 误差值 最大为 5
    Const ErrorNum As Integer = 5
    
    '索引值 + 1
    Index = Index + 1
    Debug.Print Index
    'Downing条件
    '1.当前 与 下一条 的鼠标按键一致
    '2.当前 与 下一条 的鼠标按键状态值均为 -1
    MouseDowningFlag = TranscribeResultData(Index).Button = TranscribeResultData(Index + 1).Button And _
                      (TranscribeResultData(Index).State = -1 And TranscribeResultData(Index + 1).State = -1)
    '如果满足 Downing条件 ，则 MouseDowningAdd变量 + 1
    If MouseDowningFlag Then MouseDowningAdd = MouseDowningAdd + 1
    
    'Click条件
    '1.当前 与 下一条 的鼠标按键一致
    '2.当前 状态值为 -1 并且 下一条 状态值为 1
    MouseClickFlag = TranscribeResultData(Index).Button = TranscribeResultData(Index + 1).Button And _
                    (TranscribeResultData(Index).State = -1 And TranscribeResultData(Index + 1).State = 1)
    
    'Up条件
    '1.鼠标按键状态值为 1
    MouseUpFlag = (TranscribeResultData(Index).State = 1)
    
    '如果 满足Downing条件 并且 MouseDowningAdd变量 = 1  或者  满足Up条件 并且 MouseDowningAdd变量 不在 0 - 5 之内， _
     调用 SetMouse 类模块 的 鼠标按键转换函数 MouseExchange
    '如果 满足Click条件 并且 MouseDowningAdd变量 在 0 - 5 之内， _
     调用 SetMouse 类模块 的 鼠标按键单击转换函数 MouseKeyClick
    '以上都不满足返回 0
    '最后把 结果 赋入 MouseKey变量
    MouseKey = Switch((MouseDowningFlag And MouseDowningAdd = 1) Or (MouseUpFlag And Not (0 <= MouseDowningAdd And MouseDowningAdd <= ErrorNum)), _
                      SetMouse.MouseExchange(TranscribeResultData(Index).Button, TranscribeResultData(Index).State), _
                      MouseClickFlag And (0 <= MouseDowningAdd And MouseDowningAdd <= ErrorNum), _
                      SetMouse.MouseKeyClick(TranscribeResultData(Index).Button), _
                      True, 0)
    
    '如果 取反(满足Downing条件 并且 MouseDowningAdd变量 <> 1)，则调用 SetMouse 类模块 的 SendMouseKey  发送鼠标点击
    If Not (MouseDowningFlag And MouseDowningAdd <> 1) Then SetMouse.SendMouseKey MouseKey
    '调用 SetMouse 类模块 的 API SetCursorPos 设置鼠标在屏幕的位置
    SetMouse.SetMousePosition TranscribeResultData(Index).MouseX, TranscribeResultData(Index).MouseY
    '如果 满足Click条件 并且 MouseDowningAdd变量 在 0 - 5 之内  ，则让 索引值 + 1 跳过下一条 并且 MouseDowningAdd变量 = 0
    If MouseClickFlag And (0 <= MouseDowningAdd And MouseDowningAdd <= ErrorNum) Then Index = Index + 1: MouseDowningAdd = 0
    '如果 满足Up条件，则让 MouseDowningAdd变量 = 0
    If MouseUpFlag Then MouseDowningAdd = 0
    '如果 索引值 = 索引最大值 - 1 ，则 索引值 = 0
    If Index >= IndexMax - 1 Then Index = 0
End Sub
