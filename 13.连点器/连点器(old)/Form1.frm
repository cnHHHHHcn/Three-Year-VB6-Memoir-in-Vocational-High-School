VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   11745
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer GetMousePosition 
      Interval        =   10
      Left            =   1560
      Top             =   3840
   End
   Begin VB.Timer GetKeyboardState 
      Interval        =   500
      Left            =   6240
      Top             =   0
   End
   Begin VB.Timer TranscribeRun 
      Enabled         =   0   'False
      Left            =   4800
      Top             =   360
   End
   Begin VB.CommandButton ReadMouseTranscribeFile 
      Caption         =   "读取"
      Height          =   495
      Left            =   9480
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4560
      Left            =   8640
      TabIndex        =   18
      Top             =   240
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "按键设置"
      Height          =   2535
      Left            =   5280
      TabIndex        =   9
      ToolTipText     =   "按回车键保存"
      Top             =   120
      Width           =   2895
      Begin VB.TextBox TranscribePause_Text 
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   23
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
         TabIndex        =   20
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
         TabIndex        =   12
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
         TabIndex        =   10
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "鼠标设置"
      Height          =   5175
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "右键刷新鼠标键选项"
      Top             =   120
      Width           =   4815
      Begin VB.TextBox ChangeMouseY 
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
         Left            =   1920
         TabIndex        =   30
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox ChangeMouseX 
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
         Left            =   1920
         TabIndex        =   29
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox MouseLock 
         Caption         =   "鼠标锁定"
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton OpenTranscribeForm 
         Caption         =   "打开录制界面"
         Height          =   495
         Left            =   3240
         TabIndex        =   24
         Top             =   3600
         Width           =   1455
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
         Left            =   3720
         TabIndex        =   16
         Top             =   4725
         Width           =   975
      End
      Begin VB.Timer Run 
         Enabled         =   0   'False
         Left            =   4200
         Top             =   240
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
         Left            =   2040
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
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
         TabIndex        =   31
         Top             =   2280
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
         TabIndex        =   28
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000A&
         X1              =   2400
         X2              =   2400
         Y1              =   4200
         Y2              =   4680
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
         Left            =   2520
         TabIndex        =   26
         Top             =   4320
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
         Top             =   4320
         Width           =   360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   4800
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000A&
         X1              =   3600
         X2              =   3600
         Y1              =   4680
         Y2              =   5160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   1800
         X2              =   1800
         Y1              =   4680
         Y2              =   5160
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
         Left            =   1920
         TabIndex        =   15
         Top             =   4800
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
         TabIndex        =   14
         Top             =   4800
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   4800
         Y1              =   4680
         Y2              =   4680
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1320
         Width           =   2865
      End
   End
   Begin VB.CommandButton AcceptKey 
      Caption         =   "AcceptKey"
      Height          =   495
      Left            =   10680
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ListBox MouseSettingList 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "Form1.frx":0000
      Left            =   5280
      List            =   "Form1.frx":0002
      TabIndex        =   0
      ToolTipText     =   "右键刷新鼠标设置项"
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "鼠标设置项显示："
      Height          =   180
      Left            =   5280
      TabIndex        =   17
      Top             =   2760
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'定义 模块级变量

'引用 类 OperateKey 操作键值 分别赋入 ProjectDispose 和 MouseSetting 变量
Public ProjectDispose As New OperateKey   '作用域：公用
Private MouseSetting As New OperateKey    '作用域：私用

'MouseKeyReturn    鼠标按键  类型：字符
'IntervalReturn    时间间隔  类型：数值整形
'ReadyTimeReturn   预备时间  类型：数值整形
'IndexMax  索引最大值 与 TranscribeResultData 连用 类型：数值整形
'MouseLocked    鼠标锁定标识符  类型：布尔
'BootFlag    启动标识符  类型：布尔
'KeyChangeFlag    按键更改标识符  类型：布尔
'MousePosition    鼠标位置  类型：POINTAPI(SetMouse模块)
Dim MouseKeyReturn As String, _
    IntervalReturn As Integer, ReadyTimeReturn As Integer, IndexMax As Integer, KeyChangeAdd As Integer, _
    MouseLocked As Boolean, BootFlag As Boolean, KeyChangeFlag As Boolean
Dim MousePos As POINTAPI

'定义 模块级数组

'TranscribeResultData   录制数据  类型：MouseData(SetMouse模块) 动态数组
Dim TranscribeResultData() As MouseData
'FirstCheck     第一次键盘按键检测  类型：布尔 静态数组
'SecondCheck    第二次键盘按键检测  类型：布尔 静态数组
Dim FirstCheck(25) As Boolean, SecondCheck(25) As Boolean

'以下定义的变量均为 过程级变量

'发送鼠标点击子程序
Private Sub SendMouseKey(ByVal Key As MouseKey)
    '调用 SetMouse模块 的 发送鼠标点击API mouse_event
    SetMouse.mouse_event Key, 0, 0, 0, 0
    '如果 鼠标锁定的复选框的值为 1 ，则调用 SetMouse模块 的 设置鼠标位置 API SetCursorPos
    If MouseLock.Value = 1 Then SetMouse.SetCursorPos Mid(MousePosXList.Caption, 3), Mid(MousePosYList.Caption, 3)
End Sub

Private Sub AcceptKey_KeyPress(KeyAscii As Integer)
    '让接收按键(LCase(Chr(KeyAscii))) 与 下面进行匹配
    Select Case LCase(Chr(KeyAscii))
        '启动
        Case LCase(BootText.Text)
            '如果 模块级 BootFlag变量 为 False ，则执行下列代码
            If Not BootFlag Then
                '调用 启动自检函数 BootCheck 自检通过 BootFlag 为 True 否则 为 False
                If BootCheck() Then Ready.Enabled = True: BootFlag = True Else BootFlag = False
            End If
        '暂停
        Case LCase(PauseText.Text)
            BootFlag = False
            Run.Enabled = False
            TranscribeRun.Enabled = False
    End Select
End Sub

'启动自检函数
Private Function BootCheck() As Boolean
    '常量 Enter 赋入 13(键盘 Enter 的 Ascii码)
    '定义变量
    'MFlag    鼠标键值标识符  类型：布尔
    'IFlag    时间间隔标识符  类型：布尔
    'RFlag    预备时间标识符  类型：布尔
    'ErrorText    报错信息  类型：字符串
    Const Enter = 13
    Dim MFlag As Boolean, IFlag As Boolean, RFlag As Boolean
    Dim ErrorText As String
    
    '如果用户未按 Enter键，则程序自动输入
    If IntervalTime.Text <> "" Then IntervalTime_KeyPress Enter
    If ReadyTime.Text <> "" Then ReadyTime_KeyPress Enter
    If ChangeMouseX.Text <> "" Then ChangeMouseX_KeyPress Enter
    If ChangeMouseY.Text <> "" Then ChangeMouseY_KeyPress Enter
    
    '默认为 False
    BootCheck = False
    '如果 鼠标键值、时间间隔、预备时间 有值，则相对应的 标识符为 True
    If Mouse.Text <> "" Then MFlag = True Else ErrorText = ErrorText & "鼠标键值，"
    If Val(Mid(IntervalList.Caption, 4)) <> 0 Then IFlag = True Else ErrorText = ErrorText & "时间间隔，"
    If Val(Mid(ReadyList.Caption, 6)) <> 0 Then RFlag = True Else ErrorText = ErrorText & "预备时间，"
    '报错继续执行
    On Error Resume Next
    ErrorText = Left(ErrorText, Len(ErrorText) - 1) '启动自检报错信息
    '三个选项的值都为 True 则，返回值 为 True
    If MFlag And IFlag And RFlag Then BootCheck = True Else MsgBox "BootError：程序自检错误！" & Chr(10) & ErrorText & "未设置！", 48
End Function

'ChangeMouseX_KeyPress 子过程代码注释 与 下面 ChangeMouseY_KeyPress 同理
Private Sub ChangeMouseX_KeyPress(KeyAscii As Integer)
    '如果按下按键为 Enter，则执行下列代码
    If KeyAscii = 13 Then
        '把 名为ChangeMouseX的文本框的文本 赋入 模块级 对象MousePos的X中
        MousePos.X = Val(ChangeMouseX.Text)
        '名为ChangeMouseX的文本框的文本清空
        ChangeMouseX.Text = ""
        '显示鼠标X位置
        MousePosXList.Caption = "X：" & MousePos.X
        '如果 鼠标锁定复选框 可用，则 鼠标锁定复选框的值 = 1
        If MouseLock.Enabled Then MouseLock.Value = 1
        '调用 鼠标锁定复选框 单击事件过程
        Call MouseLock_Click
    End If
End Sub

Private Sub ChangeMouseY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MousePos.Y = Val(ChangeMouseY.Text)
        ChangeMouseY.Text = ""
        MousePosYList.Caption = "Y：" & MousePos.Y
        If MouseLock.Enabled Then MouseLock.Value = 1
        Call MouseLock_Click
    End If
End Sub

'鼠标锁定复选框
Private Sub MouseLock_Click()
    '把 鼠标锁定复选框的值 转化成布尔型 再取反 赋入 Flag变量
    Flag = CBool(MouseLock.Value)
    '根据 Flag变量 取反 再判定 是否开启 GetMousePosition定时器
    GetMousePosition.Enabled = Not Flag
    '把 Flag变量 赋入 模块级 MouseLocked变量
    MouseLocked = Flag
End Sub

'打开录制界面按钮
Private Sub OpenTranscribeForm_Click()
    Form2.Show: Form3.Show: Form1.Hide
End Sub

'保存按键设置
'--------------------------------------------------------------------------------------------------------------------------------------
'按键接收通用子过程 KeyPressGeneral
Private Sub KeyPressGeneral(Text As TextBox, KeyName As String)
    KeyChangeFlag = True
    '调用 按键重复检查函数 KeyPressRepeatCheck 返回 True ，则 代表有重复 报错 ，否则 执行 Else后语句
    If KeyPressRepeatCheck(BootText.Text, PauseText.Text, TranscribeText.Text, TranscribePause_Text.Text) Then
        MsgBox "按键有重复！", 48, "连点器"
        Text = ""
    Else
        '调用 按键检查函数 KeyPressCheck
        Result = KeyPressCheck(Text.Text)
        '如果返回结果为 False，则 报错 ，否则 更改项目配置
        If Result = False Then
            MsgBox "按键值无效！", 48
        Else
            KeyChangeFlag = False
            AcceptKey.SetFocus
            Text.Text = Result
            ProjectDispose.SetKey Result, , , KeyName     '写入
        End If
    End If
End Sub

'按键检查函数
Private Function KeyPressCheck(CheckKey As String)
    '如果字符为大写，则转小写
    If 65 <= Asc(CheckKey) And Asc(CheckKey) <= 90 Then
        KeyPressCheck = LCase(CheckKey)
    '否则如果字符为小写，则返回
    ElseIf 97 <= Asc(CheckKey) And Asc(CheckKey) <= 122 Then
        KeyPressCheck = CheckKey
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
            If KeyText(i) = KeyText(j) Then KeyPressRepeatCheck = True
        Next
    Next
End Function

Private Sub BootText_KeyPress(KeyAscii As Integer)  '启动按键
    '如果按下按键为 Enter，则 调用 KeyPressGeneral 子程序 ，以下同理
    If KeyAscii = 13 Then KeyPressGeneral BootText, "Boot"
End Sub

Private Sub TranscribePause_Text_KeyPress(KeyAscii As Integer)  '录制暂停按键
    If KeyAscii = 13 Then KeyPressGeneral TranscribePause_Text, "TranscribePause"
End Sub

Private Sub TranscribeText_KeyPress(KeyAscii As Integer) '录制按键
    If KeyAscii = 13 Then KeyPressGeneral TranscribeText, "Transcribe"
End Sub

Private Sub PauseText_KeyPress(KeyAscii As Integer) '暂停按键
    If KeyAscii = 13 Then KeyPressGeneral PauseText, "Pause"
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------

Private Sub ReadMouseTranscribeFile_Click()
    '定义变量 Data    鼠标数据  类型：MouseData(SetMouse模块)
    Dim Data As MouseData
    '判定 名为 Mouse的组合框的文本是否为 录制文件
    If CBool(InStr(1, Mouse.Text, ".MDS")) Then
        '打开 本工程下的 Transcribe-Data文件夹 的 录制文件
        Open App.Path & "\Transcribe-Data\" & Mouse.Text For Random As #1
            '当 文件指针 指到 文件尾部时 退出循环
            Do Until EOF(1)
                '索引值 + 1
                Index = Index + 1
                '获取鼠标数据
                Get #1, Index, Data
                '在列表框中显示
                List1.AddItem Data.MouseX & " " & Data.MouseY & " " & Data.Button & " " & Data.State
            Loop
        '关闭文件
        Close #1
    Else
        '报错
        MsgBox "此选项不是录制文件！", 16
    End If
End Sub


'鼠标设置项
'----------------------------------------------------------------
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
            'MouseSettingKeyNameArray = Split(MouseSetting.KeyList(), ",")   '暂定
            
            '加载 鼠标设置项 中的 MouseKey , Interval , ReadyTime , Lock , MouseX , MouseY 的键值
            Mouse.Text = MouseSetting.GetKey(, , "MouseKey")
            Run.Interval = MouseSetting.GetKey(, , "Interval")
            IntervalList.Caption = "间隔：" & Run.Interval & "ms"
            ReadyList.Caption = "预备时间：" & MouseSetting.GetKey(, , "ReadyTime") & "s"
            LockFlag = MouseSetting.GetKey(, , "Lock")
            If LockFlag = "True" Then
                MousePosXList.Caption = "X：" & MouseSetting.GetKey(, , "MouseX")
                MousePosYList.Caption = "Y：" & MouseSetting.GetKey(, , "MouseY")
                MouseLock.Value = 1
                GetMousePosition.Enabled = False
            Else
                MouseLock.Value = 0
                GetMousePosition.Enabled = True
            End If
        End If
    End If
End Sub
'----------------------------------------------------------------

Private Sub Mouse_Click()
    '把 名为 Mouse 的文本框的值 赋入 模块级 MouseKeyReturn 变量
    MouseKeyReturn = Mouse.Text
    MouseLock.Enabled = Not CBool(InStr(1, Mouse.Text, ".MDS"))
    MouseLock.Value = 0
    AcceptKey.SetFocus
End Sub

'保存鼠标设置项
Private Sub SaveMouseSetting_Click()
    Dim MouseSettingName As String
    '调用 启动自检函数 BootCheck
    If BootCheck() Then
GetName:
        '向用户询问 鼠标设置名称 赋入 MouseSettingName 变量
        MouseSettingName = InputBox("名称：", "连点器")
        '如果 MouseSettingName 为 空串 ，则接着向用户询问
        If MouseSettingName = "" Then GoTo GetName
        
        '调用 SetMouse模块 的 检查名称重复函数 NameRepeatCheck 判定 名称是否重复
        If SetMouse.NameRepeatCheck(MouseSettingName, MouseSetting.ClassList(), ",") Then MsgBox "鼠标设置项名称重复！", 48: GoTo GetName
        
        '创建 鼠标设置项
        MouseSetting.CreateClass MouseSettingName
        '在已创建的鼠标设置项中，添加 键名 键值(模块级变量)
        MouseSetting.CreateKey "", MouseSettingName, "MouseKey;" & MouseKeyReturn, "Interval;" & IntervalReturn, "ReadyTime;" & ReadyTimeReturn, _
                                                     "Lock;" & MouseLocked, "MouseX;" & MousePos.X, "MouseY;" & MousePos.Y
    End If
End Sub

'------------------------------------------------------------------------------------------------
'加载连点器配置 子程序
Private Sub LoadProjectDispose()
    Dim i As Integer
    '加载按键
    BootText.Text = ProjectDispose.GetKey(, , "Boot")
    PauseText.Text = ProjectDispose.GetKey(, , "Pause")
    TranscribeText.Text = ProjectDispose.GetKey(, , "Transcribe")
    TranscribePause_Text.Text = ProjectDispose.GetKey(, , "TranscribePause")
    
    '清除名为 Mouse 的组合框中的元素
    Mouse.Clear
    '加载鼠标录制项
    '获取 鼠标录制项 以数组的形式 赋入 MouseTranscribeOption 变量
    MouseTranscribeOption = Split(ProjectDispose.KeyList(, "MouseTranscribe"), ",")
    '历遍 MouseTranscribe 数组
    For i = LBound(MouseTranscribeOption) To UBound(MouseTranscribeOption)
        '向 名称为 Mouse 的组合框 添加 鼠标录制项
        Mouse.AddItem MouseTranscribeOption(i)
    Next
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '调用 子过程 LoadProjectDispose
    '右键刷新 加载连点器配置
    If Button = 2 Then LoadProjectDispose
End Sub

Private Sub Form_Load()
    AcceptKey.Left = -1000
    AcceptKey.Top = -1000
    '设置 类 ProjectDispose 和 MouseSetting 的属性
    ProjectDispose.File = App.Path & "\Mouse.dispose"
    ProjectDispose.Class = "Mouse"
    MouseSetting.File = App.Path & "\MouseSetting.Mouse"
    
    '程序加载时，记录 “a” - “z”的按键状态
    For i = 0 To 25
        SecondCheck(i) = SetMouse.GetKeyState(65 + i)
    Next
    '调用 子过程 LoadProjectDispose 来加载连点器配置
    LoadProjectDispose

    '调用 子过程 MouseSettingRefresh 来加载鼠标设置项
    MouseSettingRefresh
End Sub
'------------------------------------------------------------------------------------------------

'时间间隔设置
Private Sub IntervalTime_KeyPress(KeyAscii As Integer)
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
            '为 名为AcceptKey的按钮 设置焦点
            AcceptKey.SetFocus
        Else
            '报错
            MsgBox "时间间隔值无效！", 48
        End If
        IntervalTime.Text = ""
    End If
End Sub

'预备时间设置
Private Sub ReadyTime_KeyPress(KeyAscii As Integer)
    '如果按下按键为 Enter，则保存
    If KeyAscii = 13 Then
        '把 名称为 ReadyTime的文本框 的值赋入 RT 变量
        RT = Val(ReadyTime.Text)
        '判定值 是否在 1 - 10 之间
        If 0 < RT And RT < 11 Then
            '显示 RT 变量的值 并赋入 模块级 ReadyTimeReturn 变量
            ReadyList.Caption = "预备时间：" & RT & "s"
            ReadyTimeReturn = RT
            '为 名为AcceptKey的按钮 设置焦点
            AcceptKey.SetFocus
        Else
            '报错
            MsgBox "预备时间值无效！", 48
        End If
        ReadyTime.Text = ""
    End If
End Sub

'运行定时器
Private Sub Run_Timer()
    '让名称为 Mouse的组合框 与下列结果进行匹配 _
     匹配成功，则 调用 子过程SendMouseKey
    Select Case Mouse.Text
        
        Case "左键点击"
            SendMouseKey MouseLeftKey
            
        Case "左键按下"
            SendMouseKey MouseLeftKey_Down
            
        Case "左键弹起"
            SendMouseKey MouseLeftKey_Up
            
        Case "右键点击"
            SendMouseKey MouseRightKey
            
        Case "右键按下"
            SendMouseKey MouseRightKey_Down
            
        Case "右键弹起"
            SendMouseKey MouseRightKey_Up
            
        Case "滚轮键点击"
            SendMouseKey MouseMiddleKey
            
        Case "滚轮键按下"
            SendMouseKey MouseMiddleKey_Down
            
        Case "滚轮键弹起"
            SendMouseKey MouseMiddleKey_Up
        Case Else
            '如果都不匹配，则打开 Transcribe-Data 文件夹下的 录制文件
            Num = FreeFile
            FileName = Mouse.Text
            '清除 TranscribeResultData 数组元素
            ReDim TranscribeResultData(0)
            '并让索引最大值为 0
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


Private Sub TranscribeRun_Timer()
    '常量 ErrorNum    误差最大值  类型：数值整形  值 = 5
    '定义变量
    'Index    索引值  类型：数值长整形  静态变量
    'MouseDowningAdd    鼠标按下累积值  类型：数值长整形  静态变量
    'Mouse    鼠标键值  类型：数值长整形
    'MouseClickFlag    鼠标单击标识符  类型：布尔
    'MouseDowningFlag  鼠标长按标识符  类型：布尔
    'MouseUpFlag       鼠标抬起标识符  类型：布尔
    Const ErrorNum As Integer = 5
    Static Index As Long, MouseDowningAdd As Long
    Dim Mouse As Long, _
        MouseClickFlag As Boolean, MouseUpFlag As Boolean, MouseDowningFlag As Boolean
    
    '如果 模块级 BootFlag变量为 False ，则索引值 = 0
    If Not BootFlag Then Index = 0
    
    '索引值 + 1
    Index = Index + 1
    
    'Downing条件
    '1.当前 与 下一条 的按钮值一致
    '2.当前 与 下一条 的按钮状态值均为 -1
    MouseDowningFlag = TranscribeResultData(Index).Button = TranscribeResultData(Index + 1).Button And _
                      (TranscribeResultData(Index).State = -1 And TranscribeResultData(Index + 1).State = -1)
    '如果 满足Downing条件 ，则 MouseDowning变量 + 1
    If MouseDowningFlag Then MouseDowningAdd = MouseDowningAdd + 1
    
    'Click条件
    '1.当前 与 下一条 的按钮值一致
    '2.当前 的按钮状态值为 1 与 下一条 的按钮状态值为 -1
    MouseClickFlag = TranscribeResultData(Index).Button = TranscribeResultData(Index + 1).Button And _
                    (TranscribeResultData(Index).State = -1 And TranscribeResultData(Index + 1).State = 1)
    
    'Up条件
    '1.当前 的按钮状态值均为 1
    MouseUpFlag = TranscribeResultData(Index).State = 1
    
    '如果 (满足Downing条件 并且 MouseDowning变量 = 1)  或者  (满足Up条件 并且 MouseDowning变量 不在 0 - 5 的误差值之内) _
     调用 SetMouse模块 的 鼠标按键转换函数 MouseExchange
    '如果 满足Click条件 并且 MouseDowning变量 在 0 - 5 的误差值之内 _
     调用 SetMouse模块 的 鼠标按键单击转换函数 MouseKeyClick
    '如果以上都不满足 返回 0
    '最后赋入 MouseKey变量
    Mouse = Switch((MouseDowningFlag And MouseDowningAdd = 1) Or (MouseUpFlag And Not (0 <= MouseDowningAdd And MouseDowningAdd <= ErrorNum)), _
                      SetMouse.MouseExchange(TranscribeResultData(Index).Button, TranscribeResultData(Index).State), _
                      MouseClickFlag And (0 <= MouseDowningAdd And MouseDowningAdd <= ErrorNum), _
                      SetMouse.MouseKeyClick(TranscribeResultData(Index).Button), _
                      True, 0)
    
    '如果 取反(满足Downing条件 并且 MouseDowningAdd变量不为 1) 调用 SetMouse模块 的 发送鼠标点击API Mouse_event
    If Not (MouseDowningFlag And MouseDowningAdd <> 1) Then SetMouse.mouse_event Mouse, 0, 0, 0, 0
    '调用 SetMouse模块 的 设置鼠标位置API SetCursorPos
    SetMouse.SetCursorPos TranscribeResultData(Index).MouseX, TranscribeResultData(Index).MouseY
    '如果 满足Click条件 并且 MouseDowning变量 在 0 - 5 的误差值之内，则 MouseDonwing变量 = 0 并且 索引值 + 1 ，跳过下一条
    If MouseClickFlag And (0 <= MouseDowningAdd And MouseDowningAdd <= ErrorNum) Then MouseDowningAdd = 0: Index = Index + 1
    '如果满足Up条件，则 MouseDonwing变量 = 0
    If MouseUpFlag Then MouseDowningAdd = 0
    '如果 索引值 = 索引最大值 - 1 ，则 索引值 = 0
    If Index = IndexMax - 1 Then Index = 0
End Sub

Private Sub GetKeyboardState_Timer()
    Dim KeyCode As Integer
    If Not KeyChangeFlag Then
        For i = 0 To 25
            '“A” - “Z” 的Ascii码 65 - 90
            KeyCode = 65 + i
            '调用 SetMouse模块 的 获取按键状态API 赋入 模块级 FirstCheck数组
            FirstCheck(i) = SetMouse.GetKeyState(KeyCode)
            '如果 当前 与 上此 的值 不一样 ，则执行下列代码
            If FirstCheck(i) Xor SecondCheck(i) Then
                '调用 按键接受子过程
                AcceptKey_KeyPress KeyCode
                If Form2.Visible Then Form2.Form_KeyPress KeyCode
            End If
            '把 当前值 当作 上次的值
            SecondCheck(i) = FirstCheck(i)
        Next
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
    If Second = 0 Then Run.Enabled = True: Ready.Enabled = False
    '显示运算后的值
    ReadyList.Caption = "预备时间：" & Second & "s"
End Sub

Private Sub GetMousePosition_Timer()
    'MousePosition    鼠标位置  类型：POINTAPI(SetMouse模块)
    Dim MousePosition As POINTAPI
    '调用 SetMouse模块 的 获取鼠标位置API 赋入 MousePosition
    SetMouse.GetCursorPos MousePosition
    '显示
    MousePosXList.Caption = "X:" & MousePosition.X
    MousePosYList.Caption = "Y:" & MousePosition.Y
End Sub
