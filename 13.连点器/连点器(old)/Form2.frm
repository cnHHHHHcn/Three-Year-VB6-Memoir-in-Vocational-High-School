VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form2"
   ScaleHeight     =   8640
   ScaleWidth      =   11895
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'引用 API
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

'定义模块级变量

'TranscribeResultData   录制数据  类型：MouseData
Dim Transcribe As MouseData

'Index   记录号索引  类型：数值整形
Dim Index As Integer

'FileName   文件名称  类型：字符
'Comment    解释说明  类型：字符
Dim FileName As String, Comment As String

'WriteFlag    文件写入标识符  类型：布尔
'DownFlag     鼠标按下标识符  类型：布尔
'UpFlag       鼠标抬起标识符  类型：布尔
Dim WriteFlag As Boolean, _
    DownFlag As Boolean, UpFlag As Boolean
    
'MousePos     鼠标位置  类型：POINTAPI
Dim MousePos As POINTAPI

Private Sub Form_Load()
    Form2.WindowState = 2
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE) '获取窗体风格
    rtn = rtn Or WS_EX_LAYERED '设置风格标志为扩展
    SetWindowLong Me.hwnd, GWL_EXSTYLE, rtn '设置风格标志为扩展
    SetLayeredWindowAttributes Me.hwnd, 50, 50, LWA_ALPHA '设置窗体为透明，透明度200，（0-不可见，255-不透明）
End Sub

Public Sub Form_KeyPress(KeyAscii As Integer)
    '与下列结果进行匹配
    Select Case LCase(Chr(KeyAscii))
        '录制启动
        Case LCase(Form1.TranscribeText.Text)
            '如果 文件写入标识符(WriteFlag) 为 True ，则执行下列代码
            If Not WriteFlag Then
GetName:
                '向用户询问 鼠标录制名称 赋入 FileName 变量
                FileName = InputBox("请输入名称：", "录制")
                '如果 FileName 为 空串 ，则接着向用户询问
                If FileName = "" Then GoTo GetName
                '如果 SetMouse模块 的 检查重复函数 NameRepeatCheck 返回 True ，则接着向用户询问
                If SetMouse.NameRepeatCheck(FileName, Form1.ProjectDispose.KeyList(, "MouseTranscribe"), ",") Then MsgBox "鼠标录制名称重复！", 48: GoTo GetName
                
GetComment:
                '向用户询问 鼠标录制说明 赋入 Comment 变量
                Comment = InputBox("说明：", "录制")
                '如果 Comment 为 空串 ，则接着向用户询问
                If Comment = "" Then GoTo GetComment
                
                '以随机文件的方式打开 以 FileName变量的制 为名称的文件，并标识 文件记录号 为 #1
                Open App.Path & "\Transcribe-Data\" & FileName & ".MDS" For Random As #1
                
                '文件写入标识符(WriteFlag) 标记为 True
                WriteFlag = True
            End If
        '录制暂停
        Case LCase(Form1.TranscribePause_Text.Text)
            '文件写入标识符(WriteFlag) 标记为 False
            WriteFlag = False
            '关闭 文件记录号 为 #1 的文件
            Close #1
            
            '如果 SetMouse模块 的 检查重复函数 NameRepeatCheck 返回 False ，则执行下面代码块
            If Not SetMouse.NameRepeatCheck(FileName & ".MDS", Form1.ProjectDispose.KeyList(, "MouseTranscribe"), ",") Then
                '将此 鼠标录制数据文件 记录到 Mouse.Dispose 文件中
                Form1.ProjectDispose.CreateKey "", "MouseTranscribe", FileName & ".MDS" & ";" & Comment
            End If
    End Select
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '如果 文件写入标识符(WriteFlag) 为 True ，则执行下列语句块
    If WriteFlag Then
        '记录 记录号索引 + 1
        Index = Index + 1
        
        '调用 SetMouse模块 的 获取鼠标位置API GetCursorPos 记录 鼠标的 X、Y坐标 和 按键
        SetMouse.GetCursorPos MousePos
        Transcribe.MouseX = MousePos.X
        Transcribe.MouseY = MousePos.Y
        '如果当鼠标抬起时 不记录按键值
        Transcribe.Button = Button
        
        ' 与下列结果进行匹配
        Select Case True
            '鼠标按下
            Case DownFlag
                Transcribe.State = -1
            '鼠标抬起
            Case UpFlag
                Transcribe.State = 1
            '鼠标 无
            Case Else
                Transcribe.State = 0
        End Select
        '显示鼠标数据在 Form3窗体中
        With Form3
            .MouseButton.Caption = "Button：" & Transcribe.Button
            .MouseState.Caption = "State：" & Transcribe.State
            .MouseX.Caption = "MouseX：" & Transcribe.MouseX
            .MouseY.Caption = "MouseY：" & Transcribe.MouseY
        End With
        '记录 鼠标数据
        Put #1, Index, Transcribe
        '当鼠标抬起时 鼠标 按下、抬起 标识符均为 False
        If UpFlag Then
            UpFlag = False
            DownFlag = False
        End If
    Else
        With Form3
            .MouseButton.Caption = "Button："
            .MouseState.Caption = "State："
            .MouseX.Caption = "MouseX："
            .MouseY.Caption = "MouseY："
    End With
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '当鼠标按下时 _
     鼠标抬起标识符(UpFlag) 为 False _
     鼠标按下标识符(DownFlag) 为 True
    UpFlag = False
    DownFlag = True
    '调用 Form_MouseMove子过程
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '当鼠标按下时 _
     鼠标抬起标识符(UpFlag) 为 True _
     鼠标按下标识符(DownFlag) 为 False
    UpFlag = True
    DownFlag = False
    '调用 Form_MouseMove子过程
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show: Form2.Hide: Form3.Hide
End Sub
