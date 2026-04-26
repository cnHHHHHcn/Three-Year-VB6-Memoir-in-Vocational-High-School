VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'引用 API

'GetKeyState 获取按键状态(全局检测)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'SetDoubleClickTime 设置鼠标左键双击时间间隔
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
'GetDoubleClickTime 获取鼠标左键双击时间间隔
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
'SetCursorPos 设置鼠标在屏幕位置
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'GetCursorPos 获取鼠标在屏幕位置
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'mouse_event 鼠标点击
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'获取窗口分辨率大小
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'获取桌面句柄
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'限制鼠标窗口
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

'类型 RECT 窗口窗口
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'创建类型 MouseButtonData 长度：12
Private Type MouseButtonData
    MouseX As Long
    MouseY As Long
    Button As Integer
    State As Integer
End Type

'创建类型 POINTAPI
'在 GetCurorPos、SetCursorPos 会用到
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum MouseMessage
    MouseMove = &H200

    LeftButtonDown = &H201
    LeftButtonUp = &H202

    RightButtonDown = &H204
    RightButtonUp = &H205

    MiddleButtonDown = &H207
    MiddleButtonUp = &H208
End Enum

'创建 枚举类 MouseKey
Public Enum MouseKey
    
    MouseNullKey = &H0           '鼠标初始状态
    
    MouseLeftKey = &H2 + &H4     '左键点击
    MouseLeftKey_Down = &H2      '左键按下
    MouseLeftKey_Up = &H4        '左键弹起
    
    MouseRightKey = &H8 + &H10   '右键点击
    MouseRightKey_Down = &H8     '右键按下
    MouseRightKey_Up = &H10      '右键弹起
    
    MouseMiddleKey = &H20 + &H40 '滚轮键点击
    MouseMiddleKey_Down = &H20   '滚轮键按下
    MouseMiddleKey_Up = &H40     '滚轮键弹起

End Enum

Public Event MouseChange(Button As Integer, State As Integer, X As Long, Y As Long)

Private Function Is_in(ByVal Part, ParamArray PartList()) As Boolean
    For i = LBound(PartList) To UBound(PartList)
        If Part = PartList(i) Then Is_in = True: Exit For
    Next
End Function

'获取桌面大小(即屏幕分辨率)
Public Sub GetDeskTopArea(Width As Long, Height As Long)
    Dim Area As RECT
    GetWindowRect GetDesktopWindow, Area
    Width = Area.Right
    Height = Area.Bottom
End Sub

'鼠标锁定
Public Sub MouseLock(PosX As Long, PosY As Long, IsLock As Boolean)
    Dim Area As RECT
    Area.Left = PosX
    Area.Right = PosX
    Area.Top = PosY
    Area.Bottom = PosY
    If IsLock Then ClipCursor Area Else ClipCursor 0
End Sub

'获取鼠标位置
Public Sub GetMousePosition(X As Long, Y As Long)
    Dim Pos As POINTAPI
    GetCursorPos Pos
    X = Pos.X
    Y = Pos.Y
End Sub

'设置鼠标位置
Public Sub SetMousePosition(X As Long, Y As Long)
    SetCursorPos X, Y
End Sub

'获取鼠标双击时间
Public Sub GetMouseDblClick(ClickTime As Long)
    Time = GetDoubleClickTime
End Sub

'设置鼠标双击时间
Public Sub SetMouseDblClick(ClickTime As Long)
    SetDoubleClickTime Time
End Sub


'鼠标  按键值  | 鼠标状态 值
'左键    +1    | 按下      -1
'右键    +2    | 无         0
'中键    +4    | 抬起      +1
'无键     0

'发送鼠标点击子程序
Public Sub SendMouseKey(ByVal MouseKey As MouseKey)
    '调用 SetMouse 类模块 的 API mouse_event
    mouse_event MouseKey, 0, 0, 0, 0
End Sub

'鼠标按键转换函数
Public Function MouseExchange(ByVal ButtonKey As Integer, ByVal State As Integer) As Long
    '检测 鼠标当前状态
    Select Case State
        '鼠标按键状态：按下
        Case -1
            '检测 鼠标当前按键，下面同理
            Select Case ButtonKey
                '鼠标按键：左键
                Case 1
                   MouseExchange = &H2
                '鼠标按键：右键
                Case 2
                   MouseExchange = &H8
                '鼠标按键：中键
                Case 4
                   MouseExchange = &H20
            End Select
        '鼠标按键状态：抬起
        Case 1
            Select Case ButtonKey
                Case 1
                   MouseExchange = &H4
                Case 2
                   MouseExchange = &H10
                Case 4
                   MouseExchange = &H40
            End Select
        '鼠标按键状态：无
        Case 0
           MouseExchange = &H0
    End Select
End Function

'鼠标按键单击转换函数
Public Function MouseKeyClick(ByVal ClickButton As Integer)
    '查看单击按钮
    Select Case ClickButton
        '左键单击
        Case 1
            MouseKeyClick = &H2 + &H4
        '右键单击
        Case 2
            MouseKeyClick = &H8 + &H10
        '中键单击
        Case 4
            MouseKeyClick = &H20 + &H40
    End Select
End Function

'鼠标消息转化
Public Sub MouseMessageExchange(MouseMSG As MouseMessage)
    Dim Button As Integer, State As Integer, X As Long, Y As Long
    Static LeftButton As Boolean, RightButton As Boolean, MiddleButton As Boolean
    '如果鼠标当前没有状态，则转到 NOClick 标签下
    If Not (LeftButton Or RightButton Or MiddleButton) And MouseMSG = MouseMove Then GoTo NOClick
    '查看鼠标哪个按键抬起，并将变量 State = 1
    Select Case MouseMSG
        Case LeftButtonUp
            LeftButton = False
            Button = 1
        Case RightButtonUp
            RightButton = False
            Button = 2
        Case MiddleButtonUp
            MiddleButton = False
            Button = 4
    End Select
    State = 1
    '如果鼠标一直在按住某个按键 或 刚按下某个按键时，则执行下列代码
    If (LeftButton Or RightButton Or MiddleButton) Or Is_in(MouseMSG, LeftButtonDown, MiddleButtonDown, RightButtonDown) Then
    '查看鼠标哪个按键第一次按下，并将变量 State = -1
        Select Case MouseMSG
            Case LeftButtonDown
                LeftButton = True
            Case RightButtonDown
                RightButton = True
            Case MiddleButtonDown
                MiddleButton = True
        End Select
        '查看鼠标一直在按住某个按键，并返回对应按钮值 赋予 变量 Button
        Button = Switch(LeftButton, 1, RightButton, 2, MiddleButton, 4)
        State = -1
    End If
NOClick:
    Me.GetMousePosition X, Y
    '触发 MouseChange 事件
    RaiseEvent MouseChange(Button, State, X, Y)
End Sub
'可以用以下数值(与枚举 MouseMessage 结合观看)试试
'eg: 513  512  512  512  512  512  514
