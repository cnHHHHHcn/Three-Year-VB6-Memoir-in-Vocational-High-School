Attribute VB_Name = "SetMouse"
'模块 SetMouse 作用域：公用

'引用 API

'GetKeyState 获取按键状态(全局检测)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'SetDoubleClickTime 设置鼠标左键双击时间间隔
Public Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
'GetDoubleClickTime 获取鼠标左键双击时间间隔
Public Declare Function GetDoubleClickTime Lib "user32" () As Long
'SetCursorPos 设置鼠标在屏幕位置
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'GetCursorPos 获取鼠标在屏幕位置
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'mouse_event 鼠标点击
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'创建类型 MouseData  记录鼠标的位置、按键、状态值
Public Type MouseData
    MouseX As Long
    MouseY As Long
    Button As Integer
    State As Integer
End Type

'创建类型 POINTAPI  鼠标在屏幕上的 X、Y
'在 GetCurorPos、SetCursorPos 会用到
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'创建 枚举类 MouseKey  鼠标按键值
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

'鼠标  按键值  | 鼠标状态 值
'左键    +1    | 按下      -1
'右键    +2    | 无         0
'中键    +4    | 抬起      +1
'无键     0

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

'检查名称重复函数
Public Function NameRepeatCheck(NewName As String, OldName As String, Delimiter As String) As Boolean
    Dim NameArray() As String
    '默认为 False
    NameRepeatCheck = False
    '把 OldName变量 以 Delimiter变量 为分隔符，分割字符串。以数组的形式 赋入 NameArray变量
    NameArray = Split(OldName, Delimiter)
    '历遍 NameArray数组
    For i = LBound(NameArray) To UBound(NameArray)
        '如果 名称有重复 ，则报错，并返回 True
        If NameArray(i) = NewName Then NameRepeatCheck = True
    Next
End Function
