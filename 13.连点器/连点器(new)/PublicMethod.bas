Attribute VB_Name = "PublicMethod"
'GetKeyState 获取按键状态(全局检测)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long


Public Type ScreenPercent
    Width As Single
    Height As Single
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'创建类型 MouseButtonData 长度：7
Public Type MouseButtonData
    MouseX As Long
    MouseY As Long
    Button As Integer
    State As Integer
End Type

'创建类型 POINTAPI
'在 GetCurorPos、SetCursorPos 会用到
Public Type POINTAPI
    X As Long
    Y As Long
End Type


'检查重复函数
Public Function NameRepeatCheck(NewName As String, OldName As String, Delimiter As String) As Boolean
    Dim NameArray() As String
    '默认为 False
    NameRepeatCheck = False
    '把 OldName变量 以 Delimiter变量 为分隔符，分割字符串。 _
     以数组的形式 赋入 NameArray变量
    NameArray = Split(OldName, Delimiter)
    '历遍 NameArray数组
    For i = LBound(NameArray) To UBound(NameArray)
        '如果 名称有重复 ，则报错，并返回 True
        If NameArray(i) = NewName Then NameRepeatCheck = True
    Next
End Function

