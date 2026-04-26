Attribute VB_Name = "OperaKey"

'获取键值
Function GetKey(File As String, ByVal Class As String, ByVal KeyLabel As String) As String

Dim StartClass As Boolean, EndClass As Boolean, Flag As Boolean
Dim KeyLabelPosition As Integer
Dim Str As String

Num = FreeFile
Open File For Input As #Num
    Do Until EOF(Num)
        Line Input #Num, Str
        
        '如果起始含有类标签，则返回 True
        StartClass = InStr(1, Str, "[+" & Class & "-]")
        If StartClass Then Flag = True
        
        '如果末尾含有类标签，则返回 False
        EndClass = InStr(1, Str, "[-" & Class & "+]")
        If EndClass Then Flag = False
        
        '在类模块中提取键值
        If Not (StartClass Xor EndClass) And Flag Then
            KeyLabelPosition = InStr(Str, KeyLabel)
            If CBool(KeyLabelPosition) Then GetKey = Mid(Str, KeyLabelPosition + Len(KeyLabel) + 1): Exit Do
        End If
    Loop
Close #Num
End Function

'显示类名称
Function ClassList(File As String) As String
Dim Str As String

Num = FreeFile
Open File For Input As #Num
    Do Until EOF(Num)
        Line Input #Num, Str
        
        '寻找类模块标记，并返回位置
        Start = InStr(Str, "[+")
        ends = InStr(Str, "-]")
        
        '如果 Start 与 Ends 变量的值不为 0 ，则提取类模块 标签
        If CBool(Start) And CBool(ends) Then
            Class = ""
            For i = Start + 2 To ends - 1
                Class = Class & Mid(Str, i, 1)
            Next
            ClassList = ClassList & Class & ","
        End If
        
    Loop
Close #Num
'返回值
ClassList = Left(ClassList, Len(ClassList) - 1)
End Function

'显示类中的键名称
Function KeyList(File As String, ClassLabel As String) As String

Dim Str As String
Dim StartClass As Boolean, EndClass As Boolean, Flag As Boolean
Dim StartKeyPosition As Integer, EndKeyPosition As Integer

Num = FreeFile
Open File For Input As #Num
    Do Until EOF(Num)
        Line Input #Num, Str
        
        '如果起始含有类标签，则返回 True
        StartClass = InStr(1, Str, "[+" & ClassLabel & "-]")
        If StartClass Then Flag = True
        
        '如果末尾含有类标签，则返回 False
        EndClass = InStr(1, Str, "[-" & ClassLabel & "+]")
        If EndClass Then Flag = False
        
        '如果条件返回 True ，则提取类模块中的 键值标签
        If Not (StartClass Xor EndClass) And Flag Then
            Key = ""
            StartKeyPosition = InStr(1, Str, "<")
            EndKeyPosition = InStr(1, Str, ">")
            For i = StartKeyPosition + 1 To EndKeyPosition - 1
                Key = Key & Mid(Str, i, 1)
            Next
            KeyList = KeyList & Key & ","
        End If
        
    Loop
Close #Num
'返回值
KeyList = Left(KeyList, Len(KeyList) - 1)
End Function

'设置类中的键值
Sub SetKey(File As String, Class As String, Key As String, KeyValue As String)

Dim StrText As String, LeftText As String, RightText As String, NewText As String, Str As String
Dim StartPosition As Integer, EndPosition As Integer, KeyPosition As Integer, KeyValuePosition As Integer

'提取配置文件中的内容
Num = FreeFile
Open File For Input As #Num
    Do Until EOF(Num)
        Line Input #Num, Str
        StrText = StrText & Chr(10) & Str
    Loop
Close #Num

'查找类模块的 起始位 与 终止位
StartPosition = InStr(1, StrText, "[+" & Class & "-]")
EndPosition = InStr(1, StrText, "[-" & Class & "+]")

'设置 Key 的 标签、标签的长度、标签的位置
Key = "<" & Key & ">"
keylen = Len(Key)
KeyPosition = 1

'查找 类模块 中的 Key标签的位置
Do
    KeyPosition = InStr(KeyPosition, StrText, Key)
    KeyPosition = KeyPosition + keylen
    If StartPosition < KeyPosition And KeyPosition < EndPosition Then Exit Do
Loop

'提取 KeyValue 的 左边与右边 的文本
LeftText = Left(StrText, KeyPosition - 1)
KeyValuePosition = InStr(KeyPosition, StrText, Chr(10))
RightText = Mid(StrText, KeyValuePosition)

'获取 新键值 赋给 NewText 变量
NewText = LeftText & KeyValue & RightText
MsgBox NewText

'输出
NewTextLen = Len(NewText)
EndPosition = 1
Open File For Output As #Num
    Do
        StartPosition = InStr(EndPosition + 1, NewText, Chr(10))
        EndPosition = InStr(StartPosition + 1, NewText, Chr(10))
        TextLen = IIf(EndPosition = 0, NewTextLen - EndPosition + 1, EndPosition - StartPosition)
        Print #Num, Mid(NewText, StartPosition, TextLen)
        Debug.Print Mid(NewText, StartPosition, TextLen)
        If EndPosition >= NewTextLen Or EndPosition = 0 Then Exit Do
    Loop
Close #Num
End Sub

