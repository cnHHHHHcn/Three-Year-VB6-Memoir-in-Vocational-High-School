VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperateKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "用于编写操作配置文件\r\n注：\r\n\r\n属性：File 、 Class 、 Key\r\n若可选形参不写值，自动用属性值\r\n"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


'类模块 OperateKey 代码如下

'关键字    作用域说明

'Private   仅在本模块中使用，不向对外公布使用

'Public    可在本模块中使用，也可对外公布使用

'属性
Public File As String
Public Class As String
Public Key As String

'对内使用的方法
'*********************************************************************

'类模块方法中 形参值的优先级 要高于 属性值 _
 优先级 形参值 > 属性值
'---------------------------------------------------------------------
'获取属性File(文件)函数
Private Function GetPropertyFile(ByVal File As String) As String
    GetPropertyFile = IIf(File = "", Me.File, File)
End Function

'获取属性Class(模块名)函数
Private Function GetPropertyClass(ByVal Class As String) As String
    GetPropertyClass = IIf(Class = "", Me.Class, Class)
End Function

'获取属性Key(键名)函数
Private Function GetPropertyKey(ByVal Key As String) As String
    GetPropertyKey = IIf(Key = "", Me.Key, Key)
End Function
'---------------------------------------------------------------------

'输出
Private Sub PrintLine(ByVal File As String, ByVal Text As String)
    
    Dim Num As Integer
    Dim TextLine() As String
    
    '获取空闲文件号
    Num = FreeFile
    '把 Text变量 中 以 “|” 为分隔符，分割字符串。 _
     以数组的形式 赋入 TextLine变量
    TextLine = Split(Right(Text, Len(Text) - 1), "|")
    '以新建文件的方式 打开以 File变量 为名的文件
    Open File For Output As #Num
        '历遍 TextLine数组
        For i = LBound(TextLine) To UBound(TextLine)
            '输出到文件
            Print #1, TextLine(i)
        Next
    '关闭文件
    Close #Num
End Sub
'*********************************************************************


'对外公布的方法
'*********************************************************************

'获取键值
Public Function GetKey(Optional ByVal File As String, Optional ByVal Class As String, Optional ByVal Key As String) As String
Attribute GetKey.VB_Description = "Function函数\r\n用来获取文件中模块的键的键值\r\n形参 File 可选\r\n形参 Class 可选\r\n形参 Key 可选\r\n若可选形参不写自动用属性值"
    
    Dim Flag As Boolean, _
        StartClass As Boolean, EndClass As Boolean
    Dim Num As Integer, KeyPosition As Integer
    Dim Str As String, _
        StartLabel As String, EndLabel As String
    
    '若形参无值，则获取该类的属性值
    File = GetPropertyFile(File)
    Class = GetPropertyClass(Class)
    Key = "<" & GetPropertyKey(Key) & ">"
    
    '获取 起始、结束模块标签
    StartLabel = "[+" & Class & "-]"
    EndLabel = "[-" & Class & "+]"
    
    '获取空闲文件号
    Num = FreeFile
    '以 读取方式 打开以 File变量 为名称的文件
    Open File For Input As #Num
        '当 文件指针 指到 文件尾部 则结束循环
        Do Until EOF(Num)
            '读取每行的信息
            Line Input #Num, Str
            
            '如果含有起始模块标签，则返回 True
            StartClass = InStr(1, Str, StartLabel)
            If StartClass Then Flag = True
            
            '如果含有末尾模块标签，则返回 False
            EndClass = InStr(1, Str, EndLabel)
            If EndClass Then Flag = False
            
            '在模块中提取键值
            '--------------------------------------------------------------------------------------------
            '判定文件指针是否在该模块区域
            If Not (StartClass Xor EndClass) And Flag Then
                '查找键名
                KeyPosition = InStr(1, Str, Key)
                '如果找到 则输出键值，并退出循环
                If CBool(KeyPosition) Then GetKey = Mid(Str, KeyPosition + Len(Key)): Exit Do
            End If
            '--------------------------------------------------------------------------------------------
        Loop
    Close #Num
End Function

'显示模块名称
Public Function ClassList(Optional ByVal File As String) As String
Attribute ClassList.VB_Description = "Function函数\r\n用来显示文件中的模块\r\n形参 File 可选\r\n若可选形参不写自动用属性值"

    On Error Resume Next
    Dim Num As Integer, _
        StartPosition As Integer, endPosition As Integer
    Dim Str As String
    
    '若形参无值，则获取该类的属性值
    File = GetPropertyFile(File)
    
    '获取空闲文件号
    Num = FreeFile
    '以 读取方式 打开以 File变量 为名称的文件
    Open File For Input As #Num
        '当 文件指针 指到 文件尾部 则结束循环
        Do Until EOF(Num)
            '读取每行的信息
            Line Input #Num, Str
            
            '寻找模块标记，并返回位置
            StartPosition = InStr(Str, "[+")
            endPosition = InStr(Str, "-]")
            
            '如果 Start 与 Ends 变量的值不为 0 ，则提取模块 标签
            If CBool(StartPosition) And CBool(endPosition) Then
                '清空上一个模块名
                Class = ""
                '截取这个模块名
                For i = StartPosition + 2 To endPosition - 1
                    Class = Class & Mid(Str, i, 1)
                Next
                '添加
                ClassList = ClassList & Class & ","
            End If
        Loop
    Close #Num
    '返回值
    ClassList = Left(ClassList, Len(ClassList) - 1)
End Function

'显示模块中的键名称
Public Function KeyList(Optional ByVal File As String, Optional ByVal Class As String) As String
Attribute KeyList.VB_Description = "Function函数\r\n用来显示文件中模块的键\r\n形参 File 可选\r\n形参 Class 可选\r\n若可选形参不写自动用属性值"
    On Error Resume Next
    Dim Str As String, _
        StartLabel As String, EndLabel As String
        
    Dim Flag As Boolean, _
        StartClass As Boolean, EndClass As Boolean
    Dim Num As Integer, _
        StartKeyPosition As Integer, EndKeyPosition As Integer
    
    '若形参无值，则获取该类的属性值
    File = GetPropertyFile(File)
    Class = GetPropertyClass(Class)
    
    StartLabel = "[+" & Class & "-]"
    EndLabel = "[-" & Class & "+]"
    
    '获取空闲文件号
    Num = FreeFile
    '以 读取方式 打开以 File变量 为名称的文件
    Open File For Input As #Num
        '当 文件指针 指到 文件尾部 则结束循环
        Do Until EOF(Num)
            '读取每行的信息
            Line Input #Num, Str
            
            '如果含有起始模块标签，则返回 True
            StartClass = InStr(1, Str, StartLabel)
            If StartClass Then Flag = True
            
            '如果含有末尾模块标签，则返回 False
            EndClass = InStr(1, Str, EndLabel)
            If EndClass Then Flag = False
            
            '提取模块中的 键值标签
            '------------------------------------------------------------
            '判定文件指针是否在该模块区域
            If Not (StartClass Xor EndClass) And Flag Then
                '清空上一个键名
                Key = ""
                '获取这个键名的 起始与结束 位置
                StartKeyPosition = InStr(1, Str, "<")
                EndKeyPosition = InStr(1, Str, ">")
                '截取键名
                For i = StartKeyPosition + 1 To EndKeyPosition - 1
                    Key = Key & Mid(Str, i, 1)
                Next
                '添加
                KeyList = KeyList & Key & ","
            End If
            '-------------------------------------------------------------
        Loop
    Close #Num
    '返回值
    KeyList = Left(KeyList, Len(KeyList) - 1)
End Function

'设置模块中的键值
Public Sub SetKey(ByVal KeyValue As String, Optional ByVal File As String, Optional ByVal Class As String, Optional ByVal Key As String)
Attribute SetKey.VB_Description = "Sub子程序\r\n用来更改文件中模块的键的键值\r\n形参 KeyValue 必选\r\n形参 File 可选\r\n形参 Class 可选\r\n形参 Key 可选\r\n若可选形参不写自动用属性值"
    
    Dim StrText As String, Str As String, _
        LeftText As String, RightText As String, NewText As String
    Dim Num As Integer, _
        StartPosition As Integer, endPosition As Integer, _
        KeyPosition As Integer, KeyValuePosition As Integer
    
    '若形参无值，则获取该类的属性值
    File = GetPropertyFile(File)
    Class = GetPropertyClass(Class)
    Key = GetPropertyKey(Key)
    
    '提取配置文件中的内容
    '-----------------------------------------------------
    '获取空闲文件号
    Num = FreeFile
    '以 读取方式 打开以 File变量 为名称的文件
    Open File For Input As #Num
        '当 文件指针 指到 文件尾部 则结束循环
        Do Until EOF(Num)
            '读取每行的信息
            Line Input #Num, Str
            StrText = StrText & "|" & Str
        Loop
    Close #Num
    '-----------------------------------------------------
    
    '查找模块的 起始位 与 终止位
    StartPosition = InStr(1, StrText, "[+" & Class & "-]")
    endPosition = InStr(1, StrText, "[-" & Class & "+]")
    
    '设置 Key 的 标签、标签的长度、标签的位置
    Key = "<" & Key & ">"
    keylen = Len(Key)
    KeyPosition = 1
    
    '查找 模块 中的 Key标签的位置
    Do
        KeyPosition = InStr(KeyPosition, StrText, Key)
        KeyPosition = KeyPosition + keylen
        If StartPosition < KeyPosition And KeyPosition < endPosition Then Exit Do
    Loop
    
    '提取 KeyValue 的 左边与右边 的文本
    LeftText = Left(StrText, KeyPosition - 1)
    KeyValuePosition = InStr(KeyPosition, StrText, "|")
    RightText = Mid(StrText, KeyValuePosition)
    
    '获取 新键值 赋给 NewText 变量
    NewText = LeftText & KeyValue & RightText
    '调用 PrintLine 子程序
    Call PrintLine(File, NewText)
End Sub

'在文件中创建 模块
Public Sub CreateClass(ClassName As String, Optional ByVal File As String)
Attribute CreateClass.VB_Description = "Sub子程序\r\n用来在文件中创建模块\r\n形参 ClassName 必选\r\n形参 File 可选\r\n若可选形参不写自动用属性值"
    
    Dim Num As Integer
    
    '若形参无值，则获取该类的属性值
    File = GetPropertyFile(File)
    
    
    '获取空闲文件号
    Num = FreeFile
    '以 追加方式 打开以 File变量 为名称的文件
    Open File For Append As #Num
        '创建模块
        Print #Num,
        Print #Num, "[+" & ClassName & "-]"
        Print #Num, "[-" & ClassName & "+]"
    Close #Num
End Sub

'在文件的模块中创建 键名 与 键值
Public Sub CreateKey(ByVal File As String, ByVal Class As String, ParamArray Key_NameValue())
Attribute CreateKey.VB_Description = "Sub子程序\r\n用来在文件的模块中创建 键与键值\r\n形参 File 可选\r\n形参 Class 可选\r\n形参 Key_NameValue 必选\r\n若可选形参不写自动用属性值\r\n\r\n形参 Key_NameValue 写法\r\n形参 Key_NameValue 是一个数组，可以写多个参数。\r\n语法：键名;键值\r\n注：\r\n    1.键名 与 键值 中间用英文分号分隔。\r\n    2.类型必须是字符串。"
    
    Dim Str As String, _
            StartLabel As String, EndLable As String, _
            Text As String, LText As String, MText As String, RText As String, _
            KeyName As String, KeyValue As String, NewKey As String
        
    Dim Num As Integer, StartPosition As Integer, endPosition As Integer
    Dim Flag As Boolean, StartClass As Boolean, EndClass As Boolean
    
    '若形参无值，则获取该模块的属性值
    File = GetPropertyFile(File)
    Class = GetPropertyClass(Class)
    
    '提取文件中的内容
    '-----------------------------------------------
    '获取空闲文件号
    Num = FreeFile
    '以 读取方式 打开以 File变量 为名称的文件
    Open File For Input As #Num
        '当 文件指针 指到 文件尾部 则结束循环
        Do Until EOF(Num)
            '读取每行的信息
            Line Input #Num, Str
            Text = Text & "|" & Str
        Loop
    Close #Num
    '-----------------------------------------------
    
    '把 起始 与 结束 标签 赋入变量 StartLabel 与 EndLabel
    StartLabel = "[+" & Class & "-]"
    EndLabel = "[-" & Class & "+]"
    '把起始 与 结束 标签的位置 赋入变量 StartPosition 与 EndPosition
    StartPosition = InStr(1, Text, StartLabel) + Len(StartLabel)
    endPosition = InStr(1, Text, EndLabel)
    '截取模块标签标签 以上 及 以下 的内容
    LText = Left(Text, StartPosition)
    Debug.Print Text
    RText = Mid(Text, endPosition)
    '截取模块中的 键名 及 键值
    MText = Mid(Text, StartPosition + 1, endPosition - StartPosition)
    
    '赋入 键名 及 键值
    For i = LBound(Key_NameValue) To UBound(Key_NameValue)
        '获取键名
        KeyName = Left(Key_NameValue(i), InStr(1, Key_NameValue(i), ";") - 1)
        '获取键值
        KeyValue = Mid(Key_NameValue(i), InStr(1, Key_NameValue(i), ";") + 1)
        '编写格式 并赋入 NewKey 变量中
        NewKey = "    <" & KeyName & ">" & KeyValue
        '添加
        MText = MText & NewKey & "|"
    Next
    
    MText = Mid(MText, 2)
    
    '连接
    Text = LText & MText & RText
    
    '调用 PrintLine 子程序
    Call PrintLine(File, Text)
    
End Sub

Public Sub DeleteKey(Optional File As String, Optional Class As String, Optional Key As String)
    Dim Str As String, _
        StartCLabel As String, EndLabel As String
    Dim Flag As Boolean, _
        StartClass As Boolean, EndClass As Boolean
    
    '若形参无值，则获取该类的属性值
    File = GetPropertyFile(File)
    Class = GetPropertyClass(Class)
    Key = GetPropertyKey(Key)
    
    StartLabel = "[+" & Class & "-]"
    EndLabel = "[-" & Class & "+]"
    
    '获取空闲文件号
    Num = FreeFile
    '以 读取方式 打开以 File变量 为名称的文件
    Open File For Input As #Num
        '当 文件指针 指到 文件尾部 则结束循环
        Do Until EOF(Num)
            '读取每行的信息
            Line Input #Num, Str
            '如果含有起始模块标签，则返回 True
            StartClass = InStr(1, Str, StartLabel)
            If StartClass Then Flag = True
            '如果含有末尾模块标签，则返回 False
            EndClass = InStr(1, Str, EndLabel)
            If EndClass Then Flag = False
            '判定文件指针是否在该模块区域
            If Not (StartClass Xor EndClass) And Flag Then
                '如果 InStr(1, Str, Key) 返回值为 0(即 False)再取反 为 True 则把 Str变量 追加入 Text变量
                If Not CBool(InStr(1, Str, Key)) Then Text = Text & "|" & Str
            Else
                '否则把 Str变量 追加入 Text变量
                Text = Text & "|" & Str
            End If
        Loop
    Close #Num
    '调用 PrintLine 子程序
    Call PrintLine(File, Text)
End Sub

Public Sub DeleteClass(Optional File As String, Optional Class As String)
    Dim Str As String, _
        StartLabel As String, EndLabel As String
    Dim Flag As Boolean, _
        StartClass As Boolean, EndClass As Boolean
    
    '若形参无值，则获取该类的属性值
    File = GetPropertyFile(File)
    Class = GetPropertyClass(Class)
    Key = GetPropertyKey(Key)
    
    '获取 起始、结束 标签
    StartLabel = "[+" & Class & "-]"
    EndLabel = "[-" & Class & "+]"
    
    '获取空闲文件号
    Num = FreeFile
    '以 读取方式 打开以 File变量 为名称的文件
    Open File For Input As #Num
        '当 文件指针 指到 文件尾部 则结束循环
        Do Until EOF(Num)
            '读取每行的信息
            Line Input #Num, Str
            '如果含有起始模块标签，则返回 True
            StartClass = InStr(1, Str, StartLabel)
            If StartClass Then Flag = True: GoTo RemoveClassLabel
            '如果含有末尾模块标签，则返回 False
            EndClass = InStr(1, Str, EndLabel)
            If EndClass Then Flag = False: GoTo RemoveClassLabel
            '如果文件指针不在该模块区域，则把 Str变量 追加入 Text变量
            If Not (Not (StartClass Xor EndClass) And Flag) Then Text = Text & "|" & Str
RemoveClassLabel:
        Loop
    Close #Num
    '调用 PrintLine 子程序
    Call PrintLine(File, Text)
End Sub
'*********************************************************************
