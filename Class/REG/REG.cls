VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

'-注册表 API 声明...
'RegOpenKeyEx(根键句柄,子键路径,0,访问级别,返回子键句柄)
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'RegCloseKey(子键句柄)
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'RegCreateKeyEx(根键句柄,子键路径,0,子键类型,0,访问级别,没用,返回子键句柄,返回子键状态)
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
'RegQueryValueEx(子键句柄,键值名称,0,返回类型,返回键值数据,键值数据长度)
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'RegSetValueEx（子键句柄,子键名称,0,类型,数据,数据长度）
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'RegDeleteKey(子键句柄,子键名称)
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'RegDeleteValue(子键句柄,键值名称)
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'RegEnumKeyEx（子键句柄,索引值,子键名称,子键名称长度,0,子键类型,子键类型,没用)
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FileTime) As Long
'RegEnumValue(子键句柄,索引值,子键名称,子键名称长度,0,键值类型,键值数据,键值数据长度)
Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'以下这俩没用过
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long

' 注册表创建类型值...
Private Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留
Private Const Name_Buffer_Size = 255
Private Const Data_Buffer_Size = 1024

' 注册表关键字根类型
Private Enum REG_HKEY
    Classes_Root = &H80000000
    Current_User = &H80000001
    Local_Machine = &H80000002
    Users = &H80000003
    Performance_Data = &H80000004
    Current_Config = &H80000005
End Enum

' 注册表关键字安全选项
Public Enum REG_KeyOpera
    Synchronize = &H100000
    Standard_Rights_All = &H1F0000
    READ_CONTROL = &H20000
    Query_Value = &H1
    Set_Value = &H2
    Create_Sub_Key = &H4
    Enumerate_Sub_Keys = &H8
    Notify = &H10
    Create_Link = &H20
    KeyRead = Query_Value + Enumerate_Sub_Keys + Notify + READ_CONTROL
    KeyWrite = Set_Value + Create_Sub_Key + READ_CONTROL
    Standard_Rights_Read = (READ_CONTROL)
    Read = ((Standard_Rights_Read Or Query_Value Or Enumerate_Sub_Keys Or Notify) And (Not Synchronize))
    KeyEexcute = Read
    All_Access = ((Standard_Rights_All Or Query_Value Or Set_Value Or Create_Sub_Key Or Enumerate_Sub_Keys Or Notify Or Create_Link) And (Not Synchronize))
End Enum

' Reg Data Types...
Public Enum REG_KeyType
    NONE = 0                       ' 未定义类型
    SZ = 1                         ' Unicode空终结字符串
    EXPAND_SZ = 2                  ' Unicode空终结字符串
    DWORD = 4                      ' 32-bit 数字
End Enum

' 返回值...
Public Enum REG_ReturnValue
    ERROR_NONE = 0
    ERROR_BADKEY = 2
    ERROR_ACCESS_DENIED = 8
    ERROR_SUCCESS = 0
    ERROR_NO_MORE_ITEMS = 259&
End Enum

Public Enum REG_SubKeyState
    Created_New_Key = &H1                      ' New Registry Key created
    Opened_Existing_Key = &H2                      ' Existing Key opened
End Enum

Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


' 注册表安全属性类型
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'子键信息类型
Private Type SubKeyInfo
    Handle As Long
    State As REG_SubKeyState
End Type

'子键路径信息类型
Private Type REGPathInfo
    RootKeyName As String
    RootKeyHandle As Long
    SubKeyPath As String
End Type


Private Local_REGKeyPath As String

'
Private SubKey As SubKeyInfo, REGInfo As REGPathInfo
Private RootKeyHandle As Long
Private ErrorInfo As String
Private Result

'注册表路径(可读写属性)
Public Property Let REGKeyPath(ByVal New_REGKeyPath As String)
    Local_REGKeyPath = New_REGKeyPath
End Property

Public Property Get REGKeyPath() As String
    If Local_REGKeyPath = "" Then
        ErrorInfo = "注册表路径未设置！"
        Exit Property
    Else
        REGKeyPath = Local_REGKeyPath
    End If
End Property

'返回报错信息(只读属性)
Public Property Get Error() As String
    Error = ErrorInfo
End Property

'获取 注册表错误信息
Private Function GetError() As String
    Dim ErrorCode As Long
    ErrorInfo = Space(256)
    ErrorCode = GetLastError()
    ErrorLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, ErrorCode, 0, ErrorInfo, 256, 0)
    If ErrorLength > 0 Then ErrorInfo = Trim(ErrorInfo): GetError = ErrorInfo
End Function

'返回子键路径信息(只读属性)
Public Property Get KeyPathInformation() As String
    Dim Info As String
    KeyPathInformation = "根键名：" & REGInfo.RootKeyName & vbCrLf & _
                         "根键句柄：" & REGInfo.RootKeyHandle & vbCrLf & _
                         "子键路径：" & REGInfo.SubKeyPath & vbCrLf & _
                         "子键句柄：" & SubKey.Handle & vbCrLf & _
                         "子键状态：" & IIf(SubKey.State = Created_New_Key, "已创建新子键", "已打开子键")
End Property

'返回 文件名 用于 WinStartRun方法
Private Function GetFileName(ByVal Path_String As String) As String
    Do While Path_String <> filestr
        filestr = Path_String
        Path_String = Right(Path_String, Len(Path_String) - InStr(Path_String, "\"))
    Loop
    If InStr(1, Path_String, ".") <> 0 Then GetFileName = Path_String Else GetFileName = "None"
End Function

'返回根键句柄
Private Function GetRootKeyHandle(KeyPath As String) As REG_HKEY
    Dim RootKeyName As String
    If CBool(InStr(1, KeyPath, "计算机\")) Then KeyPath = Mid(KeyPath, 5)
    '获取根键名称
    REGInfo.RootKeyName = UCase(Left(KeyPath, InStr(1, KeyPath, "\") - 1))
    '返回除根键的路径
    KeyPath = Mid(KeyPath, InStr(1, KeyPath, "\") + 1): REGInfo.SubKeyPath = KeyPath
    '获取根键句柄
    Select Case REGInfo.RootKeyName
        Case "HKEY_CLASSES_ROOT"
            GetRootKeyHandle = Classes_Root
        Case "HKEY_CURRENT_USER"
            GetRootKeyHandle = Current_User
        Case "HKEY_LOCAL_MACHINE"
            GetRootKeyHandle = Local_Machine
        Case "HKEY_USERS"
            GetRootKeyHandle = Users
        Case "HKEY_CURRENT_CONFIG"
            GetRootKeyHandle = Current_Config
        Case "PERFORMANCE_DATA"
            GetRootKeyHandle = Performance_Data
    End Select
    REGInfo.RootKeyHandle = GetRootKeyHandle
End Function

'关闭子键句柄
Private Function CloseSubKey(SubKeyHandle As Long) As Boolean
    '关闭子键
    Result = RegCloseKey(SubKeyHandle)
    '查看是否关闭
    CloseSubKey = IIf(Result = 0, True, False)
End Function

'获取子键信息(句柄，路径，状态)
Private Function GetSubKeyHandle(RootKey As Long, Path As String, Result As SubKeyInfo) As Boolean
    Dim REGSecurity As SECURITY_ATTRIBUTES
    Dim REGReturn
    Dim KeyHandle As Long, KeyState As REG_SubKeyState
    
    ' 设置安全属性为缺省值...
    REGSecurity.nLength = 50
    REGSecurity.lpSecurityDescriptor = 0
    REGSecurity.bInheritHandle = False
    '创建 或 打开子键，如果成功 返回 0 ，否则 返回 1
    REGReturn = RegCreateKeyEx(RootKey, Path, 0, "", 0, All_Access, REGSecurity, KeyHandle, KeyState)
    '如果变量 REGReturn = 0 则 返回 子键句柄 和 状态 并将返回值设为 True 否则 为 False
    If REGReturn = 0 Then
        Result.Handle = KeyHandle
        Result.State = KeyState
        GetSubKeyHandle = True
    Else
        GetSubKeyHandle = False
    End If
End Function

'添加子键
Public Function ModifySubKey(KeyName As String, Optional ByVal KeyPath As String) As Boolean
    '判定 注册表路径是否 合法 ，不合法 则退出 此方法
    KeyPath = IIf(KeyPath <> "", KeyPath, Me.REGKeyPath)
    If KeyPath = "" Then Exit Function
    '获取根键句柄
    RootKeyHandle = GetRootKeyHandle(KeyPath)
    '使子键路径规范化
    KeyPath = IIf(Right(KeyPath, 1) = "\", KeyPath, KeyPath & "\")
    '添加子键 至 子键路径
    KeyPath = KeyPath & KeyName
    '获取子键 句柄 和 状态
    Result = GetSubKeyHandle(RootKeyHandle, KeyPath, SubKey)
    GetError
    '如果调用成功 则 关闭
    If Result Then CloseSubKey (SubKey.Handle)
    '返回结果
    ModifySubKey = IIf(Result = 0, True, False)
End Function

'删除子键
Public Function DeleteSubKey(KeyName As String, Optional ByVal KeyPath As String) As Boolean
    '判定 注册表路径是否 合法 ，不合法 则退出 此方法
    KeyPath = IIf(KeyPath <> "", KeyPath, Me.REGKeyPath)
    If KeyPath = "" Then Exit Function
    '获取根键句柄
    RootKeyHandle = GetRootKeyHandle(KeyPath)
    '获取子键 句柄 和 状态
    Result = GetSubKeyHandle(RootKeyHandle, KeyPath, SubKey)
    '如果 返回为 True 则执行下列代码 否则返回 False
    If Result Then
        '删除子键
        Result = RegDeleteKey(SubKey.Handle, IIf(KeyName <> "", KeyName, ""))
        GetError
        '调用成功 关闭子键
        If Result = 0 Then
            CloseSubKey (SubKey.Handle)
            DeleteSubKey = True
            Exit Function
        End If
    End If
    DeleteSubKey = False
End Function

Public Function EnumSubKey(ReturnSubKey(), Optional ByVal KeyPath As String) As Boolean
    Dim SubKeyName As String, SubKeySize As Long, _
        Count As Integer
    Dim ChangeFileTime As FileTime
    
    '报错执行 NONE 标签后的语句
    On Error GoTo NONE
    
    '判定 注册表路径是否 合法 ，不合法 则退出 此方法
    KeyPath = IIf(KeyPath <> "", KeyPath, Me.REGKeyPath)
    If KeyPath = "" Then Exit Function
    '获取根键句柄
    RootKeyHandle = GetRootKeyHandle(KeyPath)
    '获取子键 句柄 和 状态
    Result = GetSubKeyHandle(RootKeyHandle, KeyPath, SubKey)
    If Result Then
        '设置 SubKeyName 的缓存(256个空格)
        SubKeyName = Space(Name_Buffer_Size)
        '设置 SubKeySize 的长度(256)
        SubKeySize = Name_Buffer_Size
        '清除 旧数组的所有元素
        Erase ReturnSubKey
        '开始枚举 当前注册表子键 下的 子键
        Do While RegEnumKeyEx(SubKey.Handle, Count, SubKeyName, SubKeySize, 0, "", 0, ChangeFileTime) <> ERROR_NO_MORE_ITEMS
            GetError
            '数组 ReturnSubKey 元素个数 + 1 ，并保存元素中的值
            ReDim Preserve ReturnSubKey(Count)
            '存入 子键名称
            ReturnSubKey(Count) = Trim(SubKeyName)
            '索引值 + 1
            Count = Count + 1
            '设置 SubKeyName 的缓存(256个空格)
            SubKeyName = Space(Name_Buffer_Size)
        Loop
        '检测 是否报错，报错为 False 即没有子键(变量 p 没用)
        p = LBound(ReturnSubKey)
        '返回 True ，并退出 此方法
        EnumSubKey = True
        Exit Function
    End If
NONE:
    EnumSubKey = False
End Function

Public Function ModifyKeyValue(ValueName As String, ValueData As Variant, Optional ByVal KeyPath As String) As Boolean
    Dim ValueType As REG_KeyType
    Dim ValueLen As Long
    
    '报错继续运行
    On Error Resume Next
    '判定 注册表路径是否 合法 ，不合法 则退出 此方法
    KeyPath = IIf(KeyPath <> "", KeyPath, Me.REGKeyPath)
    If KeyPath = "" Then Exit Function
    '获取根键句柄
    RootKeyHandle = GetRootKeyHandle(KeyPath)
    '获取子键 句柄 和 状态
    Result = GetSubKeyHandle(RootKeyHandle, KeyPath, SubKey)
    '如果 返回为 True 则执行下列代码 否则返回 False
    If Result Then
        '查看 输入数据 为 什么类型(Visual Basic to Register)， _
        根据对应类型进行数据转化 赋入 变量ValueData
        Select Case VarType(ValueData)
            '字符串
            Case vbString
                '如果字符串中含有 “%” 则为 可扩充字符串(注册表类型) 否则为 字符串(注册表类型)
                ValueType = IIf(CBool(InStr(1, ValueData, "%")), EXPAND_SZ, SZ)
                '获取 数据的字节长度
                ValueLen = LenB(ValueData)
            '长整型，整型，十进制，字节，布尔
            Case vbLong, vbInteger, vbDecimald, vbByte, vbBoolean
                '设置 类型为 32位整形数值(注册表类型)
                ValueType = DWORD
                '让变量 ValueData 为字符串类型，让变量 Temp 变为正数(负数不支持)
                ValueData = "": Temp = Abs(Temp)
                '当 Temp 的值为 0 时停止循环
                Do While Temp <> 0
                    '利用 AND运算 获取 低位 8 bit数据(从右向左)， _
                     并用 Chr函数 转成 字符，存入 变量ValueData
                    ValueData = ValueData & Chr(Temp And 255)
                    '让 Temp 除以 256 并去掉 小数点部分
                    Temp = Fix(Temp / 256)
                Loop
                '数据的字节长度为 4
                ValueLen = 4
            '其他类型，返回 False 退出 此方法
            Case Else
                ModifyKeyValue = False
                '关闭子键句柄
                CloseSubKey (SubKey.Handle)
                ErrorInfo = "注册表不支持此类型数据！"
                Exit Function
        End Select
        '判断 子键状态 是否为 创建新子键 或者 打开已有子键
        If SubKey.State = Created_New_Key Or SubKey.State = Opened_Existing_Key Then
            '设置 值名称 和 值数据
            Result = RegSetValueEx(SubKey.Handle, ValueName, 0, ValueType, ValueData, ValueLen)
            GetError
            '调用成功 返回为 True ，否则为FALSE
            ModifyKeyValue = IIf(Result = 0, True, False)
            '关闭子键句柄
            CloseSubKey (SubKey.Handle)
        End If
    End If
End Function

Public Function DeleteKeyValue(ValueName As String, Optional ByVal KeyPath As String) As Boolean
    '判定 注册表路径是否 合法 ，不合法 则退出 此方法
    KeyPath = IIf(KeyPath <> "", KeyPath, Me.REGKeyPath)
    If KeyPath = "" Then Exit Function
    '获取根键句柄
    RootKeyHandle = GetRootKeyHandle(KeyPath)
    '获取子键 句柄 和 状态
    Result = GetSubKeyHandle(RootKeyHandle, KeyPath, SubKey)
    '如果 返回为 True 则执行下列代码 否则返回 False
    If Result Then
        '删除 值名称 和 值数据
        Result = RegDeleteValue(SubKey.Handle, ValueName)
        GetError
        '调用成功 返回为 True ，否则为FALSE
        DeleteKeyValue = IIf(Result = 0, True, False)
        '关闭子键句柄
        CloseSubKey (SubKey.Handle)
    End If
End Function

Public Function QueryKeyValue(ValueName As String, ValueData As Variant, Optional ByVal KeyPath As String) As Boolean
    Dim ValueTempData As String, ValueDataType As Long, ValueDataSize As Long
    '判定 注册表路径是否 合法 ，不合法 则退出 此方法
    KeyPath = IIf(KeyPath <> "", KeyPath, Me.REGKeyPath)
    If KeyPath = "" Then Exit Function
    '获取根键句柄
    RootKeyHandle = GetRootKeyHandle(KeyPath)
    '获取子键 句柄 和 状态
    Result = GetSubKeyHandle(RootKeyHandle, KeyPath, SubKey)
    '如果 返回为 True 则执行下列代码 否则返回 False
    If Result Then
        '设置 ValueDataSize 的长度(1024)，ValueTempData 的缓存(1024个空格)
        ValueDataSize = Data_Buffer_Size: ValueTempData = Space(ValueDataSize)
        '查询 值数据
        Result = RegQueryValueEx(SubKey.Handle, ValueName, 0, ValueDataType, ValueTempData, ValueDataSize)
        GetError
        '调用成功 执行下列数据转化代码，并返回为True，否则 返回为False
        If Result = 0 Then
            '去掉多余空格
            ValueTempData = Trim(ValueTempData)
            '查看 数据 为 什么类型(Register to Visual Basic)， _
             根据对应类型进行数据转化 赋入 变量ValueData
            Select Case ValueDataType
                '32位整形数值
                Case DWORD
                    '读者自行理解，理解不了也没关系(反正能用)
                    For i = Len(ValueTempData) To 1 Step -1
                        ValueData = ValueData + Hex(Asc(Mid(ValueTempData, i, 1)))
                    Next
                    ValueData = Format("&H" + ValueData)
                '可扩充字符串,字符串
                Case SZ, EXPAND_SZ
                    '原样返回
                    ValueData = ValueTempData
                '其他类型
                Case Else
                    ValueData = ValueTempData
            End Select
            QueryKeyValue = True
        Else
            QueryKeyValue = False
        End If
    End If
End Function

Public Function EnumKeyValue(ReturnKeyValue(), Optional ByVal KeyPath As String) As Boolean
    Dim ValueName As String, ValueNameSize As Long, _
        ValueData As String, ValueTempData As String, _
        ValueDataType As Long, ValueDataSize As Long, _
        Count As Integer
    
    '报错执行 NONE 标签后的语句
    On Error GoTo NONE
    
    '判定 注册表路径是否 合法 ，不合法 则退出 此方法
    KeyPath = IIf(KeyPath <> "", KeyPath, Me.REGKeyPath)
    If KeyPath = "" Then Exit Function
    '获取根键句柄
    RootKeyHandle = GetRootKeyHandle(KeyPath)
    '获取子键 句柄 和 状态
    Result = GetSubKeyHandle(RootKeyHandle, Local_REGKeyPath, SubKey)
    '如果 返回为 True 则执行下列代码 否则返回 False
    If Result Then
        'ValueName 的缓存(256个空格)，设置 ValueName 的长度(256)
        ValueName = Space(Name_Buffer_Size): ValueNameSize = Name_Buffer_Size
        'ValueTempData 的缓存(1024个空格)，设置 ValueDataSize 的长度(1024)
        ValueTempData = Space(Data_Buffer_Size): ValueDataSize = Data_Buffer_Size
        '清除 旧数组的所有元素
        Erase ReturnKeyValue
        '开始枚举 当前注册表子键 下的 值名称 和 值数据
        Do While RegEnumValue(SubKey.Handle, Count, ValueName, ValueNameSize, 0, ValueDataType, ValueTempData, ValueDataSize) <> ERROR_NO_MORE_ITEMS
            GetError
            '数组 ReturnKeyValue 元素个数 + 2 ，并保存元素中的值
            ReDim Preserve ReturnKeyValue(1, Count)
            '去掉多余空格，存入 值名称
            ReturnKeyValue(0, Count) = Trim(ValueName)
            '让变量ValueTempData 去掉多余空格
            ValueTempData = Trim(ValueTempData)
            '同上，见 QueryKeyValue 方法
            Select Case ValueDataType
                Case DWORD
                    For i = Len(ValueTempData) To 1 Step -1
                        ValueData = ValueData + Hex(Asc(Mid(ValueTempData, i, 1)))
                    Next
                Case SZ, EXPAND_SZ
                    ValueData = ValueTempData
                Case Else
                    ValueData = ValueTempData
            End Select
            '存入 值数据
            ReturnKeyValue(1, Count) = ValueData
            '设置 ValueName 的缓存(256个空格)
            ValueName = Space(Name_Buffer_Size)
            '设置 ValueTempName 的缓存(1024个空格)
            ValueTempData = Space(Data_Buffer_Size)
            '索引值 + 1
            Count = Count + 1
        Loop
        '检测 是否报错，报错为 False 即没有子键(变量 p 没用)
        p = LBound(ReturnSubKey)
        '返回 True ，并退出 此方法
        EnumKeyValue = TrueKeyValue
        Exit Function
    End If
NONE:
    EnumKeyValue = False
End Function

Public Function WinStartRun(ProgramPath As String, Is_Current As Boolean) As Boolean
    Const StartRunPath As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    Dim KeyPath As String, ProgramName As String
    KeyPath = IIf(Is_Current, "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE") & "\" & StartRunPath
    Name = GetFileName(ProgramPath)
    ProgramName = Left(Name, InStr(1, Name, ".") - 1)
    WinStartRun = Me.ModifyKeyValue(KeyPath, ProgramName, ProgramPath)
End Function
