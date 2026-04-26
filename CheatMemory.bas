Attribute VB_Name = "CheatMemory"
Option Explicit

' API 21
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, _
    ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, _
    ByRef lpNumberOfBytesRead As Long) As Long

Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, _
    ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, _
    ByRef lpNumberOfBytesWritten As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


' 常量定义
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const PROCESS_VM_OPERATION = &H8



Public Function APIread(ByVal PID As Long, rValue, ByVal BaseAddress As Long, Offset() As Long) As Variant
    Dim Buffer() As Byte
    Dim hProcess As Long, bytesRead As Long, Address As Long
    Dim AddrTmp As String
    hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_VM_OPERATION, 0, PID)
    If hProcess = 0 Then Exit Function
    ReDim Buffer(Len(rValue) - 1)
    Address = BaseAddress
    For i = LBound(Offest) To UBound(Offest) - 1
        ReadProcessMemory hProcess, Address + Offset(i), Buffer(0), 4, bytesRead
        For i = LBound(Buffer) To UBound(Buffer)
            AddrTmp = Hex(Buffer(i)) & AddrTmp
        Next
        Address = CLng("&H" & AddrTmp)
    Next
    'ReadProcessMemory hProcess, Address + Offset(i), Buffer(0), Len(rValue), bytesRead
    CopyMemory VarPtr(rValue), Address + Offset(UBound(Offset)), Len(rValue)
    
End Function

Public Function APIwrite(ByVal PID As Long, wValue, ByVal BaseAddress As Long, Offset() As Long)

End Function

Private Function ReadAddressPoint() As Long

End Function

' 读取内存函数
Private Function ReadMemory(ByVal ProcessID As Long, ByVal Address As Long, ByVal Size As Long) As Byte()
    Dim hProcess As Long
    Dim Buffer() As Byte
    Dim bytesRead As Long
    ' 打开进程
    hProcess = OpenProcess(PROCESS_VM_READ Or PROCESS_VM_OPERATION, 0, ProcessID)
    If hProcess = 0 Then Exit Function
    ReDim Buffer(Size - 1)
    ' 读取内存
    If ReadProcessMemory(hProcess, Address, Buffer(0), Size, bytesRead) Then ReadMemory = Buffer
    ' 关闭句柄
    CloseHandle hProcess
End Function

' 写入内存函数
Private Function WriteMemory(ByVal ProcessID As Long, ByVal Address As Long, ByRef Data() As Byte) As Boolean
    Dim hProcess As Long
    Dim bytesWritten As Long
    ' 打开进程
    hProcess = OpenProcess(PROCESS_VM_WRITE Or PROCESS_VM_OPERATION, 0, ProcessID)
    If hProcess = 0 Then Exit Function
    ' 写入内存
    WriteMemory = WriteProcessMemory(hProcess, Address, Data(0), UBound(Data) + 1, bytesWritten)
    ' 关闭句柄
    CloseHandle hProcess
End Function
