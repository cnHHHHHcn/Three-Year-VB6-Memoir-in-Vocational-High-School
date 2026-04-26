VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type ConfigInfo
    Host As String
    IP As String
    Port As Long
End Type

Public ControlReleasePackage() As Byte
Public ChatPackage() As Byte
Public VirtualOnline() As Byte
Public TeacherInfo As ConfigInfo
Public Sock As Winsock

Public Sub ReadPackage(File As String, Bytes() As Byte)
    On Error Resume Next
    Dim Temp As String, TempArray() As String
    Erase Bytes: freenum = FreeFile
    Open File For Input As #freenum
        Do Until EOF(freenum)
            Line Input #1, Temp
            TempArray = Split(Temp, ":")
            ReDim Bytes(UBound(TempArray) - LBound(TempArray) + 1)
            For i = LBound(TempArray) To UBound(TempArray)
                Bytes(i + 1) = CLng("&H" & TempArray(i))
            Next
        Loop
    Close #freenum
End Sub

Public Function ListenTeacher() As Boolean
    Sock.Close
    Sock.Protocol = sckUDPProtocol
    Sock.Bind TeacherInfo.Port, TeacherInfo.IP
End Function
