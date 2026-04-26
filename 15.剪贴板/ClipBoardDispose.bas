Attribute VB_Name = "ClipBoardDispose"
Public Config As Config
Public OperaDispose As New OperateKey

Public Type HotKey
    Control As String
    Word As String
End Type

Public Type Config
    ListMax As Byte
    OutputInterval As Single
    HotKey As HotKey
    REG As Boolean
End Type

Public Sub SaveDispose()
    With OperaDispose
        .File = App.Path & "\config.Dispose"
        .Class = "Dispose"
        .SetKey Config.ListMax, , , "ListMax"
        .SetKey Config.OutputInterval, , , "OutputInterval"
        .SetKey Config.HotKey.Control & "+" & Config.HotKey.Word, , , "HotKey"
    End With
End Sub

Public Sub LoadDispose(Result As Config)
    Dim Dispose As New OperateKey, ReturnResult As Variant
    Dim Keys() As String
    With Dispose
    .File = App.Path & "\Config.Dispose"
    .Class = "Dispose"
    Keys = Split(.KeyList, ",")
    For i = LBound(Keys) To UBound(Keys)
        .Key = Keys(i)
        ReturnResult = .GetKey
        Select Case Keys(i)
            Case "ListMax"
                Result.ListMax = ReturnResult
            Case "OutputInterval"
                Result.OutputInterval = ReturnResult
            Case "HotKey"
                HotKeyString = ReturnResult
                Result.HotKey.Control = Left(HotKeyString, InStr(1, HotKeyString, "+") - 1)
                Result.HotKey.Word = Mid(HotKeyString, InStr(1, HotKeyString, "+") + 1)
            Case "REG"
                Result.REG = IIf(ReturnResult = "True", True, False)
        End Select
    Next
    End With
    Config.ListMax = Result.ListMax
    Config.OutputInterval = Result.OutputInterval
    Config.HotKey.Control = Result.HotKey.Control
    Config.HotKey.Word = Result.HotKey.Word
    Config.REG = Result.REG
End Sub

Public Function ReturnControlKeyCode(Key As String) As vKeyCode
    Select Case Key
        Case "Control"
            ReturnControlKeyCode = ControlKey
        Case "LeftControl"
            ReturnControlKeyCode = LControl
        Case "RightControl"
            ReturnControlKeyCode = RControl
        Case "Shift"
            ReturnControlKeyCode = ShiftKey
        Case "LeftShift"
            ReturnControlKeyCode = LShift
        Case "RightShift"
            ReturnControlKeyCode = RSshft
    End Select
End Function
