Attribute VB_Name = "KeyBoardProc"
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private BeforeKeyFlag(11) As Boolean

Private Second As Boolean

Public Enum vKeyCode
    F1 = &H70
    F2 = &H71
    F3 = &H72
    F4 = &H73
    F5 = &H74
    F6 = &H75
    F7 = &H76
    F8 = &H77
    F9 = &H78
    F10 = &H79
    F11 = &H7A
    F12 = &H7B
    ControlKey = &H11
    LControl = &HA2
    RControl = &HA3
    ShiftKey = &H10
    LShift = &HA0
    RSshft = &HA1
End Enum

Public Function KeyChange(KeyCode As vKeyCode) As Boolean
    Dim CurrentKeyFlag As Boolean
    i = KeyCode - F1
    CurrentKeyFlag = GetKeyState(KeyCode)
    KeyChange = CurrentKeyFlag Xor BeforeKeyFlag(i)
    BeforeKeyFlag(i) = CurrentKeyFlag
End Function

Public Function HotKey(Mode As vKeyCode, WordKey As String) As Boolean
    Dim WordKeyCode As Integer
    Dim First As Boolean, Result As Boolean
    WordKeyCode = Asc(UCase(WordKey))
    ModeFlag = GetKeyState(Mode)
    If ModeFlag <= -127 Then
        First = GetKeyState(WordKeyCode)
        Result = First Xor Second
        Second = First
    End If
    HotKey = Result
End Function
