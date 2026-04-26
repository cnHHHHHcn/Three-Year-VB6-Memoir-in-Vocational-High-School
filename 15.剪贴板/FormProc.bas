Attribute VB_Name = "FormProc"
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const WM_HOTKEY = &H312
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40
Public Const HTCAPTION = 2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4

Public Function FormMove(ByVal hwnd As Long) As Boolean
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Function

Public Function AlwaysTopOn(ByVal hwnd As Long) As Boolean
    Result = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW + SWP_NOSIZE + SWP_NOMOVE)
    AlwaysTopOn = IIf(Result <> 0, True, False)
End Function
