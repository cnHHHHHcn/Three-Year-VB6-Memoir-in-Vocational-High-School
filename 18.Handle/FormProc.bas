Attribute VB_Name = "FormProc"
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Type Window
    ExStyle As WindowExStyle
    ClassName As String
    Caption As String
    Style As WindowStyle
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    FatherHandle As Long
End Type


Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Enum WindowStyle
    WS_GROUP = &H20000
    WS_TABSTOP = &H10000
    WS_BORDER = &H800000
    WS_CLIPSIBLINGS = &H4000000
    WS_CLIPCHILDREN = &H2000000
    WS_CHILD = &H40000000              '子窗体
    WS_DLGFRAME = &H400000             '有边框
    WS_OVERLAPPED = &H0&
    WS_THICKFRAME = &H40000
    
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_SIZEBOX = WS_THICKFRAME
    WS_ICONIC = WS_MINIMIZE
    
    WS_VSCROLL = &H200000
    WS_HSCROLL = &H100000
    
    WS_VISIBLE = &H10000000            '显示
    WS_DISABLED = &H8000000
    
    WS_SYSMENU = &H80000
    WS_POPUP = &H80000000
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
    
    WS_CAPTION = &HC00000              'WS_BORDER Or WS_DLGFRAME   有边框、标题
    
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_TILED = WS_OVERLAPPED
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
End Enum
Public Enum WindowExStyle
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_TRANSPARENT = &H20&
End Enum
Public Enum WindowZPos
    HWND_BOTTOM = 1
    HWND_BROADCAST = &HFFFF&
    HWND_DESKTOP = 0
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum
Private Const HTCAPTION = 2
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED



Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Type WindowInfo
    ExStyle As Long
    Style As Long
    ProcAddr As Long
    hInstance As Long
    id As Long
    UserData As Long
    DialogProcAddr As Long
    DialogResult As Long
    DialogUserData As Long
End Type
Public Enum WindowInfoFlag
    GWL_EXSTYLE = (-20)
    GWL_STYLE = (-16)
    GWL_WNDPROC = (-4)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_USERDATA = (-21)
    DWL_DLGPROC = 4
    DWL_MSGRESULT = 0
    DWL_USER = 8
End Enum


Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long



Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Enum WindowMessage
    WM_SETTEXT = &HC
    WM_PAINT = &HF
End Enum

Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type


Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Private Const WM_HOTKEY = &H312
Private Const WM_NCLBUTTONDOWN = &HA1


Private Const MaxLength = 255

Private Handle As Long

Public Enum WinType
    Normal = 0
    Active = 1
    Father = 2
    Desktop = 3
End Enum

Public Type WinSize
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Function GetHandle(WinHandle) As Long
    GetHandle = IIf(TypeName(WinHandle) = "Long", WinHandle, CLng(WinHandle))
End Function

Public Function GetWindowHandle(WindowType As WinType) As Long
    Dim Pos As POINTAPI
    GetCursorPos Pos
    Handle = WindowFromPoint(Pos.X, Pos.Y)
    Select Case WindowType
        Case Normal
            GetWindowHandle = Handle
        Case Father
            GetWindowHandle = GetParent(Handle)
        Case Active
            GetWindowHandle = GetForegroundWindow
        Case Desktop
            GetWindowHandle = GetDesktopWindow
    End Select
End Function

Public Function CreateWindow(Windows As Window) As Long
    CreateWindowEx Windows.ExStyle, Windows.ClassName, Windows.Caption, Windows.Style, Windows.Left, Windows.Top, Windows.Width, Windows.Height, Windows.FatherHandle, 0, 0, 0
End Function

Public Function GetWindowProcressID(WinHandle) As Long
    Handle = GetHandle(WinHandle)
    GetWindowThreadProcessId Handle, GetWindowProcressID
End Function

Public Function GetWindowClass(WinHandle) As String
    Dim TempStr As String
    Handle = GetHandle(WinHandle)
    TempStr = Space(MaxLength)
    GetClassName Handle, TempStr, MaxLength
    GetWindowClass = Trim(TempStr)
End Function

Public Function FindWindowHandle(WindowClass As String, WindowTitle As String) As Long
    FindWindowHandle = FindWindow(IIf(Len(Trim(WindowClass)) = 0, vbNullString, WindowClass), IIf(Len(Trim(WindowTitle)) = 0, vbNullString, WindowTitle))
End Function

Public Function FormMove(WinHandle) As Boolean
    Handle = GetHandle(WinHandle)
    ReleaseCapture
    SendMessage Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Function

Public Function AlwaysTopOn(WinHandle) As Boolean
    Handle = GetHandle(WinHandle)
    Result = SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW + SWP_NOSIZE + SWP_NOMOVE)
    AlwaysTopOn = IIf(Result <> 0, True, False)
End Function

Public Sub SetWindowLight(WinHandle, Light As Byte)
    Handle = GetHandle(WinHandle)
    rtn = GetWindowLong(Handle, GWL_EXSTYLE) '获取窗体风格
    rtn = rtn Or WS_EX_LAYERED '设置风格标志为扩展
    SetWindowLong Handle, GWL_EXSTYLE, rtn '设置风格标志为扩展
    SetLayeredWindowAttributes Handle, 0, Light, LWA_ALPHA '设置窗体为透明，透明度200，（0-不可见，255-不透明）
End Sub

Public Function GetWindowTitle(WinHandle) As String
    Dim TempStr As String
    Handle = GetHandle(WinHandle)
    TempStr = Space(MaxLength)
    GetWindowText Handle, TempStr, MaxLength
    GetWindowTitle = Trim(TempStr)
End Function

Public Sub SetWindowTitle(WinHandle, Title As String)
    Handle = GetHandle(WinHandle)
    SetWindowText Handle, Title
End Sub

Public Sub GetWindowArea(WinHandle, Area As WinSize)
    Dim AreaRect As RECT
    Handle = GetHandle(WinHandle)
    GetWindowRect Handle, AreaRect
    Area.Left = AreaRect.Left: Area.Top = AreaRect.Top
    Area.Width = (AreaRect.Right - AreaRect.Left)
    Area.Height = (AreaRect.Bottom - AreaRect.Top)
End Sub

Public Sub GetDesktopArea(DesktopArea As WinSize)
    Handle = GetDesktopWindow
    GetWindowArea Handle, DesktopArea
End Sub

Public Sub SetWindowArea(WinHandle, Area As WinSize)
    Handle = GetHandle(WinHandle)
    SetWindowPos Handle, 0, Area.Left, Area.Top, Area.Width, Area.Height, 0
End Sub

Public Sub SetWindowStyle(WinHandle, Style As WindowStyle)
    Handle = GetHandle(WinHandle)
    WinStyle = GetWindowLong(Handle, GWL_STYLE)
    WinStyle = WinStyle Or Style
    SetWindowLong Handle, GWL_STYLE, WinStyle
    SetWindowPos Handle, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub

Public Sub SetWindowExStyle(WinHandle, ExStyle As WindowExStyle)
    Handle = GetHandle(WinHandle)
    WinExStyle = GetWindowLong(Handle, GWL_EXSTYLE)
    WinExStyle = WinExStyle Or ExStyle
    SetWindowLong Handle, GWL_EXSTYLE, WinExStyle
    SetWindowPos Handle, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub

Public Sub GetWindowInfo(WinHandle, WinInfo As WindowInfo)
    Handle = GetHandle(WinHandle)
    With WinInfo
        .ExStyle = GetWindowLong(Handle, GWL_EXSTYLE)
        .Style = GetWindowLong(Handle, GWL_STYLE)
        .ProcAddr = GetWindowLong(Handle, GWL_hwndproc)
        .hInstance = GetWindowLong(Handle, GWL_HINSTANCE)
        .id = GetWindowLong(Handle, GWL_ID)
        .UserData = GetWindowLong(Handle, GWL_USERDATA)
        .DialogProcAddr = GetWindowLong(Handle, DWL_DLGPROC)
        .DialogResult = GetWindowLong(Handle, DWL_MSGRESULT)
        .DialogUserData = GetWindowLong(Handle, DWL_USER)
    End With
End Sub

