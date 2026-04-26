Attribute VB_Name = "ImageProc"
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long

Private Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
Private Const OCR_NORMAL = 32512

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Private Const NIM_ADD = &H0

Private Const NIM_DELETE = &H2

Private Const NIM_MODIFY = &H1

Private Const NIF_ICON = &H2

Private Const NIF_MESSAGE = &H1

Private Const NIF_TIP = &H4

Private Const IMAGE_ICON = 1

Private Const IMAGE_CURSOR = 2

Private Const LR_DEFAULTSIZE = 0

Private Const LR_LOADFROMFILE = 16

Private Const ICON_SMALL = 0

Enum FromICON
    File = 0
    DLL = 1
    EXE = 1
    WinICON = 2
End Enum

Private Handle As Long

Public Sub SetMouseIcon(File As String)
    Handle = LoadCursorFromFile(File)
    SetSystemCursor Handle, OCR_NORMAL
End Sub

Public Function GetImageHWND(Name As String) As Long
    Dim Result As Long
    Select Case LCase(Right(Name, 4))
        Case ".ico"
            Result = LoadImage(0, Name, IMAGE_ICON, 0, 0, LR_LOADFROMFILE + LR_DEFAULTSIZE)
        Case ".cur", ".ani"
            Result = LoadImage(0, Name, IMAGE_CURSOR, 0, 0, LR_LOADFROMFILE + LR_DEFAULTSIZE)
    End Select
    GetImageHWND = Result
End Function

Public Function GetDLLImageHWND(Name As String, Index As Integer) As Long
                                'ÓÐÃ«²¡
    GetDLLImageHWND = ExtractIcon(0, Name, Index)
End Function

Public Function TaskBarAddICON(FormHWND As Long, ImageName As String, Flag As FromICON, Optional DLLIndex As Integer, Optional TipText As String) As Boolean
    Dim ICONData As NOTIFYICONDATA
    With ICONData
        .cbSize = 88
        .hwnd = FormHWND
        .uID = 1
        .uFlags = NIF_ICON + NIF_TIP + NIF_MESSAGE
        .uCallbackMessage = 2048
        .szTip = TipText
        Select Case Flag
            Case 0
                .hIcon = GetImageHWND(ImageName)
            Case 1
                .hIcon = GetDLLImageHWND(ImageName, DLLIndex)
            Case 2
                .hIcon = SendMessage(FormHWND, wm_getICon, ICON_SMALL, 0)
        End Select
    End With
    If CBool(Shell_NotifyIcon(NIM_ADD, ICONData)) Then
        TaskBarAddICON = True
    Else
        TaskBarAddICON = False
    End If
End Function

Public Function TaskBarModifyICON(FormHWND As Long, ImageName As String) As Boolean
    Dim ICONData As NOTIFYICONDATA
    With ICONData
        .cbSize = 88
        .hwnd = FormHWND
        .uID = 1
        .uFlags = NIF_ICON
        .hIcon = GetImageHWND(ImageName)
    End With
    TaskBarModifyICON = CBool(Shell_NotifyIcon(NIM_MODIFY, ICONData))
End Function


Public Function TaskBarDeleteICON(FormHWND As Long) As Boolean
    Dim ICONData As NOTIFYICONDATA
    With ICONData
        .cbSize = 88
        .hwnd = FormHWND
        .uID = 1
    End With
    TaskBarDeleteICON = CBool(Shell_NotifyIcon(NIM_DELETE, ICONData))
End Function

