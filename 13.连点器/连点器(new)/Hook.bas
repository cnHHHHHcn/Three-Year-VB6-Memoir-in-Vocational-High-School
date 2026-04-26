Attribute VB_Name = "Hook"
'模块 Hook

'设置钩子
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As HookID, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'释放钩子
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

'钩子类型
Public Enum HookID
    WH_MOUSE_LL = 14
    WH_CALLWNDPROC = 4
    WH_CBT = 5
    WH_DEBUG = 9
    WH_FOREGROUNDIDLE = 11
    WH_GETMESSAGE = 3
    WH_HARDWARE = 8
    WH_JOURNALPLAYBACK = 1
    WH_JOURNALRECORD = 0
    WH_KEYBOARD = 2
    WH_MAX = 11
    WH_MIN = (-1)
    WH_MOUSE = 7
    WH_MSGFILTER = (-1)
    WH_SHELL = 10
    WH_SYSMSGFILTER = 6
End Enum

'钩子句柄
Private HookHandle As Long

Private Function MouseHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    MouseClick.SetMouse.MouseMessageExchange wParam
    
    Call CallNextHookEx(hHook, nCode, wParam, lParam)
End Function

Public Function SetHook()
    HookHandle = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHook, App.hInstance, 0)
End Function

Public Function UnHook()
    UnhookWindowsHookEx (HookHandle)
End Function
