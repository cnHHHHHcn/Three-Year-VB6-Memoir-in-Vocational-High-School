VERSION 5.00
Object = "{198702C4-CFCF-4077-AE03-B087916B52ED}#2.0#0"; "系统托盘.ocx"
Begin VB.Form HandleForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Handle"
   ClientHeight    =   5610
   ClientLeft      =   8160
   ClientTop       =   7770
   ClientWidth     =   7005
   Icon            =   "HandleForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7005
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3960
      TabIndex        =   42
      Top             =   4440
      Width           =   1215
   End
   Begin 系统托盘.cSysTray cSysTray1 
      Left            =   5760
      Top             =   1560
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "HandleForm.frx":0442
      TrayTip         =   "双击显示"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "创建窗口"
      Height          =   375
      Left            =   2880
      TabIndex        =   36
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame 
      Caption         =   "窗口信息"
      Height          =   2415
      Left            =   2400
      TabIndex        =   25
      Top             =   120
      Width           =   2055
      Begin VB.TextBox WinStyleTemp 
         Height          =   290
         Left            =   720
         TabIndex        =   41
         Top             =   600
         Width           =   925
      End
      Begin VB.TextBox WinExStyleTemp 
         Height          =   290
         Left            =   960
         TabIndex        =   40
         Top             =   240
         Width           =   700
      End
      Begin VB.ComboBox WinExStyle 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   39
         Top             =   240
         Width           =   990
      End
      Begin VB.ComboBox WinStyle 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   705
         TabIndex        =   38
         Top             =   600
         Width           =   1230
      End
      Begin VB.TextBox WinUserData 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox WinID 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox WinHinstance 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox WinProcAddr 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "用户数据："
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   2085
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "标识："
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   1725
         Width           =   540
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "事例句柄："
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   1365
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "过程地址："
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "风格："
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   645
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "扩展风格："
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   285
         Width           =   900
      End
   End
   Begin VB.Frame General 
      Caption         =   "窗口通用"
      Height          =   3375
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   2055
      Begin VB.ListBox WinHideList 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox WinPID 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox WinClass 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   600
         TabIndex        =   22
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox WinHandleType 
         Height          =   300
         ItemData        =   "HandleForm.frx":0894
         Left            =   960
         List            =   "HandleForm.frx":08A4
         TabIndex        =   20
         Text            =   "正常"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox WinHandle 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   18
         Text            =   "&H0"
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox WinRequest 
         Height          =   300
         ItemData        =   "HandleForm.frx":08C0
         Left            =   960
         List            =   "HandleForm.frx":08CA
         TabIndex        =   17
         Text            =   "查找"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton WinVisiable 
         Caption         =   "隐藏"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Title 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Light 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   720
         TabIndex        =   12
         Text            =   "255"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "PID："
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   2445
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "类名："
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   2080
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "句柄："
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "标题："
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "亮度："
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1000
         Width           =   540
      End
   End
   Begin VB.Frame Size 
      Caption         =   "窗口区域"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
      Begin VB.TextBox WinWidth 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox WinHeight 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox WinTop 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   765
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox WinLeft 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   765
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "宽度："
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "高度："
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "上边距："
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "左边距："
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Text            =   "VFP66400000"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Tip 
      Caption         =   "提示"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Timer GetKeyBoard 
      Interval        =   10
      Left            =   5400
      Top             =   1080
   End
   Begin VB.Timer Finding 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   1080
   End
End
Attribute VB_Name = "HandleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Dim ShowFlag As Byte, _
    WinExStyleFlag As Long, WinStyleFlag As Long

Private Sub Form_DblClick()
    HandleForm.Hide
    cSysTray1.InTray = True
End Sub

Private Sub Form_Load()
    ShowFlag = 255: WinHideList.Clear
    With WinStyle
        .AddItem "WS_BORDER"
        .AddItem "WS_CAPTION"
        .AddItem "WS_CHILD"
        .AddItem "WS_CLIPCHILDREN"
        .AddItem "WS_CLIPSIBLINGS"
        .AddItem "WS_DISABLED"
        .AddItem "WS_DLGFRAME"
        .AddItem "WS_GROUP"
        .AddItem "WS_HSCROLL"
        .AddItem "WS_ICONIC"
        .AddItem "WS_MAXIMIZE"
        .AddItem "WS_MAXIMIZEBOX"
        .AddItem "WS_MINIMIZE"
        .AddItem "WS_MINIMIZEBOX"
        .AddItem "WS_OVERLAPPED"
        .AddItem "WS_OVERLAPPEDWINDOW"
        .AddItem "WS_POPUP"
        .AddItem "WS_POPUPWINDOW"
        .AddItem "WS_SIZEBOX"
        .AddItem "WS_SYSMENU"
        .AddItem "WS_TABSTOP"
        .AddItem "WS_THICKFRAME"
        .AddItem "WS_TILED"
        .AddItem "WS_TILEDWINDOW"
        .AddItem "WS_VISIBLE"
        .AddItem "WS_VSCROLL"
    End With
    Dim a As WindowExStyle
    With WinExStyle
        .AddItem "WS_EX_ACCEPTFILES"
        .AddItem "WS_EX_DLGMODALFRAME"
        .AddItem "WS_EX_NOPARENTNOTIFY"
        .AddItem "WS_EX_TOPMOST"
        .AddItem "WS_EX_TRANSPARENT"
    End With
End Sub

Private Sub Tip_Click()
    MsgBox "1.按住 Alt 键，指向窗口标题获取句柄" & Chr(13) & "2.在标题框输入窗口标题获取句柄" & Chr(13) & "3.修改亮度让窗口透明。" & Chr(13) & "4.按 ←、→ 键调节窗口透明度" & Chr(13) & "5.也可以调节其他参数", 64, "Tip"
End Sub

Private Sub cSysTray1_MouseDblClick(Button As Integer, id As Long)
    HandleForm.Show
End Sub

Private Sub WinClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Text1.Text = FormProc.FindWindowHandle(WinClass.Text, "")
End Sub

Private Sub WinExStyle_Click()
    WinExStyleFlag = WinExStyleFlag Or Choose(WinExStyle.ListIndex + 1, WS_EX_ACCEPTFILES, WS_EX_DLGMODALFRAME, WS_EX_NOPARENTNOTIFY, _
                                                                        WS_EX_TOPMOST, WS_EX_TRANSPARENT)
    WinExStyleTemp.Text = WinExStyleTemp.Text & "+" & WinExStyle.Text
    WinExStyle.Text = Mid(WinExStyleTemp, 2)
    WinExStyleTemp.Text = WinExStyle.Text
    WinExStyleTemp.ToolTipText = WinExStyle.Text
    WinExStyle.ToolTipText = WinExStyle.Text
End Sub

Private Sub WinExStyle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then FormProc.SetWindowExStyle WinHandle.Text, WinExStyleFlag
End Sub

Private Sub WinExStyleTemp_Change()
    If Len(Trim(WinExStyleTemp.Text)) = 0 Then WinExStyleFlag = 0
End Sub

Private Sub WinStyle_Click()
    WinStyleFlag = WinStyleFlag Or Choose(WinStyle.ListIndex + 1, WS_BORDER, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN, WS_CLIPSIBLINGS, _
                                                                  WS_DISABLED, WS_DLGFRAME, WS_GROUP, WS_HSCROLL, WS_ICONIC, WS_MAXIMIZE, _
                                                                  WS_MAXIMIZEBOX, WS_MINIMIZE, WS_MINIMIZEBOX, WS_OVERLAPPED, WS_OVERLAPPEDWINDOW, _
                                                                  WS_POPUP, WS_POPUPWINDOW, WS_SIZEBOX, WS_SYSMENU, WS_TABSTOP, WS_THICKFRAME, _
                                                                  WS_TILED, WS_TILEDWINDOW, WS_VISIBLE, WS_VSCROLL)
    WinStyleTemp.Text = WinStyleTemp.Text & "+" & WinStyle.Text
    WinStyleTemp.Text = Mid(WinStyleTemp, 2)
    
    WinStyle.ToolTipText = WinStyle.Text
End Sub

Private Sub WinStyle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then FormProc.SetWindowStyle WinHandle, WinStyleFlag
End Sub

Private Sub WinStyleTemp_Change()
    If Len(Trim(WinStyleTemp.Text)) = 0 Then WinStyleFlag = 0
End Sub

Private Sub WinHideList_Click()
    Light.Text = 0
    WinHandle.Text = WinHideList.Text
End Sub

Private Sub Light_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 Or KeyCode = 39 Then
        If KeyCode = 37 Then
            Light.Text = Val(Light.Text) - 5
            If Val(Light.Text) <= 15 Then Light.Text = 5
        End If
        If KeyCode = 39 Then
            Light.Text = Val(Light.Text) + 5
            If Val(Light.Text) >= 255 Then Light.Text = 255
        End If
        FormProc.SetWindowLight WinHandle.Text, Val(Light.Text)
    End If
End Sub

Private Sub Light_KeyPress(KeyAscii As Integer)
    On Error GoTo Max
    GoTo NextLine
Max: Light.Text = 255
NextLine:
    If KeyAscii = 13 Then
        FormProc.SetWindowLight WinHandle.Text, Val(Light.Text)
    End If
End Sub

Private Sub WinHeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ChangeWinSize
End Sub

Private Sub WinLeft_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ChangeWinSize
End Sub

Private Sub WinTop_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ChangeWinSize
End Sub

Private Sub WinWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ChangeWinSize
End Sub

Private Sub ChangeWinSize()
    Dim Area As WinSize
    Area.Left = WinLeft.Text
    Area.Top = WinTop.Text
    Area.Height = WinHeight.Text
    Area.Width = WinWidth.Text
    FormProc.SetWindowArea WinHandle.Text, Area
End Sub

Private Sub WinVisiable_Click()
    WinVisiable.Caption = IIf(Not CBool(ShowFlag), "隐藏", "显示")
    ShowFlag = Not ShowFlag
    Light.Text = ShowFlag
    If "&H" & Hex(Me.hwnd) <> WinHandle.Text Then
        If Len(Trim(WinHandle.Text)) <> 0 Then
            FormProc.SetWindowLight WinHandle.Text, ShowFlag
            If Not CBool(ShowFlag) Then
                WinHideList.AddItem WinHandle.Text
            Else
                For i = WinHideList.ListCount - 1 To 0 Step -1
                    If WinHideList.List(i) = WinHandle.Text Then WinHideList.RemoveItem i
                Next
            End If
        End If
    Else
        Light.Text = 255
    End If
End Sub

Private Sub Title_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WinRequest.Text = "查找" Then WinHandle.Text = "&H" & Hex(FormProc.FindWindowHandle(WinClass.Text, Title.Text))
        If WinRequest.Text = "更改" Then FormProc.SetWindowTitle WinHandle.Text, Title.Text
    End If
End Sub

Private Sub Finding_Timer()
    Dim Area As WinSize, Info As WindowInfo
    WinHandle.Text = "&H" & Hex(FormProc.GetWindowHandle(Switch(WinHandleType.Text = "正常", Normal, WinHandleType.Text = "父级", Father, WinHandleType.Text = "活动", Active, WinHandleType.Text = "桌面", Desktop)))
    Title.Text = FormProc.GetWindowTitle(WinHandle.Text)
    WinClass.Text = FormProc.GetWindowClass(WinHandle.Text)
    FormProc.GetWindowArea WinHandle.Text, Area
    FormProc.GetWindowInfo WinHandle.Text, Info
    With Info
        WinExStyleTemp.Text = .ExStyle
        WinStyleTemp.Text = "&H" & Hex(.Style)
        WinProcAddr.Text = "&H" & Hex(.ProcAddr)
        WinHinstance.Text = .hInstance
        WinID.Text = .id
        WinUserData.Text = .UserData
    End With
    WinPID.Text = FormProc.GetWindowProcressID(WinHandle.Text)
    WinLeft.Text = Area.Left
    WinTop.Text = Area.Top
    WinWidth.Text = Area.Width
    WinHeight.Text = Area.Height
End Sub

Private Sub GetKeyBoard_Timer()
    Flag = GetKeyState(18) <= -127 Or (WinHandleType.Text = "活动" Or WinHandleType.Text = "桌面")
    Finding.Enabled = Flag
    ImageProc.SetMouseIcon App.Path & "\Picture\" & Abs(Flag) & ".cur"
End Sub
