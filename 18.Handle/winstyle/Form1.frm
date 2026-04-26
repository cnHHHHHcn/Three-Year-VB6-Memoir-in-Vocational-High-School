VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   720
      ScaleHeight     =   1635
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_DRAWFRAME = &H20
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const WS_DLGFRAME = &H400000
Private Const WS_POPUP = &H80000000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_MAXIMIZE = &H1000000


'例子一：任何一个控件（只要有窗口，这是我们的前提，下同），你可以在运行时随便更改它的大小。
Private Sub ControlSize(ControlName As Control, SetTrue As Boolean)
Dim dwStyle As Long
dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
If SetTrue Then
dwStyle = dwStyle Or WS_THICKFRAME
Else
dwStyle = dwStyle - WS_THICKFRAME
End If
dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub
'用法：ControlSize picture1,true；设置第二个参数为False取消这种设置，下同


'例子二：任何一个控件，我们都可以控制其显示风格为对话框的风格。
Private Sub ControlDialog(ControlName As Control, SetTrue As Boolean)
Dim dwStyle As Long
dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
If SetTrue Then
dwStyle = dwStyle Or WS_DLGFRAME
Else
dwStyle = dwStyle - WS_DLGFRAME
End If
dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub
'用法：ControlSize picture1,true



'例子三：任何一个控件，我们都可以控制其显示风格为模式对话框的风格
Private Sub ControlModal(ControlName As Control, SetTrue As Boolean)
Dim dwStyle As Long
dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
If SetTrue Then
dwStyle = dwStyle Or WS_POPUP
Else
dwStyle = dwStyle - WS_POPUP
End If
dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub
'用法：ControlModal Picture1,true


'例子四：任何一个控件，我们都可以给它加上标题栏，通过拖动标题栏，可以实现控件的运行时移动。
Private Sub ControlCaption(ControlName As Control, SetTrue As Boolean)
Dim dwStyle As Long
dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
If SetTrue Then
dwStyle = dwStyle Or WS_CAPTION
Else
dwStyle = dwStyle - WS_CAPTION
End If
dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub
'用法：ControlCaption picture1,true



'例子五：任何一个控件，我们都可以给它加上ControlBox（所谓ControlBox，就是窗体的图标＋最小化＋最大化＋关闭按钮）。
Private Sub ControlSysMenu(ControlName As Control, SetTrue As Boolean)
Dim dwStyle As Long
dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
If SetTrue Then
dwStyle = dwStyle Or WS_SYSMENU
Else
dwStyle = dwStyle - WS_SYSMENU
End If
dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
End Sub
'用法：ControlCaption picture1,true:ControlSysmenu picture1,true

Private Sub Command1_Click()
ControlSize Picture1, True
End Sub
