VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3420
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   16
      Text            =   "5"
      Top             =   2835
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "33"
      Top             =   1560
      Width           =   495
   End
   Begin VB.OptionButton Option7 
      Caption         =   "自定义（十进制VK键码）"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1560
      Width           =   2505
   End
   Begin VB.OptionButton Option6 
      Caption         =   "中键"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1200
      Width           =   825
   End
   Begin VB.OptionButton Option5 
      Caption         =   "下侧键"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   1200
      Width           =   945
   End
   Begin VB.OptionButton Option4 
      Caption         =   "上侧键"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   945
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Home键"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   945
   End
   Begin VB.OptionButton Option2 
      Caption         =   "F10键"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   825
   End
   Begin VB.OptionButton Option1 
      Caption         =   "F1键"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   720
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "16"
      Top             =   2850
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   3120
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   2640
      Top             =   -15
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   0
   End
   Begin VB.Label Label10 
      Caption         =   "缓释为左键按下到弹起之间的时间"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2550
      Width           =   2985
   End
   Begin VB.Label Label9 
      Caption         =   "缓释"
      Height          =   255
      Left            =   1770
      TabIndex        =   18
      Top             =   2895
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "毫秒"
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "52pojie UndefinedPanda 免费软件"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3270
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "请设置一个开关连点器的热键："
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   2970
   End
   Begin VB.Label Label8 
      Caption         =   "毫秒"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   2895
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "间隔"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2910
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "间隔为2次点击之间时间"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2235
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "连点器当前状态：关"
      Height          =   255
      Left            =   255
      TabIndex        =   2
      Top             =   120
      Width           =   2370
   End
   Begin VB.Label Label2 
      Caption         =   "开启后按住左键即可连点"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

Private Sub Form_Load()
switch = vbKeyF10
Randomize
rndFlag = Int(Rnd * (4096 - 1024 + 1)) + 1024   '防检测
IsHooked = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If IsHooked = True Then UnHooK
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        switch = vbKeyF1
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        switch = vbKeyF10
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then
        switch = vbKeyHome
    End If
End Sub

Private Sub Option4_Click()
    If Option4.Value = True Then
        switch = 6
    End If
End Sub

Private Sub Option5_Click()
    If Option5.Value = True Then
        switch = 5
    End If
End Sub

Private Sub Option6_Click()
    If Option6.Value = True Then
        switch = 4
    End If
End Sub

Private Sub Option7_Click()
    If Option7.Value = True Then
        switch = CInt(Text2.Text)
    End If
End Sub

Private Sub Text1_Change()
    If Text1.Text = vbNullString Then
        Text1.Text = "1"
    End If
    If CInt(Text1.Text) = 0 Then
        Text1.Text = "1"
    End If
    Text1.Text = CInt(Text1.Text)
    Timer1.Interval = CInt(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text2_Change()
    If Text2.Text = vbNullString Then
        Text2.Text = "2"
    End If
    If CInt(Text2.Text) < 2 Then
        Text2.Text = "2"
    End If
    Text2.Text = CInt(Text2.Text)
    If Option7.Value = True Then
        switch = CInt(Text2.Text)
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text3_Change()
    If Text3.Text = vbNullString Then
        Text3.Text = "1"
    End If
    If CInt(Text3.Text) < 1 Then
        Text3.Text = "1"
    End If
    Text3.Text = CInt(Text3.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If IsHooked Then
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, rndFlag
    Delay (CInt(Text3.Text))
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, rndFlag
End If
End Sub

Private Sub Timer2_Timer()
If IsHooked = False Then
Label4.Caption = "连点器当前状态：关"
Else
Label4.Caption = "连点器当前状态：开"
End If
End Sub

Private Sub Timer3_Timer()
If GetKeyState(switch) <> 1 And GetKeyState(switch) <> 0 Then
    If IsHooked = False Then
        IsHooked = True
        HooK
    Else
        UnHooK
        IsHooked = False
    End If
    Delay (250)
End If
End Sub
