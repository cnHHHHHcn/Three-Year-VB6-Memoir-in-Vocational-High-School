VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5880
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3240
      TabIndex        =   30
      Top             =   240
      Width           =   2415
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   1320
   End
   Begin VB.Frame Frame4 
      Caption         =   "预备时间"
      Height          =   1215
      Left            =   3240
      TabIndex        =   25
      Top             =   720
      Width           =   2535
      Begin VB.CommandButton Command6 
         Caption         =   "确定"
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "预备时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   28
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.CommandButton Recording 
      Caption         =   "录制"
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "启动"
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "确定"
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   2400
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "限时"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   3015
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   840
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1440
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   840
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   180
      End
      Begin VB.Label sec 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1600
         TabIndex        =   18
         Top             =   240
         Width           =   180
      End
      Begin VB.Label min 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   900
         TabIndex        =   17
         Top             =   240
         Width           =   180
      End
      Begin VB.Label hour 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   180
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "时间间隔"
      Height          =   1335
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "确定"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "时间间隔:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "键位"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3015
      Begin VB.ListBox List1 
         Height          =   420
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         MaxLength       =   1
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "指定键位:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1680
      Top             =   240
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Text            =   "单键"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Menu Key 
      Caption         =   "键"
      Visible         =   0   'False
      Begin VB.Menu Keyboard 
         Caption         =   "键盘"
         Begin VB.Menu control 
            Caption         =   "控制"
            Begin VB.Menu ctrl 
               Caption         =   "Ctrl"
            End
            Begin VB.Menu shift 
               Caption         =   "Shift"
            End
            Begin VB.Menu tab 
               Caption         =   "Tab"
            End
            Begin VB.Menu enter 
               Caption         =   "Enter"
            End
            Begin VB.Menu alt 
               Caption         =   "Alt"
            End
         End
         Begin VB.Menu point 
            Caption         =   "箭头"
            Begin VB.Menu up 
               Caption         =   "↑"
            End
            Begin VB.Menu down 
               Caption         =   "↓"
            End
            Begin VB.Menu left 
               Caption         =   "←"
            End
            Begin VB.Menu right 
               Caption         =   "→"
            End
         End
         Begin VB.Menu bcy 
            Caption         =   "不常用"
            Begin VB.Menu insert 
               Caption         =   "Insert"
            End
            Begin VB.Menu Delete 
               Caption         =   "Delete"
            End
            Begin VB.Menu Home 
               Caption         =   "Home"
            End
            Begin VB.Menu End 
               Caption         =   "End"
            End
            Begin VB.Menu PGUP 
               Caption         =   "Page Up"
            End
            Begin VB.Menu PGDN 
               Caption         =   "Page Down"
            End
         End
         Begin VB.Menu change 
            Caption         =   "锁定键"
            Begin VB.Menu capslock 
               Caption         =   "Caps lock"
            End
            Begin VB.Menu numlock 
               Caption         =   "Num lock"
            End
            Begin VB.Menu scrolllock 
               Caption         =   "Scroll lock"
            End
         End
         Begin VB.Menu F1_12 
            Caption         =   "F1-12"
            Begin VB.Menu F1 
               Caption         =   "F1"
            End
            Begin VB.Menu F2 
               Caption         =   "F2"
            End
            Begin VB.Menu F3 
               Caption         =   "F3"
            End
            Begin VB.Menu F4 
               Caption         =   "F4"
            End
            Begin VB.Menu F5 
               Caption         =   "F5"
            End
            Begin VB.Menu F6 
               Caption         =   "F6"
            End
            Begin VB.Menu F7 
               Caption         =   "F7"
            End
            Begin VB.Menu F8 
               Caption         =   "F8"
            End
            Begin VB.Menu F9 
               Caption         =   "F9"
            End
            Begin VB.Menu F10 
               Caption         =   "F10"
            End
            Begin VB.Menu F11 
               Caption         =   "F11"
            End
            Begin VB.Menu F12 
               Caption         =   "F12"
            End
         End
         Begin VB.Menu clear 
            Caption         =   "清除"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jw() As String

Private Sub alt_Click()
If Combo1.Text = "多键" Then List1.AddItem "%" Else Label2.Caption = "%"
Call Command1_Click
End Sub

Private Sub capslock_Click()
If Combo1.Text = "多键" Then List1.AddItem "CAPSLOCK" Else Label2.Caption = "CAPSLOCK"
Call Command1_Click
End Sub

Private Sub clear_Click()
List1.clear: Label2.Caption = ""
End Sub

'选择键盘模式
Private Sub Combo1_Click()
Label2.Caption = ""
If Combo1.Text = "多键" Then List1.Visible = True: List1.clear Else List1.Visible = False
End Sub

'添加按键
Private Sub Command1_Click()
    ReDim jw(1 To List1.ListCount + 1)
    If Combo1.Text = "多键" Then
        If Text1.Text <> "" Then List1.AddItem Text1.Text
        For k = 1 To List1.ListCount
           jw(k) = List1.List(k - 1)
        Next
    Else
        Label2.Caption = Text1.Text
        ReDim jw(0 To 0)
        jw(0) = Label2.Caption
    End If
    Text1.Text = ""
End Sub

'设置按键时间间隔
Private Sub Command2_Click()
Label4.Caption = Val(Text2.Text) & "ms"
Timer1.Interval = Val(Text2.Text)
Text2.Text = ""
End Sub

'设置按键持续时间
Private Sub Command4_Click()
hour.Caption = Val(Text3.Text)
min.Caption = Val(Text4.Text)
sec.Caption = Val(Text5.Text)
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

'启动
Private Sub Command5_Click()
'暂停
If Command5.Caption = "暂停" Then Timer1.Enabled = False: Timer2.Enabled = False: Command5.Caption = "启动": GoTo ends
Label5.Caption = "0"
'启动前检测
If Label2.Caption = "" And List1.ListCount = 0 Then MsgBox "未设置键位", 48
code1 = False: code2 = False
If Val(Label6.Caption) = 0 Then MsgBox "未设置预备时间", 48 Else Timer3.Enabled = True
If Val(hour.Caption) = 0 And Val(min.Caption) = 0 And Val(sec.Caption) = 0 Then MsgBox "未设置限时", 48: code1 = True
If Val(Label4.Caption) = 0 Then MsgBox "未设置时间间隔", 48: code2 = True
If code1 Or code2 Then Timer3.Enabled = False
ends:
End Sub

Private Sub Command6_Click()
Label6.Caption = Val(Text6.Text)
Text6.Text = ""
End Sub

Private Sub ctrl_Click()
If Combo1.Text = "多键" Then List1.AddItem "^" Else Label2.Caption = "^"
Call Command1_Click
End Sub

Private Sub Delete_Click()
If Combo1.Text = "多键" Then List1.AddItem "DEL" Else Label2.Caption = "DEL"
Call Command1_Click
End Sub

Private Sub down_Click()
If Combo1.Text = "多键" Then List1.AddItem "DOWN" Else Label2.Caption = "DOWN"
Call Command1_Click
End Sub

Private Sub End_Click()
If Combo1.Text = "多键" Then List1.AddItem "END" Else Label2.Caption = "END"
Call Command1_Click
End Sub

Private Sub enter_Click()
If Combo1.Text = "多键" Then List1.AddItem "ENTER" Else Label2.Caption = "ENTER"
Call Command1_Click
End Sub

Private Sub F1_Click()
If Combo1.Text = "多键" Then List1.AddItem "F1" Else Label2.Caption = "F1"
Call Command1_Click
End Sub

Private Sub F10_Click()
If Combo1.Text = "多键" Then List1.AddItem "F10" Else Label2.Caption = "F10"
Call Command1_Click
End Sub

Private Sub F11_Click()
If Combo1.Text = "多键" Then List1.AddItem "F11" Else Label2.Caption = "F1"
Call Command1_Click
End Sub

Private Sub F12_Click()
If Combo1.Text = "多键" Then List1.AddItem "F12" Else Label2.Caption = "F12"
Call Command1_Click
End Sub

Private Sub F2_Click()
If Combo1.Text = "多键" Then List1.AddItem "F2" Else Label2.Caption = "F2"
Call Command1_Click
End Sub

Private Sub F3_Click()
If Combo1.Text = "多键" Then List1.AddItem "F3" Else Label2.Caption = "F3"
Call Command1_Click
End Sub

Private Sub F4_Click()
If Combo1.Text = "多键" Then List1.AddItem "F4" Else Label2.Caption = "F4"
Call Command1_Click
End Sub

Private Sub F5_Click()
If Combo1.Text = "多键" Then List1.AddItem "F5" Else Label2.Caption = "F5"
Call Command1_Click
End Sub

Private Sub F6_Click()
If Combo1.Text = "多键" Then List1.AddItem "F6" Else Label2.Caption = "F6"
Call Command1_Click
End Sub

Private Sub F7_Click()
If Combo1.Text = "多键" Then List1.AddItem "F7" Else Label2.Caption = "F7"
Call Command1_Click
End Sub

Private Sub F8_Click()
If Combo1.Text = "多键" Then List1.AddItem "F8" Else Label2.Caption = "F8"
Call Command1_Click
End Sub

Private Sub F9_Click()
If Combo1.Text = "多键" Then List1.AddItem "F9" Else Label2.Caption = "F9"
Call Command1_Click
End Sub

Private Sub Frame1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Key
End Sub

Private Sub Home_Click()
If Combo1.Text = "多键" Then List1.AddItem "HOME" Else Label2.Caption = "HOME"
Call Command1_Click
End Sub

Private Sub insert_Click()
If Combo1.Text = "多键" Then List1.AddItem "INS" Else Label2.Caption = "INS"
Call Command1_Click
End Sub

Private Sub left_Click()
If Combo1.Text = "多键" Then List1.AddItem "LEFT" Else Label2.Caption = "LEFT"
Call Command1_Click
End Sub

Private Sub numlock_Click()
If Combo1.Text = "多键" Then List1.AddItem "NUMLOCK" Else Label2.Caption = "NUMLOCK"
Call Command1_Click
End Sub

Private Sub PGDN_Click()
If Combo1.Text = "多键" Then List1.AddItem "PGDN" Else Label2.Caption = "PGDN"
Call Command1_Click
End Sub

Private Sub PGUP_Click()
If Combo1.Text = "多键" Then List1.AddItem "PGUP" Else Label2.Caption = "PGUP"
Call Command1_Click
End Sub

Private Sub Recording_Click()

End Sub

Private Sub right_Click()
If Combo1.Text = "多键" Then List1.AddItem "RIGHT" Else Label2.Caption = "RIGHT"
Call Command1_Click
End Sub

Private Sub scrolllock_Click()
If Combo1.Text = "多键" Then List1.AddItem "SCROLLLOCK" Else Label2.Caption = "SCROLLLOCK"
Call Command1_Click
End Sub

Private Sub shift_Click()
If Combo1.Text = "多键" Then List1.AddItem "+" Else Label2.Caption = "+"
Call Command1_Click
End Sub

Private Sub tab_Click()
If Combo1.Text = "多键" Then List1.AddItem "TAB" Else Label2.Caption = "TAB"
Call Command1_Click
End Sub

'输出按键
Private Sub Timer1_Timer()
On Error Resume Next
Set a = CreateObject("wscript.shell")
Max = UBound(jw)
If Max = 0 Then Label5.Caption = 0: GoTo zreo
If Val(Label5.Caption) = Max Then Label5.Caption = 0
Label5.Caption = Label5.Caption + 1
zreo:
If jw(Val(Label5.Caption)) <> "" Then a.SendKeys "{" & jw(Val(Label5.Caption)) & "}"
End Sub

'计时
Private Sub Timer2_Timer()
If Val(sec.Caption) = 0 And Val(min.Caption) > 0 Then min.Caption = min.Caption - 1: sec.Caption = 59
If Val(hour.Caption) > 0 And Val(min.Caption) = 0 And Val(sec.Caption) = 0 Then hour.Caption = hour.Caption - 1: sec.Caption = 59: min.Caption = 59
If Val(sec.Caption) = 0 And Val(min.Caption) = 0 And Val(hour.Caption) = 0 Then
    Timer1.Enabled = False: Timer2.Enabled = False: Command5.Caption = "启动"
    Set a = CreateObject("SAPI.SpVoice")
    a.speak "键盘宏已停止使用请注意"
    GoTo Finish
End If
sec.Caption = sec.Caption - 1
Finish:
End Sub

'暂停
Private Sub Timer3_Timer()
Label6.Caption = Label6.Caption - 1
If Val(Label6.Caption) = 0 Then
    If Command5.Caption = "暂停" Then
        Timer1.Enabled = False: Timer2.Enabled = False
        Command5.Caption = "启动"
    Else
        Timer1.Enabled = True: Timer2.Enabled = True
        Command5.Caption = "暂停"
    End If
    Timer3.Enabled = False
End If
End Sub

Private Sub up_Click()
If Combo1.Text = "多键" Then List1.AddItem "UP" Else Label2.Caption = "UP"
Call Command1_Click
End Sub
