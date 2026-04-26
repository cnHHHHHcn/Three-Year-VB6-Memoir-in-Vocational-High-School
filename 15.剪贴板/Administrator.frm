VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Administrator 
   Caption         =   "**调试模式**"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "Administrator.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   8445
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton SaveSAdmin 
      Caption         =   "保存SAdmin"
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton REG 
      Caption         =   "注册注册表文件"
      Height          =   375
      Left            =   6240
      TabIndex        =   18
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Text            =   "1000"
         ToolTipText     =   "设置ctrl+v的秒数/ms"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Text            =   "500"
         ToolTipText     =   "设置检测剪贴板的秒数的秒数/ms"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton LoadSAdmin 
      Caption         =   "导入SAdmin"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   2760
      TabIndex        =   10
      Top             =   1560
      Width           =   3375
      Begin VB.CommandButton PointList 
         Caption         =   "预览列表指针"
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Point 
         Caption         =   "预览指针"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton BackGround 
         Caption         =   "预览背景"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Text            =   "3000"
         ToolTipText     =   "ctrl按住间隔/ms"
         Top             =   240
         Width           =   615
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   4560
         Top             =   240
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Text            =   "1"
         ToolTipText     =   "False判断秒数/ms"
         Top             =   240
         Width           =   615
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5520
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5040
         Top             =   240
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Text            =   "1"
         ToolTipText     =   "True判断秒数/ms"
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "开启Ctrl模式（测试）"
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   5895
      Begin VB.CommandButton DelText 
         Caption         =   "Del"
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton AddText 
         Caption         =   "Add"
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "即将导入的字符串"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "0"
         ToolTipText     =   "导入字符串的index"
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".set"
      MaxFileSize     =   9999
   End
End
Attribute VB_Name = "Administrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim string_ As String
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Sub Check1_Click()
    If Check1.Value = 1 Then Timer3.Enabled = True Else Timer2.Enabled = False: Timer1.Enabled = False: Timer3.Enabled = False
End Sub

Private Sub LoadSAdmin_Click()
CommonDialog1.Filter = "导入文件 (*.SAdmin)|*.SAdmin"
CommonDialog1.ShowOpen
Set_SAdmin = CommonDialog1.FileName
MsgBox CommonDialog1.FileName
End Sub

Private Sub AddText_Click()
On Error GoTo Error
    Form1.ClipBoardList.AddItem Text4.Text, Val(Text3.Text)
Exit Sub

Error:
    If Err.Number = 5 Then
        For i = Form1.ClipBoardList.ListCount To Val(Text3.Text) - 1
            Form1.ClipBoardList.AddItem ""
        Next i
    End If
End Sub

Private Sub DelText_Click()
On Error GoTo Error
    Form1.ClipBoardList.RemoveItem Val(Text3.Text)
    Exit Sub

Error:
    If Err.Number = 5 Then
        For i = Form1.ClipBoardList.ListCount To Val(Text3.Text) - 1
            Form1.ClipBoardList.AddItem ""
        Next i
    End If
End Sub

Private Sub BackGround_Click()
CommonDialog1.Filter = "导入文件 (*.图片)|*.png"
CommonDialog1.ShowOpen
Set_Picture = CommonDialog1.FileName
End Sub

Private Sub Point_Click()
CommonDialog1.Filter = "导入文件 (*.Cur)|*.Cur"
CommonDialog1.ShowOpen
Set_SAdmin = CommonDialog1.FileName
End Sub

Private Sub PointList_Click()
CommonDialog1.Filter = "导入文件 (*.Cur|*.Cur"
CommonDialog1.ShowOpen
Set_SAdmin = CommonDialog1.FileName
End Sub

Public Sub REG_Click()
 Set fs = CreateObject("Scripting.FileSystemObject")
 fs.CopyFile App.Path & "\ICO\" & "minecraft.ico", "c:\windows\"
UpdateKey &H80000000, ".SAdmin\PersistentHandler", "", "{5e456d80-bf96-11ad-b589-52642b30bfrg}" '"{5e783d80-fd69-56ad-b545-52642b31bfrg}"
UpdateKey &H80000000, ".SAdmin", "", "SAdmina"
UpdateKey &H80000000, "SAdmina\DefaultIcon", "", "c:\windows\minecraft.ico"
UpdateKey &H80000000, "SAdmina\shell\edit\command", "", "C:\Windows\System32\notepad.exe %1"
UpdateKey &H80000000, "SAdmina\shell\open\command", "", "C:\Windows\System32\notepad.exe %1"

fs.CopyFile App.Path & "\ICO\" & "ClipBoard.ico", "c:\windows\"
UpdateKey &H80000000, ".ClipText\PersistentHandler", "", "{5e783d80-fd69-56ad-b545-52642b31bfrg}"
UpdateKey &H80000000, ".ClipText", "", "ClipTexta"
UpdateKey &H80000000, "ClipTexta\DefaultIcon", "", "c:\windows\ClipBoard.ico"
UpdateKey &H80000000, "ClipTexta\shell\edit\command", "", "C:\Windows\System32\notepad.exe %1"
UpdateKey &H80000000, "ClipTexta\shell\open\command", "", "C:\Windows\System32\notepad.exe %1"
'开机自启动
UpdateKey &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ClipBoardLook", App.Path & "\" & "ClipBoardLook.exe"

End Sub

Private Sub Form_Load()
    Form1.Show
End Sub

Private Sub Text1_Change()
Form1.PrintClipBoardText.Interval = Val(Text1.Text)
End Sub

Private Sub Text2_Change()
    Form1.CilpboardCheck.Interval = Val(Text2.Text)
End Sub

Private Sub Text5_Change()
    Timer2.Interval = Val(Text5.Text)
End Sub

Private Sub Text6_Change()
    Timer1.Interval = Val(Text6.Text)
End Sub

Private Sub Text7_Change()
    Timer3.Interval = Val(Text7.Text)
End Sub

Private Sub Timer1_Timer()
    If (GetKeyState(17) = -127) Or (GetKeyState(17) = -128) Then
        If Clipboard.GetText <> "" Then
            string_ = string_ + Clipboard.GetText
            Clipboard.Clear
        End If
    Else
        If string_ <> "" Then
            Form1.ClipBoardList.AddItem string_, 0
            string_ = ""
        End If
        Timer2.Enabled = True
        Timer1.Enabled = False
    End If
    
    
End Sub

Private Sub Timer2_Timer()
    If (GetKeyState(17) = -127) Or (GetKeyState(17) = -128) Then
        Timer3.Enabled = True
        Timer2.Enabled = False
    Else
        Timer3.Enabled = False
    End If
End Sub

Private Sub Timer3_Timer()
    Timer1.Enabled = True
    Clipboard.Clear
    Timer3.Enabled = False
End Sub
