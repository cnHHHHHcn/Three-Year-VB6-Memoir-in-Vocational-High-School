VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   5790
   ClientLeft      =   5115
   ClientTop       =   3615
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Choose 
      Interval        =   100
      Left            =   1800
      Top             =   720
   End
   Begin VB.CommandButton Clear 
      Height          =   495
      Left            =   2760
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Down 
      Height          =   495
      Left            =   2160
      Picture         =   "Form1.frx":02BA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Delete 
      Height          =   495
      Left            =   1560
      Picture         =   "Form1.frx":06FC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Up 
      Height          =   495
      Left            =   960
      Picture         =   "Form1.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   495
   End
   Begin VB.Timer PrintClipBoardText 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   720
   End
   Begin VB.Timer CilpboardCheck 
      Interval        =   500
      Left            =   600
      Top             =   720
   End
   Begin VB.ListBox ClipBoardList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   3960
      Picture         =   "Form1.frx":0C70
      Top             =   120
      Width           =   420
   End
   Begin VB.Menu HotKey 
      Caption         =   "粘贴快键"
      Visible         =   0   'False
      Begin VB.Menu Frist 
         Caption         =   "第一位"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Second 
         Caption         =   "第二位"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Third 
         Caption         =   "第三位"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentText As String, BeforeText As String
Dim FileName As String
Dim Index As Integer

Private Enum ListMove
    Up = -1
    Down = 1
End Enum

Private Sub Choose_Timer()
    On Error GoTo ExitSub
    For i = 0 To 11
        If KeyBoardProc.KeyChange(F1 + i) Then
            Text = ClipBoardList.List(i)
            If Len(Text) <> 0 Then
                Clipboard.Clear
                Clipboard.SetText Text
            End If
        End If
    Next
    GoTo Run
ExitSub:
    Exit Sub
Run:
    If KeyBoardProc.HotKey(ClipBoardDispose.ReturnControlKeyCode(Config.HotKey.Control), Config.HotKey.Word) Then _
    Form1.Visible = Not Form1.Visible
End Sub

Private Sub CilpboardCheck_Timer()
    Dim NotRepeatFlag As Boolean
    NotRepeatFlag = True
    CurrentText = Clipboard.GetText
    If CurrentText = "/Set Administrator" Then
        Administrator.Show
        Clipboard.Clear
    End If
    If Not CurrentText = BeforeText And Len(Trim(CurrentText)) <> 0 Then
        For i = 1 To ClipBoardList.ListCount
            If ClipBoardList.List(i - 1) = CurrentText Then NotRepeatFlag = False: Exit For
        Next
        If NotRepeatFlag Then ClipBoardList.AddItem CurrentText, 0
    End If
    BeforeText = CurrentText
    Form2.Label1.Caption = "剪贴板当前内容：" & Chr(13) & Clipboard.GetText
    If ClipBoardList.ListCount = Config.ListMax + 1 Then ClipBoardList.RemoveItem Config.ListMax
End Sub

Private Sub Clear_Click()
    If MsgBox("是否清楚全部数据项？", 32 + vbYesNo, "ClipBoardList") = 6 Then ClipBoardList.Clear
End Sub

Private Sub ClipBoardList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case True
        Case (Button = 1 And Shift = 0)
            Clipboard.Clear
            Clipboard.SetText ClipBoardList.Text
            PrintClipBoardText.Enabled = True
        Case (Button = 2 And Shift = 4 And Len(ClipBoardList.List(ClipBoardList.ListIndex)) <> 0)
            Index = ClipBoardList.ListIndex
            Form2.Text.Text = ClipBoardList.List(Index)
            Form2.Show
    End Select
End Sub

Private Sub Delete_Click()
    On Error GoTo ExitSub
    GoTo Run
ExitSub:
    Exit Sub
Run:
    Index = ClipBoardList.ListIndex
    ClipBoardList.RemoveItem Index
    ClipBoardList.Selected(Index) = True
End Sub

Private Sub Form_Load()
    Dim a As Config
    ClipBoardDispose.LoadDispose a
    BeforeText = Clipboard.GetText
    ClipBoardList.AddItem BeforeText
    FormProc.AlwaysTopOn Form1.hwnd
    For i = 0 To 11
        KeyBoardProc.KeyChange F1 + i
    Next
    If Not Config.REG Then
        Administrator.REG_Click
        OperaDispose.SetKey "True", App.Path & "\Config.Dispose", "Dispose", "REG"
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        FormProc.FormMove Form1.hwnd
    End If
    FormProc.AlwaysTopOn Form1.hwnd
End Sub

Private Sub Image1_Click()
    Form3.Show
    'ClipBoardDispose.ProgramClosed
End Sub

Private Sub PrintClipBoardText_Timer()
    Set a = CreateObject("wscript.shell")
    a.SendKeys "^v"
    PrintClipBoardText.Enabled = False
End Sub

Private Sub Up_Click()
    ListMoveProc Up
End Sub

Private Sub Down_Click()
    ListMoveProc Down
End Sub

Private Sub ListMoveProc(Direction As ListMove)
    Text = ClipBoardList.Text
    Index = ClipBoardList.ListIndex
    
    If (Index = 0 And Direction = Up) Or (Index = ClipBoardList.ListCount - 1 And Direction = Down) Then Exit Sub
    
    If 0 <= Index And Index <= ClipBoardList.ListCount - 1 Then
        ClipBoardList.RemoveItem Index
        ClipBoardList.AddItem Text, Index + Direction
        ClipBoardList.Selected(Index + Direction) = True
    End If
End Sub
