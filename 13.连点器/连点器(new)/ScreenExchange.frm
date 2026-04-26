VERSION 5.00
Begin VB.Form ScreenExchange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ScreenExchange"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7335
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton GetScreenArea 
      Caption         =   "获取屏幕分辨率"
      Height          =   495
      Left            =   480
      TabIndex        =   25
      Top             =   4560
      Width           =   1575
   End
   Begin VB.FileListBox TranscribeFileList 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Finish 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "完成"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame WriteTransrcibeArea 
      Caption         =   "录制文件编写区"
      Enabled         =   0   'False
      Height          =   3615
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6855
      Begin VB.TextBox FileLocate 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         ToolTipText     =   "点击即可复制"
         Top             =   3240
         Width           =   6615
      End
      Begin VB.TextBox NewFileName 
         Appearance      =   0  'Flat
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   4920
         TabIndex        =   24
         Text            =   "请输入文件名"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox CurrentScreenWidth 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox CurrentScreenHeight 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton Exchange 
         Caption         =   "转换"
         Height          =   495
         Left            =   4920
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton Options 
         Caption         =   "替换当前录制文件"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   20
         Top             =   2160
         Width           =   1815
      End
      Begin VB.OptionButton Options 
         Caption         =   "生成新的录制文件"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   19
         Top             =   2640
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox NeedScreenHeight 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1440
         TabIndex        =   18
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox NeedScreenWidth 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1440
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ListBox ScreenYList 
         Appearance      =   0  'Flat
         Height          =   1110
         ItemData        =   "ScreenExchange.frx":0000
         Left            =   1560
         List            =   "ScreenExchange.frx":0002
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.ListBox ButtonList 
         Appearance      =   0  'Flat
         Height          =   1110
         ItemData        =   "ScreenExchange.frx":0004
         Left            =   2880
         List            =   "ScreenExchange.frx":0006
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.ListBox StateList 
         Appearance      =   0  'Flat
         Height          =   1110
         ItemData        =   "ScreenExchange.frx":0008
         Left            =   4200
         List            =   "ScreenExchange.frx":000A
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin VB.ListBox ScreenXList 
         Appearance      =   0  'Flat
         Height          =   1110
         ItemData        =   "ScreenExchange.frx":000C
         Left            =   240
         List            =   "ScreenExchange.frx":000E
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "转换成分辨率："
         Height          =   180
         Left            =   1440
         TabIndex        =   16
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "文件分辨率："
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label CurrentRecord 
         AutoSize        =   -1  'True
         Caption         =   "当前选中："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   5520
         TabIndex        =   11
         Top             =   480
         Width           =   750
      End
      Begin VB.Label TransrcibeList 
         AutoSize        =   -1  'True
         Caption         =   "总计："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   450
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   5400
         X2              =   5400
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   6840
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "State:"
         Height          =   180
         Left            =   4200
         TabIndex        =   6
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ScreenY:"
         Height          =   180
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Button:"
         Height          =   180
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ScreenX:"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.TextBox OpenFile 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      TabIndex        =   1
      Top             =   300
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "文件名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "ScreenExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Percent As ScreenPercent, _
    ScreenWidth As Long, ScreenHeight As Long

Private Sub Exchange_Click()
    Dim Data As MouseButtonData, Note As String
    Note = MouseClick.ProjectDispose.GetKey(, "MouseTranscribe", OpenFile.Text)
    Percent.Width = Round(Val(NeedScreenWidth.Text) / CurrentScreenWidth, 5)
    Percent.Height = Round(Val(NeedScreenHeight.Text) / CurrentScreenHeight, 5)
    ExchangeFile Percent
    Select Case True
        Case Options(0).Value
            NewFileName.Enabled = False
            Open TranscribeFileList.Path & "\" & OpenFile.Text For Random As #2
                For i = 0 To ScreenXList.ListCount - 1
                    Data.MouseX = ScreenXList.List(i)
                    Data.MouseY = ScreenYList.List(i)
                    Data.Button = ButtonList.List(i)
                    Data.State = StateList.List(i)
                    Put #2, i + 1, Data
                Next
            Close #2
            OldName = OpenFile.Text
            MouseClick.ProjectDispose.DeleteKey , "MouseTranscribe", OpenFile.Text
            OpenFile.Text = Left(OpenFile.Text, InStr(1, OpenFile.Text, "(") - 1) & "(" & NeedScreenWidth.Text & "-" & NeedScreenHeight.Text & ").MDS"
            Name TranscribeFileList.Path & "\" & OldName As TranscribeFileList.Path & "\" & OpenFile.Text
        Case Options(1).Value
            If Len(Trim(NewFileName.Text)) = 0 Then MsgBox "请输入新的文件名！", 48: Exit Sub
            OpenFile = NewFileName.Text & "(" & NeedScreenWidth.Text & "-" & NeedScreenHeight.Text & ").MDS"
            Open TranscribeFileList.Path & "\" & OpenFile.Text For Random As #2
                For i = 0 To ScreenXList.ListCount - 1
                    Data.MouseX = ScreenXList.List(i)
                    Data.MouseY = ScreenYList.List(i)
                    Data.Button = ButtonList.List(i)
                    Data.State = StateList.List(i)
                    Put #2, i + 1, Data
                Next
            Close #2
    End Select
    MouseClick.ProjectDispose.CreateKey "", "MouseTranscribe", OpenFile.Text & ";" & Note
End Sub

Private Sub FileLocate_Click()
    Clipboard.Clear
    Clipboard.SetText Mid(FileLocate.Text, 6)
End Sub

Private Sub Finish_Click()
    Finish.Enabled = False
End Sub

Private Sub Form_Load()
    TranscribeFileList.Path = App.Path & "\Transcribe-Data\"
    GetScreenArea_Click
End Sub

Private Sub ExchangeFile(Percent As ScreenPercent)
    For i = 0 To ScreenXList.ListCount - 1
        If i Mod 500 = 0 Then DoEvents
        ScreenXList.List(i) = Fix(ScreenXList.List(i) * Percent.Width)
        ScreenYList.List(i) = Fix(ScreenYList.List(i) * Percent.Height)
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MouseClick.GetKeyCheck.Enabled = True
End Sub

Private Sub GetScreenArea_Click()
    MouseClick.SetMouse.GetDeskTopArea ScreenWidth, ScreenHeight
    NeedScreenWidth.Text = ScreenWidth
    NeedScreenHeight.Text = ScreenHeight
End Sub

Private Sub NewFileName_GotFocus()
    NewFileName.ForeColor = &H80000012
    If NewFileName.Text <> "请输入文件名" Then Exit Sub
    NewFileName.Text = ""
End Sub

Private Sub NewFileName_LostFocus()
    If NewFileName.Text <> "" Then Exit Sub
    NewFileName.ForeColor = &H808080
    NewFileName.Text = "请输入文件名"
End Sub

Private Sub OpenFile_KeyPress(KeyAscii As Integer)
    Dim Index As Long, MouseTranscribe As MouseButtonData
    If KeyAscii = 13 Then
        If Len(Trim(OpenFile.Text)) <> 0 Then
            FileLocate.Text = "文件位于:" & TranscribeFileList.Path & "\" & OpenFile.Text
            Flag = InStr(1, OpenFile.Text, ".MDS")
            If Flag = 0 Then Flag = 1
            FileName = IIf(Flag - 1 = 0, OpenFile.Text, Left(OpenFile.Text, Flag - 1))
            FileName = App.Path & "\Transcribe-Data\" & FileName & ".MDS"
            WriteTransrcibeArea.Enabled = True
            Finish.Enabled = True
            FileNum = FreeFile
            StartPos = InStr(1, OpenFile.Text, "(") + 1
            EndPos = InStr(1, OpenFile.Text, "-")
            CurrentScreenWidth.Text = Mid(OpenFile.Text, StartPos, EndPos - StartPos)
            StartPos = EndPos + 1
            EndPos = InStr(1, OpenFile.Text, ")")
            CurrentScreenHeight.Text = Mid(OpenFile.Text, StartPos, EndPos - StartPos)
            Open FileName For Random As #FileNum
                ListClear
                Do Until EOF(FileNum)
                    Index = Index + 1
                    If Index Mod 500 = 0 Then DoEvents
                    Get #FileNum, Index, MouseTranscribe
                    If Not (MouseTranscribe.Button = 0 And MouseTranscribe.State = 0 And MouseTranscribe.MouseX = 0 And MouseTranscribe.MouseY = 0) Then
                        ScreenXList.AddItem MouseTranscribe.MouseX
                        ScreenYList.AddItem MouseTranscribe.MouseY
                        ButtonList.AddItem MouseTranscribe.Button
                        StateList.AddItem MouseTranscribe.State
                    End If
                TransrcibeList.Caption = "总计：" & Index
                Loop
            Close #1
        Else
            MsgBox "文件名不能为空！", 48
            WriteTransrcibeArea.Enabled = False
            Finish.Enabled = False: ListClear
        End If
    End If
End Sub

Private Sub ListClear()
    ScreenXList.Clear: ScreenYList.Clear
    ButtonList.Clear: StateList.Clear
End Sub

Private Sub Options_Click(Index As Integer)
    NewFileName.Enabled = CBool(Index)
End Sub

Private Sub TranscribeFileList_Click()
    OpenFile.Text = TranscribeFileList.FileName
    TranscribeFileList.Refresh
End Sub

Private Sub Record(Index As Integer)
    ScreenXList.Selected(Index) = True
    ScreenYList.Selected(Index) = True
    ButtonList.Selected(Index) = True
    StateList.Selected(Index) = True
    CurrentRecord.Caption = "当前选中：" & Index + 1
End Sub

Private Sub ScreenXList_Click()
    Record ScreenXList.ListIndex
End Sub

Private Sub ScreenYList_Click()
    Record ScreenYList.ListIndex
End Sub

Private Sub ButtonList_Click()
    Record ButtonList.ListIndex
End Sub

Private Sub StateList_Click()
    Record StateList.ListIndex
End Sub
