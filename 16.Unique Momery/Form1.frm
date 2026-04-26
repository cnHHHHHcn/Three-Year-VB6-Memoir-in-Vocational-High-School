VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UniqueMemory"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox FontList 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      TabIndex        =   17
      Text            =   "9"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll 
      Height          =   300
      Left            =   2040
      Max             =   20
      Min             =   8
      TabIndex        =   16
      Top             =   5880
      Value           =   9
      Width           =   4215
   End
   Begin VB.CheckBox FontWidth 
      Height          =   375
      Index           =   2
      Left            =   7200
      TabIndex        =   12
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox FontWidth 
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   11
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox FontWidth 
      Height          =   375
      Index           =   0
      Left            =   7200
      TabIndex        =   10
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton WriterShow 
      Caption         =   "说明"
      Height          =   495
      Left            =   8160
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox LocateDateList 
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
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Tomorrow 
      Caption         =   "明天"
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Yesterday 
      Caption         =   "昨天"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4440
      Width           =   5535
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3240
      Width           =   5535
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   1095
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   5535
   End
   Begin VB.FileListBox MemoryFileList 
      Height          =   810
      Left            =   1440
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Left            =   4440
      TabIndex        =   18
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label FontSet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "字体大小："
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
      Left            =   480
      TabIndex        =   15
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "***:"
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
      Index           =   2
      Left            =   600
      TabIndex        =   14
      Top             =   4800
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "***:"
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
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   3480
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "请选择回忆："
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
      Left            =   1440
      TabIndex        =   8
      Top             =   360
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "***:"
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
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ReadMemory As New OperateKey
Private DateArray() As String
Private Index As Byte

Private Enum ListMove
    Up = -1
    Down = 1
End Enum

Private Sub Form_Load()
    MemoryFileList.Path = App.Path
    Label5.Caption = "注：缺少的日期为放假日期" & vbCrLf & "文中有一些错别字是人为的。"
End Sub

Private Sub HScroll_Change()
    FontList.Text = HScroll.Value
    For i = 0 To 2
        If FontWidth(i).Value = 1 Then Text(i).FontSize = HScroll.Value
    Next
End Sub

Private Sub LocateDateList_Click()
    Read LocateDateList.Text
    Index = LocateDateList.ListIndex
    LocateDateList.ToolTipText = LocateDateList.Text
End Sub

Private Sub MemoryFileList_Click()
    Label5.Caption = "注：缺少的日期为放假日期" & vbCrLf & "文中有一些错别字是人为的。"
    If MemoryFileList.FileName = "2025-06.txt" Or MemoryFileList.FileName = "2025-05.txt" Then Label5.Caption = Label5.Caption & vbCrLf & "YC:Z～Z～Z～  为最近为写作业太忙，以至于没写。"
    Index = 0: LocateDateList.Clear
    ReadMemory.File = MemoryFileList.Path & "\" & MemoryFileList.FileName
    DateArray = Split(ReadMemory.ClassList, ",")
    For i = LBound(DateArray) To UBound(DateArray)
        LocateDateList.AddItem DateArray(i)
    Next
    LocateDateList.Text = DateArray(0)
    Read DateArray(0)
End Sub

Private Sub Read(StrDate As String)
    Dim Person() As String
    ReadMemory.Class = StrDate
    For i = 0 To 2
        Label1(i).Caption = "": Text(i).Text = ""
    Next
    Person = Split(ReadMemory.KeyList, ",")
    For i = LBound(Person) To UBound(Person)
        Label1(i).Caption = UCase(Person(i)) & ":"
        Text(i).Text = ReadMemory.GetKey(, , Person(i))
    Next

End Sub

Private Sub Tomorrow_Click()
    ListMoveProc Down
End Sub

Private Sub WriterShow_Click()
    Form2.Show
End Sub

Private Sub Yesterday_Click()
    ListMoveProc Up
End Sub

Private Sub ListMoveProc(Direction As ListMove)
    If (Index = 0 And Direction = Up) Or (Index = LocateDateList.ListCount And Direction = Down) Then Exit Sub
    Index = Index + Direction
    If 0 <= Index And Index <= LocateDateList.ListCount - 1 Then
        Read DateArray(Index)
        LocateDateList.Text = LocateDateList.List(Index)
    Else
        If Index = 8 And MemoryFileList.FileName = "2025-06.txt" Then MsgBox "                                                    新手教程-第五章《中专》" & vbCrLf & _
                                                                             "任务: 对口高考-已完成" & vbCrLf & _
                                                                             "任务奖励：90+天自由探索时间、更多支线任务已开放、全球地图已下载" & vbCrLf & _
                                                                             "特殊奖励(评级)：将在1至2周内完成评级，玩家可根据评级决定是否开启新手教程进阶任务『大学』" & vbCrLf & _
                                                                             "成就: 青春逐梦·Ⅻ -已达成" & vbCrLf & _
                                                                             "当前人生进度：20%" & vbCrLf & _
                                                                             "                                             下一任务: 主线任务-第六章《大学》(未接取)" & vbCrLf & _
                                                                             "        [欢迎来到地球online!]" & vbCrLf & _
                                                                             "        你准备好了吗，玩家？" & vbCrLf & _
                                                                             "        祝你好运！！！" & vbCrLf & _
                                                                             "                                                                                                   ——地球online运营部 "
        Index = Index - 1
    End If
    LocateDateList.ToolTipText = LocateDateList.Text
End Sub

