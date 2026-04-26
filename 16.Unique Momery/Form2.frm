VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "说明"
   ClientHeight    =   6915
   ClientLeft      =   4170
   ClientTop       =   1815
   ClientWidth     =   9150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox WriterText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   5295
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form2.frx":0000
      Top             =   7800
      Width           =   6015
   End
   Begin VB.TextBox WriterText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   5295
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form2.frx":018A
      Top             =   720
      Width           =   6015
   End
   Begin VB.Timer LabelMoveTime 
      Interval        =   5
      Left            =   8640
      Top             =   1320
   End
   Begin VB.ListBox LaptopDateList 
      Height          =   420
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer ClickTime 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8640
      Top             =   720
   End
   Begin VB.Label ClickLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim add As Byte
Dim Text(2) As String

Private Sub ClickTime_Timer()
    add = add + 1
    ClickLabel.Top = ClickLabel.Top + 10
    If add = 20 Then add = 0: ClickTime.Enabled = False: ClickLabel.Visible = False
End Sub

Private Sub Form_Load()
    WriterText(1).Text = WriterText(0).Text
    WriterText(0).Top = Me.Height \ 2
    Text(0) = "YC": Text(1) = "YH": Text(2) = "RH"
    LaptopDateList.AddItem "第一册：2024年9月25日 - 2024年11月5日"
    LaptopDateList.AddItem "第二册：2024年11月6日 - 2024年11月29日"
    LaptopDateList.AddItem "第三册：2024年12月2日 - 2024年12月29日"
    LaptopDateList.AddItem "第四册：2024年12月30日 - 2025年1月23日"
    LaptopDateList.AddItem "第五册：2025年2月5日 - 2025年2月28日"
    LaptopDateList.AddItem "第六册：2025年3月3日 - 2025年3月31日"
    LaptopDateList.AddItem "第七册：2025年4月1日 - 2025年4月29日"
    LaptopDateList.AddItem "第八册：2025年4月30日 - 2025年6月8日"
    For i = 0 To 1
        WriterText(i).Text = WriterText(i).Text & vbCrLf & "小日志册个数及日期"
        For j = 0 To LaptopDateList.ListCount - 1
            WriterText(i).Text = WriterText(i).Text & vbCrLf & LaptopDateList.List(j)
        Next
        WriterText(i).Left = Me.Width \ 2 - WriterText(i).Width \ 2
    Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickLabelShow Me, X, Y
End Sub

Private Sub ClickLabelShow(Object As Object, X As Single, Y As Single)
    If Object Is Form2 Then
        ObjectLeft = 0: ObjectTop = 0
    Else
        ObjectLeft = Object.Left: ObjectTop = Object.Top
    End If
    add = 0: ClickLabel.Visible = True: ClickLabel.Caption = Text(Int(Rnd * 3))
    ClickLabel.Left = ObjectLeft + X - ClickLabel.Width \ 2: ClickLabel.Top = ObjectTop + Y - ClickLabel.Height \ 2
    ClickLabel.ForeColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
    ClickTime.Enabled = True
End Sub

Private Sub LabelMoveTime_Timer()
    Static Flag As Boolean
    WriterText(Abs(Flag)).Top = WriterText(Abs(Flag)).Top - 10
    If WriterText(Abs(Flag)).Top < 0 Then WriterText(Abs(Not Flag)).Top = WriterText(Abs(Not Flag)).Top - 10
    If WriterText(Abs(Flag)).Top + WriterText(Abs(Flag)).Height <= 0 Then WriterText(Abs(Flag)).Top = Me.Height: Flag = Not Flag
End Sub

Private Sub WriterText_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    For i = 0 To 1
        If WriterText(i).Left < 0 Then WriterText(i).Left = 0: Exit Sub
        If WriterText(i).Left + WriterText(i).Width > Me.Width Then WriterText(i).Left = Me.Width - WriterText(i).Width: Exit Sub
        If KeyCode = 37 Then WriterText(i).Left = WriterText(i).Left - 25
        If KeyCode = 39 Then WriterText(i).Left = WriterText(i).Left + 25
    Next
End Sub
