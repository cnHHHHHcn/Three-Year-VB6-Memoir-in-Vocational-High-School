VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton TextSet 
      Caption         =   "确定"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   4215
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Label1"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then FormProc.FormMove Form2.hwnd
End Sub

Private Sub TextSet_Click()
    BeforeText = Form1.ClipBoardList.Text
    For i = 0 To Form1.ClipBoardList.ListCount - 1
        If Form1.ClipBoardList.List(i) = BeforeText Then
            With Form1.ClipBoardList
                .RemoveItem i
                .AddItem Text.Text, i
            End With
        End If
    Next
    Text.Text = ""
    Form2.Hide
End Sub
