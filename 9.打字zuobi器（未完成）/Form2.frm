VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form2"
   ScaleHeight     =   5400
   ScaleWidth      =   6990
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   4935
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form2.frx":0000
      Top             =   0
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6240
      Top             =   3960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b%
Dim a%
Dim num%

Private Sub Command1_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Command1.Enabled = False
num = 20
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
num = num - 1
a = a + 1
Command1.Caption = "了解(" & num & ")"
If a = 20 Then
Timer1.Enabled = False
Command1.Enabled = True
End If
End Sub
