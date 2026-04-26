VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4380
   LinkTopic       =   "Form6"
   ScaleHeight     =   2655
   ScaleWidth      =   4380
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "关闭"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发送"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = vbCrLf & Form2.Label3.Caption & "：" & vbCrLf & Text2.Text
Text1.Text = Text2.Text
Text2.Text = ""
Open Form2.Label8.Caption For Append As 3
   Print #3, a
Close
End Sub

Private Sub Command2_Click()
Text2.Text = ""
End Sub

Private Sub Command3_Click()
Form6.Hide
End Sub
