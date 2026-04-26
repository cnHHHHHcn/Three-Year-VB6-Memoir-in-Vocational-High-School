VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   15630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   28500
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   15630
   ScaleWidth      =   28500
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   13680
      Top             =   7560
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "判题"
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出题"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label12 
      Height          =   495
      Left            =   18720
      TabIndex        =   11
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   495
      Left            =   18960
      TabIndex        =   10
      Top             =   11400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   9
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   8
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "错误："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   7
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "正确："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   6
      Top             =   2760
      Width           =   1485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1800
      Width           =   255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Label2.Caption = Int(Rnd * 101)
Label4.Caption = Int(Rnd * 101)
Command1.Enabled = False
Command2.Enabled = True
a = Int(Rnd * 4)
Select Case a
Case 0
Label3.Caption = "+"
Case 1
Label3.Caption = "-"
Case 2
Label3.Caption = "×"
Case 3
Label3.Caption = "÷"
End Select
End Sub

Private Sub Command2_Click()
Select Case Label3.Caption
Case "+"
If Val(Text1.Text) = Val(Label2.Caption) + Val(Label4.Caption) Then
Label6.Caption = "对"
Label9.Caption = Label9.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
Else
Label6.Caption = "错"
Label10.Caption = Label10.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
End If

Case "-"

If Val(Text1.Text) = Val(Label2.Caption) - Val(Label4.Caption) Then
Label6.Caption = "对"
Label9.Caption = Label9.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
Else
Label6.Caption = "错"
Label10.Caption = Label10.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
End If

Case "×"

If Val(Text1.Text) = Val(Label2.Caption) * Val(Label4.Caption) Then
Label6.Caption = "对"
Label9.Caption = Label9.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
Else
Label6.Caption = "错"
Label10.Caption = Label10.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
End If

Case "÷"

If Val(Text1.Text) = Round(Val(Label2.Caption) / Val(Label4.Caption), 2) Then
Label6.Caption = "对"
Label9.Caption = Label9.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
Else
Label6.Caption = "错"
Label10.Caption = Label10.Caption + 1
Command1.Enabled = True
Command2.Enabled = False
Text1.Text = ""
End If

End Select

If Label6.Caption = "对" Then GoTo a
If Val(Label10.Caption) Mod 5 = 0 Then
MsgBox "这都做不对，那在加一道题吧。"
Form1.Text1.Text = Form1.Text1.Text + 1
End If
a:
If Form1.Text1.Text = Label9.Caption Then End
End Sub

Private Sub Form_Load()
On Error Resume Next
Command2.Enabled = False
Label11.Caption = Form3.Width
Label12.Caption = Form3.Height
End Sub

Private Sub Form_Resize()
On Error Resume Next
a = Label11.Caption: b = Label12.Caption
If 0 < Form3.Width And Form3.Width < a Or 0 < Form3.Height And Form3.Height < b Then MsgBox "你想干嘛？", , "系统"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
a = Label11.Caption: b = Label12.Caption
Form3.Top = 0: Form3.Width = a
Form3.Left = 0: Form3.Height = b
End Sub
