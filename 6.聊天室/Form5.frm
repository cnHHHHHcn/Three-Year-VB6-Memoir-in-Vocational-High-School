VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "被@设置"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   ScaleHeight     =   3825
   ScaleWidth      =   5790
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Text            =   "5"
      ToolTipText     =   "最少为3s，最多为10s"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   2
      Left            =   4200
      TabIndex        =   10
      Text            =   "230"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   1
      Left            =   4200
      TabIndex        =   9
      Text            =   "230"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   270
      Index           =   0
      Left            =   4200
      TabIndex        =   8
      Text            =   "230"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   1080
      Max             =   230
      TabIndex        =   2
      Top             =   1680
      Value           =   230
      Width           =   3000
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   1080
      Max             =   230
      TabIndex        =   1
      Top             =   1200
      Value           =   230
      Width           =   3000
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1080
      Max             =   230
      TabIndex        =   0
      Top             =   720
      Value           =   230
      Width           =   3000
   End
   Begin VB.Label Label4 
      Caption         =   "5"
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "持续时间："
      Height          =   180
      Left            =   600
      TabIndex        =   11
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "蓝："
      Height          =   180
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "绿："
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "红："
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "被@提醒屏幕颜色"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Hide
End Sub

Private Sub Command2_Click()
If Not (3 <= Val(Text2.Text) And Val(Text2.Text) <= 10) Then Text2.Text = 3
If Val(Text2.Text) Mod 2 = 0 Then Label4.Caption = Text2.Text + 1 Else Label4.Caption = Text2.Text
End Sub

Private Sub Form_Load()
Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll1_Change()
Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1(0).Text = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1(0).Text = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1(1).Text = HScroll2.Value
End Sub

Private Sub HScroll2_Scroll()
Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1(1).Text = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1(2).Text = HScroll3.Value
End Sub

Private Sub HScroll3_Scroll()
Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text1(2).Text = HScroll3.Value
End Sub

Private Sub Text1_Change(Index As Integer)
On Error GoTo a
GoTo b
a:
Select Case Index
Case 0
    Text1(0).Text = 230
Case 1
    Text1(1).Text = 230
Case 2
    Text1(2).Text = 230
End Select
b:
Select Case Index
Case 0
    HScroll1.Value = Text1(0).Text
    Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Case 1
    HScroll2.Value = Text1(1).Text
    Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Case 2
    HScroll3.Value = Text1(2).Text
    Form5.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Select

End Sub

Private Sub Text2_Change()
Command2.Enabled = True
End Sub
