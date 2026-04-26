VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11340
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7710
   ScaleWidth      =   11340
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   7440
      Top             =   5040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清零"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "减"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "你的智商："
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
      Left            =   2040
      TabIndex        =   0
      Top             =   1680
      Width           =   2475
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = Text1.Text + 1
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text - 1
End Sub

Private Sub Command3_Click()
Text1.Text = 0
Command3.Visible = False
End Sub

Private Sub Text1_Change()
Set a = CreateObject("wscript.shell")
Select Case Val(Text1.Text)
Case 2
MsgBox "你确定你有那么高的智商？", , "系统"
Case 5
MsgBox "系统：你这智商智商不对啊？！", 48, "系统"
Case 10
a.run "1.vbs"
a.run "完美的童年.exe"
End
Case -10
MsgBox "系统：不对吧？应该是……等一下，我给你调调", , "系统"
a.run "240.vbs"
Timer1.Interval = 2000
End Select
End Sub

Private Sub Timer1_Timer()
MsgBox "系统：好了，这才是你的智商。", 64, "系统"
MsgBox "按钮2：我你*&，$%……敲了我240下！我真的$%……", , "按钮2（减）"
Command1.Enabled = True
MsgBox "按钮1：发生肾磨石了？！", 32, "按钮1（加）"
MsgBox "系统：按钮2禁止发言", 64, "系统"
MsgBox "按钮2:我你……（禁止发言中）", , "按钮2（减）"
Command2.Enabled = False
MsgBox "按钮2:……（愤怒中说不出话来）", , "按钮2（减）"
MsgBox "按钮1:哈哈哈！以前叫你少说几句非不听，遭报应了吧。", , "按钮2（减）"
MsgBox "按钮2:……（愤怒中说不出话来）", , "按钮2（减）"
Command3.Visible = True
Timer1.Interval = 0
End Sub
