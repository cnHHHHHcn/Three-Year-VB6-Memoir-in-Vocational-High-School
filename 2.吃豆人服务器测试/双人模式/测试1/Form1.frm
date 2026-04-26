VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   ScaleHeight     =   10890
   ScaleWidth      =   15525
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   12240
      Top             =   5640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 87, 119
    Command1.Top = Command1.Top - 50
    
Case 65, 97
    Command1.Left = Command1.Left - 50
    
Case 83, 115
    Command1.Top = Command1.Top + 50
    
Case 68, 100
    Command1.Left = Command1.Left + 50
    
End Select
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 87, 119
    Command2.Top = Command2.Top - 50
    
Case 65, 97
    Command2.Left = Command2.Left - 50
    
Case 83, 115
    Command2.Top = Command2.Top + 50
    
Case 68, 100
    Command2.Left = Command2.Left + 50
    
End Select
End Sub

Private Sub Timer1_Timer()
Dim zb(2)
i = 0
Open "\\XSH47\Users\2.txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, zb(i)
        i = i + 1
    Loop
Close
Command2.Left = zb(0)
Command2.Top = zb(1)
Open "\\XSH47\Users\1.txt" For Output As #2
    Print #2, Command1.Left
    Print #2, Command1.Top
Close
End Sub
