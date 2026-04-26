VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   9810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form5"
   ScaleHeight     =   9810
   ScaleWidth      =   10020
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "·µ»Ø"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   9240
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form1.Show: Form5.Hide
End Sub

