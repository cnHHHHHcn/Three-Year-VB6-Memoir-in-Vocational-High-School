VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "对话框标题"
   ClientHeight    =   2025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label TransferInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Local_Value As Boolean

Public Property Get Value() As Boolean
    Value = Local_Value
End Property

Private Sub CancelButton_Click()
    Local_Value = False
End Sub

Public Sub Msg(TextCaption As String, Text As String)
    Dialog.Caption = Text
    TransferInfo.Caption = TextCaption
    Dialog.Show
End Sub

Private Sub OKButton_Click()
    Local_Value = True
End Sub
