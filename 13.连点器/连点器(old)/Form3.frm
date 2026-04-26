VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   1605
   ClientLeft      =   16800
   ClientTop       =   645
   ClientWidth     =   4260
   LinkTopic       =   "Form3"
   ScaleHeight     =   1605
   ScaleWidth      =   4260
   Begin VB.Label Remine 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   165
   End
   Begin VB.Label MouseState 
      AutoSize        =   -1  'True
      Caption         =   "State："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label MouseX 
      AutoSize        =   -1  'True
      Caption         =   "MouseX："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label MouseY 
      AutoSize        =   -1  'True
      Caption         =   "MouseY："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label MouseButton 
      AutoSize        =   -1  'True
      Caption         =   "Button："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   1305
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Remine.Caption = "慢一点单击鼠标。" & Chr(10) & "鼠标移动、点击时才会录制。"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show: Form3.Hide: Form2.Hide
End Sub
