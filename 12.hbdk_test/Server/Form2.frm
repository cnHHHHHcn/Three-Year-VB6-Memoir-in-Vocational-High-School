VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   5880
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.FileListBox File 
      Height          =   2610
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.DirListBox Dir 
      Height          =   2610
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.DriveListBox Drive 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form1.List_Jion_File.Text = Text1.Text
End Sub

Private Sub Dir_Change()
If InStr(Len(Dir.Path) - 1, Dir.Path, "\") Then Text1.Text = Dir.Path Else Text1.Text = Dir.Path & "\"
File.Path = Text1.Text
End Sub

Private Sub Drive_Change()
Text1.Text = GetDOS.GetDrive(Drive.Drive) & "\"
Dir.Path = Text1.Text
End Sub

Private Sub File_Click()
flag = GetDOS.GetFileName(Text1.Text)
If (Not CBool(InStr(1, flag, "."))) Or flag <> File.FileName Then Text1.Text = GetDOS.GetPath(Text1.Text) & File.FileName
End Sub

Private Sub Form_Load()
Text1.Text = Dir.Path & "\"
End Sub
