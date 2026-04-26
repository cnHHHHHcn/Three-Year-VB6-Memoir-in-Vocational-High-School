VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8430
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Path 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   6255
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1830
      Left            =   0
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.ListBox DNS_list 
      Appearance      =   0  'Flat
      Height          =   3090
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申请区"
      Height          =   180
      Left            =   5760
      TabIndex        =   4
      Top             =   2640
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "路径："
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Server As New Server
Dim Public_Code As New Public_Code

Private Sub Text1_Change()
On Error Resume Next
File1.Path = Path.Text
End Sub

Private Sub File1_Click()
On Error Resume Next
DNS_list.Clear
Open File1.Path & "\" & File1.FileName For Input As #1
    Do Until EOF(1)
        Line Input #1, strDNS
        DNS_list.AddItem strDNS
    Loop
Close
End Sub

Private Sub Path_Change()

End Sub
