VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   3120
   ClientTop       =   3840
   ClientWidth     =   14535
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   18
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   14535
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9360
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "锁定"
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      Height          =   480
      Left            =   7080
      TabIndex        =   8
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "检查"
      Height          =   615
      Left            =   11160
      TabIndex        =   6
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "说明"
      Height          =   495
      Left            =   12600
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   615
      Left            =   11160
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   13095
   End
   Begin VB.Label Label4 
      Caption         =   "最大截取字符数："
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "设置按键："
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "输入文档路径:"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a() As String
'Private Declare Function dll1 Lib "f:\dll1.dll" (startfile As String, endfile As String)
Private Sub Command1_Click()
Dim b%
b = 1
Open "f:\cs.txt" For Output As #1
ReDim a(999999)
Print #1, Text2.Text & "," & 1 & "=" & alltrim(Mid(Text1.Text, 1, Text3.Text))
For i = 1 To 999999
    a(i) = alltrim(Mid(Text1.Text, i * 100, Text3.Text))
    If a(i) = "" Then
        ReDim Preserve a(i)
        Exit For
    End If
    Print #1, Text2.Text & "," & i + 1 & "="; a(i)
Next i
Close

Call filecopy("f:\cs.txt", "C:\Program Files\SogouWBInput\3.1.0.1751\phrases.ini")
MsgBox "成功！！", , "恭喜"
End Sub

Private Sub Command3_Click()
If Text2.Text = "" Or Text3.Text = "" Or Len(Text2.Text) < 4 Or Len(Text2.Text) > 300 Then
    MsgBox "错误请查看说明！！", 48, "致命错误"
    Command1.Visible = False
    Command3.Visible = True

    Else
        Command1.Visible = True
        Command3.Visible = False
        Text2.Enabled = False
        Text3.Enabled = False

End If

 '   MsgBox "截取不能为空！！", 48, "致命错误"

'    MsgBox "按键不能超过300字符！！", 48, "致命错误"



'按键不能小于4位字符！！


End Sub

Private Sub Command4_Click()
File1.Path = Text4.Text
Text4.Text = ""
End Sub

Private Sub File1_Click()
Text1.Text = ""
Open File1.Path & "\" & File1.FileName For Input As #1
    Do Until EOF(1)
        Line Input #1, ab
        Text1.Text = Text1.Text & ab
    Loop
Close
End Sub

Private Sub Form_Load()
Label1.BackStyle = 0
File1.Path = "F:\金山打字2006（原金山打字通）正式版\jsdz2006\Data\Chinese\E_General"
End Sub

Function alltrim(strings) As String
For i = 1 To Len(strings)
    strs = Mid(strings, i, 1)
    If Not Asc(strs) = 13 Then alltrim = alltrim & strs
Next i
End Function
Sub filecopy(startfile, endfile)

Open startfile For Input As #1
Open endfile For Append As #2
Do Until EOF(1)
    Line Input #1, strs
    Print #2, strs
Loop
Close #1, #2

End Sub

