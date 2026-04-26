VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.FileListBox File1 
      Height          =   990
      Left            =   840
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Open File1.Path & "\" & File1.FileName For Input As #1
        Dim Str As String, Temp As String
        Dim Str_Array() As String
        Do Until EOF(1)
            Line Input #1, Str
            Str_Array = Split(Str, " ")
            For i = LBound(Str_Array) To UBound(Str_Array)
                Temp = Temp & ":" & Str_Array(i)
            Next
        Loop
        MsgBox Temp
    Close #1
    Open File1.Path & "\" & File1.FileName For Output As #1
        Print #1, Temp
    Close
End Sub

Private Sub Form_Load()
    File1.Path = "D:\工程\Visual-basic-6.0\21.极域破解"
End Sub
