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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '&H5578
    Dim pid As Long, Addr As String
    Dim Buffer() As Byte
    pid = 8492
    Buffer = CheatMemory.ReadMemory(pid, &H755E0C, 4)
    Addr = ""
    For i = LBound(Buffer) To UBound(Buffer)
        Addr = Hex(Buffer(i)) + Addr
    Next
    Buffer = CheatMemory.ReadMemory(pid, CLng("&H" + Addr) + &H868, 4)
    Addr = ""
    For i = LBound(Buffer) To UBound(Buffer)
        Addr = Hex(Buffer(i)) + Addr
    Next
    Buffer = CheatMemory.ReadMemory(pid, CLng("&H" + Addr) + &H5578, 4)
    Addr1 = ""
    For i = LBound(Buffer) To UBound(Buffer)
        Addr1 = Hex(Buffer(i)) + Addr1
    Next
    Debug.Print CLng("&H" + Addr1)
    Value = 1000
    For i = 0 To 3
        Buffer(i) = Value And 255
        Value = Fix(Value \ 255)
    Next
    CheatMemory.WriteMemory pid, CLng("&H" + Addr) + &H5578, Buffer
End Sub
