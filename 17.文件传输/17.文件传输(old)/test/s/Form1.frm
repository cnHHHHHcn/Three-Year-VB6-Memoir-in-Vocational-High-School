VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3840
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a() As Byte

Private Sub Command1_Click()
    Winsock1.Close
    Winsock1.RemotePort = 10001
    Winsock1.RemoteHost = "admin015"
    Winsock1.Connect "admin015", 10001
End Sub

Private Sub Form_Load()
Open "C:\1.exe" For Binary As #2
End Sub


Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'检查控件的 State 属性是否为关闭的。
'如果不是，
'在接受新的连接之前先关闭此连接。
If Winsock1.State <> sckClosed Then _
Winsock1.Close
'接受具有 requestID 参数的
'连接。
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Static SizeCount As Long
    Static AcceptCount As Long
    DoEvents
    AcceptCount = AcceptCount + 1
    Debug.Print AcceptCount
    If AcceptCount = 1 Then SizeCount = 1
    If AcceptCount Mod 128 = 0 Then
        Close
        SizeCount = FileLen("C:\红蜘蛛.zip") '- 2
        Open "C:\红蜘蛛.zip" For Binary As #2
    End If
    ReDim a(1 To bytesTotal)
    Winsock1.GetData a, vbByte + vbArray, bytesTotal
    Put #2, SizeCount, a
    SizeCount = SizeCount + bytesTotal
    If bytesTotal < 8192 Then Close
End Sub

''下面是接收端的程序:
'
'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Dim bytData() As Byte
'Dim lLenFile As Long
'Dim f
'f = FreeFile
'Open strFileName For Binary As #f ''strFileName是文件名
'    lLenFile = LOF(f)
'    ReDim bytData(1 To bytesTotal)
'    Winsock1.GetData bytData
'    If lLenFile = 0 Then ''lLenFile=0表示是第一次打开文件，这里有个问题，就是
'        ''如果如果该文件存在的话，就会出错，应该在打开前检查文件是否存在。（这里我省略了）
'        Put #f, 1, bytData
'    Else
'        Put #f, lLenFile + 1, bytData
'    End If
'Close #f
'End Sub

