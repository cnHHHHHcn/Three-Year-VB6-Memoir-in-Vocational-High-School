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
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2760
      Top             =   1440
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB中使用WinSock控件传送文件 (2010-07-23 08:09:37)
'标签:  杂谈
'分类:  VB
'传送文件对于网络编程来说是基本的功能，比如远程控制软件。
'在编制一个软件时，我从网上下了很多传文件的程序，这些程序提供的传文件功能根本就不能用。传文本还可以，传二进制文件根本就不行。因此，作为一个基本的功能模块，有必要单独介绍一下。
'首先，在VB中要传送字符串，你可以这样写：


'Dim strData As String
'strData = "Test"
'Winsock1.SendData strData
   
'但是如果你传送的二进制文件，你还能用String变量来存放吗？从理论上分析是不行的，我也做了实验，确实是不行的。文件虽然可以传，但是接受的文件和发送的不一样，原因可能是二进制文件里可以有任何"字符"，但是不是所有的字符都可以放在String变量里。
'除了String类型的变量，VB中其他类型的变量都只有几个字节长，难道一次只能发几个字节吗？那样岂不是要累死机器了！其实，情况没有那么悲观，我们完全可以使用数组来解决这个问题，就是使用byte数组。把要传送的文件都读到数组里，然后发送出去。

'程序如下:
'FileName 为要传送的文件名，WinS为发送文件的WinSock控件。这是一个发送端的程序。

Const iMax = 65535

Public Sub SendFile1(FileName As String, WinS As Winsock)

Dim FreeF As Integer ''空闲的文件号
Dim LenFile As Long ''文件的长度
Dim bytData() As Byte ''存放数据的数组
FreeF = FreeFile ''获得空闲的文件号
Open FileName For Binary As #FreeFile ''打开文件
    DoEvents
    LenFile = LOF(FreeFile) ''获得文件长度
    ReDim bytData(1 To LenFile) ''根据文件长度重新定义数组大小
    Get #FreeFile, , bytData ''把文件读入到数组里
Close #FreeFile ''关闭文件
WinS.SendData bytData ''发送数据
End Sub

Private Sub Command1_Click()
    SendFile2 "E:\office2010破解版.rar", Winsock1
End Sub

Private Sub Form_Load()
    Winsock1.Close
    Winsock1.LocalPort = 10001
    Winsock1.Listen
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Winsock1.State
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

'接受端的程序如下:
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim bytData() As Byte
Dim f
f = FreeFile
Open strFileName For Binary As #ff
    ReDim bytData(1 To bytesTotal)
    Winsock1.GetData bytData
    Put #f, i, bytData
    i = i + bytesTotal ''保证每次写都是在文件的末尾, i是个全局变量
Close #f
End Sub

'这里有两个需要注意的地方，ReDim Preserve bytData(1 To LenFile)，下标是从1开始的，如果你写成ReDim bytData( LenFile)，下标就是从0开始了，
'数组就有LenFile+1长了。LenFile = LOF(FreeFile)中的LOF是获得文件长度的函数，是VB里带的，我见过很多例子用API，或者循环的读直到末尾来获取文件长度，这样都是很麻烦的，使用LOF函数就可以了。
'这样的程序，即可以传送文本文件，也可以传送二进制文件。但是你有没有发现这个程序的问题呢？如果我要传送一个50M的文件呢？系统可以为bytData分配50M的内存空间吗？
'于是笔者拿一个50M的文件做实验吧，接收到的文件和原来的文件不一样，比原来的大。问题出在那呢？
'首先，根据文件大小重新定义bytData数组的大小本身就有问题，系统是不可能无限制的给数组分配空间的，即使可以，也会造成系统响应变慢。在传50M文件的时候，系统就跟死机了一样。那么怎么解决这个问题呢，一个自然的想法就是把数据分段传送。

'程序如下:
'发送程序 , iPos是个全局变量, 初始值为0?
'这个变量保存着当前数据的位置?

Public Sub SendFile2(FileName As String, WinS As Winsock)
    Dim FreeF As Integer ''空闲的文件号
    Dim LenFile As Long ''文件的长度
    Dim bytData() As Byte ''存放数据的数组
    FreeF = FreeFile ''获得空闲的文件号
    Open FileName For Binary As #FreeF ''打开文件
        DoEvents
        LenFile = LOF(FreeF) ''获得文件长度
        If LenFile <= iMax Then ''如果要发送的文件小于数据块大小，直接发送
            ReDim bytData(1 To LenFile) ''根据文件长度重新定义数组大小
            Get #FreeF, , bytData
            ''把文件读入到数组里
            Close #FreeF ''关闭文件
            WinS.SendData bytData ''发送数据
            Exit Sub
        End If
    ''文件大于数据块大小，进行分块发送
    Do Until (iPos >= (LenFile - iMax)) ''发送整块数据的循环
        ReDim bytData(1 To iMax)
        Get #FreeF, iPos + 1, bytData
        WinS.SendData bytData
        iPos = iPos + iMax ''移动iPos，使它指向下来要读的数据
    Loop
    ''这里要注意的是，必须检查文件有没有剩下的数据，如果文件大小正好等于数据块大小的
    '' 整数倍，那么就没有剩下的数据了
    ReDim bytData(1 To LenFile - iPos) ''发送剩下的不够一个数据块的数据
    Get #FreeF, iPos + 1, bytData
    WinS.SendData bytData
    Close #FreeF
End Sub

'下面是接收端的程序:

Private Sub Winsock12_DataArrival(ByVal bytesTotal As Long)
Dim bytData() As Byte
Dim lLenFile As Long
Dim f
f = FreeFile
Open strFileName For Binary As #f ''strFileName是文件名
    lLenFile = LOF(f)
    ReDim bytData(1 To bytesTotal)
    Winsock1.GetData bytData
    If lLenFile = 0 Then ''lLenFile=0表示是第一次打开文件，这里有个问题，就是
        ''如果如果该文件存在的话，就会出错，应该在打开前检查文件是否存在。（这里我省略了）
        Put #f, 1, bytData
    Else
        Put #f, lLenFile + 1, bytData
    End If
Close #f
End Sub
