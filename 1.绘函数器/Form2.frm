VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   Caption         =   "绘函数器"
   ClientHeight    =   4725
   ClientLeft      =   6960
   ClientTop       =   4410
   ClientWidth     =   6735
   LinkTopic       =   "Form2"
   ScaleHeight     =   4725
   ScaleWidth      =   6735
   Begin VB.OLE OLE2 
      Class           =   "Package"
      Height          =   855
      Left            =   480
      OleObjectBlob   =   "Form2.frx":0000
      SourceDoc       =   "D:\Backup\Documents\My Videos\关静音教程.mp4"
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   855
      Left            =   360
      OleObjectBlob   =   "Form2.frx":123BDD7
      SourceDoc       =   "D:\Backup\Documents\My Videos\绘函数器讲解.mp4"
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   615
      Left            =   12240
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   1085
   End
   Begin VB.Menu czhs 
      Caption         =   "初中函数"
      Begin VB.Menu yechs 
         Caption         =   "一，二次函数"
         Shortcut        =   ^Y
      End
      Begin VB.Menu fblhs 
         Caption         =   "反比例函数"
         Shortcut        =   ^F
      End
      Begin VB.Menu czzhhs 
         Caption         =   "初中综合函数"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu gzhs 
      Caption         =   "高中函数"
      Begin VB.Menu zshs 
         Caption         =   "指数函数"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub czzhhs_Click()
Form2.Hide: Form4.Show
End Sub

Private Sub fblhs_Click()
Form3.Show: Form2.Hide
End Sub

Private Sub czhs_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\睡觉.mp3"
End Sub

Private Sub yechs_Click()
Form1.Show: Form2.Hide
End Sub

Private Sub zshs_Click()
Form5.Show: Form2.Hide
End Sub
