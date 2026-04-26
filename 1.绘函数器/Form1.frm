VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "一二次函数"
   ClientHeight    =   12675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   12675
   ScaleWidth      =   21360
   Begin VB.Timer Timer3 
      Left            =   14520
      Top             =   6840
   End
   Begin VB.Timer Timer2 
      Left            =   18480
      Top             =   3840
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   16320
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   62
      Top             =   3240
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   18120
      Top             =   2760
   End
   Begin VB.CommandButton Command7 
      Caption         =   "返回"
      Height          =   495
      Left            =   19320
      TabIndex        =   60
      Top             =   11520
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FFFF&
      Caption         =   "y轴减"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   38
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "y轴加"
      Height          =   495
      Left            =   8520
      TabIndex        =   37
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "X轴加"
      Height          =   495
      Left            =   6360
      TabIndex        =   30
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "X轴减"
      Height          =   495
      Left            =   1680
      TabIndex        =   29
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   11400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   11400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   11400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "绘函数图像并计算"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "绘坐标系"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   9960
      Width           =   1215
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer3 
      Height          =   615
      Left            =   13800
      TabIndex        =   64
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
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
      _cx             =   2355
      _cy             =   1085
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
      Height          =   5415
      Left            =   14280
      TabIndex        =   63
      Top             =   3600
      Visible         =   0   'False
      Width           =   5295
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
      _cx             =   9340
      _cy             =   9551
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   14040
      TabIndex        =   61
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
      URL             =   "fdsg"
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
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2566
      _cy             =   873
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   20
      Left            =   10800
      TabIndex        =   59
      Top             =   10920
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   19
      Left            =   10800
      TabIndex        =   58
      Top             =   10440
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   18
      Left            =   10800
      TabIndex        =   57
      Top             =   9960
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   17
      Left            =   10800
      TabIndex        =   56
      Top             =   9480
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   10800
      TabIndex        =   55
      Top             =   9000
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   10800
      TabIndex        =   54
      Top             =   8520
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   10800
      TabIndex        =   53
      Top             =   8040
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   10800
      TabIndex        =   52
      Top             =   7560
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   10800
      TabIndex        =   51
      Top             =   7080
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   10800
      TabIndex        =   50
      Top             =   6600
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   10800
      TabIndex        =   49
      Top             =   6120
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   10800
      TabIndex        =   48
      Top             =   5640
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   10800
      TabIndex        =   47
      Top             =   5160
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   10800
      TabIndex        =   46
      Top             =   4680
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   10800
      TabIndex        =   45
      Top             =   4200
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   10800
      TabIndex        =   44
      Top             =   3720
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   10800
      TabIndex        =   43
      Top             =   3240
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   10800
      TabIndex        =   42
      Top             =   2760
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   10920
      TabIndex        =   41
      Top             =   2280
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   10800
      TabIndex        =   40
      Top             =   1800
      Width           =   165
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   10800
      TabIndex        =   39
      Top             =   1320
      Width           =   165
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "y轴↑"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   36
      Top             =   8520
      Width           =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8520
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8520
      TabIndex        =   34
      Top             =   8040
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "X轴→"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8640
      TabIndex        =   33
      Top             =   4200
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7920
      TabIndex        =   32
      Top             =   8520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "-10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   31
      Top             =   8520
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   22
      Left            =   12240
      TabIndex        =   28
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   21
      Left            =   11160
      TabIndex        =   27
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   20
      Left            =   11760
      TabIndex        =   26
      Top             =   10920
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   19
      Left            =   11760
      TabIndex        =   25
      Top             =   10440
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   18
      Left            =   11760
      TabIndex        =   24
      Top             =   9960
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   17
      Left            =   11760
      TabIndex        =   23
      Top             =   9480
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   11760
      TabIndex        =   22
      Top             =   9000
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   11760
      TabIndex        =   21
      Top             =   8520
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   11760
      TabIndex        =   20
      Top             =   8040
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   11760
      TabIndex        =   19
      Top             =   7560
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   11760
      TabIndex        =   18
      Top             =   7080
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   11760
      TabIndex        =   17
      Top             =   6600
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   11760
      TabIndex        =   16
      Top             =   6120
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   11760
      TabIndex        =   15
      Top             =   5640
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   11760
      TabIndex        =   14
      Top             =   5160
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   11760
      TabIndex        =   13
      Top             =   4680
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   11760
      TabIndex        =   12
      Top             =   4200
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   11760
      TabIndex        =   11
      Top             =   3720
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   11760
      TabIndex        =   10
      Top             =   3240
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   11760
      TabIndex        =   9
      Top             =   2760
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   11760
      TabIndex        =   8
      Top             =   2280
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   11760
      TabIndex        =   7
      Top             =   1800
      Width           =   165
   End
   Begin VB.Line Line2 
      Index           =   24
      X1              =   11640
      X2              =   11640
      Y1              =   720
      Y2              =   11280
   End
   Begin VB.Line Line2 
      Index           =   22
      X1              =   10800
      X2              =   13080
      Y1              =   11280
      Y2              =   11280
   End
   Begin VB.Line Line2 
      Index           =   21
      X1              =   10800
      X2              =   13080
      Y1              =   10800
      Y2              =   10800
   End
   Begin VB.Line Line2 
      Index           =   20
      X1              =   10800
      X2              =   13080
      Y1              =   10320
      Y2              =   10320
   End
   Begin VB.Line Line2 
      Index           =   19
      X1              =   10800
      X2              =   13080
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Line Line2 
      Index           =   18
      X1              =   10800
      X2              =   13080
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line2 
      Index           =   17
      X1              =   10800
      X2              =   13080
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line Line2 
      Index           =   16
      X1              =   10800
      X2              =   13080
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line2 
      Index           =   15
      X1              =   10800
      X2              =   13080
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line2 
      Index           =   14
      X1              =   10800
      X2              =   13080
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line2 
      Index           =   13
      X1              =   10800
      X2              =   13080
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line2 
      Index           =   12
      X1              =   10800
      X2              =   13080
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line2 
      Index           =   11
      X1              =   10800
      X2              =   13080
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line2 
      Index           =   10
      X1              =   10800
      X2              =   13080
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line2 
      Index           =   9
      X1              =   10800
      X2              =   13080
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   10800
      X2              =   13080
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   10800
      X2              =   13080
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   10800
      X2              =   13080
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   10800
      X2              =   13080
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   10800
      X2              =   13080
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   10800
      X2              =   13080
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   10800
      X2              =   13080
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   10800
      X2              =   13080
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   10800
      X2              =   13080
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   11760
      TabIndex        =   6
      Top             =   1320
      Width           =   165
   End
   Begin VB.Label Label1 
      Caption         =   "y=       X^2+       X+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   11400
      Width           =   4575
   End
   Begin VB.Line Line1 
      Index           =   200
      X1              =   360
      X2              =   1560
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Line Line1 
      Index           =   199
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   198
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   197
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   196
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   195
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   194
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   193
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   192
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   191
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   190
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   189
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   188
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   187
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   186
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   185
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   184
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   183
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   182
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   181
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   180
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   179
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   178
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   177
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   176
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   175
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   174
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   173
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   172
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   171
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   170
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   169
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   168
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   167
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   166
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   165
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   164
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   163
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   162
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   161
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   160
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   159
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   158
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   157
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   156
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   155
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   154
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   153
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   152
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   151
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   150
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   149
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   148
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   147
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   146
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   145
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   144
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   143
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   142
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   141
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   140
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   139
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   138
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   137
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   136
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   135
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   134
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   133
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   132
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   131
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   130
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   129
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   128
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   127
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   126
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   125
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   124
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   123
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   122
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   121
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   120
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   119
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   118
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   117
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   116
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   115
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   114
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   113
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   112
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   111
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   110
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   109
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   108
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   107
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   106
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   105
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   104
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   103
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   102
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   101
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   100
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   99
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   98
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   97
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   96
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   95
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   94
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   93
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   92
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   91
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   90
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   89
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   88
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   87
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   86
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   85
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   84
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   83
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   82
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   81
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   80
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   79
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   78
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   77
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   76
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   75
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   74
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   73
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   72
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   71
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   70
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   69
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   68
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   67
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   66
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   65
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   64
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   63
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   62
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   61
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   60
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   59
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   58
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   57
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   56
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   55
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   54
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   53
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   52
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   51
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   50
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   49
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   48
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   47
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   46
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   45
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   44
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   43
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   42
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   41
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   40
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   39
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   38
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   37
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   36
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   35
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   34
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   33
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   32
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   29
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   28
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   26
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   25
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      Index           =   0
      X1              =   720
      X2              =   1920
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   30
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   29
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   28
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   27
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   26
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   25
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   24
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   23
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   22
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   21
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   20
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   19
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   18
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   17
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   16
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   15
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   14
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   13
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   12
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   11
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   10
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   9
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   8
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   7
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   6
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   5
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   4
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape2 
      Height          =   195
      Index           =   0
      Left            =   360
      Top             =   600
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   410
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   409
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   408
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   407
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   406
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   405
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   404
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   403
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   402
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   401
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   400
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   399
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   398
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   397
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   396
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   395
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   394
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   393
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   392
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   391
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   390
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   389
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   388
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   387
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   386
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   385
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   384
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   383
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   382
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   381
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   380
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   379
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   378
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   377
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   376
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   375
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   374
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   373
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   372
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   371
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   370
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   369
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   368
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   367
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   366
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   365
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   364
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   363
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   362
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   361
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   360
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   359
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   358
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   357
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   356
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   355
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   354
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   353
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   352
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   351
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   350
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   349
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   348
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   347
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   346
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   345
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   344
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   343
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   342
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   341
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   340
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   339
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   338
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   337
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   336
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   335
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   334
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   333
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   332
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   331
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   330
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   329
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   328
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   327
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   326
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   325
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   324
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   323
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   322
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   321
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   320
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   319
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   318
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   317
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   316
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   315
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   314
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   313
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   312
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   311
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   310
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   309
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   308
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   307
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   306
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   305
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   304
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   303
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   302
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   301
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   300
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   299
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   298
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   297
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   296
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   295
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   294
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   293
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   292
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   291
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   290
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   289
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   288
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   287
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   286
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   285
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   284
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   283
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   282
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   281
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   280
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   279
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   278
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   277
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   276
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   275
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   274
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   273
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   272
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   271
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   270
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   269
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   268
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   267
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   266
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   265
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   264
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   263
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   262
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   261
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   260
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   259
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   258
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   257
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   256
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   255
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   254
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   253
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   252
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   251
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   250
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   249
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   248
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   247
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   246
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   245
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   244
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   243
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   242
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   241
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   240
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   239
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   238
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   237
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   236
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   235
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   234
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   233
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   232
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   231
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   230
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   229
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   228
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   227
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   226
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   225
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   224
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   223
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   222
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   221
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   220
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   219
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   218
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   217
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   216
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   215
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   214
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   213
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   212
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   211
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   210
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   209
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   208
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   207
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   206
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   205
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   204
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   203
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   202
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   201
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   200
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   199
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   198
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   197
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   196
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   195
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   194
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   193
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   192
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   191
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   190
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   189
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   188
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   187
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   186
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   185
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   184
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   183
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   182
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   181
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   180
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   179
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   178
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   177
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   176
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   175
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   174
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   173
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   172
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   171
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   170
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   169
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   168
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   167
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   166
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   165
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   164
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   163
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   162
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   161
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   160
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   159
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   158
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   157
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   156
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   155
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   154
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   153
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   152
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   151
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   150
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   149
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   148
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   147
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   146
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   145
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   144
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   143
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   142
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   141
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   140
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   139
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   138
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   137
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   136
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   135
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   134
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   133
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   132
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   131
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   130
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   129
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   128
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   127
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   126
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   125
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   124
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   123
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   122
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   121
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   120
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   119
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   118
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   117
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   116
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   115
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   114
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   113
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   112
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   111
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   110
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   109
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   108
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   107
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   106
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   105
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   104
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   103
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   102
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   101
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   100
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   99
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   98
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   97
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   96
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   95
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   94
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   93
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   92
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   91
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   90
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   89
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   88
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   87
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   86
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   85
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   84
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   83
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   82
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   81
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   80
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   79
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   78
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   77
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   76
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   75
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   74
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   73
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   72
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   71
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   70
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   69
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   68
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   67
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   66
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   65
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   64
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   63
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   62
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   61
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   60
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   59
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   58
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   57
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   56
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   55
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   54
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   53
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   52
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   51
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   50
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   49
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   48
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   47
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   46
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   45
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   44
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   43
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   42
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   41
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   40
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   39
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   38
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   37
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   36
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   35
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   34
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   33
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   32
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   31
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   30
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   29
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   28
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   27
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   26
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   25
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   24
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   23
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   22
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   21
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   20
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   19
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   18
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   17
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   16
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   15
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   14
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   13
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   12
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   11
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   10
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   9
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   8
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   7
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   6
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   5
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   4
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   195
      Index           =   0
      Left            =   480
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\睡觉.mp3"
For i = 1 To 20
For j = 1 To 10
Shape1(i * 10 + j).Shape = 1
Shape1(i * 10 + j).Width = 400
Shape1(i * 10 + j).Height = 400
Shape1(i * 10 + j).Left = 400 * i
Shape1(i * 10 + j).Top = 400 * j
Next j
Next i
For i = 1 To 20
For j = 1 To 10
Shape1(200 + i * 10 + j).Shape = 1
Shape1(200 + i * 10 + j).Width = 400
Shape1(200 + i * 10 + j).Height = 400
Shape1(200 + i * 10 + j).Left = 400 * i
Shape1(200 + i * 10 + j).Top = 400 * j + 4000
Next j
Next i
For i = 0 To 10
Shape1(i).Left = 300 + 400 * i
Shape1(i).Top = 4300
Shape1(i).Shape = 3
Shape1(i).BorderColor = vbRed
Next
For i = 0 To 9
Shape2(i).Left = 4700 + 400 * i
Shape2(i).Top = 4300
Shape2(i).Shape = 3
Shape2(i).BorderColor = vbRed
Next
For i = 10 To 30
Shape2(i).Left = 4300
Shape2(i).Top = 300 + 400 * (i - 10)
Shape2(i).Shape = 3
Shape2(i).BorderColor = vbRed
Next
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Label3.Visible = True
Label4.Visible = True
Label6.Visible = True
Label7.Visible = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\你干嘛.mp3"
For i = 0 To 10
Shape1(i).Left = 300 + 400 * i
Shape1(i).Top = 4300
Shape1(i).Shape = 3
Shape1(i).BorderColor = vbRed
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Left = 4700 + 400 * i
Shape2(i).Top = 4300
Shape2(i).Shape = 3
Shape2(i).BorderColor = vbRed
Shape2(i).Visible = True
Next
For i = 10 To 30
Shape2(i).Left = 4300
Shape2(i).Top = 300 + 400 * (i - 10)
Shape2(i).Shape = 3
Shape2(i).BorderColor = vbRed
Shape2(i).Visible = True
Next
Label4.Caption = 10
Label7.Caption = 10
Label3.Caption = -10
Label6.Caption = -10
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text)
For i = -10 To 10 Step 0.1
Line1(i * 10 + 100).X1 = i * 400 + 4400
Line1(i * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + 4400
Line1(i * 10 + 100).X2 = (i + 0.1) * 400 + 4400
Line1(i * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + 4400
Line1(i * 10 + 100).BorderColor = &H40C0&
Next
For j = -10 To 10
Label2(j + 10).Caption = a * j ^ 2 + b * j + c
Next
For i = -10 To 10
Label9(i + 10).Caption = i
Next
End Sub

Private Sub Command3_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\唱.mp3"
Label3.Caption = Label3.Caption - 1: Label4.Caption = Label4.Caption - 1
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): gh = Val(Label3.Caption): ty = Val(Label4.Caption): zx = Val(Label7.Caption)
For i = gh To ty Step 0.1
Line1((i - (ty - 10)) * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + (400 * zx)
Line1((i - (ty - 10)) * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + (400 * zx)
Line1((i - (ty - 10)) * 10 + 100).Y1 = Line1((i - (ty - 10)) * 10 + 100).Y1 + 400
Line1((i - (ty - 10)) * 10 + 100).Y2 = Line1((i - (ty - 10)) * 10 + 100).Y2 + 400
Next
For i = 0 To 20
Label9(i).Caption = gh + i
Next
For j = gh To ty
Label2((j - (ty - 10)) + 10).Caption = a * j ^ 2 + b * j + c
Next
For i = 10 To 30
Shape2(i).Left = Shape2(i).Left + 400
Next
If Label3.Caption < -20 And Label4.Caption < 0 Then
For df = 10 To 30
Shape2(df).Visible = False
Next
ElseIf -20 <= Label3.Caption And Label3.Caption <= 0 And 0 <= Label4.Caption And Label4.Caption <= 20 Then
For i = 10 To 30
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Command4_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\跳.mp3"
Label3.Caption = Label3.Caption + 1: Label4.Caption = Label4.Caption + 1
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): gh = Val(Label3.Caption): ty = Val(Label4.Caption): zx = Val(Label7.Caption)
For i = gh To ty Step 0.1
Line1((i - (ty - 10)) * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + (400 * zx + 800)
Line1((i - (ty - 10)) * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + (400 * zx + 800)
Line1((i - (ty - 10)) * 10 + 100).Y1 = Line1((i - (ty - 10)) * 10 + 100).Y1 - 400
Line1((i - (ty - 10)) * 10 + 100).Y2 = Line1((i - (ty - 10)) * 10 + 100).Y2 - 400
Next
For i = 0 To 20
Label9(i).Caption = gh + i
Next
For j = gh To ty
Label2((j - (ty - 10)) + 10).Caption = a * j ^ 2 + b * j + c
Next
For i = 10 To 30
Shape2(i).Left = Shape2(i).Left - 400
Next
If Label4.Caption > 20 And Label3.Caption > 0 Then
For i = 10 To 30
Shape2(i).Visible = False
Next
ElseIf 0 <= Label4.Caption And Label4.Caption <= 20 And -20 <= Label3.Caption And Label3.Caption <= 0 Then
For i = 10 To 30
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Command5_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\rap2.mp3"
Label7.Caption = Label7.Caption + 1: Label6.Caption = Label6.Caption + 1
For i = 0 To 10
Shape1(i).Top = Shape1(i).Top + 400
Next
For i = 0 To 9
Shape2(i).Top = Shape2(i).Top + 400
Next
For i = -10 To 10 Step 0.1
Line1(i * 10 + 100).Y1 = Line1(i * 10 + 100).Y1 + 400
Line1(i * 10 + 100).Y2 = Line1(i * 10 + 100).Y2 + 400
Next
If Label7.Caption > 20 And Label6.Caption > 0 Then
For i = 0 To 10
Shape1(i).Visible = False
Next
For i = 0 To 9
Shape2(i).Visible = False
Next
ElseIf -20 <= Label6.Caption And Label6.Caption <= 0 And 0 <= Label7.Caption And Label7.Caption <= 20 Then
For i = 0 To 10
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Command6_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\篮球.mp3"
Label7.Caption = Label7.Caption - 1: Label6.Caption = Label6.Caption - 1
For i = 0 To 10
Shape1(i).Top = Shape1(i).Top - 400
Next
For i = 0 To 9
Shape2(i).Top = Shape2(i).Top - 400
Next
For i = -10 To 10 Step 0.1
Line1(i * 10 + 100).Y1 = Line1(i * 10 + 100).Y1 - 400
Line1(i * 10 + 100).Y2 = Line1(i * 10 + 100).Y2 - 400
Next
If Label7.Caption < 0 And Label6.Caption < -20 Then
For i = 0 To 10
Shape1(i).Visible = False
Next
For i = 0 To 9
Shape2(i).Visible = False
Next
ElseIf -20 <= Label6.Caption And Label6.Caption <= 0 And 0 <= Label7.Caption And Label7.Caption <= 20 Then
For i = 0 To 10
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Command7_Click()
Form1.Hide: Form2.Show
End Sub

Private Sub Form_Load()
For i = -10 To 10
Label9(i + 10).Caption = i
Next
Form1.Hide: Form2.Show
End Sub

Private Sub Label1_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\你干嘛   哎呦.mp3"
fj:
ui = Val(InputBox("输入1则进入一次函数转换区；输入2则进入二次函数转化区；输入0退出"))
If ui = 1 Then
ao:
    '点斜式、一般式、截距式代码
    er = Val(InputBox("输入1为点斜式；输入2为一般式；输入3为截距式；输入0为退出"))
    If er = 1 Then
        d = -(Val(InputBox("y-y0=k(X-X0)" & Chr(10) & "y0:", "请输入数值")))
        e = Val(InputBox("y-y0=k(X-X0)" & Chr(10) & "k:", "请输入数值"))
        f = -(Val(InputBox("y-y0=k(X-X0)" & Chr(10) & "X0:", "请输入数值")))
        Text2.Text = e
        Text3.Text = e * f - d
    ElseIf er = 2 Then
        d = Val(InputBox("AX+By+C=0 (a*a+b*b≠0)" & Chr(10) & "A:", "请输入数值"))
        e = Val(InputBox("AX+By+C=0 (a*a+b*b≠0)" & Chr(10) & "B:", "请输入数值"))
        f = Val(InputBox("AX+By+C=0 (a*a+b*b≠0)" & Chr(10) & "C:", "请输入数值"))
        Text2.Text = -(d / e)
        Text3.Text = -(f / e)
    ElseIf er = 3 Then
        d = Val(InputBox("X/A+y/B=1" & Chr(10) & "A:", "请输入数值"))
        e = Val(InputBox("X/A+y/B=1" & Chr(10) & "B:", "请输入数值"))
        f = d * e
        Text2.Text = -(e / d)
        Text3.Text = f / d
    ElseIf er = 0 Then GoTo o:
    Else: GoTo ao
    End If
ElseIf ui = 2 Then
gh:
    '顶点式和交点式代码
    op = Val(InputBox("输入1为顶点式；输入2为交点式；输入0为退出", "请输入编码。", 1))
    If op = 1 Then
        a = Val(InputBox("y=a(X+h)^2+k" & Chr(10) & "a:", "请输入数值", 1))
        b = Val(InputBox("y=a(X+h)^2+k" & Chr(10) & "h:", "请输入数值"))
        c = Val(InputBox("y=a(X+h)^2+k" & Chr(10) & "k:", "请输入数值"))
        Text1.Text = a
        Text2.Text = a * (b * 2)
        Text3.Text = a * (b ^ 2) + c
    ElseIf op = 2 Then
        a = Val(InputBox("y=a(X+X1)(X+X2)" & Chr(10) & "a:", "请输入数值", 1))
        b = Val(InputBox("y=a(X+X1)(X+X2)" & Chr(10) & "X1:", "请输入数值"))
        c = Val(InputBox("y=a(X+X1)(X+X2)" & Chr(10) & "X2:", "请输入数值"))
        Text1.Text = a
        Text2.Text = a * (b + c)
        Text3.Text = a * (b * c)
    ElseIf op = 0 Then: GoTo o
    Else: GoTo gh:
    End If
ElseIf ui = 0 Then GoTo o
Else: GoTo fj
End If
o:
End Sub

Private Sub Label3_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\鸡.mp3"
kl = Label3.Caption
op = Val(InputBox("请输入数值", "X轴", Label3.Caption))
Label3.Caption = op
Label4.Caption = op + 20
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): gh = Val(Label3.Caption): ty = Val(Label4.Caption): zx = Val(Label7.Caption)
For i = gh To ty Step 0.1
Line1((i - (ty - 10)) * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + (400 * zx + 400)
Line1((i - (ty - 10)) * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + (400 * zx + 400)
Line1((i - (ty - 10)) * 10 + 100).X1 = Line1((i - (ty - 10)) * 10 + 100).X1 - (5 * (kl - op))
Line1((i - (ty - 10)) * 10 + 100).X2 = Line1((i - (ty - 10)) * 10 + 100).X2 - (5 * (kl - op))
Next
For i = 10 To 30
Shape2(i).Left = Shape2(i).Left + (400 * (kl - op))
Next
For i = gh To ty
Label9((i - (ty - 10)) + 10).Caption = i
Label2((i - (ty - 10)) + 10).Caption = a * i ^ 2 + b * i + c
Next
If ty > -20 And gh > 0 Then
For i = 10 To 30
Shape2(i).Visible = False
Next
ElseIf -20 <= gh And gh <= 0 And 0 <= ty And ty <= 20 Then
For i = 10 To 30
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Label4_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\你.mp3"
kl = Label4.Caption
op = Val(InputBox("请输入数值", "X轴", Label4.Caption))
Label4.Caption = op
Label3.Caption = op - 20
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): gh = Val(Label3.Caption): ty = Val(Label4.Caption): zx = Val(Label7.Caption)
For i = gh To ty Step 0.1
Line1((i - (ty - 10)) * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + (400 * zx + 400)
Line1((i - (ty - 10)) * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + (400 * zx + 400)
Line1((i - (ty - 10)) * 10 + 100).X1 = Line1((i - (ty - 10)) * 10 + 100).X1 + (5 * (kl - op))
Line1((i - (ty - 10)) * 10 + 100).X2 = Line1((i - (ty - 10)) * 10 + 100).X2 + (5 * (kl - op))
Next
For i = 10 To 30
Shape2(i).Left = Shape2(i).Left + (400 * (kl - op))
Next
For i = gh To ty
Label9((i - (ty - 10)) + 10).Caption = i
Label2((i - (ty - 10)) + 10).Caption = a * i ^ 2 + b * i + c
Next
If gh < -20 And ty < 0 Then
For i = 10 To 30
Shape2(i).Visible = False
Next
ElseIf -20 <= gh And gh <= 0 And 0 <= ty And ty <= 20 Then
For i = 10 To 30
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Label6_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\太.mp3"
b = Label6.Caption
a = Val(InputBox("请输入数值", "y轴", Label6.Caption))
Label6.Caption = a
Label7.Caption = a + 20
For i = 0 To 10
Shape1(i).Top = Shape1(i).Top - (400 * (b - a))
Next
For i = 0 To 9
Shape2(i).Top = Shape2(i).Top - (400 * (b - a))
Next
For i = -10 To 10 Step 0.1
Line1(i * 10 + 100).Y1 = Line1(i * 10 + 100).Y1 - (400 * (b - a))
Line1(i * 10 + 100).Y2 = Line1(i * 10 + 100).Y2 - (400 * (b - a))
Next
If Label7.Caption < 20 And Label6.Caption < 0 Then
For i = 0 To 10
Shape1(i).Visible = False
Next
For i = 0 To 9
Shape2(i).Visible = False
Next
ElseIf -20 <= Label6.Caption And Label6.Caption <= 0 And 0 <= Label7.Caption And Label7.Caption <= 20 Then
For i = 0 To 10
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Label7_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\没.mp3"
b = Label7.Caption
a = Val(InputBox("请输入数值", "y轴", Label7.Caption))
Label7.Caption = a
Label6.Caption = a - 20
For i = 0 To 10
Shape1(i).Top = Shape1(i).Top - (400 * (b - a))
Next
For i = 0 To 9
Shape2(i).Top = Shape2(i).Top - (400 * (b - a))
Next
For i = -10 To 10 Step 0.1
Line1(i * 10 + 100).Y1 = Line1(i * 10 + 100).Y1 - (400 * (b - a))
Line1(i * 10 + 100).Y2 = Line1(i * 10 + 100).Y2 - (400 * (b - a))
Next
If Label7.Caption > 0 And Label6.Caption > -20 Then
For i = 0 To 10
Shape1(i).Visible = False
Next
For i = 0 To 9
Shape2(i).Visible = False
Next
ElseIf -20 <= Label6.Caption And Label6.Caption <= 0 And 0 <= Label7.Caption And Label7.Caption <= 20 Then
For i = 0 To 10
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Picture1_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\哇哈哈.mp3"
Timer3.Interval = 3000
Timer1.Interval = 2500
End Sub

Private Sub Text1_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\大家好.m4a"
End Sub

Private Sub Text2_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\个人练习生.mp3"
End Sub

Private Sub Text3_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\练习坤.mp3"
End Sub

Private Sub Timer1_Timer()
WindowsMediaPlayer2.URL = App.Path & "\音频\0.gif"
Timer2.Interval = 5000
Timer1.Interval = 0
End Sub

Private Sub Timer2_Timer()
WindowsMediaPlayer2.URL = App.Path & "  "
Timer2.Interval = 0
End Sub

Private Sub Timer3_Timer()
WindowsMediaPlayer1.URL = App.Path & "\音频\鸡你太美.mp3"
Timer3.Interval = 0
End Sub


