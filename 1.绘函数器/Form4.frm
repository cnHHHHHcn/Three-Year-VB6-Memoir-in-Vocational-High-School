VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form4 
   Caption         =   "初中综合函数"
   ClientHeight    =   12915
   ClientLeft      =   -60
   ClientTop       =   285
   ClientWidth     =   21360
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   Begin VB.CommandButton Command7 
      Caption         =   "返回"
      Height          =   495
      Left            =   19800
      TabIndex        =   112
      Top             =   11640
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "y轴减"
      Height          =   495
      Left            =   8640
      TabIndex        =   111
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "y轴加"
      Height          =   495
      Left            =   8640
      TabIndex        =   110
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "X轴加"
      Height          =   495
      Left            =   6360
      TabIndex        =   109
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "X轴减"
      Height          =   495
      Left            =   1200
      TabIndex        =   108
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "绘坐标系"
      Height          =   495
      Left            =   1680
      TabIndex        =   103
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   12120
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   11400
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   10560
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   10560
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   10560
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   9960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   9960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   9960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "绘函数图像并计算"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   9360
      Width           =   1815
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   615
      Left            =   10680
      TabIndex        =   113
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
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
      _cx             =   1508
      _cy             =   1085
   End
   Begin VB.Label Label11 
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
      Left            =   8640
      TabIndex        =   107
      Top             =   8160
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label10 
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
      Left            =   8640
      TabIndex        =   106
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label9 
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
      Left            =   8160
      TabIndex        =   105
      Top             =   8640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label8 
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
      TabIndex        =   104
      Top             =   8640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label7 
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
      Height          =   495
      Left            =   8760
      TabIndex        =   102
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Height          =   495
      Left            =   3840
      TabIndex        =   101
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   87
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   86
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   85
      Left            =   0
      TabIndex        =   98
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   84
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   83
      Left            =   0
      TabIndex        =   96
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   82
      Left            =   0
      TabIndex        =   95
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   81
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   80
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   79
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   78
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   77
      Left            =   0
      TabIndex        =   90
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   76
      Left            =   0
      TabIndex        =   89
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   75
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   74
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   73
      Left            =   0
      TabIndex        =   86
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   72
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   71
      Left            =   0
      TabIndex        =   84
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   70
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   69
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   68
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   67
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   65
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   66
      Left            =   18720
      TabIndex        =   78
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   64
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   63
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   62
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   61
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   60
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   59
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   58
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   57
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   56
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   55
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   54
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   53
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   52
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   51
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   50
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   49
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   48
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   47
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   46
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   45
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   44
      Left            =   15840
      TabIndex        =   57
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   43
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   42
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   41
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   40
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   39
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   38
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   37
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   36
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   35
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   34
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   33
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   32
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   31
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   30
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   29
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   28
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   27
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   26
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   25
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   24
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   23
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H000040C0&
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   22
      Left            =   13080
      TabIndex        =   35
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   21
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   20
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   19
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   18
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   17
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   16
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   15
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   14
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   13
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   12
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   11
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   10
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   9
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   8
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   7
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   6
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   5
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   4
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   3
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   2
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   11280
      TabIndex        =   13
      Top             =   360
      Width           =   195
   End
   Begin VB.Line Line3 
      Index           =   25
      X1              =   17400
      X2              =   17400
      Y1              =   240
      Y2              =   10800
   End
   Begin VB.Line Line3 
      Index           =   24
      X1              =   14640
      X2              =   14640
      Y1              =   240
      Y2              =   10800
   End
   Begin VB.Line Line3 
      Index           =   23
      X1              =   12120
      X2              =   12120
      Y1              =   240
      Y2              =   10800
   End
   Begin VB.Line Line3 
      Index           =   22
      X1              =   6120
      X2              =   4320
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   21
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   20
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   19
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   18
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   17
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   16
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   15
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   14
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   13
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   12
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   11
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   10
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   9
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   8
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   7
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   6
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   5
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   4
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   11160
      X2              =   20400
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      Index           =   605
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   604
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   603
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   602
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   601
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   600
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   599
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   598
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   597
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   596
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   595
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   594
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   593
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   592
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   591
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   590
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   589
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   588
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   587
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   586
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   585
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   584
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   583
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   582
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   581
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   580
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   579
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   578
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   577
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   576
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   575
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   574
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   573
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   572
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   571
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   570
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   569
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   568
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   567
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   566
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   565
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   564
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   563
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   562
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   561
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   560
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   559
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   558
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   557
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   556
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   555
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   554
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   553
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   552
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   551
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   550
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   549
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   548
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   547
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   546
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   545
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   544
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   543
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   542
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   541
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   540
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   539
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   538
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   537
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   536
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   535
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   534
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   533
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   532
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   531
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   530
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   529
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   528
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   527
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   526
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   525
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   524
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   523
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   522
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   521
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   520
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   519
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   518
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   517
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   516
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   515
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   514
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   513
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   512
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   511
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   510
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   509
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   508
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   507
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   506
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   505
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   504
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   503
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   502
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   501
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   500
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   499
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   498
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   497
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   496
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   495
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   494
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   493
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   492
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   491
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   490
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   489
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   488
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   487
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   486
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   485
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   484
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   483
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   482
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   481
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   480
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   479
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   478
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   477
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   476
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   475
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   474
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   473
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   472
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   471
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   470
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   469
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   468
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   467
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   466
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   465
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   464
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   463
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   462
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   461
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   460
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   459
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   458
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   457
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   456
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   455
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   454
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   453
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   452
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   451
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   450
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   449
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   448
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   447
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   446
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   445
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   444
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   443
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   442
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   441
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   440
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   439
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   438
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   437
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   436
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   435
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   434
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   433
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   432
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   431
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   430
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   429
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   428
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   427
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   426
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   425
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   424
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   423
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   422
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   421
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   420
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   419
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   418
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   417
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   416
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   415
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   414
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   413
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   412
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   411
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   410
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   409
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   408
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   407
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   406
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   405
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   404
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   403
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   402
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   401
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   400
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   399
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   398
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   397
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   396
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   395
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   394
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   393
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   392
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   391
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   390
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   389
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   388
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   387
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   386
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   385
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   384
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   383
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   382
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   381
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   380
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   379
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   378
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   377
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   376
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   375
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   374
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   373
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   372
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   371
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   370
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   369
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   368
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   367
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   366
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   365
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   364
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   363
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   362
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   361
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   360
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   359
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   358
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   357
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   356
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   355
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   354
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   353
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   352
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   351
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   350
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   349
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   348
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   347
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   346
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   345
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   344
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   343
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   342
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   341
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   340
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   339
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   338
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   337
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   336
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   335
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   334
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   333
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   332
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   331
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   330
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   329
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   328
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   327
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   326
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   325
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   324
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   323
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   322
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   321
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   320
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   319
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   318
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   317
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   316
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   315
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   314
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   313
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   312
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   311
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   310
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   309
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   308
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   307
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   306
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   305
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   304
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   303
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   302
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   301
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   300
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   299
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   298
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   297
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   296
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   295
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   294
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   293
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   292
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   291
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   290
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   289
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   288
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   287
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   286
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   285
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   284
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   283
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   282
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   281
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   280
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   279
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   278
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   277
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   276
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   275
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   274
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   273
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   272
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   271
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   270
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   269
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   268
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   267
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   266
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   265
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   264
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   263
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   262
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   261
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   260
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   259
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   258
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   257
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   256
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   255
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   254
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   253
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   252
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   251
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   250
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   249
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   248
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   247
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   246
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   245
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   244
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   243
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   242
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   241
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   240
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   239
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   238
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   237
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   236
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   235
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   234
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   233
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   232
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   231
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   230
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   229
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   228
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   227
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   226
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   225
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   224
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   223
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   222
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   221
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   220
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   219
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   218
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   217
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   216
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   215
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   214
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   213
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   212
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   211
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   210
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   209
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   208
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   207
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   206
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   205
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   204
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   203
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   202
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   201
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   200
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   199
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   198
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   197
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   196
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   195
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   194
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   193
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   192
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   191
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   190
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   189
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   188
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   187
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   186
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   185
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   184
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   183
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   182
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   181
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   180
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   179
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   178
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   177
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   176
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   175
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   174
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   173
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   172
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   171
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   170
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   169
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   168
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   167
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   166
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   165
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   164
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   163
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   162
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   161
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   160
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   159
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   158
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   157
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   156
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   155
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   154
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   153
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   152
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   151
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   150
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   149
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   148
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   147
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   146
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   145
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   144
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   143
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   142
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   141
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   140
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   139
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   138
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   137
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   136
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   135
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   134
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   133
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   132
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   131
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   130
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   129
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   128
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   127
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   126
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   125
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   124
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   123
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   122
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   121
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   120
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   119
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   118
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   117
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   116
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   115
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   114
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   113
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   112
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   111
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   110
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   109
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   108
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   107
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   106
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   105
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   104
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   103
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   102
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   101
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   100
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   99
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   98
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   97
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   96
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   95
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   94
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   93
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   92
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   91
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   90
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   89
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   88
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   87
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   86
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   85
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   84
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   83
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   82
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   81
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   80
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   79
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   78
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   77
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   76
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   75
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   74
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   73
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   72
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   71
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   70
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   69
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   68
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   67
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   66
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   65
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   64
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   63
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   62
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   61
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   60
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   59
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   58
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   57
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   56
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   55
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   54
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   53
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   52
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   51
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   50
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   49
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   48
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   47
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   46
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   45
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   44
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   43
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   42
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   41
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   40
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   39
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   38
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   37
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   36
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   35
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   34
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   33
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   32
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   31
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   30
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   29
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   28
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   27
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   26
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   25
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   24
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   23
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   22
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   21
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   20
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   19
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   18
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   17
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   16
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   15
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   14
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   13
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   12
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   11
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   10
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   9
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   720
      X2              =   1920
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   12120
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   2760
      Y1              =   12000
      Y2              =   12000
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "y="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   11760
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
      Caption         =   "y=     X^2+     X+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   10560
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Caption         =   "y=     X^2+     X+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   9960
      Width           =   5415
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
      Left            =   1320
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
      Left            =   720
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\睡觉.mp3"
'坐标系绘制
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
'显示按钮和标签
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\造就完了.mp3"
Label8.Caption = -10: Label9.Caption = 10: Label10.Caption = 10: Label11.Caption = -10
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): a1 = Val(Text4.Text): b1 = Val(Text5.Text): c1 = Val(Text6.Text): op = Val(Text7.Text): lp = Val(Text8.Text)
'判定反比例函数框内是否有数
If op <= -1 Or lp >= 1 Then
Label8.Enabled = False
Label9.Enabled = False
Label10.Enabled = False
Label11.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
For i = 400 To 600
Line2(i).Visible = True
Next
Else
Label8.Enabled = True
Label9.Enabled = True
Label10.Enabled = True
Label11.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
For i = 400 To 600
Line2(i).Visible = False
Next
For i = -10 To -1
Label5(i + 77).Caption = 0
Next
For i = 1 To 10
Label5(i + 77).Caption = 0
Next
End If
'重新让函数图像回到-10至10区间
For i = -10 To 10 Step 0.1
Line2(i * 10 + 100).X1 = i * 400 + 4400
Line2(i * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + 4400
Line2(i * 10 + 100).X2 = (i + 0.1) * 400 + 4400
Line2(i * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + 4400
Line2(i * 10 + 100).BorderColor = &H40C0&
Next
For i = -10 To 10 Step 0.1
Line2(i * 10 + 300).X1 = i * 400 + 4400
Line2(i * 10 + 300).Y1 = -(a1 * i ^ 2 + b1 * i + c1) * 400 + 4400
Line2(i * 10 + 300).X2 = (i + 0.1) * 400 + 4400
Line2(i * 10 + 300).Y2 = -(a1 * (i + 0.1) ^ 2 + b1 * (i + 0.1) + c1) * 400 + 4400
Line2(i * 10 + 300).BorderColor = &HC000C0
Next
'x轴，y轴重新归位
For i = 0 To 10
Shape1(i).Left = 300 + 400 * i
Shape1(i).Top = 4300
Next
For i = 0 To 9
Shape2(i).Left = 4700 + 400 * i
Shape2(i).Top = 4300
Next
For i = 10 To 30
Shape2(i).Left = 4300
Shape2(i).Top = 300 + 400 * (i - 10)
Next
'数值重新调回-10至10区间
For j = -10 To 10
Label5(j + 33).Caption = a * j ^ 2 + b * j + c
Label5(j + 55).Caption = a1 * j ^ 2 + b1 * j + c1
Next
'判定反比例函数框内是否有数值
If op = 0 Or lp = 0 Then

Else
For i = -10 To -0.2 Step 0.1
Line2(i * 10 + 500).X1 = i * 400 + 4400
Line2(i * 10 + 500).X2 = (i + 0.1) * 400 + 4400
Line2(i * 10 + 500).Y1 = -(op / (lp * i)) * 400 + 4400
Line2(i * 10 + 500).Y2 = -(op / (lp * (i + 0.1))) * 400 + 4400
Line2(i * 10 + 500).BorderColor = &HFF&
Next
For i = 0.1 To 10 Step 0.1
Line2(i * 10 + 500).X1 = i * 400 + 4400
Line2(i * 10 + 500).X2 = (i + 0.1) * 400 + 4400
Line2(i * 10 + 500).Y1 = -(op / (lp * i)) * 400 + 4400
Line2(i * 10 + 500).Y2 = -(op / (lp * (i + 0.1))) * 400 + 4400
Line2(i * 10 + 500).BorderColor = &HFF&
Next
For i = -10 To -1
Label5(i + 77).Caption = Round(-(op / (lp * -i)), 7)
Next
For i = 1 To 10
Label5(i + 77).Caption = Round(-(op / (lp * -i)), 7)
Next
End If
'标签数值重置-10至10区间
For i = 0 To 20
Label5(i).FontSize = 18
Label5(i + 1).Caption = i - 10
Next
End Sub

Private Sub Command3_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\黄瓜丝.mp3"
Label8.Caption = Label8.Caption - 1: Label9.Caption = Label9.Caption - 1
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): a1 = Val(Text4.Text): b1 = Val(Text5.Text): c1 = Val(Text6.Text): op = Val(Text7.Text): lp = Val(Text8.Text): gh = Val(Label8.Caption): ty = Val(Label9.Caption): zx = Val(Label10.Caption)
For i = 10 To 30
Shape2(i).Left = Shape2(i).Left + 400
Next
'让y轴向右移动一个单位
For i = gh To ty Step 0.1
Line2((i - (ty - 10)) * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + (400 * zx)
Line2((i - (ty - 10)) * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + (400 * zx)
Line2((i - (ty - 10)) * 10 + 100).Y1 = Line2((i - (ty - 10)) * 10 + 100).Y1 + 400
Line2((i - (ty - 10)) * 10 + 100).Y2 = Line2((i - (ty - 10)) * 10 + 100).Y2 + 400
Line2((i - (ty - 10)) * 10 + 300).Y1 = -(a1 * i ^ 2 + b1 * i + c1) * 400 + (400 * zx)
Line2((i - (ty - 10)) * 10 + 300).Y2 = -(a1 * (i + 0.1) ^ 2 + b1 * (i + 0.1) + c1) * 400 + (400 * zx)
Line2((i - (ty - 10)) * 10 + 300).Y1 = Line2((i - (ty - 10)) * 10 + 300).Y1 + 400
Line2((i - (ty - 10)) * 10 + 300).Y2 = Line2((i - (ty - 10)) * 10 + 300).Y2 + 400
Next
'让变量x整体-1
For i = 0 To 20
Label5(i + 1).Caption = gh + i
Next
'让函数得数随变量x变化(不包含反比例函数)
For j = gh To ty
Label5((j - (ty - 10)) + 33).Caption = a * j ^ 2 + b * j + c
Label5((j - (ty - 10)) + 55).Caption = a1 * j ^ 2 + b1 * j + c1
Next
'判定y轴是否超出右边
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

Private Sub Command4_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\老干妈.mp3"
Label8.Caption = Label8.Caption + 1: Label9.Caption = Label9.Caption + 1
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): a1 = Val(Text4.Text): b1 = Val(Text5.Text): c1 = Val(Text6.Text): op = Val(Text7.Text): lp = Val(Text8.Text): gh = Val(Label8.Caption): ty = Val(Label9.Caption): zx = Val(Label10.Caption)
'y轴向左减1个单位长度
For i = 10 To 30
Shape2(i).Left = Shape2(i).Left - 400
Next
'函数图像向左减1个单位长度
For i = gh To ty Step 0.1
Line2((i - (ty - 10)) * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + (400 * zx + 800)
Line2((i - (ty - 10)) * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + (400 * zx + 800)
Line2((i - (ty - 10)) * 10 + 100).Y1 = Line2((i - (ty - 10)) * 10 + 100).Y1 - 400
Line2((i - (ty - 10)) * 10 + 100).Y2 = Line2((i - (ty - 10)) * 10 + 100).Y2 - 400
Line2((i - (ty - 10)) * 10 + 300).Y1 = -(a1 * i ^ 2 + b1 * i + c1) * 400 + (400 * zx + 800)
Line2((i - (ty - 10)) * 10 + 300).Y2 = -(a1 * (i + 0.1) ^ 2 + b1 * (i + 0.1) + c1) * 400 + (400 * zx + 800)
Line2((i - (ty - 10)) * 10 + 300).Y1 = Line2((i - (ty - 10)) * 10 + 300).Y1 - 400
Line2((i - (ty - 10)) * 10 + 300).Y2 = Line2((i - (ty - 10)) * 10 + 300).Y2 - 400
Next
'让变量x整体+1
For i = 0 To 20
Label5(i + 1).Caption = gh + i
Next
'让函数得数随变量x变化(不包含反比例函数)
For j = gh To ty
Label5((j - (ty - 10)) + 33).Caption = a * j ^ 2 + b * j + c
Label5((j - (ty - 10)) + 55).Caption = a1 * j ^ 2 + b1 * j + c1
Next
'判定y轴是否超出左边
If Label9.Caption > 20 And Label8.Caption > 0 Then
For i = 10 To 30
Shape2(i).Visible = False
Next
ElseIf 0 <= Label9.Caption And Label9.Caption <= 20 And -20 <= Label8.Caption And Label8.Caption <= 0 Then
For i = 10 To 30
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Command5_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\加柠檬.mp3"
Label10.Caption = Label10.Caption + 1: Label11.Caption = Label11.Caption + 1
'x轴向下加1个单位长度
For i = 0 To 10
Shape1(i).Top = Shape1(i).Top + 400
Next
For i = 0 To 9
Shape2(i).Top = Shape2(i).Top + 400
Next
'函数向下加1个单位长度
For i = -10 To 10 Step 0.1
Line2(i * 10 + 100).Y1 = Line2(i * 10 + 100).Y1 + 400
Line2(i * 10 + 100).Y2 = Line2(i * 10 + 100).Y2 + 400
Line2(i * 10 + 300).Y1 = Line2(i * 10 + 300).Y1 + 400
Line2(i * 10 + 300).Y2 = Line2(i * 10 + 300).Y2 + 400
Next
'判定x轴是否超出下边
If Label10.Caption > 20 And Label11.Caption > 0 Then
For i = 0 To 10
Shape1(i).Visible = False
Next
For i = 0 To 9
Shape2(i).Visible = False
Next
ElseIf -20 <= Label11.Caption And Label11.Caption <= 0 And 0 <= Label10.Caption And Label10.Caption <= 20 Then
For i = 0 To 10
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Command6_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\臭豆腐.mp3"
Label10.Caption = Label10.Caption - 1: Label11.Caption = Label11.Caption - 1
'x轴向上加1个单位长度
For i = 0 To 10
Shape1(i).Top = Shape1(i).Top - 400
Next
For i = 0 To 9
Shape2(i).Top = Shape2(i).Top - 400
Next
'函数向上加1个单位长度
For i = -10 To 10 Step 0.1
Line2(i * 10 + 100).Y1 = Line2(i * 10 + 100).Y1 - 400
Line2(i * 10 + 100).Y2 = Line2(i * 10 + 100).Y2 - 400
Line2(i * 10 + 300).Y1 = Line2(i * 10 + 300).Y1 - 400
Line2(i * 10 + 300).Y2 = Line2(i * 10 + 300).Y2 - 400
Next
'判定x轴是否超出上边
If Label10.Caption > 20 And Label11.Caption > 0 Then
For i = 0 To 10
Shape1(i).Visible = False
Next
For i = 0 To 9
Shape2(i).Visible = False
Next
ElseIf -20 <= Label11.Caption And Label11.Caption <= 0 And 0 <= Label10.Caption And Label10.Caption <= 20 Then
For i = 0 To 10
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Command7_Click()
Form4.Hide: Form2.Show
End Sub

Private Sub Form_Load()
'线型控件排列
For i = 0 To 21
Line3(i + 1).Y1 = Line3(i).Y1 + 480
Line3(i + 1).Y2 = Line3(i).Y2 + 480
Line3(i + 1).X1 = 11160
Line3(i + 1).X2 = 20400
Next
'标签控件排列(三行)
For i = 0 To 20
Label5(i).FontSize = 18
Label5(i + 1).Top = Label5(i).Top + 480
Label5(i + 1).Caption = i - 10
Label5(i + 1).Left = 11280
Label5(i + 23).Top = Label5(i + 22).Top + 480
Label5(i + 23).FontSize = 18
Label5(i + 23).Left = 12240
Label5(i + 45).Top = Label5(i + 22).Top + 480
Label5(i + 45).FontSize = 18
Label5(i + 45).Left = 14760
Label5(i + 67).Top = Label5(i + 22).Top + 480
Label5(i + 67).FontSize = 18
Label5(i + 67).Left = 17520
Next
Label5(77).Caption = "/(0)"
End Sub

Private Sub Label1_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\嘿嘿,来啦.mp3"
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

Private Sub Label10_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\还管饱.mp3"
'存数据，改变标签数值
b = Label10.Caption
a = Val(InputBox("请输入数值", "y轴", Label10.Caption))
Label10.Caption = a
Label11.Caption = a - 20
'x轴向上移动
For i = 0 To 10
Shape1(i).Top = Shape1(i).Top - (400 * (b - a))
Next
For i = 0 To 9
Shape2(i).Top = Shape2(i).Top - (400 * (b - a))
Next
'让函数图像向上移动(不包含反比例函数)
For i = -10 To 10 Step 0.1
Line2(i * 10 + 100).Y1 = Line2(i * 10 + 100).Y1 - (400 * (b - a))
Line2(i * 10 + 100).Y2 = Line2(i * 10 + 100).Y2 - (400 * (b - a))

Line2(i * 10 + 300).Y1 = Line2(i * 10 + 300).Y1 - (400 * (b - a))
Line2(i * 10 + 300).Y2 = Line2(i * 10 + 300).Y2 - (400 * (b - a))
Next
'判定x轴是否隐藏
If Label10.Caption > 0 And Label11.Caption > -20 Then
For i = 0 To 10
Shape1(i).Visible = False
Next
For i = 0 To 9
Shape2(i).Visible = False
Next
ElseIf -20 <= Label11.Caption And Label11.Caption <= 0 And 0 <= Label10.Caption And Label10.Caption <= 20 Then
For i = 0 To 10
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Label11_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\既实惠.mp3"
'存数据，改变标签数值
b = Label11.Caption
a = Val(InputBox("请输入数值", "y轴", Label11.Caption))
Label11.Caption = a
Label10.Caption = a + 20
'x轴向下移动
For i = 0 To 10
Shape1(i).Top = Shape1(i).Top + (400 * (b - a))
Next
For i = 0 To 9
Shape2(i).Top = Shape2(i).Top + (400 * (b - a))
Next
'让函数图像向下移动(不包含反比例函数)
For i = -10 To 10 Step 0.1
Line2(i * 10 + 100).Y1 = Line2(i * 10 + 100).Y1 - (400 * (b - a))
Line2(i * 10 + 100).Y2 = Line2(i * 10 + 100).Y2 - (400 * (b - a))

Line2(i * 10 + 300).Y1 = Line2(i * 10 + 300).Y1 - (400 * (b - a))
Line2(i * 10 + 300).Y2 = Line2(i * 10 + 300).Y2 - (400 * (b - a))
Next
'判定x轴是否隐藏
If Label10.Caption < 20 And Label11.Caption < 0 Then
For i = 0 To 10
Shape1(i).Visible = False
Next
For i = 0 To 9
Shape2(i).Visible = False
Next
ElseIf -20 <= Label11.Caption And Label11.Caption <= 0 And 0 <= Label10.Caption And Label10.Caption <= 20 Then
For i = 0 To 10
Shape1(i).Visible = True
Next
For i = 0 To 9
Shape2(i).Visible = True
Next
End If
End Sub

Private Sub Label2_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\一日三餐没烦恼.mp3"
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

Private Sub Label8_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\撤硕.mp3"
'存数据，改变标签数值
kl = Val(Label8.Caption)
qw = Val(InputBox("请输入数值", "X轴", Label8.Caption))
Label8.Caption = qw
Label9.Caption = qw + 20
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): a1 = Val(Text4.Text): b1 = Val(Text5.Text): c1 = Val(Text6.Text): op = Val(Text7.Text): lp = Val(Text8.Text): gh = Val(Label8.Caption): ty = Val(Label9.Caption): zx = Val(Label10.Caption)
'y轴向左移动
For i = 10 To 30
Shape2(i).Left = Shape2(i).Left + (400 * -(qw - kl))
Next
'函数图像向左移动并随标签数值改变图像
For i = gh To ty Step 0.1
Line2((i - (ty - 10)) * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + (400 * zx)
Line2((i - (ty - 10)) * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + (400 * zx)
Line2((i - (ty - 10)) * 10 + 100).Y1 = Line2((i - (ty - 10)) * 10 + 100).Y1 + 400
Line2((i - (ty - 10)) * 10 + 100).Y2 = Line2((i - (ty - 10)) * 10 + 100).Y2 + 400
Line2((i - (ty - 10)) * 10 + 300).Y1 = -(a1 * i ^ 2 + b1 * i + c1) * 400 + (400 * zx)
Line2((i - (ty - 10)) * 10 + 300).Y2 = -(a1 * (i + 0.1) ^ 2 + b1 * (i + 0.1) + c1) * 400 + (400 * zx)
Line2((i - (ty - 10)) * 10 + 300).Y1 = Line2((i - (ty - 10)) * 10 + 300).Y1 + 400
Line2((i - (ty - 10)) * 10 + 300).Y2 = Line2((i - (ty - 10)) * 10 + 300).Y2 + 400
Next
'标签数值整体减1
For i = 0 To 20
Label5(i + 1).Caption = gh + i
Next
'函数图像随标签数值变化
For j = gh To ty
Label5((j - (ty - 10)) + 33).Caption = a * j ^ 2 + b * j + c
Label5((j - (ty - 10)) + 55).Caption = a1 * j ^ 2 + b1 * j + c1
Next
'判定y轴是否隐藏
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

Private Sub Label9_Click()
WindowsMediaPlayer1.URL = App.Path & "\音频\还是那句话.mp3"
'存数据，改变标签数值
kl = Val(Label9.Caption)
qw = Val(InputBox("请输入数值", "X轴", Label9.Caption))
Label9.Caption = qw
Label8.Caption = qw - 20
a = Val(Text1.Text): b = Val(Text2.Text): c = Val(Text3.Text): a1 = Val(Text4.Text): b1 = Val(Text5.Text): c1 = Val(Text6.Text): op = Val(Text7.Text): lp = Val(Text8.Text): gh = Val(Label8.Caption): ty = Val(Label9.Caption): zx = Val(Label10.Caption)
'y轴向右移动
For i = 10 To 30
Shape2(i).Left = Shape2(i).Left + (400 * -(qw - kl))
Next
'函数图像向右移动并随标签数值改变图像
For i = gh To ty Step 0.1
Line2((i - (ty - 10)) * 10 + 100).Y1 = -(a * i ^ 2 + b * i + c) * 400 + (400 * zx + 800)
Line2((i - (ty - 10)) * 10 + 100).Y2 = -(a * (i + 0.1) ^ 2 + b * (i + 0.1) + c) * 400 + (400 * zx + 800)
Line2((i - (ty - 10)) * 10 + 100).Y1 = Line2((i - (ty - 10)) * 10 + 100).Y1 - 400
Line2((i - (ty - 10)) * 10 + 100).Y2 = Line2((i - (ty - 10)) * 10 + 100).Y2 - 400
Line2((i - (ty - 10)) * 10 + 300).Y1 = -(a1 * i ^ 2 + b1 * i + c1) * 400 + (400 * zx + 800)
Line2((i - (ty - 10)) * 10 + 300).Y2 = -(a1 * (i + 0.1) ^ 2 + b1 * (i + 0.1) + c1) * 400 + (400 * zx + 800)
Line2((i - (ty - 10)) * 10 + 300).Y1 = Line2((i - (ty - 10)) * 10 + 300).Y1 - 400
Line2((i - (ty - 10)) * 10 + 300).Y2 = Line2((i - (ty - 10)) * 10 + 300).Y2 - 400
Next
'标签数值整体加1
For i = 0 To 20
Label5(i + 1).Caption = gh + i
Next
'函数图像随标签数值变化
For j = gh To ty
Label5((j - (ty - 10)) + 33).Caption = a * j ^ 2 + b * j + c
Label5((j - (ty - 10)) + 55).Caption = a1 * j ^ 2 + b1 * j + c1
Next
'判定y轴是否隐藏
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

Private Sub Text1_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\今天就吃.mp3"
End Sub

Private Sub Text2_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\老八秘制小汉堡儿.mp3"
End Sub

Private Sub Text3_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\老八汉堡升级啦.mp3"
End Sub

Private Sub Text4_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\喇叭撤硕二战语录.mp3"
End Sub

Private Sub Text5_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\吸溜.mp3"
End Sub

Private Sub Text6_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\呕.mp3"
End Sub

Private Sub Text7_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\兄弟们2.mp3"
End Sub

Private Sub Text8_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\奥利给,干就完了.mp3"
End Sub
