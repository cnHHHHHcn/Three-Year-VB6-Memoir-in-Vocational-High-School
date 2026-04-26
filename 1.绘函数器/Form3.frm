VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   Caption         =   "反比例函数"
   ClientHeight    =   12915
   ClientLeft      =   -60
   ClientTop       =   465
   ClientWidth     =   21360
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      Height          =   495
      Left            =   18600
      TabIndex        =   51
      Top             =   11400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "绘坐标系"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   10200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "绘函数图像并计算"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   10200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
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
      Left            =   2640
      TabIndex        =   3
      Top             =   11640
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
      Left            =   2640
      TabIndex        =   1
      Top             =   10920
      Width           =   1215
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
      TabIndex        =   54
      Top             =   4200
      Width           =   1140
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
      TabIndex        =   53
      Top             =   8640
      Width           =   1140
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   8880
      TabIndex        =   52
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
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
      _cx             =   1720
      _cy             =   873
   End
   Begin VB.Label Label3 
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
      Index           =   44
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   43
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   42
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   41
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   40
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   39
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   38
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   37
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   36
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   35
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   34
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   33
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   32
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   31
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   30
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   29
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   28
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   27
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   26
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   25
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   24
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   23
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   12240
      TabIndex        =   28
      Top             =   600
      Width           =   195
   End
   Begin VB.Label Label3 
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
      TabIndex        =   27
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   20
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   19
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   18
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   17
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   16
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   15
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   14
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   13
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   12
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   11
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   10
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   9
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   8
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   7
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   6
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label3 
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
      Left            =   10680
      TabIndex        =   6
      Top             =   600
      Width           =   195
   End
   Begin VB.Line Line4 
      X1              =   11280
      X2              =   11280
      Y1              =   480
      Y2              =   11040
   End
   Begin VB.Line Line3 
      Index           =   22
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   21
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   20
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   19
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   18
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   17
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   16
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   15
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   14
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   13
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   12
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   11
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   10
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   9
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   8
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   7
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   6
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   5
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   4
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   3
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   0
      X2              =   2280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   -360
      X2              =   1920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      Index           =   0
      X1              =   10440
      X2              =   13440
      Y1              =   480
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
      X1              =   960
      X2              =   2160
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Label Label2 
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
      Left            =   3840
      TabIndex        =   2
      Top             =   11640
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   4560
      Y1              =   11520
      Y2              =   11520
   End
   Begin VB.Label Label1 
      Caption         =   "y="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   11280
      Width           =   615
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
      Left            =   240
      Top             =   600
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   410
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   409
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   408
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   407
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   406
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   405
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   404
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   403
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   402
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   401
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   400
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   399
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   398
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   397
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   396
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   395
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   394
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   393
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   392
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   391
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   390
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   389
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   388
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   387
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   386
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   385
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   384
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   383
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   382
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   381
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   380
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   379
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   378
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   377
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   376
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   375
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   374
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   373
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   372
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   371
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   370
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   369
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   368
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   367
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   366
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   365
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   364
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   363
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   362
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   361
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   360
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   359
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   358
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   357
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   356
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   355
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   354
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   353
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   352
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   351
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   350
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   349
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   348
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   347
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   346
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   345
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   344
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   343
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   342
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   341
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   340
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   339
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   338
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   337
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   336
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   335
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   334
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   333
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   332
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   331
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   330
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   329
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   328
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   327
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   326
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   325
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   324
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   323
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   322
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   321
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   320
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   319
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   318
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   317
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   316
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   315
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   314
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   313
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   312
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   311
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   310
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   309
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   308
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   307
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   306
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   305
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   304
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   303
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   302
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   301
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   300
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   299
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   298
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   297
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   296
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   295
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   294
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   293
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   292
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   291
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   290
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   289
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   288
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   287
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   286
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   285
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   284
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   283
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   282
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   281
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   280
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   279
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   278
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   277
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   276
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   275
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   274
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   273
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   272
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   271
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   270
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   269
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   268
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   267
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   266
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   265
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   264
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   263
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   262
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   261
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   260
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   259
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   258
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   257
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   256
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   255
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   254
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   253
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   252
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   251
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   250
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   249
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   248
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   247
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   246
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   245
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   244
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   243
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   242
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   241
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   240
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   239
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   238
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   237
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   236
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   235
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   234
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   233
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   232
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   231
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   230
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   229
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   228
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   227
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   226
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   225
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   224
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   223
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   222
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   221
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   220
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   219
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   218
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   217
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   216
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   215
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   214
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   213
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   212
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   211
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   210
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   209
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   208
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   207
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   206
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   205
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   204
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   203
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   202
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   201
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   200
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   199
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   198
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   197
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   196
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   195
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   194
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   193
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   192
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   191
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   190
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   189
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   188
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   187
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   186
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   185
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   184
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   183
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   182
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   181
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   180
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   179
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   178
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   177
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   176
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   175
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   174
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   173
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   172
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   171
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   170
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   169
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   168
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   167
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   166
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   165
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   164
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   163
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   162
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   161
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   160
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   159
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   158
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   157
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   156
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   155
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   154
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   153
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   152
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   151
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   150
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   149
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   148
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   147
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   146
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   145
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   144
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   143
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   142
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   141
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   140
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   139
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   138
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   137
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   136
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   135
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   134
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   133
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   132
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   131
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   130
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   129
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   128
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   127
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   126
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   125
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   124
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   123
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   122
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   121
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   120
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   119
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   118
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   117
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   116
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   115
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   114
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   113
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   112
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   111
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   110
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   109
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   108
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   107
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   106
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   105
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   104
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   103
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   102
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   101
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   100
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   99
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   98
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   97
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   96
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   95
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   94
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   93
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   92
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   91
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   90
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   89
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   88
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   87
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   86
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   85
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   84
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   83
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   82
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   81
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   80
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   79
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   78
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   77
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   76
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   75
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   74
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   73
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   72
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   71
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   70
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   69
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   68
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   67
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   66
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   65
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   64
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   63
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   62
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   61
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   60
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   59
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   58
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   57
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   56
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   55
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   54
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   53
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   52
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   51
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   50
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   49
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   48
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   47
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   46
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   45
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   44
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   43
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   42
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   41
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   40
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   39
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   38
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   37
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   36
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   35
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   34
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   33
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   32
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   31
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   30
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   29
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   28
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   27
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   26
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   25
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   24
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   23
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   22
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   21
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   20
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   19
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   18
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   17
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   16
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   15
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   14
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   13
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   12
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   11
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   10
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   9
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   8
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   7
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   6
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   5
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   4
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   105
      Index           =   0
      Left            =   600
      Top             =   120
      Width           =   105
   End
End
Attribute VB_Name = "Form3"
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
Shape1(i).Height = 200
Shape1(i).Width = 200
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
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
a = Val(Text1.Text): b = Val(Text2.Text)
If Text1.Text = "" Or Text2.Text = "" Then
WindowsMediaPlayer1.URL = App.Path & "\音频\你干嘛.mp3"
GoTo a
End If
WindowsMediaPlayer1.URL = App.Path & "\音频\上分.mp3"
For i = -10 To -0.2 Step 0.1
Line2(i * 10 + 100).X1 = i * 400 + 4400
Line2(i * 10 + 100).X2 = (i + 0.1) * 400 + 4400
Line2(i * 10 + 100).Y1 = -(a / (b * i)) * 400 + 4400
Line2(i * 10 + 100).Y2 = -(a / (b * (i + 0.1))) * 400 + 4400
Line2(i * 10 + 100).BorderColor = &H40C0&
Next
For i = 0.1 To 10 Step 0.1
Line2(i * 10 + 100).X1 = i * 400 + 4400
Line2(i * 10 + 100).X2 = (i + 0.1) * 400 + 4400
Line2(i * 10 + 100).Y1 = -(a / (b * i)) * 400 + 4400
Line2(i * 10 + 100).Y2 = -(a / (b * (i + 0.1))) * 400 + 4400
Line2(i * 10 + 100).BorderColor = &H40C0&
Next
Line2(99).Visible = False
Line2(100).Visible = False
For i = 23 To 32
Label3(i).Caption = Round(-(a / (b * -(i - 33))), 7)
Next
Label3(33).Caption = "/(0)"
For i = 34 To 43
Label3(i).Caption = Round(-(a / (b * -(i - 33))), 7)
Next
GoTo rt
a:
MsgBox "分子和分母不能等于0。", 48, "警告"
rt:
End Sub

Private Sub Command3_Click()
Form3.Hide: Form2.Show
End Sub

Private Sub Form_Load()
For i = 0 To 21
Line3(i + 1).Y1 = Line3(i).Y1 + 480
Line3(i + 1).Y2 = Line3(i).Y2 + 480
Line3(i + 1).X1 = 10440
Line3(i + 1).X2 = 13440
Next
For i = 0 To 20
Label3(i + 1).Top = Label3(i).Top + 480
Label3(i + 1).Caption = i - 10
Label3(i + 1).Left = 10560
Label3(i + 23).Top = Label3(i + 22).Top + 480
Label3(i + 23).Left = 11400
Next
End Sub

Private Sub Text1_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\太香了.mp3"
End Sub

Private Sub Text2_Change()
WindowsMediaPlayer1.URL = App.Path & "\音频\狒狒叫.mp3"
End Sub
