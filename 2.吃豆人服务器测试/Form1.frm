VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "游戏"
   ClientHeight    =   10650
   ClientLeft      =   2055
   ClientTop       =   1425
   ClientWidth     =   17670
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   17670
   Begin VB.Timer readying 
      Enabled         =   0   'False
      Left            =   14400
      Top             =   6960
   End
   Begin VB.Timer player1and2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15120
      Top             =   6960
   End
   Begin VB.CommandButton Command8 
      Caption         =   "游戏记录"
      Height          =   495
      Left            =   14280
      TabIndex        =   24
      Top             =   9960
      Width           =   1215
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15720
      Top             =   8760
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15720
      Top             =   8160
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   15720
      Top             =   7560
   End
   Begin VB.CommandButton Command7 
      Caption         =   "p(&G)"
      Height          =   180
      Left            =   2.45745e5
      TabIndex        =   21
      Top             =   11520
      Width           =   90
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13680
      Top             =   9960
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13680
      Top             =   9360
   End
   Begin VB.CommandButton Command6 
      Caption         =   "继续"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   495
      Left            =   14280
      TabIndex        =   11
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "设置"
      Height          =   495
      Left            =   14280
      TabIndex        =   10
      Top             =   9240
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13680
      Top             =   8760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   13680
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13680
      Top             =   7560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Height          =   200
      Left            =   4920
      Picture         =   "Form1.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   200
   End
   Begin VB.CommandButton Command2 
      DisabledPicture =   "Form1.frx":259D0
      DownPicture     =   "Form1.frx":4A55E
      Height          =   200
      Left            =   4920
      Picture         =   "Form1.frx":6F0EC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   200
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   15120
      TabIndex        =   29
      Top             =   3240
      Width           =   90
   End
   Begin VB.Label lplayer2 
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   28
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label lplayer1 
      AutoSize        =   -1  'True
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
      Left            =   13560
      TabIndex        =   27
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "player1:"
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
      Left            =   11520
      TabIndex        =   26
      Top             =   4680
      Width           =   2040
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "player2:"
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
      Left            =   11520
      TabIndex        =   25
      Top             =   5400
      Width           =   2040
   End
   Begin VB.Label Label16 
      Height          =   495
      Left            =   10800
      TabIndex        =   23
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   495
      Left            =   10800
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   495
      Left            =   16800
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "0"
      Height          =   495
      Left            =   16800
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
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
      Left            =   12840
      TabIndex        =   18
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "暂停："
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
      Left            =   11520
      TabIndex        =   17
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   495
      Left            =   9240
      TabIndex        =   16
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   13800
      TabIndex        =   15
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "无敌时间："
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
      Left            =   11520
      TabIndex        =   14
      Top             =   3240
      Width           =   2475
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   495
      Left            =   9240
      TabIndex        =   12
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "1"
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
      Left            =   13800
      TabIndex        =   9
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "复活次数："
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
      Left            =   11520
      TabIndex        =   8
      Top             =   2520
      Width           =   2475
   End
   Begin VB.Image Image5 
      Height          =   195
      Left            =   8760
      Picture         =   "Form1.frx":795AA
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   8760
      Picture         =   "Form1.frx":81954
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   8520
      Picture         =   "Form1.frx":81FA4
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   8520
      Picture         =   "Form1.frx":8A980
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Left            =   12840
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "时间:"
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
      Left            =   11520
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   375
      Left            =   12840
      TabIndex        =   5
      Top             =   12360
      Visible         =   0   'False
      Width           =   1095
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
      _cx             =   1931
      _cy             =   661
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   400
      Left            =   0
      Picture         =   "Form1.frx":90E6B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   399
      Left            =   0
      Picture         =   "Form1.frx":910B5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   398
      Left            =   0
      Picture         =   "Form1.frx":912FF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   397
      Left            =   0
      Picture         =   "Form1.frx":91549
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   396
      Left            =   0
      Picture         =   "Form1.frx":91793
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   395
      Left            =   0
      Picture         =   "Form1.frx":919DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   394
      Left            =   0
      Picture         =   "Form1.frx":91C27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   393
      Left            =   0
      Picture         =   "Form1.frx":91E71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   392
      Left            =   0
      Picture         =   "Form1.frx":920BB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   391
      Left            =   0
      Picture         =   "Form1.frx":92305
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   390
      Left            =   0
      Picture         =   "Form1.frx":9254F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   389
      Left            =   0
      Picture         =   "Form1.frx":92799
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   388
      Left            =   0
      Picture         =   "Form1.frx":929E3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   387
      Left            =   0
      Picture         =   "Form1.frx":92C2D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   386
      Left            =   0
      Picture         =   "Form1.frx":92E77
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   385
      Left            =   0
      Picture         =   "Form1.frx":930C1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   384
      Left            =   0
      Picture         =   "Form1.frx":9330B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   383
      Left            =   0
      Picture         =   "Form1.frx":93555
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   382
      Left            =   0
      Picture         =   "Form1.frx":9379F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   381
      Left            =   0
      Picture         =   "Form1.frx":939E9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   380
      Left            =   0
      Picture         =   "Form1.frx":93C33
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   379
      Left            =   0
      Picture         =   "Form1.frx":93E7D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   378
      Left            =   0
      Picture         =   "Form1.frx":940C7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   377
      Left            =   0
      Picture         =   "Form1.frx":94311
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   376
      Left            =   0
      Picture         =   "Form1.frx":9455B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   375
      Left            =   0
      Picture         =   "Form1.frx":947A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   374
      Left            =   0
      Picture         =   "Form1.frx":949EF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   373
      Left            =   0
      Picture         =   "Form1.frx":94C39
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   372
      Left            =   0
      Picture         =   "Form1.frx":94E83
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   371
      Left            =   0
      Picture         =   "Form1.frx":950CD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   370
      Left            =   0
      Picture         =   "Form1.frx":95317
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   369
      Left            =   0
      Picture         =   "Form1.frx":95561
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   368
      Left            =   0
      Picture         =   "Form1.frx":957AB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   367
      Left            =   0
      Picture         =   "Form1.frx":959F5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   366
      Left            =   0
      Picture         =   "Form1.frx":95C3F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   365
      Left            =   0
      Picture         =   "Form1.frx":95E89
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   364
      Left            =   0
      Picture         =   "Form1.frx":960D3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   363
      Left            =   0
      Picture         =   "Form1.frx":9631D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   362
      Left            =   0
      Picture         =   "Form1.frx":96567
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   361
      Left            =   0
      Picture         =   "Form1.frx":967B1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   360
      Left            =   0
      Picture         =   "Form1.frx":969FB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   359
      Left            =   0
      Picture         =   "Form1.frx":96C45
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   358
      Left            =   0
      Picture         =   "Form1.frx":96E8F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   357
      Left            =   0
      Picture         =   "Form1.frx":970D9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   356
      Left            =   0
      Picture         =   "Form1.frx":97323
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   355
      Left            =   0
      Picture         =   "Form1.frx":9756D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   354
      Left            =   0
      Picture         =   "Form1.frx":977B7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   353
      Left            =   0
      Picture         =   "Form1.frx":97A01
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   352
      Left            =   0
      Picture         =   "Form1.frx":97C4B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   351
      Left            =   0
      Picture         =   "Form1.frx":97E95
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   350
      Left            =   0
      Picture         =   "Form1.frx":980DF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   349
      Left            =   0
      Picture         =   "Form1.frx":98329
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   348
      Left            =   0
      Picture         =   "Form1.frx":98573
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   347
      Left            =   0
      Picture         =   "Form1.frx":987BD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   346
      Left            =   0
      Picture         =   "Form1.frx":98A07
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   345
      Left            =   0
      Picture         =   "Form1.frx":98C51
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   344
      Left            =   0
      Picture         =   "Form1.frx":98E9B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   343
      Left            =   0
      Picture         =   "Form1.frx":990E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   342
      Left            =   0
      Picture         =   "Form1.frx":9932F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   341
      Left            =   0
      Picture         =   "Form1.frx":99579
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   340
      Left            =   0
      Picture         =   "Form1.frx":997C3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   339
      Left            =   0
      Picture         =   "Form1.frx":99A0D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   338
      Left            =   0
      Picture         =   "Form1.frx":99C57
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   337
      Left            =   0
      Picture         =   "Form1.frx":99EA1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   336
      Left            =   0
      Picture         =   "Form1.frx":9A0EB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   335
      Left            =   0
      Picture         =   "Form1.frx":9A335
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   334
      Left            =   0
      Picture         =   "Form1.frx":9A57F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   333
      Left            =   0
      Picture         =   "Form1.frx":9A7C9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   332
      Left            =   0
      Picture         =   "Form1.frx":9AA13
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   331
      Left            =   0
      Picture         =   "Form1.frx":9AC5D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   330
      Left            =   0
      Picture         =   "Form1.frx":9AEA7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   329
      Left            =   0
      Picture         =   "Form1.frx":9B0F1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   328
      Left            =   0
      Picture         =   "Form1.frx":9B33B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   327
      Left            =   0
      Picture         =   "Form1.frx":9B585
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   326
      Left            =   0
      Picture         =   "Form1.frx":9B7CF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   325
      Left            =   0
      Picture         =   "Form1.frx":9BA19
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   324
      Left            =   0
      Picture         =   "Form1.frx":9BC63
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   323
      Left            =   0
      Picture         =   "Form1.frx":9BEAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   322
      Left            =   0
      Picture         =   "Form1.frx":9C0F7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   321
      Left            =   0
      Picture         =   "Form1.frx":9C341
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   320
      Left            =   0
      Picture         =   "Form1.frx":9C58B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   319
      Left            =   0
      Picture         =   "Form1.frx":9C7D5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   318
      Left            =   0
      Picture         =   "Form1.frx":9CA1F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   317
      Left            =   0
      Picture         =   "Form1.frx":9CC69
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   316
      Left            =   0
      Picture         =   "Form1.frx":9CEB3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   315
      Left            =   0
      Picture         =   "Form1.frx":9D0FD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   314
      Left            =   0
      Picture         =   "Form1.frx":9D347
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   313
      Left            =   0
      Picture         =   "Form1.frx":9D591
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   312
      Left            =   0
      Picture         =   "Form1.frx":9D7DB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   311
      Left            =   0
      Picture         =   "Form1.frx":9DA25
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   310
      Left            =   0
      Picture         =   "Form1.frx":9DC6F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   309
      Left            =   0
      Picture         =   "Form1.frx":9DEB9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   308
      Left            =   0
      Picture         =   "Form1.frx":9E103
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   307
      Left            =   0
      Picture         =   "Form1.frx":9E34D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   306
      Left            =   0
      Picture         =   "Form1.frx":9E597
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   305
      Left            =   0
      Picture         =   "Form1.frx":9E7E1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   304
      Left            =   0
      Picture         =   "Form1.frx":9EA2B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   303
      Left            =   0
      Picture         =   "Form1.frx":9EC75
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   302
      Left            =   0
      Picture         =   "Form1.frx":9EEBF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   301
      Left            =   0
      Picture         =   "Form1.frx":9F109
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   300
      Left            =   0
      Picture         =   "Form1.frx":9F353
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   299
      Left            =   0
      Picture         =   "Form1.frx":9F59D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   298
      Left            =   0
      Picture         =   "Form1.frx":9F7E7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   297
      Left            =   0
      Picture         =   "Form1.frx":9FA31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   296
      Left            =   0
      Picture         =   "Form1.frx":9FC7B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   295
      Left            =   0
      Picture         =   "Form1.frx":9FEC5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   294
      Left            =   0
      Picture         =   "Form1.frx":A010F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   293
      Left            =   0
      Picture         =   "Form1.frx":A0359
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   292
      Left            =   0
      Picture         =   "Form1.frx":A05A3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   291
      Left            =   0
      Picture         =   "Form1.frx":A07ED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   290
      Left            =   0
      Picture         =   "Form1.frx":A0A37
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   289
      Left            =   0
      Picture         =   "Form1.frx":A0C81
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   288
      Left            =   0
      Picture         =   "Form1.frx":A0ECB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   287
      Left            =   0
      Picture         =   "Form1.frx":A1115
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   286
      Left            =   0
      Picture         =   "Form1.frx":A135F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   285
      Left            =   0
      Picture         =   "Form1.frx":A15A9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   284
      Left            =   0
      Picture         =   "Form1.frx":A17F3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   283
      Left            =   0
      Picture         =   "Form1.frx":A1A3D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   282
      Left            =   0
      Picture         =   "Form1.frx":A1C87
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   281
      Left            =   0
      Picture         =   "Form1.frx":A1ED1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   280
      Left            =   0
      Picture         =   "Form1.frx":A211B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   279
      Left            =   0
      Picture         =   "Form1.frx":A2365
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   278
      Left            =   0
      Picture         =   "Form1.frx":A25AF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   277
      Left            =   0
      Picture         =   "Form1.frx":A27F9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   276
      Left            =   0
      Picture         =   "Form1.frx":A2A43
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   275
      Left            =   0
      Picture         =   "Form1.frx":A2C8D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   274
      Left            =   0
      Picture         =   "Form1.frx":A2ED7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   273
      Left            =   0
      Picture         =   "Form1.frx":A3121
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   272
      Left            =   0
      Picture         =   "Form1.frx":A336B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   271
      Left            =   0
      Picture         =   "Form1.frx":A35B5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   270
      Left            =   0
      Picture         =   "Form1.frx":A37FF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   269
      Left            =   0
      Picture         =   "Form1.frx":A3A49
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   268
      Left            =   0
      Picture         =   "Form1.frx":A3C93
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   267
      Left            =   0
      Picture         =   "Form1.frx":A3EDD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   266
      Left            =   0
      Picture         =   "Form1.frx":A4127
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   265
      Left            =   0
      Picture         =   "Form1.frx":A4371
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   264
      Left            =   0
      Picture         =   "Form1.frx":A45BB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   263
      Left            =   0
      Picture         =   "Form1.frx":A4805
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   262
      Left            =   0
      Picture         =   "Form1.frx":A4A4F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   261
      Left            =   0
      Picture         =   "Form1.frx":A4C99
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   260
      Left            =   0
      Picture         =   "Form1.frx":A4EE3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   259
      Left            =   0
      Picture         =   "Form1.frx":A512D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   258
      Left            =   0
      Picture         =   "Form1.frx":A5377
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   257
      Left            =   0
      Picture         =   "Form1.frx":A55C1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   256
      Left            =   0
      Picture         =   "Form1.frx":A580B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   255
      Left            =   0
      Picture         =   "Form1.frx":A5A55
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   254
      Left            =   0
      Picture         =   "Form1.frx":A5C9F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   253
      Left            =   0
      Picture         =   "Form1.frx":A5EE9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   252
      Left            =   0
      Picture         =   "Form1.frx":A6133
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   251
      Left            =   0
      Picture         =   "Form1.frx":A637D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   250
      Left            =   0
      Picture         =   "Form1.frx":A65C7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   249
      Left            =   0
      Picture         =   "Form1.frx":A6811
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   248
      Left            =   0
      Picture         =   "Form1.frx":A6A5B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   247
      Left            =   0
      Picture         =   "Form1.frx":A6CA5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   246
      Left            =   0
      Picture         =   "Form1.frx":A6EEF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   245
      Left            =   0
      Picture         =   "Form1.frx":A7139
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   244
      Left            =   0
      Picture         =   "Form1.frx":A7383
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   243
      Left            =   0
      Picture         =   "Form1.frx":A75CD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   242
      Left            =   0
      Picture         =   "Form1.frx":A7817
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   241
      Left            =   0
      Picture         =   "Form1.frx":A7A61
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   240
      Left            =   0
      Picture         =   "Form1.frx":A7CAB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   239
      Left            =   0
      Picture         =   "Form1.frx":A7EF5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   238
      Left            =   0
      Picture         =   "Form1.frx":A813F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   237
      Left            =   0
      Picture         =   "Form1.frx":A8389
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   236
      Left            =   0
      Picture         =   "Form1.frx":A85D3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   235
      Left            =   0
      Picture         =   "Form1.frx":A881D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   234
      Left            =   0
      Picture         =   "Form1.frx":A8A67
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   233
      Left            =   0
      Picture         =   "Form1.frx":A8CB1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   232
      Left            =   0
      Picture         =   "Form1.frx":A8EFB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   231
      Left            =   0
      Picture         =   "Form1.frx":A9145
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   230
      Left            =   0
      Picture         =   "Form1.frx":A938F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   229
      Left            =   0
      Picture         =   "Form1.frx":A95D9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   228
      Left            =   0
      Picture         =   "Form1.frx":A9823
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   227
      Left            =   0
      Picture         =   "Form1.frx":A9A6D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   226
      Left            =   0
      Picture         =   "Form1.frx":A9CB7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   225
      Left            =   0
      Picture         =   "Form1.frx":A9F01
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   224
      Left            =   0
      Picture         =   "Form1.frx":AA14B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   223
      Left            =   0
      Picture         =   "Form1.frx":AA395
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   222
      Left            =   0
      Picture         =   "Form1.frx":AA5DF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   221
      Left            =   0
      Picture         =   "Form1.frx":AA829
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   220
      Left            =   0
      Picture         =   "Form1.frx":AAA73
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   219
      Left            =   0
      Picture         =   "Form1.frx":AACBD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   218
      Left            =   0
      Picture         =   "Form1.frx":AAF07
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   217
      Left            =   0
      Picture         =   "Form1.frx":AB151
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   216
      Left            =   0
      Picture         =   "Form1.frx":AB39B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   215
      Left            =   0
      Picture         =   "Form1.frx":AB5E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   214
      Left            =   0
      Picture         =   "Form1.frx":AB82F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   213
      Left            =   0
      Picture         =   "Form1.frx":ABA79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   212
      Left            =   0
      Picture         =   "Form1.frx":ABCC3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   211
      Left            =   0
      Picture         =   "Form1.frx":ABF0D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   210
      Left            =   0
      Picture         =   "Form1.frx":AC157
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   209
      Left            =   0
      Picture         =   "Form1.frx":AC3A1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   208
      Left            =   0
      Picture         =   "Form1.frx":AC5EB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   207
      Left            =   0
      Picture         =   "Form1.frx":AC835
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   206
      Left            =   0
      Picture         =   "Form1.frx":ACA7F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   205
      Left            =   0
      Picture         =   "Form1.frx":ACCC9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   204
      Left            =   0
      Picture         =   "Form1.frx":ACF13
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   203
      Left            =   0
      Picture         =   "Form1.frx":AD15D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   202
      Left            =   0
      Picture         =   "Form1.frx":AD3A7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   201
      Left            =   0
      Picture         =   "Form1.frx":AD5F1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   200
      Left            =   0
      Picture         =   "Form1.frx":AD83B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   199
      Left            =   0
      Picture         =   "Form1.frx":ADA85
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   198
      Left            =   0
      Picture         =   "Form1.frx":ADCCF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   197
      Left            =   0
      Picture         =   "Form1.frx":ADF19
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   196
      Left            =   0
      Picture         =   "Form1.frx":AE163
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   195
      Left            =   0
      Picture         =   "Form1.frx":AE3AD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   194
      Left            =   0
      Picture         =   "Form1.frx":AE5F7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   193
      Left            =   0
      Picture         =   "Form1.frx":AE841
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   192
      Left            =   0
      Picture         =   "Form1.frx":AEA8B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   191
      Left            =   0
      Picture         =   "Form1.frx":AECD5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   190
      Left            =   0
      Picture         =   "Form1.frx":AEF1F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   189
      Left            =   0
      Picture         =   "Form1.frx":AF169
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   188
      Left            =   0
      Picture         =   "Form1.frx":AF3B3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   187
      Left            =   0
      Picture         =   "Form1.frx":AF5FD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   186
      Left            =   0
      Picture         =   "Form1.frx":AF847
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   185
      Left            =   0
      Picture         =   "Form1.frx":AFA91
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   184
      Left            =   0
      Picture         =   "Form1.frx":AFCDB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   183
      Left            =   0
      Picture         =   "Form1.frx":AFF25
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   182
      Left            =   0
      Picture         =   "Form1.frx":B016F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   181
      Left            =   0
      Picture         =   "Form1.frx":B03B9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   180
      Left            =   0
      Picture         =   "Form1.frx":B0603
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   179
      Left            =   0
      Picture         =   "Form1.frx":B084D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   178
      Left            =   0
      Picture         =   "Form1.frx":B0A97
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   177
      Left            =   0
      Picture         =   "Form1.frx":B0CE1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   176
      Left            =   0
      Picture         =   "Form1.frx":B0F2B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   175
      Left            =   0
      Picture         =   "Form1.frx":B1175
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   174
      Left            =   0
      Picture         =   "Form1.frx":B13BF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   173
      Left            =   0
      Picture         =   "Form1.frx":B1609
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   172
      Left            =   0
      Picture         =   "Form1.frx":B1853
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   171
      Left            =   0
      Picture         =   "Form1.frx":B1A9D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   170
      Left            =   0
      Picture         =   "Form1.frx":B1CE7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   169
      Left            =   0
      Picture         =   "Form1.frx":B1F31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   168
      Left            =   0
      Picture         =   "Form1.frx":B217B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   167
      Left            =   0
      Picture         =   "Form1.frx":B23C5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   166
      Left            =   0
      Picture         =   "Form1.frx":B260F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   165
      Left            =   0
      Picture         =   "Form1.frx":B2859
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   164
      Left            =   0
      Picture         =   "Form1.frx":B2AA3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   163
      Left            =   0
      Picture         =   "Form1.frx":B2CED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   162
      Left            =   0
      Picture         =   "Form1.frx":B2F37
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   161
      Left            =   0
      Picture         =   "Form1.frx":B3181
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   160
      Left            =   0
      Picture         =   "Form1.frx":B33CB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   159
      Left            =   0
      Picture         =   "Form1.frx":B3615
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   158
      Left            =   0
      Picture         =   "Form1.frx":B385F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   157
      Left            =   0
      Picture         =   "Form1.frx":B3AA9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   156
      Left            =   0
      Picture         =   "Form1.frx":B3CF3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   155
      Left            =   0
      Picture         =   "Form1.frx":B3F3D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   154
      Left            =   0
      Picture         =   "Form1.frx":B4187
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   153
      Left            =   0
      Picture         =   "Form1.frx":B43D1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   152
      Left            =   0
      Picture         =   "Form1.frx":B461B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   151
      Left            =   0
      Picture         =   "Form1.frx":B4865
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   150
      Left            =   0
      Picture         =   "Form1.frx":B4AAF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   149
      Left            =   0
      Picture         =   "Form1.frx":B4CF9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   148
      Left            =   0
      Picture         =   "Form1.frx":B4F43
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   147
      Left            =   0
      Picture         =   "Form1.frx":B518D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   146
      Left            =   0
      Picture         =   "Form1.frx":B53D7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   145
      Left            =   0
      Picture         =   "Form1.frx":B5621
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   144
      Left            =   0
      Picture         =   "Form1.frx":B586B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   143
      Left            =   0
      Picture         =   "Form1.frx":B5AB5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   142
      Left            =   0
      Picture         =   "Form1.frx":B5CFF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   141
      Left            =   0
      Picture         =   "Form1.frx":B5F49
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   140
      Left            =   0
      Picture         =   "Form1.frx":B6193
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   139
      Left            =   0
      Picture         =   "Form1.frx":B63DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   138
      Left            =   0
      Picture         =   "Form1.frx":B6627
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   137
      Left            =   0
      Picture         =   "Form1.frx":B6871
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   136
      Left            =   0
      Picture         =   "Form1.frx":B6ABB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   135
      Left            =   0
      Picture         =   "Form1.frx":B6D05
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   134
      Left            =   0
      Picture         =   "Form1.frx":B6F4F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   133
      Left            =   0
      Picture         =   "Form1.frx":B7199
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   132
      Left            =   0
      Picture         =   "Form1.frx":B73E3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   131
      Left            =   0
      Picture         =   "Form1.frx":B762D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   130
      Left            =   0
      Picture         =   "Form1.frx":B7877
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   129
      Left            =   0
      Picture         =   "Form1.frx":B7AC1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   128
      Left            =   0
      Picture         =   "Form1.frx":B7D0B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   127
      Left            =   0
      Picture         =   "Form1.frx":B7F55
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   126
      Left            =   0
      Picture         =   "Form1.frx":B819F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   125
      Left            =   0
      Picture         =   "Form1.frx":B83E9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   124
      Left            =   0
      Picture         =   "Form1.frx":B8633
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   123
      Left            =   0
      Picture         =   "Form1.frx":B887D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   122
      Left            =   0
      Picture         =   "Form1.frx":B8AC7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   121
      Left            =   0
      Picture         =   "Form1.frx":B8D11
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   120
      Left            =   0
      Picture         =   "Form1.frx":B8F5B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   119
      Left            =   0
      Picture         =   "Form1.frx":B91A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   118
      Left            =   0
      Picture         =   "Form1.frx":B93EF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   117
      Left            =   0
      Picture         =   "Form1.frx":B9639
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   116
      Left            =   0
      Picture         =   "Form1.frx":B9883
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   115
      Left            =   0
      Picture         =   "Form1.frx":B9ACD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   114
      Left            =   0
      Picture         =   "Form1.frx":B9D17
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   113
      Left            =   0
      Picture         =   "Form1.frx":B9F61
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   112
      Left            =   0
      Picture         =   "Form1.frx":BA1AB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   111
      Left            =   0
      Picture         =   "Form1.frx":BA3F5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   110
      Left            =   0
      Picture         =   "Form1.frx":BA63F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   109
      Left            =   0
      Picture         =   "Form1.frx":BA889
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   108
      Left            =   0
      Picture         =   "Form1.frx":BAAD3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   107
      Left            =   0
      Picture         =   "Form1.frx":BAD1D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   106
      Left            =   0
      Picture         =   "Form1.frx":BAF67
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   105
      Left            =   0
      Picture         =   "Form1.frx":BB1B1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   104
      Left            =   0
      Picture         =   "Form1.frx":BB3FB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   103
      Left            =   0
      Picture         =   "Form1.frx":BB645
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   102
      Left            =   0
      Picture         =   "Form1.frx":BB88F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   101
      Left            =   0
      Picture         =   "Form1.frx":BBAD9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   100
      Left            =   0
      Picture         =   "Form1.frx":BBD23
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   99
      Left            =   0
      Picture         =   "Form1.frx":BBF6D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   98
      Left            =   0
      Picture         =   "Form1.frx":BC1B7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   97
      Left            =   0
      Picture         =   "Form1.frx":BC401
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   96
      Left            =   0
      Picture         =   "Form1.frx":BC64B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   95
      Left            =   0
      Picture         =   "Form1.frx":BC895
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   94
      Left            =   0
      Picture         =   "Form1.frx":BCADF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   93
      Left            =   0
      Picture         =   "Form1.frx":BCD29
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   92
      Left            =   0
      Picture         =   "Form1.frx":BCF73
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   91
      Left            =   0
      Picture         =   "Form1.frx":BD1BD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   90
      Left            =   0
      Picture         =   "Form1.frx":BD407
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   89
      Left            =   0
      Picture         =   "Form1.frx":BD651
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   88
      Left            =   0
      Picture         =   "Form1.frx":BD89B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   87
      Left            =   0
      Picture         =   "Form1.frx":BDAE5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   86
      Left            =   0
      Picture         =   "Form1.frx":BDD2F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   85
      Left            =   0
      Picture         =   "Form1.frx":BDF79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   84
      Left            =   0
      Picture         =   "Form1.frx":BE1C3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   83
      Left            =   0
      Picture         =   "Form1.frx":BE40D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   82
      Left            =   0
      Picture         =   "Form1.frx":BE657
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   81
      Left            =   0
      Picture         =   "Form1.frx":BE8A1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   80
      Left            =   0
      Picture         =   "Form1.frx":BEAEB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   79
      Left            =   0
      Picture         =   "Form1.frx":BED35
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   78
      Left            =   0
      Picture         =   "Form1.frx":BEF7F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   77
      Left            =   0
      Picture         =   "Form1.frx":BF1C9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   76
      Left            =   0
      Picture         =   "Form1.frx":BF413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   75
      Left            =   0
      Picture         =   "Form1.frx":BF65D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   74
      Left            =   0
      Picture         =   "Form1.frx":BF8A7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   73
      Left            =   0
      Picture         =   "Form1.frx":BFAF1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   72
      Left            =   0
      Picture         =   "Form1.frx":BFD3B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   71
      Left            =   0
      Picture         =   "Form1.frx":BFF85
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   70
      Left            =   0
      Picture         =   "Form1.frx":C01CF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   69
      Left            =   0
      Picture         =   "Form1.frx":C0419
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   68
      Left            =   0
      Picture         =   "Form1.frx":C0663
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   67
      Left            =   0
      Picture         =   "Form1.frx":C08AD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   66
      Left            =   0
      Picture         =   "Form1.frx":C0AF7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   65
      Left            =   0
      Picture         =   "Form1.frx":C0D41
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   64
      Left            =   0
      Picture         =   "Form1.frx":C0F8B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   63
      Left            =   0
      Picture         =   "Form1.frx":C11D5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   62
      Left            =   0
      Picture         =   "Form1.frx":C141F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   61
      Left            =   0
      Picture         =   "Form1.frx":C1669
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   60
      Left            =   0
      Picture         =   "Form1.frx":C18B3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   59
      Left            =   0
      Picture         =   "Form1.frx":C1AFD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   58
      Left            =   0
      Picture         =   "Form1.frx":C1D47
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   57
      Left            =   0
      Picture         =   "Form1.frx":C1F91
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   56
      Left            =   0
      Picture         =   "Form1.frx":C21DB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   55
      Left            =   0
      Picture         =   "Form1.frx":C2425
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   54
      Left            =   0
      Picture         =   "Form1.frx":C266F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   53
      Left            =   0
      Picture         =   "Form1.frx":C28B9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   52
      Left            =   0
      Picture         =   "Form1.frx":C2B03
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   51
      Left            =   0
      Picture         =   "Form1.frx":C2D4D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   50
      Left            =   0
      Picture         =   "Form1.frx":C2F97
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   49
      Left            =   0
      Picture         =   "Form1.frx":C31E1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   48
      Left            =   0
      Picture         =   "Form1.frx":C342B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   47
      Left            =   0
      Picture         =   "Form1.frx":C3675
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   46
      Left            =   0
      Picture         =   "Form1.frx":C38BF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   45
      Left            =   0
      Picture         =   "Form1.frx":C3B09
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   44
      Left            =   0
      Picture         =   "Form1.frx":C3D53
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   43
      Left            =   0
      Picture         =   "Form1.frx":C3F9D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   42
      Left            =   0
      Picture         =   "Form1.frx":C41E7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   41
      Left            =   0
      Picture         =   "Form1.frx":C4431
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   40
      Left            =   0
      Picture         =   "Form1.frx":C467B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   39
      Left            =   0
      Picture         =   "Form1.frx":C48C5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   38
      Left            =   0
      Picture         =   "Form1.frx":C4B0F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   37
      Left            =   0
      Picture         =   "Form1.frx":C4D59
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   36
      Left            =   0
      Picture         =   "Form1.frx":C4FA3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   35
      Left            =   0
      Picture         =   "Form1.frx":C51ED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   34
      Left            =   0
      Picture         =   "Form1.frx":C5437
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   33
      Left            =   0
      Picture         =   "Form1.frx":C5681
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   32
      Left            =   0
      Picture         =   "Form1.frx":C58CB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   31
      Left            =   0
      Picture         =   "Form1.frx":C5B15
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   30
      Left            =   0
      Picture         =   "Form1.frx":C5D5F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   29
      Left            =   0
      Picture         =   "Form1.frx":C5FA9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   28
      Left            =   0
      Picture         =   "Form1.frx":C61F3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   27
      Left            =   0
      Picture         =   "Form1.frx":C643D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   26
      Left            =   0
      Picture         =   "Form1.frx":C6687
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   25
      Left            =   0
      Picture         =   "Form1.frx":C68D1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   24
      Left            =   0
      Picture         =   "Form1.frx":C6B1B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   23
      Left            =   0
      Picture         =   "Form1.frx":C6D65
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   22
      Left            =   0
      Picture         =   "Form1.frx":C6FAF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   21
      Left            =   0
      Picture         =   "Form1.frx":C71F9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   20
      Left            =   0
      Picture         =   "Form1.frx":C7443
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   19
      Left            =   0
      Picture         =   "Form1.frx":C768D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   18
      Left            =   0
      Picture         =   "Form1.frx":C78D7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   17
      Left            =   0
      Picture         =   "Form1.frx":C7B21
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   16
      Left            =   0
      Picture         =   "Form1.frx":C7D6B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   15
      Left            =   0
      Picture         =   "Form1.frx":C7FB5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   14
      Left            =   0
      Picture         =   "Form1.frx":C81FF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   13
      Left            =   0
      Picture         =   "Form1.frx":C8449
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   12
      Left            =   0
      Picture         =   "Form1.frx":C8693
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   11
      Left            =   0
      Picture         =   "Form1.frx":C88DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   10
      Left            =   0
      Picture         =   "Form1.frx":C8B27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   9
      Left            =   0
      Picture         =   "Form1.frx":C8D71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   8
      Left            =   0
      Picture         =   "Form1.frx":C8FBB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   7
      Left            =   0
      Picture         =   "Form1.frx":C9205
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   6
      Left            =   0
      Picture         =   "Form1.frx":C944F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   5
      Left            =   0
      Picture         =   "Form1.frx":C9699
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   4
      Left            =   0
      Picture         =   "Form1.frx":C98E3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   3
      Left            =   0
      Picture         =   "Form1.frx":C9B2D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":C9D77
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   200
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":C9FC1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   200
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   0
      Left            =   18240
      Picture         =   "Form1.frx":CA20B
      Stretch         =   -1  'True
      Top             =   13080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   480
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   1680
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "得分:"
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
      Left            =   11520
      TabIndex        =   4
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   12840
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   195
      Left            =   8520
      Top             =   8505
      Width           =   195
   End
   Begin VB.Line Line1 
      Index           =   760
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   759
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   758
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   757
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   756
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   755
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   754
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   753
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   752
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   751
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   750
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   749
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   748
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   747
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   746
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   745
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   744
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   743
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   742
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   741
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   740
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   739
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   738
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   737
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   736
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   735
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   734
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   733
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   732
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   731
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   730
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   729
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   728
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   727
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   726
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   725
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   724
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   723
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   722
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   721
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   720
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   719
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   718
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   717
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   716
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   715
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   714
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   713
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   712
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   711
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   710
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   709
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   708
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   707
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   706
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   705
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   704
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   703
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   702
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   701
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   700
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   699
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   698
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   697
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   696
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   695
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   694
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   693
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   692
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   691
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   690
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   689
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   688
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   687
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   686
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   685
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   684
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   683
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   682
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   681
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   680
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   679
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   678
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   677
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   676
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   675
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   674
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   673
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   672
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   671
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   670
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   669
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   668
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   667
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   666
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   665
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   664
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   663
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   662
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   661
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   660
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   659
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   658
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   657
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   656
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   655
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   654
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   653
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   652
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   651
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   650
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   649
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   648
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   647
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   646
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   645
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   644
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   643
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   642
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   641
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   640
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   639
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   638
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   637
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   636
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   635
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   634
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   633
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   632
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   631
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   630
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   629
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   628
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   627
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   626
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   625
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   624
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   623
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   622
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   621
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   620
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   619
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   618
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   617
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   616
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   615
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   614
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   613
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   612
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   611
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   610
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   609
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   608
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   607
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   606
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   605
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   604
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   603
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   602
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   601
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   600
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   599
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   598
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   597
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   596
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   595
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   594
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   593
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   592
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   591
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   590
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   589
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   588
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   587
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   586
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   585
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   584
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   583
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   582
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   581
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   580
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   579
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   578
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   577
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   576
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   575
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   574
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   573
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   572
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   571
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   570
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   569
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   568
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   567
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   566
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   565
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   564
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   563
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   562
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   561
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   560
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   559
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   558
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   557
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   556
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   555
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   554
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   553
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   552
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   551
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   550
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   549
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   548
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   547
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   546
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   545
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   544
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   543
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   542
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   541
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   540
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   539
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   538
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   537
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   536
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   535
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   534
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   533
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   532
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   531
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   530
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   529
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   528
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   527
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   526
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   525
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   524
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   523
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   522
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   521
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   520
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   519
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   518
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   517
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   516
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   515
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   514
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   513
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   512
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   511
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   510
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   509
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   508
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   507
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   506
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   505
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   504
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   503
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   502
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   501
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   500
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   499
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   498
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   497
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   496
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   495
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   494
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   493
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   492
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   491
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   490
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   489
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   488
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   487
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   486
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   485
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   484
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   483
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   482
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   481
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   480
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   479
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   478
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   477
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   476
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   475
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   474
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   473
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   472
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   471
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   470
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   469
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   468
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   467
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   466
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   465
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   464
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   463
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   462
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   461
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   460
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   459
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   458
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   457
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   456
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   455
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   454
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   453
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   452
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   451
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   450
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   449
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   448
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   447
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   446
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   445
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   444
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   443
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   442
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   441
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   440
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   439
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   438
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   437
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   436
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   435
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   434
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   433
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   432
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   431
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   430
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   429
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   428
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   427
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   426
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   425
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   424
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   423
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   422
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   421
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   420
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   419
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   418
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   417
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   416
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   415
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   414
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   413
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   412
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   411
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   410
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   409
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   408
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   407
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   406
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   405
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   404
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   403
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   402
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   401
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   400
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   399
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   398
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   397
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   396
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   395
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   394
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   393
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   392
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   391
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   390
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   389
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   388
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   387
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   386
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   385
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   384
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   383
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   382
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   381
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   380
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   379
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   378
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   377
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   376
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   375
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   374
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   373
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   372
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   371
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   370
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   369
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   368
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   367
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   366
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   365
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   364
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   363
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   362
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   361
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   360
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   359
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   358
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   357
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   356
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   355
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   354
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   353
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   352
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   351
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   350
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   349
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   348
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   347
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   346
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   345
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   344
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   343
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   342
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   341
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   340
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   339
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   338
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   337
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   336
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   335
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   334
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   333
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   332
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   331
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   330
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   329
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   328
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   327
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   326
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   325
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   324
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   323
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   322
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   321
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   320
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   319
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   318
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   317
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   316
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   315
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   314
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   313
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   312
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   311
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   310
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   309
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   308
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   307
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   306
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   305
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   304
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   303
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   302
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   301
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   300
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   299
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   298
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   297
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   296
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   295
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   294
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   293
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   292
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   291
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   290
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   289
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   288
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   287
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   286
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   285
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   284
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   283
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   282
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   281
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   280
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   279
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   278
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   277
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   276
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   275
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   274
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   273
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   272
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   271
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   270
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   269
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   268
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   267
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   266
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   265
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   264
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   263
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   262
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   261
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   260
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   259
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   258
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   257
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   256
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   255
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   254
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   253
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   252
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   251
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   250
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   249
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   248
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   247
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   246
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   245
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   244
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   243
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   242
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   241
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   240
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   239
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   238
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   237
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   236
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   235
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   234
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   233
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   232
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   231
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   230
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   229
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   228
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   227
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   226
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   225
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   224
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   223
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   222
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   221
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   220
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   219
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   218
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   217
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   216
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   215
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   214
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   213
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   212
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   211
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   210
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   209
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   208
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   207
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   206
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   205
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   204
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   203
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   202
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   201
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   200
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   480
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
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   179
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   178
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   177
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   176
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   175
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   174
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   173
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   172
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   171
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   170
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   169
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   168
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   167
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   166
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   165
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   164
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   163
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   162
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   161
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   160
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   159
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   158
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   157
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   156
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   155
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   154
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   153
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   152
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   151
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   150
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   149
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   148
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   147
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   146
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   145
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   144
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   143
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   142
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   141
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   140
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   139
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   138
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   137
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   136
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   135
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   134
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   133
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   132
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   131
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   130
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   129
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   128
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   127
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   126
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   125
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   124
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   123
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   122
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   121
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   120
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   119
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   118
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   117
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   116
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   115
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   114
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   113
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   112
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   111
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   110
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   109
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   108
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   107
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   106
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   105
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   104
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   103
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   102
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   101
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   100
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   99
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   98
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   97
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   96
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   95
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   94
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   93
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   92
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   91
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   90
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   89
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   88
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   87
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   86
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   85
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   84
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   83
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   82
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   81
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   80
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   79
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   78
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   77
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   76
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   75
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   74
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   73
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   72
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   71
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   70
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   69
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   68
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   67
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   66
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   65
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   64
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   63
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   62
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   61
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   60
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   59
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   58
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   57
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   56
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   55
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   54
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   53
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   52
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   51
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   50
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   49
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   48
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   47
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   46
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   45
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   44
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   43
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   42
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   41
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   40
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   39
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   38
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   37
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   36
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   35
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   34
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   33
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   32
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   29
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   28
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   26
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   25
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   -1080
      X2              =   120
      Y1              =   -480
      Y2              =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ai As New ADODB.Connection '声明并创建链接对象
Dim rs As ADODB.Recordset '声明链接对象类型
Dim sql$, jg

Private Sub Command1_Click()
'场景布局
For i = 0 To 18
For j = 1 To 10
Line1(i * 10 + j).X1 = 400 * j
Line1(i * 10 + j).X2 = 400 * (j + 1)
Line1(i * 10 + j).Y1 = 400 * (i + 2)
Line1(i * 10 + j).Y2 = 400 * (i + 2)

Line1(i * 10 + j + 190).X1 = 400 * j + 4000
Line1(i * 10 + j + 190).X2 = 400 * (j + 1) + 4000
Line1(i * 10 + j + 190).Y1 = 400 * (i + 2)
Line1(i * 10 + j + 190).Y2 = 400 * (i + 2)

Line1(i * 10 + j + 380).X1 = 400 * (i + 2)
Line1(i * 10 + j + 380).X2 = 400 * (i + 2)
Line1(i * 10 + j + 380).Y1 = 400 * j
Line1(i * 10 + j + 380).Y2 = 400 * (j + 1)

Line1(i * 10 + j + 570).X1 = 400 * (i + 2)
Line1(i * 10 + j + 570).X2 = 400 * (i + 2)
Line1(i * 10 + j + 570).Y1 = 400 * j + 4000
Line1(i * 10 + j + 570).Y2 = 400 * (j + 1) + 4000
Next j, i

For i = 1 To 760
Randomize
Line1(i).Visible = Int(Rnd * 1.5)
Next
'围墙
Line2.X1 = 8800
Line2.X2 = 8800
Line2.Y1 = 0
Line2.Y2 = 8800

Line3.X1 = 8800
Line3.X2 = 0
Line3.Y1 = 8800
Line3.Y2 = 8800

Command2.Enabled = True
Command3.Enabled = True
Command2.SetFocus
'豆排布，随机刷新
For i = 0 To 19
For j = 1 To 10
Image1(i * 10 + j).Left = i * 400 + 500
Image1(i * 10 + j).Top = j * 400 + 100
Image1(i * 10 + j + 200).Left = i * 400 + 500
Image1(i * 10 + j + 200).Top = j * 400 + 4100
Next j, i
For i = 1 To 400
Randomize
Image1(i).Visible = Int(Rnd * 1.5)
Next
'难度设置 控制区 label14.caption
Select Case Label14.Caption
Case 0
Timer2.Interval = 70
Case 1
Timer2.Interval = 60
Case 2
Timer2.Interval = 40
End Select
'其余设置
    '如果模式为无尽模式，则时间从0开始，否则从150开始
    If Label13.Caption = 2 Then Label4.Caption = 0 Else Label4.Caption = 150
    '重置标签内容
    Label1.Caption = 0
    Label6.Caption = 1
    Label10.Caption = 0
    Label12.Caption = 0
    '运行
    Timer1.Enabled = True
    'Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    '重置鬼的位置
    Shape1.Left = 8500: Shape1.Top = 8500
    Shape1.Height = 200: Shape1.Width = 200
    Shape1.Visible = True
    '重置人的位置和其余按钮
    Command2.Top = 200
    Command3.Top = 200
    Command2.Left = 200
    Command3.Left = 200
    Command4.Enabled = False
    Command5.Enabled = True
    Command6.Enabled = False
    Command1.Enabled = False
    '重置技能的位置
    Image2.Left = Int(Rnd * 20 + 1) * 400 + 100
    Image2.Top = Int(Rnd * 20 + 1) * 400 + 100
    Image3.Left = Int(Rnd * 20 + 1) * 400 + 100
    Image3.Top = Int(Rnd * 20 + 1) * 400 + 100
    Image4.Left = Int(Rnd * 20 + 1) * 400 + 100
    Image4.Top = Int(Rnd * 20 + 1) * 400 + 100
    Image5.Left = Int(Rnd * 20 + 1) * 400 + 100
    Image5.Top = Int(Rnd * 20 + 1) * 400 + 100
    '显示与隐藏技能
    Image2.Visible = True
    Image3.Visible = True
    Image4.Visible = True
    Image5.Visible = False
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
'键位移动
a = Asc(Form3.Text1): b = Asc(Form3.Text3)
If 64 < a And a < 90 Then
c = a + 32
Else
c = a - 32
End If
If 64 < b And b < 90 Then
D = b + 32
Else
D = b - 32
End If
If Label19.Caption = "1" Then
    Select Case KeyAscii
    Case a, c
        Command2.Top = Command2.Top - 100
        Command3.Top = Command3.Top - 100

    Case b, D
        Command2.Left = Command2.Left - 100
        Command3.Left = Command3.Left - 100
    Case Else
        Command3.SetFocus
    End Select
    Call pz2(Command2)
    Call pz2(Command3)
ElseIf Label19.Caption = "0" Then
    Select Case KeyAscii
    Case a, c
        Shape1.Top = Shape1.Top - 100
    Case b, D
        Shape1.Left = Shape1.Left - 100
    Case Else
        Command3.SetFocus
    End Select
    Call pz2(Shape1)
End If
Call jw
End Sub

Private Sub Command3_keypress(KeyAscii As Integer)
'键位移动
a = Asc(Form3.Text2): b = Asc(Form3.Text4)
If 64 < a And a < 90 Then
c = a + 32
Else
c = a - 32
End If
If 64 < b And b < 90 Then
D = b + 32
Else
D = b - 32
End If

If Label19.Caption = "1" Then
    Select Case KeyAscii
    Case a, c
        Command2.Top = Command2.Top + 100
        Command3.Top = Command3.Top + 100

    Case b, D
        Command2.Left = Command2.Left + 100
        Command3.Left = Command3.Left + 100
    Case Else
        Command2.SetFocus
    End Select
    Call pz3(Command2)
    Call pz3(Command3)
ElseIf Label19.Caption = "0" Then
    Select Case KeyAscii
    Case a, c
        Shape1.Top = Shape1.Top + 100
    Case b, D
        Shape1.Left = Shape1.Left + 100
    Case Else
        Command2.SetFocus
    End Select
    Call pz3(Shape1)
End If

Call jw

End Sub

Sub pz3(player)

'四周碰墙检测
If player.Top = 8700 And player.Left >= 0 And player.Left <= 8800 Then
player.Top = player.Top - 100
ElseIf player.Left = 8700 And player.Top >= 0 And player.Top <= 8800 Then
player.Left = player.Left - 100
ElseIf player.Top = 0 And player.Left >= 0 And player.Left <= 8800 Then
player.Top = player.Top + 100
ElseIf player.Left = 0 And player.Top >= 0 And player.Top <= 8800 Then
player.Left = player.Left + 100
End If
'横线碰墙检测
For i = 1 To 380
    If Line1(i).Visible = True Then
        If Line1(i).Y1 = Line1(i).Y2 Then
            If player.Top = Line1(i).Y1 - 100 And player.Left + 100 >= Line1(i).X1 And player.Left + 100 <= Line1(i).X2 Then
                player.Top = player.Top - 100
            End If
        End If
    End If
Next i
'竖线碰墙检测
For i = 381 To 760
    If Line1(i).Visible = True Then
        If Line1(i).X1 = Line1(i).X2 Then
            If player.Left = Line1(i).X1 - 100 And player.Top + 100 >= Line1(i).Y1 And player.Top + 100 <= Line1(i).Y2 Then
                 player.Left = player.Left - 100
            End If
        End If
    End If
Next i

End Sub

Sub pz2(player)
'四周碰墙检测
If player.Top = 8700 And player.Left >= 0 And player.Left <= 8800 Then
player.Top = player.Top - 100
ElseIf player.Left = 8700 And player.Top >= 0 And player.Top <= 8800 Then
player.Left = player.Left - 100
ElseIf player.Top = 0 And player.Left >= 0 And player.Left <= 8800 Then
player.Top = player.Top + 100
ElseIf player.Left = 0 And player.Top >= 0 And player.Top <= 8800 Then
player.Left = player.Left + 100
End If
'横线碰墙检测
For i = 1 To 380
    If Line1(i).Visible = True Then
        If Line1(i).Y1 = Line1(i).Y2 Then
            If player.Top = Line1(i).Y1 - 100 And player.Left + 100 >= Line1(i).X1 And player.Left + 100 <= Line1(i).X2 Then
                player.Top = player.Top + 100
            End If
        End If
    End If
Next i
'竖线碰墙检测
For i = 381 To 760
    If Line1(i).Visible = True Then
        If Line1(i).X1 = Line1(i).X2 Then
            If player.Left = Line1(i).X1 - 100 And player.Top + 100 >= Line1(i).Y1 And player.Top + 100 <= Line1(i).Y2 Then
                 player.Left = player.Left + 100
            End If
        End If
    End If
Next i
End Sub

Sub jw()
'无敌时间
If Image2.Visible = True Then
If Image2.Left = Command3.Left And Image2.Top = Command3.Top Then
WindowsMediaPlayer1.URL = App.Path & "\音频\挡我者死.mp3"
Command2.SetFocus
Label9.Caption = 15
Image2.Visible = False
End If
End If
'生命次数
If Image3.Visible = True Then
If Image3.Left = Command3.Left And Image3.Top = Command3.Top Then
WindowsMediaPlayer1.URL = App.Path & "\音频\我要打十个.mp3"
Command2.SetFocus
Image3.Visible = False
Label6.Caption = Label6.Caption + 1
End If
End If
'时间定格
If Image4.Visible = True Then
If Image4.Left = Command3.Left And Image4.Top = Command3.Top Then
WindowsMediaPlayer1.URL = App.Path & "\音频\焯.mp3"
Command2.SetFocus
Image4.Visible = False
Label4.Caption = Label4.Caption + 5
Label12.Caption = 5
Timer5.Enabled = True
End If
End If
'强化技能
If Image5.Visible = True Then
If Image5.Left = Command3.Left And Image5.Top = Command3.Top Then
WindowsMediaPlayer1.URL = App.Path & "\音频\无耻老贼.mp3"
Command2.SetFocus
Image5.Visible = False
Shape1.Visible = False
Timer2.Enabled = False
End If
End If
'吃到豆加一分，随机播放音乐
For i = 1 To 400
    '模式切换 控制区：label13.caption
    If Label13.Caption = 0 Or Label13.Caption = 2 Then
        If Image1(i).Visible = True Then
            If Image1(i).Top = Command3.Top And Image1(i).Left = Command3.Left Then
                Image1(i).Visible = False
                Label1.Caption = Label1.Caption + 1
                b = Int(Rnd * 9)
                Select Case b
                Case 0
                    WindowsMediaPlayer1.URL = App.Path & "\音频\哎呦.mp3"
                Case 1
                    WindowsMediaPlayer1.URL = App.Path & "\音频\篮球.mp3"
                Case 2
                    WindowsMediaPlayer1.URL = App.Path & "\音频\啊！.mp3"
                Case 3
                    WindowsMediaPlayer1.URL = App.Path & "\音频\鸡，你太美。.mp3"
                Case 4
                    WindowsMediaPlayer1.URL = App.Path & "\音频\你干嘛.mp3"
                Case 5
                    WindowsMediaPlayer1.URL = App.Path & "\音频\练习坤.mp3"
                Case 6
                    WindowsMediaPlayer1.URL = App.Path & "\音频\你好烦。.mp3"
                Case 7
                    WindowsMediaPlayer1.URL = App.Path & "\音频\哇呵呵.mp3"
                Case 8
                    WindowsMediaPlayer1.URL = App.Path & "\音频\我特别厉害.mp3"
                End Select
            End If
        End If
    Else
        If Image1(i).Visible = True Then
            If Image1(i).Top = Command3.Top And Image1(i).Left = Command3.Left Then
                Image1(i).Visible = False
                Y = Int(Rnd * 10)
                Randomize
                Select Case Y
                Case 0, 1, 2, 3, 4, 5, 6
                    Label1.Caption = Label1.Caption + 1
                Case 7
                    Label1.Caption = Label1.Caption - 1
                Case 8, 9
                End Select
                b = Int(Rnd * 9)
                Select Case b
                Case 0
                    WindowsMediaPlayer1.URL = App.Path & "\音频\哎呦.mp3"
                Case 1
                    WindowsMediaPlayer1.URL = App.Path & "\音频\篮球.mp3"
                Case 2
                    WindowsMediaPlayer1.URL = App.Path & "\音频\啊！.mp3"
                Case 3
                    WindowsMediaPlayer1.URL = App.Path & "\音频\鸡，你太美。.mp3"
                Case 4
                    WindowsMediaPlayer1.URL = App.Path & "\音频\你干嘛.mp3"
                Case 5
                    WindowsMediaPlayer1.URL = App.Path & "\音频\练习坤.mp3"
                Case 6
                    WindowsMediaPlayer1.URL = App.Path & "\音频\你好烦。.mp3"
                Case 7
                    WindowsMediaPlayer1.URL = App.Path & "\音频\哇呵呵.mp3"
                Case 8
                    WindowsMediaPlayer1.URL = App.Path & "\音频\我特别厉害.mp3"
                End Select
            End If
        End If
    End If
Next
End Sub
Private Sub Command4_Click()   '设置
Label7.Caption = Label7.Caption + 1
Form1.Hide: Form3.Show
Command4.Enabled = False
End Sub

Private Sub Command5_Click()   '暂停
Timer1.Enabled = False: Timer2.Enabled = False: Timer3.Enabled = False: Timer4.Enabled = False: Timer5.Enabled = False
Command4.Enabled = True
Command5.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command6.Enabled = True
Command7.Enabled = True
Form4.Frame1.Enabled = True
End Sub

Private Sub Command6_Click()   '继续
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = False
Command5.Enabled = True
Command6.Enabled = False
Command2.SetFocus
Form4.Frame1.Enabled = False
End Sub

Private Sub Command7_Click()   '挂
f = Len(Label1.Caption + Label4.Caption + Label13.Caption + Label14.Caption)
a = InputBox("一级密码")
If a = "878967" Then
    b = InputBox("二级密码")
    If b = "11912199" Then
        c = InputBox("三级密码")
        If c = Right(Label1.Caption + Label4.Caption + Label13.Caption + Label14.Caption, f) Then
            Form4.Show: GoTo h
        End If
    End If
End If
MsgBox "密码错误", 48
h:
End Sub

Private Sub Command8_Click()   '游戏记录
Form6.Show
End Sub

Private Sub Form_Click()
Form6.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
If Form5.Label10.Caption = "app.path" Then
    ai.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"      '链接数据库的驱动字符
    ai.Open
Else
    ai.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & Form5.Label10.Caption & "\吃豆人数据库\吃豆人.mdb;Persist Security Info=false"     '链接数据库的驱动字符
    ai.Open
End If
a = Int(Rnd * 2)
Open Form5.Label10.Caption & "\联机\" & Label10.Caption & ".txt" For Output As #1
    If a = 1 Then
        Print #1, "0"
        Label19.Caption = "1"
    Else
        Print #1, "1"
        Label19.Caption = "0"
    End If
Close
'readying.Enabled = True
End Sub

Private Sub lplayer2_Click()

End Sub

Private Sub player1and2_Timer()
Dim zb(2)
i = 0
'提取对方坐标位置
Open Form5.Label10.Caption & "\联机\" & Label10.Caption & ".txt" For Input As #1
    Do Until EOF(1)
        Line Input #1, zb(i)
        i = i + 1
    Loop
Close
Command1.Left = zb(0)
Command1.Top = zb(1)
'定入我方坐标位置
Open Form5.Label10.Caption & "\联机\" & Mid(Form2.Label9.Caption, 4) & ".txt" For Output As #2
    Print #2, Command2.Left
    Print #2, Command2.Top
Close
End Sub

Private Sub readying_Timer()
Open Form5.Label10.Caption & "\联机\" & Mid(Form2.Label9.Caption, 4) & ".txt" For Input As #2
    Line Input #1, js
Close
player1and2.Enabled = True
If js = "0" Then Label19.Caption = "0" Else Label19.Caption = "1"
End Sub

Private Sub Timer1_Timer()
'过一分钟，随机刷新豆
Label10.Caption = Label10.Caption + 1
If Val(Label10.Caption) Mod 60 = 0 Then
For i = 1 To 400
Randomize
If Image1(i).Visible = False Then Image1(i).Visible = CBool(Int(Rnd * 2))
Next
End If
If Label13.Caption = 2 Then
    If Val(Label10.Caption) Mod 60 = 0 Then
        Label15.Caption = Label15.Caption + 1
    End If
End If
End Sub

Private Sub Timer2_Timer()
a = Val(Label4.Caption)
'在无尽模式中，每过一分钟，增加鬼的移动速度
If Label13.Caption = 2 Then
    Select Case Label15.Caption
    Case 0
        Timer2.Interval = 70
    Case 1
        Timer2.Interval = 60
    Case 2
        Timer2.Interval = 50
    Case 3
        Timer2.Interval = 40
    Case 4
        Timer2.Interval = 30
    Case 5
        Timer2.Interval = 20
    Case Else
        Timer2.Interval = 10
    End Select
Else
    '难度设置 控制区 label14.caption
    Select Case Label14.Caption
    Case 0
        If a = 90 Then
            Timer2.Interval = 50
        ElseIf a = 45 Then
            Image5.Visible = True
        ElseIf a = 30 Then
            Timer2.Interval = 30
        End If
    Case 1
        If a = 90 Then
            Timer2.Interval = 40
        ElseIf a = 45 Then
            Image5.Visible = True
        ElseIf a = 30 Then
            Timer2.Interval = 20
        End If
    Case 2
        If a = 90 Then
            Timer2.Interval = 20
        ElseIf a = 45 Then
            Image5.Visible = True
        ElseIf a = 30 Then
            Timer2.Interval = 5
        End If
    End Select
End If
'判定按钮位置，让鬼追踪
If Command2.Left < Shape1.Left Then
Shape1.Left = Shape1.Left - 25
End If
If Command2.Left > Shape1.Left Then
Shape1.Left = Shape1.Left + 25
End If
If Command2.Top < Shape1.Top Then
Shape1.Top = Shape1.Top - 25
End If
If Command2.Top > Shape1.Top Then
Shape1.Top = Shape1.Top + 25
End If
End Sub

Private Sub Timer3_Timer()
'判定模式为无尽模式时，则每过一秒，增加一秒，否则减少一秒
If Label13.Caption = 2 Then
Label4.Caption = Label4.Caption + 1
Else
Label4.Caption = Label4.Caption - 1
End If
If Label9.Caption > 0 Then
'判定label9.caption不等于0时，每过一秒，减少一秒，直到为0时停止
Label9.Caption = Label9.Caption - 1
End If
'模式切换 控制区：label13.caption
If Label13.Caption = 0 Then
    '时间为0时，输出结果
    If Label4.Caption = 0 Then
        If Label1.Caption <= 10 Then
            MsgBox "你这都不行？才吃" & Label1.Caption & "个真不行。"
            jieguo = 0
        ElseIf Label1.Caption <= 50 Then
            MsgBox "努努力，加油"
            jieguo = 0
        ElseIf Label1.Caption <= 99 Then
            MsgBox "再接再厉"
            jieguo = 0
        End If
        Call zanting: Timer6.Enabled = True
    End If
    If Label1.Caption >= 100 Then
        MsgBox "你赢了"
        Call zanting: Timer7.Enabled = True: Label16 = 1
    End If
ElseIf Label13.Caption = 1 Then
    If Label4.Caption = 0 Then
        If Label1.Caption <= 10 Then
            MsgBox "你这都不行？才吃" & Label1.Caption & "个真不行。"
        ElseIf Label1.Caption <= 25 Then
            MsgBox "努努力，加油"
            ElseIf Label1.Caption <= 49 Then
            MsgBox "再接再厉"
        End If
        Call zanting: Timer6.Enabled = True
    End If
    If Label1.Caption >= 50 Then
        MsgBox "你赢了"
        Call zanting: Timer7.Enabled = True: Label16 = 1
    End If
End If
End Sub

Private Sub Timer4_Timer()
'判定暂停效果生效时，不执行下列语句
If Val(Label9.Caption) > 0 Then GoTo a
'判定人与鬼重合
If Shape1.Left = Command2.Left And Shape1.Top = Command2.Top Then
    If Val(Label6.Caption) > 0 Then
        '被鬼吃掉后，随机出生在3*3的范围内
        Label6.Caption = Label6.Caption - 1
        Command2.Left = Int(Rnd * 3 + 1) * 400 + 100
        Command3.Left = Command2.Left
        Command2.Top = Int(Rnd * 3 + 1) * 400 + 100
        Command3.Top = Command2.Top
        Shape1.Left = 8500: Shape1.Top = 8500
        GoTo a
    End If
Else: GoTo a
End If
'模式切换 控制区：label13.caption
If Label13.Caption = 0 Then
    '被鬼追到输出结果
    If Label1.Caption <= 10 Then
        MsgBox "你这都不行？才吃" & Label1.Caption & "个真不行。"
        Call zanting
        Label16 = 0
        Timer7.Enabled = True
    ElseIf Label1.Caption <= 50 Then
        MsgBox "努努力，加油"
        Call zanting
        Label16 = 0
        Timer7.Enabled = True
    ElseIf Label1.Caption <= 99 Then
        MsgBox "再接再厉"
        Call zanting
        Label16 = 0
        Timer7.Enabled = True
    End If

ElseIf Label13.Caption = 1 Then
    If Label1.Caption <= 10 Then
        MsgBox "你这都不行？才吃" & Label1.Caption & "个真不行。"
        Call zanting
        Label16 = 0
        Timer7.Enabled = True
    ElseIf Label1.Caption <= 25 Then
        MsgBox "努努力，加油"
        Call zanting
        Label16 = 0
        Timer7.Enabled = True
    ElseIf Label1.Caption <= 49 Then
        MsgBox "再接再厉"
        Call zanting
        Label16 = 0
        Timer7.Enabled = True
    End If
    If Label1.Caption >= 100 Then
        MsgBox "你赢了"
        Call zanting
        Label16 = 1
        Timer7.Enabled = True
        End If
    If Label1.Caption >= 50 Then
        Label16 = 1
        MsgBox "你赢了"
        Call zanting
        Timer7.Enabled = True
    End If
Else
    If Shape1.Left = Command2.Left And Shape1.Top = Command2.Top Then
        MsgBox "恭喜你得了" & Label1.Caption & "分", 64
        Call zanting
        Timer8.Enabled = True
    End If
End If
a:
End Sub

Sub zanting()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Command5.Enabled = False
End Sub

Private Sub Timer5_Timer()
'让鬼暂停
Timer2.Enabled = False
'每过一秒，减少一秒，直到为0时停止
Label12.Caption = Label12.Caption - 1
'当label12.caption = 0 时，鬼继续运动
If Val(Label12.Caption) = 0 Then
Timer2.Enabled = True
Timer5.Enabled = False
End If
End Sub

Private Sub Timer6_Timer()
 l = "输"
Set rs = New ADODB.Recordset '数据结果记录到<用户>数据库
sql = "insert into " & Form5.Text4 & "(模式,难度,得分,游玩时间,结果,时间) values (' " & Trim(Form2.Combo1.Text) & " ', ' " & Trim(Form2.Combo2.Text) & " ',' " & Trim(Label1.Caption) & " ',' " & Trim(Str(150 - Val(Label4.Caption))) & " ', ' " & l & " ',' " & Trim(Str(Now)) & " ')"
          'insert into <表名> ([字段1],[字段2]……) values ('字段值1','字段2'，……)
          'insert into <表名> form array <数组名>
rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
Timer6.Enabled = False
Set rs = Nothing
End Sub

Private Sub Timer7_Timer()
If Label16 = "1" Then jieguo = "赢" Else jieguo = "输"
Set rs = New ADODB.Recordset '数据结果记录到<用户>数据库
sql = "insert into " & Form5.Text4 & "(模式,难度,得分,游玩时间,结果,时间) values (' " & Trim(Form2.Combo1.Text) & " ', ' " & Trim(Form2.Combo2.Text) & " ',' " & Trim(Label1.Caption) & " ',' " & Trim(Str(150 - Val(Label4.Caption))) & " ',' " & jieguo & " ',' " & Trim(Str(Now)) & " ')"
      'insert into <表名> ([字段1],[字段2]……) values ('字段值1','字段2'，……)
      'insert into <表名> form array <数组名>
rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
Timer7.Enabled = False
Set rs = Nothing
End Sub

Private Sub Timer8_Timer()
tr = "  无"
Set rs = New ADODB.Recordset '数据结果记录到<用户>数据库
sql = "insert into " & Form5.Text4 & "(模式,难度,得分,游玩时间,时间) values (' " & Trim(Form2.Combo1.Text) & " ', ' " & tr & " ',' " & Trim(Label1.Caption) & " ',' " & Trim(Label4.Caption) & " ',' " & Trim(Str(Now)) & " ')"
      'insert into <表名> ([字段1],[字段2]……) values ('字段值1','字段2'，……)
      'insert into <表名> form array <数组名>
rs.Open sql, ai, adOpenKeyset, adLockBatchOptimistic
Timer8.Enabled = False
Set rs = Nothing
End Sub

Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)
Command2.SetFocus
End Sub
