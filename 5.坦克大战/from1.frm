VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "坦克大战"
   ClientHeight    =   11625
   ClientLeft      =   1890
   ClientTop       =   630
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   11625
   ScaleWidth      =   12150
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   8000
      ScaleHeight     =   5625
      ScaleWidth      =   4125
      TabIndex        =   23
      Top             =   5670
      Width           =   4150
      Begin VB.Label df 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Left            =   840
         TabIndex        =   40
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "得分:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   645
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4150
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "技能："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   38
         Top             =   1440
         Width           =   765
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   1920
         X2              =   1920
         Y1              =   1320
         Y2              =   5640
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   7
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   6
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   5
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   4
         Left            =   120
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "键位："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   4
         Left            =   2160
         TabIndex        =   37
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "次数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   3
         Left            =   3240
         TabIndex        =   36
         Top             =   1440
         Width           =   765
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   3000
         X2              =   3000
         Y1              =   1320
         Y2              =   5640
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   7
         Left            =   720
         TabIndex        =   35
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   6
         Left            =   720
         TabIndex        =   34
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   5
         Left            =   720
         TabIndex        =   33
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   4
         Left            =   720
         TabIndex        =   32
         Top             =   4920
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   7
         Left            =   2400
         TabIndex        =   31
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   6
         Left            =   2400
         TabIndex        =   30
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   5
         Left            =   2400
         TabIndex        =   29
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   4
         Left            =   2400
         TabIndex        =   28
         Top             =   4920
         Width           =   135
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   0
         X2              =   4150
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   0
         X2              =   4150
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   0
         X2              =   4150
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   0
         X2              =   4150
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   7
         Left            =   3480
         TabIndex        =   27
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   6
         Left            =   3480
         TabIndex        =   26
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   5
         Left            =   3480
         TabIndex        =   25
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   4
         Left            =   3480
         TabIndex        =   24
         Top             =   4920
         Width           =   135
      End
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   1
      Left            =   9480
      Top             =   9360
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   1
      Left            =   9480
      Top             =   8880
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1
      Left            =   9480
      Top             =   8400
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   5
      Left            =   9000
      Top             =   9360
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   5
      Left            =   9000
      Top             =   8880
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   5
      Left            =   9000
      Top             =   8400
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Left            =   9240
      Top             =   7440
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   9480
      Top             =   7920
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5
      Left            =   9000
      Top             =   7920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   22
      Top             =   10680
      Width           =   1215
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9000
      Top             =   6960
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   9480
      Top             =   6960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6600
      TabIndex        =   21
      Top             =   10680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   9240
      Top             =   6480
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9480
      Top             =   6000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   10680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   10680
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9000
      Top             =   6000
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   8000
      ScaleHeight     =   5625
      ScaleWidth      =   4125
      TabIndex        =   4
      Top             =   0
      Width           =   4150
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Left            =   840
         TabIndex        =   41
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   3
         Left            =   3480
         TabIndex        =   20
         Top             =   4920
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   19
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   1
         Left            =   3480
         TabIndex        =   18
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   0
         Left            =   3480
         TabIndex        =   17
         Top             =   2040
         Width           =   135
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   0
         X2              =   4150
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   0
         X2              =   4150
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   4150
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   0
         X2              =   4150
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   3
         Left            =   2400
         TabIndex        =   16
         Top             =   4920
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   2
         Left            =   2400
         TabIndex        =   15
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   1
         Left            =   2400
         TabIndex        =   14
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   0
         Left            =   2400
         TabIndex        =   13
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   4920
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   2
         Left            =   720
         TabIndex        =   11
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   1
         Left            =   720
         TabIndex        =   10
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   0
         Left            =   720
         TabIndex        =   9
         Top             =   2040
         Width           =   135
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   3000
         X2              =   3000
         Y1              =   1320
         Y2              =   5640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "次数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   2
         Left            =   3240
         TabIndex        =   8
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "键位："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   765
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   3
         Left            =   120
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   2
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   495
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   1920
         X2              =   1920
         Y1              =   1320
         Y2              =   5640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "技能："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   1440
         Width           =   765
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4150
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "坦克血量："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00660002&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   840
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   1080
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   2280
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   4800
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8000
      X2              =   8000
      Y1              =   0
      Y2              =   11640
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   400
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   399
      Left            =   -120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   398
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   397
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   396
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   395
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   394
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   393
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   392
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   391
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   390
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   389
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   388
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   387
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   386
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   385
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   384
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   383
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   382
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   381
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   380
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   379
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   378
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   377
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   376
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   374
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   373
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   372
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   371
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   370
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   369
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   368
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   367
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   366
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   365
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   364
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   363
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   362
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   361
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   360
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   359
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   358
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   357
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   356
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   355
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   354
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   353
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   352
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   351
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   350
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   349
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   348
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   347
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   346
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   345
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   344
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   343
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   342
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   341
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   340
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   339
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   338
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   337
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   336
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   334
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   333
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   331
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   330
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   329
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   328
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   327
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   326
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   325
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   324
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   323
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   322
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   321
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   320
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   319
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   318
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   317
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   316
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   315
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   314
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   313
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   312
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   311
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   310
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   309
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   308
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   307
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   306
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   305
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   304
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   303
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   302
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   301
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   300
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   299
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   298
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   297
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   296
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   295
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   294
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   293
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   292
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   291
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   290
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   289
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   288
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   287
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   286
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   285
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   284
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   283
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   282
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   281
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   280
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   279
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   278
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   277
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   276
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   275
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   274
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   273
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   272
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   271
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   270
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   269
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   268
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   267
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   266
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   265
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   264
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   263
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   262
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   261
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   260
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   259
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   258
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   257
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   256
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   254
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   253
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   252
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   251
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   250
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   249
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   248
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   247
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   246
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   245
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   244
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   243
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   242
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   241
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   240
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   239
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   238
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   237
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   236
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   235
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   234
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   233
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   232
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   231
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   230
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   229
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   228
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   227
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   226
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   225
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   224
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   223
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   222
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   221
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   220
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   219
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   218
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   217
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   216
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   214
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   213
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   211
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   210
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   209
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   208
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   207
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   206
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   205
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   204
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   203
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   202
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   201
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   200
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   199
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   198
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   197
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   196
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   195
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   194
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   193
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   192
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   191
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   190
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   189
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   188
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   187
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   186
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   185
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   184
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   183
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   182
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   181
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   180
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   179
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   178
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   177
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   176
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   175
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   174
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   173
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   172
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   171
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   170
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   169
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   168
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   167
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   166
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   165
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   164
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   163
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   162
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   161
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   160
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   159
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   158
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   157
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   156
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   155
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   154
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   153
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   152
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   151
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   150
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   149
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   148
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   147
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   146
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   145
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   144
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   143
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   142
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   141
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   140
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   139
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   138
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   137
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   136
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   135
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   134
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   133
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   132
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   131
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   130
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   129
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   128
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   127
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   126
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   125
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   124
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   123
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   122
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   121
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   120
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   119
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   118
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   117
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   116
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   115
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   114
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   113
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   112
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   111
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   110
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   109
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   108
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   107
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   106
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   105
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   104
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   103
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   102
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   101
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   100
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   99
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   98
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   97
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   96
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   95
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   94
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   93
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   92
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   91
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   90
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   89
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   88
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   87
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   86
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   85
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   84
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   83
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   82
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   81
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   80
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   79
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   78
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   77
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   76
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   75
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   74
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   73
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   72
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   71
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   70
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   69
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   68
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   67
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   66
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   65
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   64
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   63
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   62
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   61
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   60
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   59
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   58
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   57
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   56
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   55
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   54
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   53
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   52
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   51
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   50
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   49
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   48
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   47
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   46
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   45
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   44
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   43
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   42
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   41
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   40
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   39
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   38
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   37
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   36
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   35
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   34
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   33
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   32
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   31
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   30
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   29
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   28
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   27
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   26
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   25
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   24
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   23
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   22
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   21
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   20
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   19
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   18
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   17
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   16
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   15
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   14
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   13
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   12
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   11
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   10
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   9
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   8
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   7
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   6
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   5
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   4
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   3
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   400
      Left            =   3240
      Picture         =   "from1.frx":0000
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   400
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   105
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   400
      Index           =   0
      Left            =   3960
      Picture         =   "from1.frx":1863
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   400
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Index           =   1
      Left            =   2040
      Picture         =   "from1.frx":1DED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   825
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Index           =   3
      Left            =   1560
      Picture         =   "from1.frx":2377
      Stretch         =   -1  'True
      Top             =   840
      Width           =   825
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Index           =   2
      Left            =   360
      Picture         =   "from1.frx":2901
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   825
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      Height          =   105
      Index           =   3
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   105
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      Height          =   105
      Index           =   2
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   105
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      Height          =   105
      Index           =   1
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   105
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      Height          =   105
      Index           =   0
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p(1 To 400), wtk As Boolean, fxm As Integer, dtk(3) As Boolean, dfxm(3), dtkm(3)

Private Sub Command1_Click()
'Randomize
'场景布局
For i = 0 To 38 Step 2
    For j = 1 To 20
        Image1(i * 10 + j).Stretch = True
        Image1(i * 10 + j).Height = 398
        Image1(i * 10 + j).Width = 398
        Image1(i * 10 + j).Left = 400 * (i \ 2) + 1
        Image1(i * 10 + j).Top = 400 * j + 801
    Next j
Next i

For i = 1 To 400
    Image1(i).Visible = CBool(Int(Rnd * 1.5))
    '加载图片场景
    If Image1(i).Visible = True Then
        d = Int(Rnd * 3)
        If d = 0 Then Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\铁砖.jpg"): p(i) = d
        If d = 1 Then Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\草.jpg"): p(i) = d
        If d = 2 Then Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\砖.jpg"): p(i) = d
    End If
Next i
Call wtksc
Label7.Caption = 100
Timer5.Enabled = True: Timer4.Enabled = True
End Sub

Sub wtksc()   '我方坦克
Image2.Height = 400: Image2.Width = 400
Image2.Top = 9200: Image2.Left = 3600
Image2.Visible = True
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
'按键移动
Select Case KeyAscii
Case 87, 119
    Image2.Top = Image2.Top - 50: fxm = 8
    Image2.Picture = LoadPicture(App.Path & "\坦克图片\金上.jpg")
    
Case 65, 97
    Image2.Left = Image2.Left - 50: fxm = 4
    Image2.Picture = LoadPicture(App.Path & "\坦克图片\金左.jpg")
    
Case 83, 115
    Image2.Top = Image2.Top + 50: fxm = 2
    Image2.Picture = LoadPicture(App.Path & "\坦克图片\金下.jpg")
    
Case 68, 100
    Image2.Left = Image2.Left + 50: fxm = 6
    Image2.Picture = LoadPicture(App.Path & "\坦克图片\金右.jpg")
    
Case 82, 114
    If Not wtk Then
        Timer1.Enabled = True
        Shape1.Left = Image2.Left + 150: Shape1.Top = Image2.Top + 150
        Shape1.Visible = True
        wtk = True
    End If
End Select
For i = 1 To 400
    '检测image1控件数组是否有隐藏
    If Image1(i).Visible = True Then
        '检测墙体与坦克的碰撞
        Select Case pz.fx(Image1(i), Image2, p(i))
        Case 6
            Image2.Left = Image2.Left + 50: Exit For
        Case 4
            Image2.Left = Image2.Left - 50: Exit For
        Case 8
            Image2.Top = Image2.Top + 50: Exit For
        Case 2
            Image2.Top = Image2.Top - 50: Exit For
        End Select
    End If
Next
End Sub

Private Sub Command3_Click()
MsgBox pz.fx(Image1(15), Image2, p(15))
End Sub

Private Sub Form_Load()
'初始化游戏模型
Text2 = wtk: Label7.Caption = 100
Shape1.Width = 100: Shape1.Height = 100
For i = 0 To 3
    Image4(i).Visible = False
    Image4(i).Width = 400: Image4(i).Height = 400
    Image4(i).Left = 1200: Image4(i).Top = 400
    Shape5(i).Width = 400: Shape5(i).Height = 400
    Shape2(i).Width = 100: Shape2(i).Height = 100
Next
End Sub

Sub zdwz()
'子弹停止移动（我方）
Shape1.Top = 0: Shape1.Left = 0: wtk = False
Shape1.Visible = False: Timer1.Enabled = False: Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
'子弹发射方向（我方）
Select Case fxm
Case 2
    Text1.Text = "down"
Case 4
    Text1.Text = "left"
Case 6
    Text1.Text = "right"
Case 8
    Text1.Text = "up"
End Select
Timer1.Enabled = False: Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'根据子弹发射方向结果，判定子弹向哪发射（我方）
If Text1.Text = "down" Then Shape1.Top = Shape1.Top + 100
If Text1.Text = "up" Then Shape1.Top = Shape1.Top - 100
If Text1.Text = "left" Then Shape1.Left = Shape1.Left - 100
If Text1.Text = "right" Then Shape1.Left = Shape1.Left + 100

'子弹与墙体碰撞检测结果
For i = 1 To 400
    If Image1(i).Visible = True Then
        '检测物体是否为铁块
        If (Not pz.zdqt(Image1(i), Shape1) = "" And p(i) = 0) Then Call zdwz: Exit For
        '检测砖块是否为半砖
        If (Not pz.zdqt(Image1(i), Shape1) = "" And p(i) = 2) And ((0 <= Image1(i).Width And Image1(i).Width <= 200) Or (0 <= Image1(i).Height And Image1(i).Height <= 200)) Then
            Image1(i).Visible = False: Call zdwz: Exit For
        Else
            '检测砖块与子弹碰撞结果
            js = pz.zdqt(Image1(i), Shape1)
            Select Case True
            Case (js = "right" And p(i) = 2)
                Image1(i).Width = 198: Timer2.Enabled = False: Call zdwz: Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\竖半砖.jpg"): Exit For
            Case (js = "left" And p(i) = 2)
                Image1(i).Width = 199: Image1(i).Left = Image1(i).Left + 200: Timer2.Enabled = False: Call zdwz: Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\竖半砖.jpg"): Exit For
            Case (js = "down" And p(i) = 2)
                Image1(i).Height = 198: Timer2.Enabled = False: Call zdwz: Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\横半砖.jpg"): Exit For
            Case (js = "up" And p(i) = 2)
                Image1(i).Height = 199: Image1(i).Top = Image1(i).Top + 200: Timer2.Enabled = False: Call zdwz: Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\横半砖.jpg"): Exit For
            End Select
        End If
    End If
Next
For i = 0 To 3
    If pz.zdtk(Image4(i), Shape1) Then Image4(i).Visible = False: Call zdwz: df = df + 1: Call dtksw(df)
Next
End Sub

Function dtksw(bm As Integer)   '敌方坦克
Image4(bm).Top = 0: Image4(bm).Left = 0
End Function

Private Sub Timer3_Timer()
Text2.Text = wtk
'检测子弹是否穿过墙体四周
If Shape1.Left < 0 Or Shape1.Top < 0 Or Shape1.Left > 8000 Or Shape1.Top > Form1.Height Then Call zdwz
For i = 0 To 3
    If Shape2(i).Left < 0 Or Shape2(i).Top < 0 Or Shape2(i).Left > 8000 Or Shape2(i).Top > Form1.Height Then Call dzdwz(Shape2(j)): dtk(i) = False
Next
If Label7.Caption = 0 Then MsgBox "你死了。": Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Static dtkma(3)
'判定四个敌方坦克是否死亡，如果死亡则再次随机四个坦克
If Image4(0).Visible = False And Image4(1).Visible = False And Image4(2).Visible = False And Image4(3).Visible = False Then
    For i = 0 To 3
        dtkma(i) = Int(Rnd * 2)
        Image4(i).Left = Int(Rnd * 20) * 400: Image4(i).Top = Int(Rnd * 3) * 400
        Image4(i).Visible = True
    Next
End If
For j = 0 To 3
    If Image4(j).Visible Then
        '坦克移动
        Select Case dtkma(j)
        Case 0
            If Shape5(j).Left > Image4(j).Left Then
                Image4(j).Left = Image4(j).Left + 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\重坦右.jpg"): dfxm(j) = 6
            ElseIf Shape5(j).Left < Image4(j).Left Then
                Image4(j).Left = Image4(j).Left - 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\重坦左.jpg"): dfxm(j) = 4
            ElseIf Shape5(j).Top > Image4(j).Top Then
                Image4(j).Top = Image4(j).Top + 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\重坦下.jpg"): dfxm(j) = 2
            ElseIf Shape5(j).Top < Image4(j).Top Then
                Image4(j).Top = Image4(j).Top - 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\重坦上.jpg"): dfxm(j) = 8
            End If
        Case 1
            If Shape5(j).Left > Image4(j).Left Then
                Image4(j).Left = Image4(j).Left + 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\轻坦右.jpg"): dfxm(j) = 6
            ElseIf Shape5(j).Left < Image4(j).Left Then
                Image4(j).Left = Image4(j).Left - 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\轻坦左.jpg"): dfxm(j) = 4
            ElseIf Shape5(j).Top > Image4(j).Top Then
                Image4(j).Top = Image4(j).Top + 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\轻坦下.jpg"): dfxm(j) = 2
            ElseIf Shape5(j).Top < Image4(j).Top Then
                Image4(j).Top = Image4(j).Top - 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\轻坦上.jpg"): dfxm(j) = 8
            End If
        Case 2
            If Shape5(j).Left > Image4(j).Left Then
                Image4(j).Left = Image4(j).Left + 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\装甲右.jpg"): dfxm(j) = 6
            ElseIf Shape5(j).Left < Image4(j).Left Then
                Image4(j).Left = Image4(j).Left - 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\装甲左.jpg"): dfxm(j) = 4
            ElseIf Shape5(j).Top > Image4(j).Top Then
                Image4(j).Top = Image4(j).Top + 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\装甲下.jpg"): dfxm(j) = 2
            ElseIf Shape5(j).Top < Image4(j).Top Then
                Image4(j).Top = Image4(j).Top - 50: Image4(j).Picture = LoadPicture(App.Path & "\坦克图片\装甲上.jpg"): dfxm(j) = 8
            End If
        End Select
        For i = 1 To 400
            '检测image1控件数组是否有隐藏
            If Image1(i).Visible = True Then
                '检测敌方坦克与墙体碰撞
                Select Case pz.fx(Image1(i), Image4(j), p(i))
                Case 2
                    Image4(j).Top = Image4(j).Top - 50: Timer5.Enabled = True
                Case 4
                    Image4(j).Left = Image4(j).Left - 50: Timer5.Enabled = True
                Case 6
                    Image4(j).Left = Image4(j).Left + 50: Timer5.Enabled = True
                Case 8
                    Image4(j).Top = Image4(j).Top + 50: Timer5.Enabled = True
                End Select
            End If
        Next i
        Randomize
        '坦克随机发射子弹（概率为5%）
        fs = Int(Rnd * 100) + 1
        If fs < 5 Then
            If Not dtk(j) Then
                Timer7(j).Enabled = True
                Shape2(j).Left = Image4(j).Left + 150: Shape2(j).Top = Image4(j).Top + 150
                Shape2(j).Visible = True
                dtk(j) = True
            End If
        End If
        If Image4(j).Left = Shape5(j).Left And Image4(j).Top = Shape5(j).Top Then Timer5.Enabled = True
    End If
Next j
End Sub

Private Sub Timer5_Timer()
'坦克随机移动
'随机left和top
Randomize
For i = 0 To 3
    lt = CBool(Int(Rnd * 2))
    If lt Then
        'top上移动
        Shape5(i).Left = Image4(i).Left
        Shape5(i).Top = Int(Rnd * 24) * 400
    Else
        'left上移动
        Shape5(i).Top = Image4(i).Top
        Shape5(i).Left = Int(Rnd * 20) * 400
    End If
Next
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer(Index As Integer)
Call fsjg(Index)
End Sub

Function fsjg(bm As Integer)
    '根据子弹发射方向结果，判定子弹向哪发射（敌方）
    Select Case dtkm(bm)
    Case "down"
        Shape2(bm).Top = Shape2(bm).Top + 100
    Case "up"
        Shape2(bm).Top = Shape2(bm).Top - 100
    Case "left"
        Shape2(bm).Left = Shape2(bm).Left - 100
    Case "right"
        Shape2(bm).Left = Shape2(bm).Left + 100
    End Select
    '子弹与墙体碰撞检测结果
    For i = 1 To 400
        If Image1(i).Visible = True Then
        '检测物体是否为铁块
            If (Not pz.zdqt(Image1(i), Shape2(bm)) = "" And p(i) = 0) Then Call dzdwz(Shape2(bm)): dtk(bm) = False: Exit For
                '检测砖块是否为半砖
                If (Not pz.zdqt(Image1(i), Shape2(bm)) = "" And p(i) = 2) And ((0 <= Image1(i).Width And Image1(i).Width <= 200) Or (0 <= Image1(i).Height And Image1(i).Height <= 200)) Then
                    Image1(i).Visible = False: Call dzdwz(Shape2(bm)): dtk(bm) = False: Exit For
                Else
                    '检测砖块与子弹碰撞结果
                    js = pz.zdqt(Image1(i), Shape2(bm))
                    Select Case True
                    Case (js = "right" And p(i) = 2)
                        Image1(i).Width = 198: Timer6(bm).Enabled = False: Call dzdwz(Shape2(bm)): dtk(bm) = False: Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\竖半砖.jpg"): Exit For
                    Case (js = "left" And p(i) = 2)
                        Image1(i).Width = 199: Image1(i).Left = Image1(i).Left + 200: Timer6(bm).Enabled = False: Call dzdwz(Shape2(bm)): dtk(bm) = False: Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\竖半砖.jpg"): Exit For
                    Case (js = "down" And p(i) = 2)
                        Image1(i).Height = 198: Timer6(bm).Enabled = False: Call dzdwz(Shape2(bm)): dtk(bm) = False: Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\横半砖.jpg"): Exit For
                    Case (js = "up" And p(i) = 2)
                        Image1(i).Height = 199: Image1(i).Top = Image1(i).Top + 200: Timer6(bm).Enabled = False: Call dzdwz(Shape2(bm)): dtk(bm) = False: Image1(i).Picture = LoadPicture(App.Path & "\坦克图片\横半砖.jpg"): Exit For
                    End Select
                End If
        End If
    Next
    If pz.zdtk(Image2, Shape2(bm)) Then Label7.Caption = Label7.Caption - 25: Timer6(bm).Enabled = False: Call dzdwz(Shape2(bm)): dtk(bm) = False
End Function

Function dzdwz(zd As Control)
zd.Top = 0: zd.Left = 0
zd.Visible = False
End Function

Private Sub Timer7_Timer(Index As Integer)
'子弹发射方向（敌方）
Call dtkfxm(Index)
Timer7(Index).Enabled = False
End Sub

Function dtkfxm(bm As Integer)
If dtk(bm) Then
    Shape2(bm).BackColor = &HFFFFFF
    Select Case dfxm(bm)
    Case 2
        dtkm(bm) = "down": Timer6(bm).Enabled = True
    Case 4
        dtkm(bm) = "left": Timer6(bm).Enabled = True
    Case 6
        dtkm(bm) = "right": Timer6(bm).Enabled = True
    Case 8
        dtkm(bm) = "up": Timer6(bm).Enabled = True
    End Select
End If
End Function
