VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8745
   ClientLeft      =   6000
   ClientTop       =   3090
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   15150
   Begin VB.CommandButton Command1 
      Caption         =   "Éý¼¶"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   1440
      ScaleHeight     =   7425
      ScaleWidth      =   12825
      TabIndex        =   2
      Top             =   600
      Width           =   12855
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   4815
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   4215
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Caption         =   "¹ºÂò"
            Height          =   495
            Left            =   2640
            TabIndex        =   5
            Top             =   1800
            Width           =   975
         End
         Begin VB.Image Image1 
            Height          =   1575
            Left            =   600
            Top             =   960
            Width           =   1575
         End
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


