VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   ScaleHeight     =   6225
   ScaleWidth      =   7095
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   7435
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock FileTransferSock 
      Left            =   6240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    TreeView.PathSeparator = App.Path
End Sub
