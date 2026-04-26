VERSION 5.00
Begin VB.UserControl AnnouncementText 
   ClientHeight    =   9675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9960
   ScaleHeight     =   9675
   ScaleWidth      =   9960
   Begin VB.VScrollBar VScroll1 
      Height          =   8175
      LargeChange     =   5
      Left            =   9360
      Max             =   30
      Min             =   10
      SmallChange     =   5
      TabIndex        =   2
      Top             =   960
      Value           =   10
      Width           =   255
   End
   Begin VB.Timer LoadText 
      Interval        =   100
      Left            =   6840
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   240
      ScaleHeight     =   8145
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      Begin VB.Label TextStr 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2220
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   3810
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   165
   End
End
Attribute VB_Name = "AnnouncementText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'缺省属性值:
Const m_def_Enabled = False
Const m_def_LoadFile = ""
Const m_def_Title = "Title1"
'属性变量:
Dim m_Enabled As Boolean
Dim m_LoadFile As String
Dim m_Title As String
Dim a


'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,Title1
Public Property Get Title() As String
Attribute Title.VB_Description = "标题名称"
    Title = m_Title

End Property

Public Property Let Title(ByVal New_Title As String)
    m_Title = New_Title
    PropertyChanged "Title"
End Property

Private Sub Command1_Click()
MsgBox LoadText.Enabled
MsgBox a
LoadText.Enabled = True
End Sub

Private Sub LoadText_Timer()
Label2.Caption = Title
Open LoadFile For Input As #1
    Picture1.Cls
    Do Until EOF(1)
        Line Input #1, gh
        a = a & Chr(10) & gh
    Loop
Close
TextStr.Caption = a
If Picture1.Height < TextStr.Height Then VScroll1.Max = TextStr.Height - Picture1.Height Else VScroll1.Enabled = False
LoadText.Enabled = False
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_Title = m_def_Title
    m_LoadFile = m_def_LoadFile
    m_Enabled = m_def_Enabled
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Title = PropBag.ReadProperty("Title", m_def_Title)
    m_LoadFile = PropBag.ReadProperty("LoadFile", m_def_LoadFile)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    LoadFile = "c:\1.txt"
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
    Call PropBag.WriteProperty("LoadFile", m_LoadFile, m_def_LoadFile)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

Private Sub VScroll1_Change()
TextStr.Top = -(VScroll1.Value)
End Sub

Private Sub VScroll1_Scroll()
TextStr.Top = -(VScroll1.Value)
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get LoadFile() As String
Attribute LoadFile.VB_Description = "请在本工程的同级或子级目录下选择文件"
    LoadFile = m_LoadFile
End Property

Public Property Let LoadFile(ByVal New_LoadFile As String)
    m_LoadFile = New_LoadFile
    PropertyChanged "LoadFile"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,2,False
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，响应LoadFile属性"
Attribute Enabled.VB_MemberFlags = "400"
    LoadText.Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

