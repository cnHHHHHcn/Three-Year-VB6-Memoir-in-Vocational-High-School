VERSION 5.00
Begin VB.UserControl ProgressBar 
   BackStyle       =   0  '透明
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   PropertyPages   =   "ProgressBar.ctx":0000
   ScaleHeight     =   90
   ScaleWidth      =   2205
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1680
      Top             =   360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   600
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   360
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   90
      Left            =   1440
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   90
      Left            =   0
      Picture         =   "ProgressBar.ctx":001F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'缺省属性值:
Const m_def_FirstColor = &HFF00&
Const m_def_SecondColor = &H80FF&
Const m_def_Max = 100
Const m_def_Min = 0
Const m_def_Value = 100
'属性变量:
Dim m_FirstColor As Long
Dim m_SecondColor As Long
Dim m_Max As Long
Dim m_Min As Long
Dim m_Value As Long
Dim a As Integer, b As Integer, c As Integer, Ready As Integer

Sub Error_2(aa As Object)
aa.Width = 0
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,0
Public Property Get FirstColor() As Long
Attribute FirstColor.VB_ProcData.VB_Invoke_Property = "FirstColor"
    FirstColor = m_FirstColor
    Shape1.BackColor = m_FirstColor
End Property

Public Property Let FirstColor(ByVal New_FirstColor As Long)
    m_FirstColor = New_FirstColor
    PropertyChanged "FirstColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,0
Public Property Get SecondColor() As Long
Attribute SecondColor.VB_ProcData.VB_Invoke_Property = "SecondColor"
    SecondColor = m_SecondColor
    Shape2.BackColor = m_SecondColor
End Property

Public Property Let SecondColor(ByVal New_SecondColor As Long)
    m_SecondColor = New_SecondColor
    PropertyChanged "SecondColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    PropertyChanged "Value"
End Property

Private Sub Command1_Click()

MsgBox "Shape1.Left" & Shape1.Left & "   " & "Shape1.Width" & _
Shape1.Width & "" & "   Value" & m_Value _
& "   " & "  a" & a & "  b" & b & "  c" & c & "  max" & m_Max & "  min" & m_Min _
 & "  m_Max - m_Min" & m_Max - m_Min & "  Image1.Width" & Image1.Width
 a = Image1.Width / (m_Max - m_Min)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
b = a * m_Value
Shape1.Width = b
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Shape2.Left = Shape1.Left + Shape1.Width
If Ready <> b Then
    Shape2.Width = Ready - Shape1.Width
    Timer3.Enabled = True
    Ready = b
End If

End Sub

Private Sub Timer3_Timer()
On Error Resume Next
'dafkjdsafds;ljflkdsjfldsfoiwfjdsajfkdsjflkjdsa;f
Shape2.Visible = True
On Error GoTo Error_1
Shape2.Width = Shape2.Width - 10
If Shape2.Width <= 0 Then
Error_1:
    Call Error_2(Shape2)
    Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
a = Image1.Width / (m_Max - m_Min)
Timer4.Enabled = False
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
a = Image1.Width / (m_Max - m_Min)
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
 On Error Resume Next
    m_FirstColor = m_def_FirstColor
    m_SecondColor = m_def_SecondColor
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_Value = m_def_Value
    UserControl.Width = 3800
    a = Image1.Width / (m_Max - m_Min)
    b = a * m_Value
    Shape1.Width = b
    Ready = b
    
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_FirstColor = PropBag.ReadProperty("FirstColor", m_def_FirstColor)
    m_SecondColor = PropBag.ReadProperty("SecondColor", m_def_SecondColor)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

Private Sub UserControl_Resize()
Image1.Width = UserControl.Width
Image1.Height = UserControl.Height
Shape1.Height = Image1.Height
Shape2.Height = Image1.Height
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("FirstColor", m_FirstColor, m_def_FirstColor)
    Call PropBag.WriteProperty("SecondColor", m_SecondColor, m_def_SecondColor)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

