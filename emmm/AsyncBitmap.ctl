VERSION 5.00
Begin VB.UserControl UserControl1
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ScaleHeight     =   5880
   ScaleWidth      =   6930
   Begin VB.PictureBox picBitmap 
      AutoSize        =   -1  'True
      Height          =   6015
      Left            =   -120
      ScaleHeight     =   5955
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   -120
      Width           =   6975
   End
End
Attribute VB_Name = "AsyncBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'ActiveX.picture 属性
'------------------------------------------------------------------
Public Property Get Picture() As Picture
   Set Picture = picBitmap.Picture
End Property

Public Property Let Picture(ByVal NewPicture As Picture)
   Set picBitmap.Picture = NewPicture
   PropertyChanged "Picture"
End Property

Public Property Set Picture(ByVal NewPicture As Picture)
   Set picBitmap.Picture = NewPicture
   PropertyChanged "Picture"
End Property
'------------------------------------------------------------------


'ActiveX.PictureFromURL 属性
'------------------------------------------------------------------
Option Explicit
Private mstrPictureFromURL As String

Public Property Get PictureFromURL() As String
   PictureFromURL = mstrPictureFromURL
End Property
Public Property Let PictureFromURL(ByVal NewString As String)
   '（验证路径或者 URL 合法性的代码，此处省略之。）
    mstrPictureFromURL = NewString
   If (Ambient.UserMode = True) And (NewString <> "") Then
      '如果程序处于运行模式，并且 URL 字符串非空，      '则开始下载。
      AsyncRead NewString, vbAsyncTypePicture, "PictureFromURL"
   End If
   PropertyChanged "PictureFromURL"
End Property
'------------------------------------------------------------------




Private Sub UserControl_AsyncReadComplete(AsyncProp As VB.AsyncProperty)
    On Error Resume Next
    Select Case AsyncProp.PropertyName
      Case "PictureFromURL"
          Set Picture = AsyncProp.Value
          Debug.Print "Download complete"
    End Select
End Sub

Private Sub picBitmap_Resize()   '如果赋予了一个 Picture，调整大小。
    If picBitmap.Picture <> 0 Then
        UserControl.Size picBitmap.Width, picBitmap.Height
    End If
End Sub

Private Sub UserControl_Resize()
   If picBitmap.Picture = 0 Then
      picBitmap.Move 0, 0, ScaleWidth, ScaleHeight
   Else
      If (Width <> picBitmap.Width) _
            Or (Height <> picBitmap.Height) Then
         Size picBitmap.Width, picBitmap.Height
      End If
   End If
End Sub


Const DOWNLOADFILE = "file:\windows\forest.bmp"

Private Sub UserControl_InitProperties()
   ' 在初始化时使用 Nothing 作为缺省值，
   '   读、写图片属性，
   '   如果没有图片
   '   就不需要 .frx 文件。
   Set Picture = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "Picture", Picture, Nothing
End Sub

Private Sub UserControl_Click()
   Hyperlink.NavigateTo Target:=URLText
End Sub

