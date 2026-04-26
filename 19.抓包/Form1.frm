VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   8715
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Del_Item 
      Caption         =   "É¾³ý"
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Add_Item 
      Caption         =   "Ìí¼Ó"
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox RemoteIP 
      Height          =   270
      Left            =   7080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox RemoteList 
      Appearance      =   0  'Flat
      Height          =   2730
      Left            =   4680
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Data 
      Height          =   2775
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock GetPackage 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox RemotePort 
      Height          =   270
      Left            =   600
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox LocalPort 
      Height          =   270
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Private Sub Add_Item_Click()
    RemoteList.AddItem RemoteIP.Text
End Sub

Private Sub Command1_Click()
    GetPackage.Close
    GetPackage.Bind Val(LocalPort.Text)
End Sub

Private Sub Command2_Click()
    GetPackage.Close
End Sub

Private Sub Del_Item_Click()
For i = RemoteList.ListCount - 1 To 0 Step -1
    If RemoteList.Selected(i) Then RemoteList.RemoveItem i
Next
End Sub

Private Sub GetPackage_DataArrival(ByVal bytesTotal As Long)
    Dim Temp() As Byte, Flag As Boolean
    Static Index As Long
    ReDim Temp(bytesTotal)
    GetPackage.GetData Temp
    If RemoteList.ListCount = 0 Then Flag = True: GoTo NextLine
    For i = 0 To RemoteList.ListCount - 1
        If GetPackage.RemoteHostIP = RemoteList.List(i) Then Flag = True: Exit For
    Next
NextLine:
    If Flag Then
        Data.Text = "RemoteIP:" & GetPackage.RemoteHostIP & "   " & "BytesTotal:" & bytesTotal & vbCrLf & Translate(Temp)
        Open "C:\" & Index & ".txt" For Output As #1
            Print #1, Data.Text
        Close
        DoEvents
        Index = Index + 1
        'Data.Text = Translate(Temp)
        'Data.SelStart = Len(Data.Text)
        DoEvents
    End If
End Sub

Function Translate(DataByte() As Byte) As String
    On Error Resume Next
    Dim StrTemp As String, StrHexTemp As String
    For i = LBound(DataByte) To UBound(DataByte)
        Temp = Hex(DataByte(i))
        StrHexTemp = StrHexTemp & " " & IIf(Len(Temp) = 1, "0" & Temp, Temp)
'        If DataByte(i) >= 128 And DataByte(i + 1) >= 128 Then
'            LngTemp = DataByte(i) * 2 ^ 8 + DataByte(i + 1)
'            i = i + 1
'        Else
'            LngTemp = DataByte(i)
'        End If
'        StrTemp = StrTemp & StrConv(ChrW(LngTemp), vbFromUnicode)
    Next
    Translate = "HexData:" & vbCrLf & StrHexTemp
    '"HexData:" & vbCrLf & StrHexTemp & vbCrLf &
    '"Data:" & vbCrLf & StrTemp
End Function
