VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9615
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Boot 
      Caption         =   "启动"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Clear 
      Appearance      =   0  'Flat
      Caption         =   "清除"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   4815
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim Label() As HTMLLabel
Dim LabelTypeCount()

Private Type HTMLLabel
    LeftPosition As Long
    RightPosition As Long
    LabelLen As Long
    LabelType As String * 10
End Type

Private Sub Boot_Click()
    HTMLCheck
End Sub

Private Sub DelEmptyLabel(LabelStruct() As HTMLLabel, Result() As Long)
    For i = 1 To UBound(LabelStruct) - 1
        If LabelStruct(i).LabelLen < 0 Then
            Index = Index + 1
            ReDim Preserve Result(Index)
            Result(Index) = ListRange * i + j
        End If
    Next
    For i = 1 To UBound(Result)
        Debug.Print Result(i)
    Next
End Sub

Private Sub HTMLCheck()
    On Error Resume Next
    Dim a() As Long
    Position = 1
    Index = 1
Start:
    DoEvents
    ReDim Preserve Label(Index)
    With Label(Index)
        .LeftPosition = InStr(Position, Text.Text, "<")
        .RightPosition = InStr(Position, Text.Text, ">")
        Position = .RightPosition + 1
        .LabelLen = .RightPosition - .LeftPosition + 1
        If Not .LabelLen > 0 Then GoTo ExitLabel
        LabelText = Mid(Text.Text, .LeftPosition, .LabelLen)
        SpacePosition = InStr(1, LabelText, " ")
        .LabelType = Mid(Text.Text, .LeftPosition + 1, IIf(SpacePosition <> 0, SpacePosition, .LabelLen) - 2)
    End With
    Debug.Print Label(Index).LeftPosition; Label(Index).RightPosition; Label(Index).LabelLen, Label(Index).LabelType
ExitLabel:
    If Label(Index).LeftPosition <> 0 Then Index = Index + 1: GoTo Start
    DelEmptyLabel Label, a
End Sub

