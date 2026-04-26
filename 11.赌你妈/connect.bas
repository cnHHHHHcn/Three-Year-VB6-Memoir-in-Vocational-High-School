Attribute VB_Name = "games_connect"

Sub Request(objectfile)
On Error Resume Next
Open objectfile For Output As #1
    Print #1, "play" & Form1.User.Caption
Close
End Sub

Function Indication(myfile, objectfile) As Boolean
On Error Resume Next
Open myfile For Input As #1
    Line Input #1, a
Close
playername = Mid(a, 5)
a = Left(a, 4)
Open objectfile For Output As #1
    If a = "play" Then code = MsgBox("玩家: " & playername & " 邀请你玩", 32 + 4, "connect")
    If code = 6 Then
        Print #1, "Yse": Indication = True
    Else
        Print #1, "No": Indication = False
    End If
Close
End Function

Function Response(myfile, objectfile) As Boolean
On Error Resume Next
Open myfile For Input As #1
    Line Input #1, a
Close

Open myfile For Output As #1
Close

Response = False
    
    If a = "Yes" Then
        code = MsgBox("是否准备好？", 32 + 4, "connect")
        If code = 6 Then
            Open objectfile For Output As #1
                Print #1, "1Ready"
            Close
            Response = True
        Else
            Form1.PlayerWaiting.Enabled = True
        End If
    Else
        Form1.PlayerCheck_Send.Enabled = False: Form1.PlayerCheck.Enabled = True
        MsgBox "联机失败。", 16, "connect"
    End If
    
End Function

Function Confirmation(myfile, objectfile) As Boolean
On Error Resume Next
Open myfile For Input As #1
    Line Input #1, a
Close
'him1 = CBool(Val(Left(a, 1)))
movecode = Right(a, Len(a) - 1)
Confirmation = False
If movecode = "Ready" Then
    code = MsgBox("是否准备好？", 32 + 4, "connect")
    If code = 6 Then
        Open objectfile For Output As #1
            Print #1, "1"
        Close
        Confirmation = True
    Else
        Form1.PlayerWaiting.Enabled = True
    End If
End If
End Function

Sub will_start(myfile, Player2)
On Error Resume Next
Open myfile For Input As #1
    Line Input #1, a
Close
Player2 = CBool(Val(a))
End Sub
