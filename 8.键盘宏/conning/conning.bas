Attribute VB_Name = "Module1"

Sub Request(objectfile)
Open objectfile For Output As #1
    Print #1, "play"
Close
End Sub

Sub Indication(myfile, objectfile)
Open myfile For Input As #1
    Line Input #1, a
Close
Open objectfile For Output As #1
    If a = "play" Then
        Print #1, "Yse"
    Else
        Print #1, "No"
    End If
Close
End Sub

Function Response(myfile, objectfile) As Boolean
Open myfile For Input As #1
    Line Input #1, a
Close

Open myfile For Output As #1
Close

Response = False
If a = "Yes" Then
    code = MsgBox("是否准备好？", 32 + 4, "conning")
    If code = 6 Then
        Open objectfile For Output As #1
            Print #1, "1Ready"
        Close
        Response = True
    Else
        '{
        '建议用定时器控制
        '}
    End If
End If
End Function

Function Confirmation(myfile, objectfile) As Boolean
Open myfile For Input As #1
    Line Input #1, a
Close
him1 = CBool(Mid(a, 1, 1))
movecode = Right(a, Len(a) - 1)
Confirmation = False
Open objectfile For Output As #1
    If movecode = "Ready" Then
        code = MsgBox("是否准备好？", 32 + 4, "conning")
        If code = 6 Then
            Open objectfile For Output As #1
                Print #1, "1"
            Close
            Confirmation = True
        Else
            '{
            '建议用定时器控制
            '}
        End If
    End If
End Function

Sub will_start(objectfile)
Open myfile For Input As #1
    Line Input #1, a
Close
him2 = CBool(a)
End Sub
