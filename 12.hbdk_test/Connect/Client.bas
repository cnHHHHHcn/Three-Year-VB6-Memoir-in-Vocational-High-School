Attribute VB_Name = "Client_Code"

Sub ReBoot(Equipment As Winsock, RemoteHostName As String, Port As Integer)
With Equipment
    If RemoteHostName = "None" Then
        .Close
        .Connect Port
    Else
        .Close
        .Connect RemoteHostName, Port
    End If
End With
End Sub

