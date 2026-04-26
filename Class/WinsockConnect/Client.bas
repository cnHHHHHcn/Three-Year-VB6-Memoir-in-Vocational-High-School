Attribute VB_Name = "Client_Code"

Sub ReBoot(Equipment As Winsock, Port As Integer,optional RemoteHostName As String = "None")
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

