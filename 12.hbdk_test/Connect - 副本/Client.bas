Attribute VB_Name = "Client_Code"

Function ReBoot(Equipment As Winsock, RemoteHostName As String, Port As Integer)
With Equipment
    .Close
    .Connect RemoteHostName, Port
End With
End Function

