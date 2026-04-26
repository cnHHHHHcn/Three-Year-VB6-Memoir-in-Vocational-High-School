VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetWorkPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Const SendHead As String = "<Src>", AcceptHead As String = "<Des>", InformationHead As String = "<Text>"

Public Sub a(b As PackageTypeMOD.Package)
    ReDim b.AcceptUser(1)
    b.AcceptUser(0) = "adsf"
End Sub

Public Function CreatePackageText(SendUser As String, Information As String, ParamArray AcceptUser()) As String

    Dim Send As String, Accept As String, InformationData As String
    
    Send = SendHead & "{" & SendUser & "}"
    
    For i = LBound(AcceptUser) To UBound(AcceptUser)
        AcceptUsers = AcceptUsers & AcceptUser & ","
    Next
    AcceptUsers = Left(AcceptUsers, Len(AcceptUsers) - 1)
    Accept = AcceptHead & "[" & AcceptUsers & "]"
    
    InformationData = InformationHead & "(" & Information & ")"
    
    CreatePackageText = Send & Accept & InformationData
End Function

Public Sub ReadPackageText(PackageData As String, User As String, Data As String)
    Dim AcceptUser() As String, AcceptFlag As Boolean
    AcceptFlag = False
    AcceptUser = Split(Me.GetConnectLabel(PackageData, "<Des>[", "]"), ",")
    For i = LBound(AcceptUser) To UBound(AcceptUser)
        If AcceptUser(i) = Me.User Then AcceptFlag = True
    Next
    If AcceptFlag Then
        User = Me.GetConnectLabel(PackageData, "<Src>{", "}")
        Data = Me.GetConnectLabel(PackageData, "<Text>(", ")")
    End If
End Sub
