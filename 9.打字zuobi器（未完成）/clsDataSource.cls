VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' 要做: 声明本地的 ADO Recordset 对象。例如:
'Private WithEvents rs As ADODB.Recordset

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
    ' 要做:  返回基于 DataMember 的相应的 recordset。例如:
    
    'Select Case DataMember
    'Case ""             ' Default
    '    Set Data = Nothing
    'Case Else           ' Default
    '    Set Data = rs
    'End Select
End Sub

