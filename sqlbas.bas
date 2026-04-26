Attribute VB_Name = "sql"
'每个值之间用"."隔开
Function value(values, valcount) As String
ends = 1: start = 1: value = "values( "
For i = 1 To valcount
    a = InStr(start, values, ".")
    If a = 0 Then a = Len(values) + 1
    ends = a - 1
    endlen = ends - (start - 1)
    e = Mid(values, start, endlen)
    start = a + 1
    If i = valcount Then value = value & e & " " & ")" Else value = value & e & " " & ","
Next
End Function

Function fields(field, fieldcount) As String
ends = 1: start = 1: fields = "( "
For i = 1 To fieldcount
    a = InStr(start, field, ".")
    If a = 0 Then a = Len(field) + 1
    ends = a - 1
    endlen = ends - (start - 1)
    e = Mid(field, start, endlen)
    start = a + 1
    If i = fieldcount Then fields = fields & e & " " & ")" Else fields = fields & e & " " & ","
Next
End Function

Function where(wheres, wherecount) As String
' # 为 and运算符
' $ 为 not运算符
' % 为 or运算符
ends = 1: start = 1: where = "where "
For i = 1 To wherecount
    a = InStr(start, wheres, ".")
    If a = 0 Then a = Len(wheres) + 1
    ends = a - 1
    endlen = ends - (start - 1)
    e = Mid(wheres, start, endlen)
    e1 = Left(e, 1): e2 = Mid(e, 2)
    Select Case e1
    Case "#"
        e = "and " & e2
    Case "$"
        e = "not " & e2
    Case "%"
        e = "or " & e2
    End Select
    
    start = a + 1
    If i = wherecount Then where = where & e Else where = where & e & " "
Next
End Function

Function sets(setvalue, setvaluecount)
ends = 1: start = 1: sets = "set "
For i = 1 To setvaluecount
    a = InStr(start, setvalue, ".")
    If a = 0 Then a = Len(setvalue) + 1
    ends = a - 1
    endlen = ends - (start - 1)
    e = Mid(setvalue, start, endlen)
    start = a + 1
    If i = setvaluecount Then sets = sets & e & " " Else sets = sets & e & " " & ","
Next
End Function

Function delete_from(from, where) As String
delete_from = "delete from " & from & where
End Function

Function drop_table(from, where) As String
drop_table = "drop table " & form & where
End Function

'当free为true时，表为自由表，否则为表为数据表
Function create_table(table, fields, free As Boolean) As String
If free = True Then table = table & " free" & " "
create_table = "create table " & table & fields
End Function

Function insert_into(from, fields, value) As String
insert_into = "insert into " & from & fields & value
End Function

Function update(from, sets, where) As String
update = "update " & from & sets & where
End Function

Function selects(fields, from, where, order, group) As String
selects = "select " & fields & from & where & order & group
End Function

Function order() As String

End Function

Function group(fields, having) As String
group = fields
End Function

Function alter_table(form, sets_fields, alter_fields, drop_fields) As String
alter_table = "alter table " & form & sets_fields & alter_fields & drop_fields
End Function

Function sets_fields() As String

End Function

Function alter_fields() As String

End Function

Function drop_fields() As String

End Function

