Attribute VB_Name = "GetDOS"

Function GetDrive(ByVal Path_String As String) As String
GetDrive = Left(Path_String, InStr(1, Path_String, ":"))
End Function

Function GetPath(ByVal Path_String As String) As String
a = GetFileName(Path_String)
GetPath = Mid(Path_String, 1, Len(Path_String) - IIf(a <> "None", Len(a), 0))
End Function

Function GetFileName(ByVal Path_String As String) As String
Do While Path_String <> filestr
    filestr = Path_String
    Path_String = Right(Path_String, Len(Path_String) - InStr(Path_String, "\"))
Loop
If InStr(1, Path_String, ".") <> 0 Then GetFileName = Path_String Else GetFileName = "None"
End Function
