Attribute VB_Name = "C_ColorValue"
Sub Show(ShowObject As Control, ShowCodeObject As Control, Red As Long, Green As Long, Blue As Long)
ShowObject.BackColor = RGB(Red, Green, Blue)
ShowCodeObject.Caption = RGB(Red, Green, Blue)
End Sub

Sub RGBChange(Code As Long, Red As Variant, Green As Variant, Blue As Variant)
ColorData = Hex(Code)
ColorData = String(6 - Len(ColorData), "0") & ColorData

'16进制列表
H_Num = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")
'10进制列表
D_Num = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15")


For i = 0 To 15
    'Red
    If Mid(ColorData, 5, 1) = H_Num(i) Then Decade = Val(D_Num(i))
    If Mid(ColorData, 6, 1) = H_Num(i) Then The_unit = Val(D_Num(i))
Next
Red = NumChange(16, Decade, The_unit)

For i = 0 To 15
    'Green
    If Mid(ColorData, 3, 1) = H_Num(i) Then Decade = Val(D_Num(i))
    If Mid(ColorData, 4, 1) = H_Num(i) Then The_unit = Val(D_Num(i))
Next
Green = NumChange(16, Decade, The_unit)

For i = 0 To 15
    'Blue
    If Mid(ColorData, 1, 1) = H_Num(i) Then Decade = Val(D_Num(i))
    If Mid(ColorData, 2, 1) = H_Num(i) Then The_unit = Val(D_Num(i))
Next
Blue = NumChange(16, Decade, The_unit)
End Sub

Function NumChange(Cardinal_Number As Integer, ByVal Decade As Long, ByVal The_unit As Long)
NumChange = Decade * Cardinal_Number ^ 1 + The_unit * Cardinal_Number ^ 0
End Function
