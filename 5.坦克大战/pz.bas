Attribute VB_Name = "pz"
'墙体与坦克碰撞检测模块函数
Function fx(qt As Control, tk As Control, ByVal ma As Integer) As Integer
'碰撞检测向右
If (qt.Left - 49 + qt.Width = tk.Left And ((tk.Top <= qt.Top And qt.Top <= tk.Top + tk.Height) Or (tk.Top <= qt.Top + qt.Height And qt.Top + qt.Height <= tk.Top + tk.Height))) And (Not ma = 1) Then
    fx = 6
'碰撞检测向左
ElseIf (qt.Left + 49 = tk.Width + tk.Left And ((tk.Top <= qt.Top And qt.Top <= tk.Top + tk.Height) Or (tk.Top <= qt.Top + qt.Height And qt.Top + qt.Height <= tk.Top + tk.Height))) And (Not ma = 1) Then
    fx = 4
'碰撞检测向上
ElseIf (qt.Top - 49 + qt.Height = tk.Top And ((tk.Left <= qt.Left And qt.Left <= tk.Left + tk.Width) Or (tk.Left <= qt.Left + qt.Width And qt.Left + qt.Width <= tk.Left + tk.Width))) And (Not ma = 1) Then
    fx = 8
'碰撞检测向下
ElseIf (qt.Top + 49 = tk.Height + tk.Top And ((tk.Left <= qt.Left And qt.Left <= tk.Left + tk.Width) Or (tk.Left <= qt.Left + qt.Width And qt.Left + qt.Width <= tk.Left + tk.Width))) And (Not ma = 1) Then
    fx = 2
End If
End Function

'墙体与子弹碰撞检测模块函数
Function zdqt(qt As Control, zd As Control) As String

'碰撞检测向右
If ((qt.Left + qt.Width = zd.Left - 1 Or qt.Left + qt.Width = zd.Left - 51) And ((qt.Top <= zd.Top And zd.Top <= qt.Top + qt.Height) Or (qt.Top <= zd.Top + zd.Height And zd.Top + zd.Height <= qt.Top + qt.Height))) Then
    zdqt = "right"
'碰撞检测向左
ElseIf ((qt.Left = zd.Width + zd.Left + 1 Or qt.Left = zd.Width + zd.Left + 51) And ((qt.Top <= zd.Top And zd.Top <= qt.Top + qt.Height) Or (qt.Top <= zd.Top + zd.Height And zd.Top + zd.Height <= qt.Top + qt.Height))) Then
zdqt = "left"
'碰撞检测向下
ElseIf ((qt.Top + qt.Height = zd.Top - 1 Or qt.Top + qt.Height = zd.Top - 51) And ((qt.Left <= zd.Left And zd.Left <= qt.Left + qt.Width) Or (qt.Left <= zd.Left + zd.Width And zd.Left + zd.Width <= qt.Left + qt.Width))) Then
    zdqt = "down"
'碰撞检测向上
ElseIf ((qt.Top = zd.Height + zd.Top + 1 Or qt.Top = zd.Height + zd.Top + 51) And ((qt.Left <= zd.Left And zd.Left <= qt.Left + qt.Width) Or (qt.Left <= zd.Left + zd.Width And zd.Left + zd.Width <= qt.Left + qt.Width))) Then
    zdqt = "up"
End If
End Function

'子弹与坦克碰撞检测模块函数
Function zdtk(tk As Control, zd As Control) As Boolean
zdtk = False
'碰撞检测向右                          向左
If (((tk.Left + tk.Width = zd.Left) Or (tk.Left = zd.Width + zd.Left)) And ((tk.Top <= zd.Top And zd.Top <= tk.Top + tk.Height) Or (tk.Top <= zd.Top + zd.Height And zd.Top + zd.Height <= tk.Top + tk.Height))) And tk.Visible Then zdtk = True
'碰撞检测向下                         向上
If (((tk.Top + tk.Height = zd.Top) Or (tk.Top = zd.Height + zd.Top)) And ((tk.Left <= zd.Left And zd.Left <= tk.Left + tk.Width) Or (tk.Left <= zd.Left + zd.Width And zd.Left + zd.Width <= tk.Left + tk.Width))) And tk.Visible Then zdtk = True
End Function

'坦克1与坦克2碰撞检测模块函数
Function tkpz(tk1 As Control, tk2 As Control) As String
'碰撞检测向右
If (tk1.Left + tk1.Width = tk2.Left And ((tk1.Top <= tk2.Top And tk2.Top <= tk1.Top + tk1.Height) Or (tk1.Top <= tk2.Top + tk2.Height And tk2.Top + tk2.Height <= tk1.Top + tk1.Height))) Then
    tkpz = "right"
'碰撞检测向左
ElseIf (tk1.Left = tk2.Width + tk2.Left And ((tk1.Top <= tk2.Top And tk2.Top <= tk1.Top + tk1.Height) Or (tk1.Top <= tk2.Top + tk2.Height And tk2.Top + tk2.Height <= tk1.Top + tk1.Height))) Then
    tkpz = "left"
'碰撞检测向下
ElseIf (tk1.Top + tk1.Height = tk2.Top And ((tk1.Left <= tk2.Left And tk2.Left <= tk1.Left + tk1.Width) Or (tk1.Left <= tk2.Left + tk2.Width And tk2.Left + tk2.Width <= tk1.Left + tk1.Width))) Then
    tkpz = "down"
'碰撞检测向上
ElseIf (tk1.Top = tk2.Height + tk2.Top And ((tk1.Left <= tk2.Left And tk2.Left <= tk1.Left + tk1.Width) Or (tk1.Left <= tk2.Left + tk2.Width And tk2.Left + tk2.Width <= tk1.Left + tk1.Width))) Then
    tkpz = "up"
End If
End Function
