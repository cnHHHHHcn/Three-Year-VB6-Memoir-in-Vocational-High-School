Attribute VB_Name = "du_gameing"

'¼¼ÄÜ×é
'***********************************************************************
Function Hurt_Add(player) As Boolean     'code 1
Hurt_Add = True
End Function

Sub Blood_Add(player)   'code 2
player.Caption = player.Caption + 1
End Sub

Sub Look_Ammo(a As Boolean, AmmoColor As Control)   'code 3
If a Then AmmoColor.FillColor = &H8080FF Else AmmoColor.FillColor = &HFF8080
End Sub

Sub Exit_Ammo(a As Boolean, AmmoColor As Control)   'code 4
If a Then AmmoColor.FillColor = &H8080FF Else AmmoColor.FillColor = &HFF8080
End Sub

Function HandCuff(player) As Boolean   'code 5
HandCuff = True
End Function
'***********************************************************************
