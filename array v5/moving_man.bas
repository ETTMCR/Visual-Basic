Attribute VB_Name = "mdl_moving_man"

Option Explicit


Public Sub moving_man(sci, pc_mn)

Select Case sci

Case 49: pc_mn.Left = pc_mn.Left - 50 'man right down
        pc_mn.Top = pc_mn.Top + 50
Case 50: pc_mn.Top = pc_mn.Top + 50 'man down
Case 51: pc_mn.Left = pc_mn.Left + 50 'man left down
        pc_mn.Top = pc_mn.Top + 50
Case 52: pc_mn.Left = pc_mn.Left - 50 'man right
'Case 53:fire
Case 54: pc_mn.Left = pc_mn.Left + 50 'man left
Case 55: pc_mn.Left = pc_mn.Left - 50 'man right up
        pc_mn.Top = pc_mn.Top - 50
Case 56: pc_mn.Top = pc_mn.Top - 50 'man up
Case 57: pc_mn.Left = pc_mn.Left + 50 'man left up
        pc_mn.Top = pc_mn.Top - 50

End Select
End Sub

