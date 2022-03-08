Attribute VB_Name = "mdl_bull_mon"
Option Explicit

Public Sub bull_mon_hit(pic_man, pic_bull_mon, tmr_bull_mon, shot_bl_mon, mone_clicks, Text1, which_al_i, shot_bl_mon_i, pic_bull_mon_i) ', pic_bull_man, shot_bl_man)
'Dim distance As Integer ' ???
'If pic_bull_mon_i = 0 Then ' the parameter for the different bull moves
'    distance = 100
'End If
'If pic_bull_mon_i = 1 Then ' the parameter for the different bull moves
'    distance = 50
'End If
    shot_bl_mon(shot_bl_mon_i) = False
     pic_bull_mon(pic_bull_mon_i).Visible = True
    pic_bull_mon(pic_bull_mon_i).Left = pic_bull_mon(pic_bull_mon_i).Left - 50 'distance 'moving the bull

        If ((pic_man.Top > pic_bull_mon(pic_bull_mon_i).Top) And _
            ((pic_man.Top) < (pic_bull_mon(pic_bull_mon_i).Top + pic_bull_mon(pic_bull_mon_i).Height))) Then
            If (((pic_bull_mon(pic_bull_mon_i).Left) < (pic_man.Left)) And _
            ((pic_bull_mon(pic_bull_mon_i).Left + pic_bull_mon(pic_bull_mon_i).Width) > (pic_man.Left))) Or _
            (((pic_bull_mon(pic_bull_mon_i).Left) < (pic_man.Left + pic_man.Width - 600)) And ((pic_bull_mon(pic_bull_mon_i).Left) > (pic_man.Left))) Then
                    MsgBox "u clicked " & mone_clicks & " times. lose !", , Text1.Text
                    End ' man banged
       End If
    End If
    'upper corner


   If ((pic_man.Top < pic_bull_mon(pic_bull_mon_i).Top) And _
        ((pic_man.Top + pic_man.Height) > (pic_bull_mon(pic_bull_mon_i).Top + pic_bull_mon(pic_bull_mon_i).Height))) Then
        If (((pic_bull_mon(pic_bull_mon_i).Left) < (pic_man.Left)) And _
            ((pic_bull_mon(pic_bull_mon_i).Left + pic_bull_mon(pic_bull_mon_i).Width) > (pic_man.Left))) Or _
            (((pic_bull_mon(pic_bull_mon_i).Left) < (pic_man.Left + pic_man.Width - 600)) And ((pic_bull_mon(pic_bull_mon_i).Left) > (pic_man.Left))) Then
                MsgBox "u clicked " & mone_clicks & " times ", , Text1.Text
                 End ' man banged
        End If
    End If
   'middle


   If ((pic_man.Top < pic_bull_mon(pic_bull_mon_i).Top) And _
        ((pic_man.Top + pic_man.Height) > (pic_bull_mon(pic_bull_mon_i).Top))) Then
          If (((pic_bull_mon(pic_bull_mon_i).Left) < (pic_man.Left)) And _
             ((pic_bull_mon(pic_bull_mon_i).Left + pic_bull_mon(pic_bull_mon_i).Width) > (pic_man.Left))) Or _
             (((pic_bull_mon(pic_bull_mon_i).Left) < (pic_man.Left + pic_man.Width - 600)) And ((pic_bull_mon(pic_bull_mon_i).Left) > (pic_man.Left))) Then
                    MsgBox "u clicked " & mone_clicks & " times ", , Text1.Text '
                    End ' man banged
          End If
    End If
    'lower corner
    'why to die if u fired in your legs ?
    
    
    '***********         'after
   ' If ((pic_bull_mon(pic_bull_mon_i).Left) < (pic_man.Left) And _
   '    (pic_bull_mon(pic_bull_mon_i).Left + pic_bull_mon(pic_bull_mon_i).Width) > (pic_man.Left)) Or _
   '    (((pic_bull_mon(pic_bull_mon_i).Left) < (pic_man.Left + pic_man.Width)) And ((pic_bull_mon(pic_bull_mon_i).Left) > (pic_man.Left))) Then
   '        End 'front and on
   ' End If
    '***********

    If (pic_bull_mon(pic_bull_mon_i).Left < (0 - pic_bull_mon(pic_bull_mon_i).Width)) Then
        pic_bull_mon(pic_bull_mon_i).Visible = False
        tmr_bull_mon.Enabled = True ' ???=false
        shot_bl_mon(shot_bl_mon_i) = True
        
    End If


  '  If pic_bull_man.Left + pic_bull_man.Width = pic_bull_mon(pic_bull_mon_i).Left And _
  '      pic_bull_man.Top = pic_bull_mon(pic_bull_mon_i).Top Then
  '      shot_bl_mon(shot_bl_mon_i) = True 'collusion beetween bulls
  '      shot_bl_man = True
  '   End If


End Sub
Public Sub which_al_fire(which_al_i, pic_al, shot_bl_mon, dead_al, dead_al_i, tmr_bull_mon, pic_bull_mon, shot_bl_mon_i, pic_bull_mon_i)


Randomize    'which al fire
   which_al_i = (Rnd * 5)
   If which_al_i > 4 Then
   which_al_i = 4
   End If
    dead_al_i = which_al_i
   
If (shot_bl_mon(shot_bl_mon_i)) Then
If dead_al(dead_al_i) = False Then
    pic_bull_mon(pic_bull_mon_i).Top = (pic_al(which_al_i).Top + pic_al(which_al_i).Height / 2)
    pic_bull_mon(pic_bull_mon_i).Left = pic_al(which_al_i).Left 'looks osem !
    tmr_bull_mon.Enabled = True
End If
End If


    If dead_al(0) And dead_al(1) And dead_al(2) And dead_al(3) And dead_al(4) Then
        MsgBox ("you beat all the aliens קונגרשלויישנס")
        End
    End If


End Sub
