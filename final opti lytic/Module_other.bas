Attribute VB_Name = "Module_other"

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public max_cells As Integer
Public max_vi As Long
Public ary_vi_x_y() As Integer
Public ary_cells() As Integer
Public img_cell_hgt As Byte
Public img_cell_wdt As Byte

Public Sub does_pix_clr_0(last_x, last_y)
'אם הפיקסל שאלי הולך להגיע , צבעו אינו לבן ז"א שיש שם תא
If Not (Form1.Point(last_x, last_y) = vbWhite) Then

' ריצה על כל התמונות תא שיש
          For h = 1 To max_cells 'Form1.img_cell(0) '(UBound(Form1.img_cell())) 'till the max of cells
          
                    If Form1.img_cell(h).Tag = "healthy" Then  'not dead cell and not sick
                              If (Form1.img_cell(h).Left = last_x Or Form1.img_cell(h).Left < last_x) Then
                                        If (Form1.img_cell(h).Left + img_cell_wdt = last_x Or Form1.img_cell(h).Left + img_cell_wdt > last_x) Then
                                                  If (Form1.img_cell(h).Top = last_y Or Form1.img_cell(h).Top < last_y) Then
                                                            If (Form1.img_cell(h).Top + img_cell_hgt = last_y Or Form1.img_cell(h).Top + img_cell_hgt > last_y) Then
                                                                      'Form1.img_cell(h).Picture = LoadPicture(App.Path & "\myvi_sick.ico")
                                                                      Form1.img_cell(h).Picture = Form1.img_cell_sick.Picture
                                                                      Form1.img_cell(h).Tag = "sick"
                                                                     ' GoTo here 'just for speeding - if it does it
                                                            End If
                                                  End If
                                        End If
                              End If
                    End If
          
          Next 'h
here:
End If 'Not (Form1.Point(last_x, last_y) = 0)

End Sub

Public Sub mitozis(cell_ID1, pos_x, pos_y) 'cell_ID= the older one
  'first randoming pos ' then check if no one there ' and if so , only then create new one,
          
           max_cells = (UBound(ary_cells)) + 1
          ReDim Preserve ary_cells(max_cells) As Integer 'must  redim Preserve , else zeroes all the ary cells
          Load Form1.img_cell(max_cells)
          up_down_or_rgt_lft = Int(Rnd * 2)
          
'          Select Case up_down_or_rgt_lft
'          Case 0 ' lft rgt
'                    new_pos_x = pos_x - img_cell_wdt '-50
'                    Form1.img_cell(max_cells).Left = pos_x + img_cell_wdt  '+ 50
'                    Form1.img_cell(max_cells).Top = pos_y
'          Case 1 ' down up
'                    new_pos_y = pos_y - img_cell_hgt  '-50
'                    Form1.img_cell(max_cells).Top = pos_y + img_cell_hgt  '+ 50
'                    Form1.img_cell(max_cells).Left = pos_x
'          End Select
          
          
                    If up_down_or_rgt_lft = 0 Then ' handeling cell id max_cell+1 'the new one
                              If Not (Form1.Point(pos_x + img_cell_wdt / 2 + img_cell_wdt, pos_y + img_cell_hgt / 2) = &HFFFFFF) Then ' blank aree
                                        Unload Form1.img_cell(max_cells)
                                        max_cells = max_cells - 1 ' why making behaiot?
                                        ReDim Preserve ary_cells(max_cells) As Integer 'must  redim Preserve , else zeroes all the ary cells
                                        'Beep
                              Else 'Case 0 ' ' will be at right of his of his cell_ID old pos_x
                                        Form1.img_cell(max_cells).Left = pos_x + img_cell_wdt  '+ 50
                                        Form1.img_cell(max_cells).Top = pos_y
                                        Form1.img_cell(max_cells).Visible = True
                                        Form1.img_cell(max_cells).Tag = "healthy"
                                        ary_cells(UBound(ary_cells)) = Int(Rnd * 10) + 1
                               End If
                    Else ' Case 1 ' down up
                              If Not (Form1.Point(pos_x + img_cell_wdt / 2, pos_y + (img_cell_hgt + img_cell_hgt / 2)) = &HFFFFFF) Then ' blank area
                                        Unload Form1.img_cell(max_cells)
                                        max_cells = max_cells - 1 ' why making behaiot?
                                        ReDim Preserve ary_cells(max_cells) As Integer 'must  redim Preserve , else zeroes all the ary cells
                                        'Beep
                              Else 'Case 1 ' down up ' will be at bottom of his cell_ID old pos_y
                                        Form1.img_cell(max_cells).Top = pos_y + img_cell_hgt  '+ 50
                                        Form1.img_cell(max_cells).Left = pos_x
                                        Form1.img_cell(max_cells).Visible = True
                                        Form1.img_cell(max_cells).Tag = "healthy"
                                        ary_cells(UBound(ary_cells)) = Int(Rnd * 10) + 1
                               End If
                              'Form1.img_cell(cell_ID1).Top = pos_y - img_cell_hgt  '-50
                    End If
                    
          
'         'בדיקה האם האזור בו הולך להיות התא החדש יש בו כבר מישהו ,ואז הורגים אותו
'          If Not (Form1.Point(Form1.img_cell(max_cells).Left + img_cell_wdt / 2, Form1.img_cell(max_cells).Top + img_cell_hgt / 2) = &HFFFFFF) Then ' blank area
'          ' תנאי זה כולל בתוכו בדיקה שהתא לא ימצא מחוץ למסך. התוצאה של פוינט נותנת  מינוס1
'                    '( !
'                    'Form1.img_cell(max_cells).Tag = "dead"
'                    'Form1.img_cell(max_cells).Visible = False
'                    Unload Form1.img_cell(max_cells)
'                    max_cells = max_cells - 1 ' why making behaiot?
'                    ReDim Preserve ary_cells(max_cells) As Integer 'must  redim Preserve , else zeroes all the ary cells
'          Else
'                    Form1.img_cell(max_cells).Visible = True
'                    Form1.img_cell(max_cells).Tag = "healthy"
'          End If
          
          'בדיקה האם האזור בו הולך להיות התא ישן יש בו כבר מישהו ,ואז הורגים אותו
          'If Not Form1.Point(Form1.img_cell(cell_ID).Left + img_cell_wdt / 2, Form1.img_cell(cell_ID).Top + img_cell_hgt / 2) = &HFFFFFF Then ' blank area
'          If Not (Form1.Point(pos_x + img_cell_wdt / 2, pos_y + img_cell_hgt / 2) = &HFFFFFF) Then ' blank area
''                                                 new_
'                    '( !
'                    Call killing(cell_ID1)
          'Else 'of If Not
                    If up_down_or_rgt_lft = 0 Then ' handeling cell_ID1
                              If Not (Form1.Point(pos_x - img_cell_wdt / 2, pos_y + img_cell_hgt / 2) = &HFFFFFF) Then ' blank area
                                        Call killing(cell_ID1)
                              Else 'Case 0 ' ' will be at left of his old pos_x
                                        Form1.img_cell(cell_ID1).Left = pos_x - img_cell_wdt  '-50
                                        ary_cells(cell_ID1) = Int(Rnd * 10) + 1 'timing sec for mitozis
                               End If
                    Else ' Case 1 ' down up
                              If Not (Form1.Point(pos_x - img_cell_wdt / 2, pos_y - (img_cell_hgt - img_cell_hgt / 2)) = &HFFFFFF) Then ' blank area
                                        Call killing(cell_ID1)
                              Else 'Case 0 ' lft rgt ' ' will be at top of his old pos_y
                                        ary_cells(cell_ID1) = Int(Rnd * 10) + 1 'timing sec for mitozis
                                        Form1.img_cell(cell_ID1).Top = pos_y - img_cell_hgt  '-50
                               End If
                              'Form1.img_cell(cell_ID1).Top = pos_y - img_cell_hgt  '-50
                    End If
                              ' Form1.img_cell(cell_ID1).Tag = "healthy" ' already happen at form load
          'End If ''of If Not

'Form1.lbl_cells_mone.Caption = UBound(ary_cells)
                 
End Sub

 Public Sub killing(cell_ID2)
 If Not (cell_ID2 = 1 And UBound(ary_cells) = 1) Then ' the last cell on screen
         ' If Not cell_ID2 = max_cells Then 'התא באמצע המערך ,נעביר את האחרון אליו, ואת האחרון נעשה אנלוד
          If Not cell_ID2 = UBound(ary_cells) Then 'התא באמצע המערך ,נעביר את האחרון אליו, ואת האחרון נעשה אנלוד
          
                    Form1.img_cell(cell_ID2).Left = Form1.img_cell(UBound(ary_cells)).Left
                    Form1.img_cell(cell_ID2).Top = Form1.img_cell(UBound(ary_cells)).Top
                    Form1.img_cell(cell_ID2).Tag = Form1.img_cell(UBound(ary_cells)).Tag
                    Form1.img_cell(cell_ID2).Picture = Form1.img_cell(UBound(ary_cells)).Picture
                    ary_cells(cell_ID2) = ary_cells(UBound(ary_cells))
'                    Form1.img_cell(cell_ID2).Tag = "dead" ' NO use for thet
'                    Form1.img_cell(cell_ID2).Visible = False
                     Unload Form1.img_cell(UBound(ary_cells))
                     max_cells = UBound(ary_cells) - 1
                    ReDim Preserve ary_cells(max_cells) As Integer
          Else 'תא אינדקס  אחרון מת
                    Unload Form1.img_cell(cell_ID2) 'why he jumping on the rest of the code ? ?????  when
                    max_cells = UBound(ary_cells) - 1
                    ReDim Preserve ary_cells(max_cells) As Integer
          End If
                    'beep
Else 'תא ינדקס 1 מת, ולמעשה האחרון במשחק
                    Unload Form1.img_cell(cell_ID2) 'why he jumping on the rest of the code ? ?????  when
                    max_cells = UBound(ary_cells) - 1
                    ReDim Preserve ary_cells(max_cells) As Integer
                    MsgBox "GAME OVER!", , "VIRUSES !"
                    Beep
End If
 End Sub
