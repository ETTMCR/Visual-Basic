Attribute VB_Name = "mdl_ti_pa_area_tt"
Public new_ti_pa As Boolean
Public go_paitiont_name  As String '����� ����� �� �� ������ ������ ����� �� �����
Public taking_name_drug_2_lst_ti_pa As Boolean '����� ������ ����� ������ ����� ������ ������ ����

'@@@@
Public db_prj As Database
Public sDBloc As String
Public rs_tbl_n_d As Recordset
Public rs_dyn_n_d As Recordset
Public rs_tbl_ti_pa As Recordset
Public rs_dyn_ti_pa As Recordset
Public rs_tbl_comps As Recordset
Public rs_dyn_comps As Recordset
Public rs_tbl_drugs_of_pa As Recordset
Public rs_dyn_drugs_of_pa As Recordset
Public rs_dyn_uses_of_drug As Recordset


Public Sub area_to_city(area As String, ByVal city As ComboBox)
'������ ���� ���� ����� ����� ����
Select Case area

    Case "�������": city.Clear
                     city.AddItem "��� ���"
                    city.AddItem "�������"
                   city.AddItem " �������"
                  '  city.AddItem
                  '   city.AddItem
                  '   city.AddItem
                          
    Case "�� ����": city.Clear
                 city.AddItem "��� ��� "
                 city.AddItem "�� �� "
                 city.AddItem "�������"
                 city.AddItem "��� ���� "
                 city.AddItem "����""�"
                 city.AddItem "��� ��"
                 city.AddItem "�� ����"
                 
    Case "����": city.Clear
            city.AddItem "����"
            city.AddItem "�����"
            city.AddItem "�����"
             city.AddItem "���"
            city.AddItem "�����"
            city.AddItem "���"
           city.AddItem "���� �����"
                                     
    Case "��� ���": city.Clear
            city.AddItem "����"
            city.AddItem "�����"
             city.AddItem "������"
             city.AddItem "��� ���"
            city.AddItem "������"
                   
    Case "�����": city.Clear
       city.AddItem "������"
        city.AddItem "����"
        city.AddItem "��� ���"
       city.AddItem "�����"
       city.AddItem "������"
       city.AddItem "�����"
      
    End Select
city.AddItem "���"

End Sub
Public Sub add_item_put_area(cbo As ComboBox)
' ������ �� ������ ���� �� ����� �����
        cbo.Clear
        cbo.AddItem "��� ���"
        cbo.AddItem "�����"
         cbo.AddItem "����"
          cbo.AddItem "�������"
         cbo.AddItem "�� ����"
         
End Sub

Public Sub add_item_put_phone_tt(cbo As ComboBox)
' ������ �� ������ ���� �� ����� �����

    cbo.Clear
    cbo.AddItem "02"
    cbo.AddItem "03"
    cbo.AddItem "04"
    cbo.AddItem "08"
    cbo.AddItem "09"
    
End Sub

Public Function a_00_to_area(tt As String) As String
'���� ����  ����� ����� ���� �����/����
 Select Case tt
        Case "03": a_00_to_area = "�� ����"
        Case "02": a_00_to_area = "�������"
        Case "04": a_00_to_area = "����"
        Case "08": a_00_to_area = "��� ���"
        Case "09": a_00_to_area = "�����"
            
    End Select

End Function

Public Function a_area_to_00(area As String) As String
'���� ���� ���� ����� ����� �����

 Select Case area
        Case "�� ����": a_area_to_00 = "03"
        Case "�������": a_area_to_00 = "02"
        Case "����": a_area_to_00 = "04"
        Case "��� ���": a_area_to_00 = "08"
        Case "�����": a_area_to_00 = "09"
 End Select

End Function
Public Sub ok_check_new_paitiont(cell, name_first, name_last, phone_2, phone_tt, st, num_home, zip, city, area, ByRef ok_ti_pa, Chk_cl, web, age)
''���� ����� ���� ��� �� ����� ����� ������� ������ �����
If Chk_cl = 1 Then
    If (Not (cell.Text <> "") Or (Len(cell) < 7)) Then
        MsgBox "���� ��' ������ ", , " ! ��� ��"
        cell.SetFocus
         cell.BackColor = &HFFC0FF
        cell = False
    Else
          cell = True
    End If
    
    If Not (web.Text <> "") Then
        MsgBox "�� ����� ��' ��� ������", , " ! ��� ��"
        web.SetFocus
        web.BackColor = &HFFC0FF
        web = False
    Else
        web = True
    End If
    
Else
    cell = True
    web = True
End If

If Not (name_first.Text <> "") Then
MsgBox "�� ���� �� �����", , " ! ��� ��"
  name_first.SetFocus
  name_first.BackColor = &HFFC0FF
name_first = False
Else
name_first = True
End If

If Not (name_last.Text <> "") Then
MsgBox "�� ���� �� ����", , " ! ��� ��"
  name_last.SetFocus
  name_last.BackColor = &HFFC0FF
name_last = False
Else
name_last = True
End If

If Not (phone_tt.Text <> "") Then
MsgBox "�� ���� ���� ����", , " ! ��� ��"
phone_tt.SetFocus
phone_tt.BackColor = &HFFC0FF
phone_tt = False
Else
phone_tt = True
End If

'If cbo_put_phone_tt.Text <> "02" Or "03" Or "04" Or "08" Or "09" Then
'MsgBox "���� ����� ����", , "! ��� ��"
'End If


If Not (st.Text <> "") Then
MsgBox "�� ���� ����", , " ! ��� ��"
  st.SetFocus
  st.BackColor = &HFFC0FF
st = False
Else
st = True
End If

If Not (num_home.Text <> "") Then
MsgBox "�� ���� ��' ���", , " ! ��� ��"
  num_home.SetFocus
  num_home.BackColor = &HFFC0FF
num_home = False
Else
num_home = True
End If

If (Not zip.Text <> "") Or (Len(zip.Text) < 5) Then
MsgBox "���� ����� ", , " ! ��� ��"
  zip.SetFocus
  zip.BackColor = &HFFC0FF
zip = False
Else
zip = True
End If


If (Not (phone_2.Text <> "")) Or (Len(phone_2.Text) < 7) Then
MsgBox "���� ��' ����� ", , " ! ��� ��"
  phone_2.SetFocus
  phone_2.BackColor = &HFFC0FF
phone_2 = False
Else
phone_2 = True
End If

If Not (area.Text <> "") Then '
MsgBox "�� ����� ����", , " ! ��� ��"
area.SetFocus
area.BackColor = &HFFC0FF
area = False
Else
area = True
End If

If Not (city.Text <> "") Then
MsgBox "�� ����� ���", , " ! ��� ��"
city.SetFocus
city.BackColor = &HFFC0FF
city = False
Else
city = True
End If

If Not (age.Text <> "") Then
MsgBox "�� ����� ���", , " ! ��� ��"
age.SetFocus
age.BackColor = &HFFC0FF
age = False
Else
age = True
End If

If web And cell And name_first And name_last _
And phone_2 And phone_tt And st And num_home _
And zip And city And area And age Then
'����� �� ������ �� �� "���" ��� ����� �� ������ �� ����
ok_ti_pa = True
End If
End Sub

Public Sub chr_chk_ltr(KeyAscii, cmd)
'���� �� ��� ��� ������ ����� �� ��� ���
Char = Chr(KeyAscii)
 If ((Asc(Char) > 250) Or (Asc(Char) < 224) And (Asc(Char) <> 8)) Then
 Beep
 KeyAscii = 0
' MsgBox "! ���� ��� ", , " ! ��� �� "
 End If
cmd.BackColor = &H80000005
End Sub
Public Sub chr_chk_num(KeyAscii, ByVal cmd)
'���� �� ��� ��� ������ ����� �� ��� ����
Char = Chr(KeyAscii)
If (((Asc(Char)) > 57) Or ((Asc(Char)) < 48) And (Asc(Char) <> 8)) Then
Beep
KeyAscii = 0
' MsgBox "! ���� ��� ", , " ! ��� �� "
 End If
 cmd.BackColor = &H80000005
End Sub
 Public Sub add_item_put_num_web(cbo As ComboBox)
 '����� �� ���� ����� ������ ����������
    cbo.Clear
    cbo.AddItem "050"
    cbo.AddItem "054"
    cbo.AddItem "052"
    cbo.AddItem "057"
 End Sub
