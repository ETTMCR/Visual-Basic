Attribute VB_Name = "mdl_new_drug"
Public go_drug_name  As String '����� ����� �� �� ������ ������ ����� �� �����
Public go_comp_name  As String '����� ����� �� �� ����� ������ ����� �� ����� ��� ����� ����� ���� ����� ����
Public taking_comp As Boolean '�� ���� ���� ����� ����� ����� ���� ����� ����� �� �����
Public taking_name_comp_2_drug As Boolean '����� ������ ����� �� ���� ����� ����� ������
Public taking_name_comp_2_find_drug As Boolean '����� ������ ����� �� ���� ����� ����� ������ �����

Public new_drug As Boolean
Public new_comp As Boolean
Public Sub which_to_use(which_use As String, ByVal use As ComboBox)
'����� �� ���� ������� ����� ���� ������ �����"which"
use.AddItem Clear
Select Case which_use

Case "��� �����":
use.Clear
use.AddItem "���"
use.AddItem "�����"
use.AddItem "���� ����"
use.AddItem "���� ������"
use.AddItem "���� �����"
use.AddItem "�������"

Case "����� ����":
use.Clear
use.AddItem "���"
use.AddItem "����� ��������"
use.AddItem "������ A"
use.AddItem "������ B"
use.AddItem "������ C"
use.AddItem "����"
use.AddItem "����"

Case "��� �����":
use.Clear
use.AddItem "���"
use.AddItem "�����"
use.AddItem "��� ���"
use.AddItem "��� ���� "
use.AddItem "��� �����"
use.AddItem "��� �����"
use.AddItem "��� ��� "
use.AddItem "��� ������"

Case "�����":
use.Clear
use.AddItem "���"
use.AddItem "�����"
use.AddItem "�����"
use.AddItem "����� ��� ��"
use.AddItem "����� ��"
use.AddItem "����� ������ ���"
use.AddItem "����� ��� ��� ����"
use.AddItem "��� ����� �����"
use.AddItem "��������"
use.AddItem "������"
use.AddItem "�������"

'Case "":
'Case "":
'use.AddItem ""

End Select

End Sub

Public Sub add_item_cbo_put_which_use(cbo As ComboBox)
'���� �� ������ �� ������� ���� ������ ��� ����� �����
cbo.Clear
 cbo.AddItem "��� �����":
 cbo.AddItem "����� ����":
cbo.AddItem "��� �����"
cbo.AddItem "�����":

End Sub

Public Sub ok_check_new_drug(name, comp, cost, which, lst_use, ByRef ok_n_d As Boolean, opt_type_else, Option1, Option2, Option3, Option4, Option5, fra_type)
'???? ��� ���� ����� �� �� ������� �� ������ ��� ��� ����� �� ���
'���� ����� ���� ��� �� ����� ����� ������� ������ �����
'����� ��� ������ ������
Dim opt As Boolean

If Not (name <> "") Then
MsgBox "�� ���� �� �����", , " ! ��� ��"
name.SetFocus
name.BackColor = &HFFC0FF
name = False
Else
name = True
End If

If Not (comp <> "") Then
MsgBox "�� ���� �� ����", , " ! ��� ��"
comp.SetFocus
comp.BackColor = &HFFC0FF
comp = False
Else
comp = True
End If

If Not (cost <> "") Then
MsgBox "�� ���� ����", , " ! ��� ��"
cost.SetFocus
cost.BackColor = &HFFC0FF
cost = False
Else
cost = True
End If

If Not (which <> "") Then
MsgBox "�� ����� ��� ����� ����", , " ! ��� ��"
which.SetFocus
which.BackColor = &HFFC0FF
which = False
Else
which = True
End If


If (lst_use.ListCount = 0) Then
MsgBox "�� ����� ��� ����� ", , " ! ��� ��"
lst_use.SetFocus
lst_use.BackColor = &HFFC0FF
lst_use = False
Else
lst_use = True
End If

If (opt_type_else.Value = 0) And (Option1.Value = 0) And (Option2.Value = 0) _
    And (Option3.Value = 0) And (Option4.Value = 0) And (Option5.Value = 0) Then
        MsgBox "�� ����� ��� ����� ", , " ! ��� ��"
        fra_type.BackColor = &HFFC0FF
        opt_ok = False
    Else
        opt_ok = True
End If

If name And comp And cost And lst_use And which And opt_ok Then
ok_n_d = True
Else
ok_n_d = False
End If
End Sub
