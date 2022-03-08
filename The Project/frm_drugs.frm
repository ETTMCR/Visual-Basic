VERSION 5.00
Begin VB.Form frm_drugs 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "מאגר התרופות"
   ClientHeight    =   4335
   ClientLeft      =   5760
   ClientTop       =   2535
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   4425
   Begin VB.ComboBox cbo_find_which_use 
      Height          =   315
      Left            =   2520
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmd_add_new_drug 
      Caption         =   "הו&ספת תרופה חדשה"
      Height          =   495
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmd_del_drug 
      Caption         =   "&מחיקת תרופה"
      Height          =   495
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox lst_drugs 
      Height          =   2985
      Left            =   600
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "לצפייה/עריכה לחץ פעמיים"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lbl_find_comp 
      Alignment       =   2  'Center
      Caption         =   "שם חברה"
      Height          =   255
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "חיפוש לפי:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbl_caption_frm_drugs 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "מאגר התרופות"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2070
   End
End
Attribute VB_Name = "frm_drugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
sDBloc = App.Path & "\db_prj.mdb"
Set db_prj = OpenDatabase(sDBloc)
Set rs_tbl_n_d = db_prj.OpenRecordset("tbl_n_d", dbOpenTable)
Set rs_dyn_n_d = db_prj.OpenRecordset("tbl_n_d", dbOpenDynaset)
MySQL = "SELECT FROM WHERE ORDER BY"

rs_tbl_n_d.MoveFirst

 Do While Not rs_tbl_n_d.EOF
    lst_drugs.AddItem rs_tbl_n_d.Fields("txt_put_name_drug")
    rs_tbl_n_d.MoveNext
Loop

Call add_item_cbo_put_which_use(cbo_find_which_use)
cbo_find_which_use.AddItem "חברה"
cbo_find_which_use.AddItem "כולן"

End Sub

Private Sub cbo_find_which_use_Click()
rs_tbl_n_d.MoveFirst
lst_drugs.Clear

Select Case cbo_find_which_use.Text
    Case "נגד דלקות":
    Do While Not rs_tbl_n_d.EOF
        If rs_tbl_n_d.Fields("txt_cbo_put_which_use") = "נגד דלקות" Then
            lst_drugs.AddItem rs_tbl_n_d.Fields("txt_put_name_drug")
        End If
        rs_tbl_n_d.MoveNext
    Loop
    Case "תוספי מזון":
    Do While Not rs_tbl_n_d.EOF
        If rs_tbl_n_d.Fields("txt_cbo_put_which_use") = "תוספי מזון" Then
            lst_drugs.AddItem rs_tbl_n_d.Fields("txt_put_name_drug")
        End If
        rs_tbl_n_d.MoveNext
    Loop
    Case "נגד כאבים":
    Do While Not rs_tbl_n_d.EOF
        If rs_tbl_n_d.Fields("txt_cbo_put_which_use") = "נגד כאבים" Then
            lst_drugs.AddItem rs_tbl_n_d.Fields("txt_put_name_drug")
        End If
        rs_tbl_n_d.MoveNext
    Loop
    Case "שונות":
    Do While Not rs_tbl_n_d.EOF
        If rs_tbl_n_d.Fields("txt_cbo_put_which_use") = "שונות" Then
            lst_drugs.AddItem rs_tbl_n_d.Fields("txt_put_name_drug")
        End If
        rs_tbl_n_d.MoveNext
    Loop
    
    Case "חברה":
'        מכניס את שם החברה מתוך המאגר לתיבת הטקסט
        taking_name_comp_2_find_drug = True
        frm_comps.Show
    Case "כולן":
        Do While Not rs_tbl_n_d.EOF
            lst_drugs.AddItem rs_tbl_n_d.Fields("txt_put_name_drug")
            rs_tbl_n_d.MoveNext
        Loop
End Select
lbl_find_comp.Caption = "שם חברה"
End Sub

Private Sub cmd_add_new_drug_Click()
'מעלה את מסך יצירת תרופה חדשה
new_drug = True
Unload frm_drugs
frm_new_drug.Show
End Sub

Private Sub cmd_del_drug_Click()
'מוחק מהרשימה תרופה נבחרת
del_name = lst_drugs.Text
'yn = MsgBox("?האם אתה בטוח שברצונך למחוק את התרופה", vbYesNo + vbCritical, "!שים לב") = vbYes
yn = MsgBox("? האם ברצונך למחוק את התרופה : " + del_name, vbYesNo + vbCritical, "!שים לב") = vbYes
Set rs_dyn_uses_of_drug = db_prj.OpenRecordset("tbl_drugs_of_pa", dbOpenDynaset)

If yn Then
   If lst_drugs.ListIndex > -1 Then
      lst_drugs.RemoveItem lst_drugs.ListIndex
       rs_dyn_n_d.MoveFirst 'מונע טעות כיוון ריצה על בסיס הנתונים
       rs_dyn_n_d.Edit
       Do While Not rs_dyn_n_d.Fields("txt_put_name_drug") = del_name
           rs_dyn_n_d.MoveNext  'מגיע לתרופה בבסיס הנתונים תרופות
       Loop
       rs_dyn_n_d.Delete
'עדכון הטבלה בבסיס הנתונים של תרופות ושימושהן
        rs_dyn_uses_of_drug.MoveFirst
        rs_dyn_uses_of_drug.Edit
        Do While Not rs_dyn_uses_of_drug.EOF 'מוחק את כל השדות של לקוח זה ובונה אות םחדש לא יעיל אבל הרבה יותר פשוט
        If rs_dyn_uses_of_drug.Fields("name") = del_name Then
            rs_dyn_uses_of_drug.Delete
        End If
        rs_dyn_uses_of_drug.MoveNext
        Loop
    End If
End If

End Sub

Private Sub lbl_find_comp_Change()
'אם משתנה הכותרת יודע שצריך לחפש תרופות של החברה בכותרת
rs_tbl_n_d.MoveFirst

    Do While Not rs_tbl_n_d.EOF
        If rs_tbl_n_d.Fields("txt_put_comp") = lbl_find_comp.Caption Then
            lst_drugs.AddItem rs_tbl_n_d.Fields("txt_put_name_drug")
        End If
        rs_tbl_n_d.MoveNext
    Loop
End Sub

Private Sub lst_drugs_dblClick()
'פותח את הכרטיס של התרופה שעליה לחצן

If taking_name_drug_2_lst_ti_pa = False Then 'מאפשר לקיחת שם תרופה מהמאגר לרשימת תרופות הלקוח שנקנו
    new_drug = False
    go_drug_name = lst_drugs.Text
    frm_new_drug.Show
    Unload frm_drugs
Else 'לצורך הוספה בתוך הרשימה של התרופות שקנה הלקוח
    frm_mini_ti_pa.lst_trufot_used.AddItem lst_drugs.Text & " - " & (Format(Date, "m/d/yy"))
End If

End Sub
