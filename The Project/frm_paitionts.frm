VERSION 5.00
Begin VB.Form frm_paitionts 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "מאגר הלקוחות"
   ClientHeight    =   6405
   ClientLeft      =   5160
   ClientTop       =   2535
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   3540
   Begin VB.CommandButton cmd_add_new_ti_pa 
      Caption         =   "הו&ספת לקוח חדש"
      Height          =   495
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmd_del_pa 
      Caption         =   "&מחק לקוח"
      Height          =   495
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.ListBox lst_paitionts 
      Height          =   3960
      Left            =   720
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "לצפייה לחץ פעמיים"
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl_caption_frm_paitionts 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "מאגר הלקוחות"
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
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2040
   End
End
Attribute VB_Name = "frm_paitionts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_add_new_ti_pa_Click()
'מעלה את מסך יצירת כרטיס לקוח חדש
new_ti_pa = True
Unload frm_paitionts
frm_ticket_paitiont.Show
End Sub

Private Sub Form_Load()
sDBloc = App.Path & "\db_prj.mdb"
Set db_prj = OpenDatabase(sDBloc)
Set rs_tbl_ti_pa = db_prj.OpenRecordset("tbl_ti_pa", dbOpenTable)
Set rs_dyn_ti_pa = db_prj.OpenRecordset("tbl_ti_pa", dbOpenDynaset)
Do While Not rs_tbl_ti_pa.EOF
    lst_paitionts.AddItem (rs_tbl_ti_pa.Fields("txt_put_name_first") + " " + rs_tbl_ti_pa.Fields("txt_put_name_last"))
    rs_tbl_ti_pa.MoveNext
Loop
End Sub

Private Sub cmd_del_pa_Click()
'מוחק מהרשימה לקוח מסומן
Set rs_dyn_drugs_of_pa = db_prj.OpenRecordset("tbl_drugs_of_pa", dbOpenDynaset)

del_name = lst_paitionts.Text
yn = MsgBox("? האם ברצונך למחוק את הלקוח : " + del_name + " ", vbYesNo + vbCritical, "!שים לב") = vbYes
If yn Then
    If lst_paitionts.ListIndex > -1 Then
        lst_paitionts.RemoveItem lst_paitionts.ListIndex
       rs_dyn_ti_pa.MoveFirst 'מונע טעות כיוון ריצה על בסיס הנתונים
       rs_dyn_ti_pa.Edit
       Do While Not (rs_dyn_ti_pa.Fields("txt_put_name_first") + " " + rs_dyn_ti_pa.Fields("txt_put_name_last")) = del_name
           rs_dyn_ti_pa.MoveNext
       Loop
       rs_dyn_ti_pa.Delete
       
       'מחיקת הלקוח מטבלת התרופות שקנה
        rs_dyn_drugs_of_pa.MoveFirst
        rs_dyn_drugs_of_pa.Edit
        Do While Not rs_dyn_drugs_of_pa.EOF
        If rs_dyn_drugs_of_pa.Fields("name") = go_paitiont_name Then
            rs_dyn_drugs_of_pa.Delete
        End If
        rs_dyn_drugs_of_pa.MoveNext
        Loop
        'rs_dyn_drugs_of_pa.Update
        
    End If
End If 'of yn
End Sub

Private Sub lst_paitionts_dblClick()
'פותח את הכרטיס של הלקוח שעליו לחצו
new_ti_pa = False
go_paitiont_name = lst_paitionts.Text
frm_mini_ti_pa.Show
Unload frm_paitionts
End Sub
