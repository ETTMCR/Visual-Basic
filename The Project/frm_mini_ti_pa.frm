VERSION 5.00
Begin VB.Form frm_mini_ti_pa 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "רשימת תרופות שקנה הלקוח"
   ClientHeight    =   4935
   ClientLeft      =   4500
   ClientTop       =   3195
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_ok 
      Caption         =   "אישור "
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmd_look 
      Caption         =   "צפייה/עריכה פרטי הלקוח"
      Height          =   495
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ListBox lst_trufot_used 
      Height          =   2985
      Left            =   600
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "לחץ על תרופה לבחירה"
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmd_put_drug 
      Caption         =   "ה&כנס תרופה מתוך המאגר"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd_del_drug 
      Caption         =   "&מחק תרופה מהרשימה"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd_clr_lst 
      Caption         =   "מחק &רשימה"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   1440
      Picture         =   "frm_mini_ti_pa.frx":0000
      RightToLeft     =   -1  'True
      ScaleHeight     =   495
      ScaleWidth      =   1830
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label lbl_frm_mini 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "שם לקוח"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frm_mini_ti_pa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
sDBloc = App.Path & "\db_prj.mdb"
Set db_prj = OpenDatabase(sDBloc)
Set rs_tbl_ti_pa = db_prj.OpenRecordset("tbl_ti_pa", dbOpenTable)
Set rs_dyn_ti_pa = db_prj.OpenRecordset("tbl_ti_pa", dbOpenTable)
Set rs_dyn_drugs_of_pa = db_prj.OpenRecordset("tbl_drugs_of_pa", dbOpenDynaset)

    rs_dyn_ti_pa.MoveFirst
    rs_tbl_ti_pa.MoveFirst
    Do While Not rs_dyn_ti_pa.EOF 'מגיעה ללקוח שבו רצינו לצפות
        If go_paitiont_name = (rs_tbl_ti_pa.Fields("txt_put_name_first") + " " + rs_tbl_ti_pa.Fields("txt_put_name_last")) Then
            Exit Do
        End If
        rs_tbl_ti_pa.MoveNext
        rs_dyn_ti_pa.MoveNext 'לחסוך ריצה מהתחלה מתי שארצה למחוק תרופה
    Loop
    
lbl_frm_mini.Caption = go_paitiont_name 'כותרת החלון כשם הלקוח
    
rs_dyn_ti_pa.MoveFirst
rs_dyn_drugs_of_pa.MoveFirst
Do While Not rs_dyn_drugs_of_pa.EOF 'אין סיבה להשתמש בבוק מרק ממילא מתי שמוסיפים שדה אז הוא לא מתמיין
If rs_dyn_drugs_of_pa.Fields("name") = go_paitiont_name Then
    For i = 0 To rs_dyn_drugs_of_pa.Fields("how_much") 'מכניס את התרופות של הלקוח לרשימה
        If rs_dyn_drugs_of_pa.Fields(i) <> "" Then  'למנוע שגיאה -מוסיף כלום, במקרה שהוספנו לרשימה לקוח חדש וטרם יש לו תרופות
            lst_trufot_used.AddItem rs_dyn_drugs_of_pa.Fields(i)
        End If
    Next
End If
    rs_dyn_drugs_of_pa.MoveNext
Loop

End Sub

Private Sub cmd_del_drug_Click()
'מוחק תרופה מרשימת התרופות שהשתמש הלקוח
Dim yn As Integer
yn = MsgBox("?האם ברצונך למחוק את התרופה מרשימת הלקוח", vbYesNo + vbCritical, "!שים לב") = vbYes
If yn Then
    If lst_trufot_used.ListIndex > -1 Then 'אכן ניתן לבצע המחיקה
        lst_trufot_used.RemoveItem (lst_trufot_used.ListIndex)
    End If
End If
    If (lst_trufot_used.ListCount) = 0 Then
        cmd_del_drug.Enabled = False
    End If
End Sub

Private Sub cmd_clr_lst_Click()
'כפתור מחיקת רשימת תרופות שקנה הלקוח
Dim yn As Integer
    yn = MsgBox("?האם ברצונך למחוק את רשימת התרופות", vbYesNo + vbCritical, " ! שים לב ") = vbYes
    If yn Then
        lst_trufot_used.Clear
        cmd_del_drug.Enabled = False
    End If

End Sub

Private Sub cmd_put_drug_Click()
'מכניס את שם התרופה מתוך הקומבו לרשימת התרופות שנקנו
    taking_name_drug_2_lst_ti_pa = True
    cmd_del_drug.Enabled = True
    'משנה סטטוס הפריים מאגר התרופות
    frm_drugs.lst_drugs.ToolTipText = "להוספה לחץ פעמיים על תרופה"
    frm_drugs.cmd_add_new_drug.Enabled = False
    frm_drugs.cmd_del_drug.Enabled = False
    frm_drugs.Show
    
''    If Not cbo_put_use.Text = "אחר" Then
'        lst_trufut.AddItem =cbo_put_drug.Text
'        cmd_del_drug.Enabled = True
''    End If
End Sub

Private Sub cmd_look_Click()
new_ti_pa = False
Unload frm_mini_ti_pa
frm_ticket_paitiont.Show
End Sub

Private Sub cmd_ok_Click()
'עדכון הטבלה בבסיס הנתונים

        rs_dyn_drugs_of_pa.MoveFirst
        rs_dyn_drugs_of_pa.Edit
        Do While Not rs_dyn_drugs_of_pa.EOF 'מוחק את כל השדות של לקוח זה ובונה אות םחדש לא יעיל אבל הרבה יותר פשוט
        If rs_dyn_drugs_of_pa.Fields("name") = go_paitiont_name Then
            rs_dyn_drugs_of_pa.Delete
        End If
        rs_dyn_drugs_of_pa.MoveNext
        Loop
        
Countd = 0
If Not lst_trufot_used.ListCount = 0 Then 'בניית השדות מחדש במקרה שלא נמחקה כל הרשימה
'@@@@@ הסדר הנכון של השימוש באלה
            rs_dyn_drugs_of_pa.AddNew
            For i = 0 To lst_trufot_used.ListCount - 1
                If i Mod 5 = 0 Then 'מעבר למס העמודות האפשרי לשדה בודד
                    rs_dyn_drugs_of_pa.Update
                    Countd = 0
                    rs_dyn_drugs_of_pa.AddNew
                End If
                    rs_dyn_drugs_of_pa.Fields("name") = go_paitiont_name
                    rs_dyn_drugs_of_pa.Fields(Countd) = lst_trufot_used.List(i)
                    rs_dyn_drugs_of_pa.Fields("how_much") = Countd
                    Countd = Countd + 1
            Next
        rs_dyn_drugs_of_pa.Update
'Else 'לא צריך לשמור את שם הלקוח בטבלה אם אין לו תרופות כי ממילא הוא מוסיף אותו מחדש ברגע שקנה
'        rs_dyn_drugs_of_pa.Fields("name") = go_paitiont_name
End If

new_ti_pa = False
Unload frm_mini_ti_pa
frm_welcome.Show
End Sub

