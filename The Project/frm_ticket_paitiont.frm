VERSION 5.00
Begin VB.Form frm_ticket_paitiont 
   BackColor       =   &H80000013&
   Caption         =   "כרטיס לקוח"
   ClientHeight    =   7170
   ClientLeft      =   2880
   ClientTop       =   1470
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   9585
   Begin VB.CommandButton cmd_edit_ticket 
      Caption         =   "ל&עריכת פרטים"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Aharoni"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   25
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&אישור נתונים"
      BeginProperty Font 
         Name            =   "Aharoni"
         Size            =   14.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Picture         =   "frm_ticket_paitiont.frx":0000
      TabIndex        =   13
      ToolTipText     =   "אישור נתונים"
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame fra_personal_data 
      BackColor       =   &H80000003&
      Caption         =   "פרטי הלקוח"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1080
      Width           =   7695
      Begin VB.TextBox txt_put_age 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox cbo_put_num_web 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "בחירת רשת סלולארית"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txt_put_cell 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   11
         Text            =   "לא חובה"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox Chk_cell 
         Caption         =   "Check1"
         Height          =   255
         Left            =   6120
         TabIndex        =   26
         Top             =   1680
         Width           =   255
      End
      Begin VB.ComboBox cbo_put_city 
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "בחירת עיר"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ComboBox cbo_put_area 
         Height          =   315
         Left            =   4440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "בחירת מחוז"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txt_put_zip 
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   8
         ToolTipText     =   "חמש ספרות"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txt_put_num_home 
         Height          =   285
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   7
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txt_put_name_last 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txt_put_name_first 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txt_put_st 
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox cbo_put_phone_tt 
         DataField       =   "03"
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "בחירת אזור חיוג"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txt_put_phone_2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3960
         MaxLength       =   7
         TabIndex        =   4
         ToolTipText     =   "שבע ספרות"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbl_put_age 
         Alignment       =   2  'Center
         Caption         =   "גיל"
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   600
         Width           =   280
      End
      Begin VB.Label lbl_put_num_web 
         Alignment       =   2  'Center
         Caption         =   "מס' רשת"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lbl_put_cell 
         Alignment       =   2  'Center
         Caption         =   "מס' פלאפון"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lbl_put_phone_tt 
         Alignment       =   2  'Center
         Caption         =   "אזור חיוג"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl_put_city 
         Alignment       =   2  'Center
         Caption         =   "עיר"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lbl_put_area 
         Alignment       =   2  'Center
         Caption         =   "מחוז"
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label lbl_put_zip 
         Alignment       =   2  'Center
         Caption         =   "מיקוד"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lbl_put_num_home 
         Alignment       =   1  'Right Justify
         Caption         =   "מספר "
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lbl_put_st 
         Alignment       =   2  'Center
         Caption         =   "רחוב"
         Height          =   255
         Left            =   6000
         TabIndex        =   19
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label lbl_put_name_last 
         Alignment       =   2  'Center
         Caption         =   "שם פרטי"
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl_put_name_first 
         Alignment       =   2  'Center
         Caption         =   "שם משפחה"
         Height          =   375
         Left            =   5760
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lbl_put_phone 
         Alignment       =   2  'Center
         Caption         =   "מספר הטלפון"
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lbl_put_adrs 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "כתובת"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   177
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   3720
      Picture         =   "frm_ticket_paitiont.frx":040E
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2040
   End
   Begin VB.Label lbl_caption_frm_ti_pa 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "לקוח חדש"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frm_ticket_paitiont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
sDBloc = App.Path & "\db_prj.mdb"
Set db_prj = OpenDatabase(sDBloc)
Set rs_tbl_ti_pa = db_prj.OpenRecordset("tbl_ti_pa", dbOpenTable)
Set rs_dyn_ti_pa = db_prj.OpenRecordset("tbl_ti_pa", dbOpenDynaset)
Set rs_dyn_drugs_of_pa = db_prj.OpenRecordset("tbl_drugs_of_pa", dbOpenDynaset)

'Set rs_dyn_pts = db_prj.OpenRecordset("tbl_pts", dbOpenDynaset)

    Call add_item_put_area(cbo_put_area)
    Call add_item_put_phone_tt(cbo_put_phone_tt)
    Call add_item_put_num_web(cbo_put_num_web)
    
        If new_ti_pa Then 'אם נלחץ במסך הפתיחה הכפתור הוספת לקוח חדש
            rs_dyn_ti_pa.AddNew
        Else
            rs_dyn_ti_pa.Edit
        End If

If new_ti_pa = False Then 'הזנת הנתונים בהתאם ללקוח שנבחר מתוך המאגר

    rs_tbl_ti_pa.MoveFirst
     rs_dyn_ti_pa.MoveFirst
    Do While Not rs_tbl_ti_pa.EOF 'מגיעה לתרופה שבה רצינו לצפות
        If go_paitiont_name = (rs_tbl_ti_pa.Fields("txt_put_name_first") + " " + rs_tbl_ti_pa.Fields("txt_put_name_last")) Then
            Exit Do
        End If
        rs_tbl_ti_pa.MoveNext
        rs_dyn_ti_pa.MoveNext
    Loop
            cmd_edit_ticket.Enabled = True
'             MsgBox "       לחץ על כפתור העריכה" + vbCrLf + _
'            "! אם רצונך לערוך את פרטי הלקוח", , "!שים לב"
           lbl_put_adrs.Enabled = False
           fra_personal_data.Enabled = False 'מבטל אפשרות עריכה
           fra_personal_data.Caption = "פרטי הלקוח - מצב צפייה"
          txt_put_age.Text = rs_dyn_ti_pa.Fields("txt_put_age")
          txt_put_name_first.Text = rs_tbl_ti_pa.Fields("txt_put_name_first")
          txt_put_name_last.Text = rs_tbl_ti_pa.Fields("txt_put_name_last")
          txt_put_phone_2.Text = rs_tbl_ti_pa.Fields("txt_put_phone_2")
          txt_put_cell.Text = rs_tbl_ti_pa.Fields("txt_put_cell")
          txt_put_st.Text = rs_tbl_ti_pa.Fields("txt_put_st")
          txt_put_num_home.Text = rs_tbl_ti_pa.Fields("txt_put_num_home")
          txt_put_zip.Text = rs_tbl_ti_pa.Fields("txt_put_zip")
          lbl_caption_frm_ti_pa.Caption = txt_put_name_first.Text + " " + txt_put_name_last.Text  'כותרת הטופס כשם הלקוח
          cbo_put_num_web.Text = rs_tbl_ti_pa.Fields("txt_cbo_put_num_web")
          cbo_put_city.Text = rs_tbl_ti_pa.Fields("txt_cbo_put_city")
          cbo_put_area.Text = rs_tbl_ti_pa.Fields("txt_cbo_put_area")
          cbo_put_phone_tt.Text = rs_tbl_ti_pa.Fields("txt_cbo_put_phone_tt")
          
          If rs_tbl_ti_pa.Fields("txt_put_cell") <> "לא חובה" Then 'יש צורך באפשרות לערוך אותם
            Chk_cell.Value = 1
            txt_put_cell.Enabled = True
            lbl_put_cell.Enabled = True
            cbo_put_num_web.Enabled = True
            lbl_put_num_web.Enabled = True
           End If
           
End If 'of new_ti_pa = False
'@@@@@

    yes_num = False
End Sub

Private Sub cmd_ok_Click()
'בדיקות תקינות לקראת סיום
'עדכון נתונים עם הכל תקין

Dim ok_ticket_paitiont As Boolean

If cmd_edit_ticket.Enabled = True Then 'זאת אומרת אין צורך בבדיקת תקינות כי זה במצב צפייה
    Unload frm_new_drug 'המשתמש יוצא ממצב צפייה
    frm_welcome.Show
    Exit Sub
End If

'קורא לפונקציה שתבדוק תקינות לקראת סיום
Call ok_check_new_paitiont(txt_put_cell, txt_put_name_first, txt_put_name_last, txt_put_phone_2, cbo_put_phone_tt, txt_put_st, txt_put_num_home, txt_put_zip, cbo_put_city, cbo_put_area, ok_ticket_paitiont, Chk_cell.Value, cbo_put_num_web, txt_put_age)
    
'בדיקה האם השתנה שם הלקוח-אינו זהה לשם שאיתו נכנסנו למצב צפייה
'אם כן אז מעדכן שם זה בטבלה של תרופות שקנה הלקוח
If go_paitiont_name <> (txt_put_name_first.Text + " " + txt_put_name_last.Text) Then
Do While Not rs_dyn_drugs_of_pa.EOF
    If rs_dyn_drugs_of_pa.Fields("name") = go_paitiont_name Then 'משמע שתא זה צריך לעבור שינוי
        rs_dyn_drugs_of_pa.Edit
        rs_dyn_drugs_of_pa.Fields("name") = (txt_put_name_first.Text + " " + txt_put_name_last.Text)
        'rs_dyn_drugs_of_pa.Fields(i) = "null to fix"
        rs_dyn_drugs_of_pa.Update
    End If
    rs_dyn_drugs_of_pa.MoveNext
Loop
'rs_dyn_drugs_of_pa.Update
End If 'השם ישתנה
    
If ok_ticket_paitiont Then 'אם כל הנתונים תקינים מעדכן את טבלת הנתונים
        If new_ti_pa = False Then 'זאת אומרת שזו עריכה
            rs_dyn_ti_pa.Edit
            MsgBox "! פרטי הלקוח נערכו בהצלחה", , "עריכת לקוח "
        Else
            rs_dyn_drugs_of_pa.AddNew 'שם חדש
            rs_dyn_drugs_of_pa.Fields("name") = (txt_put_name_first.Text + " " + txt_put_name_last.Text) 'הוספת הלקוח בטבלה של התרופות שקנה
            rs_dyn_drugs_of_pa.Update 'ומיד סוגר
            MsgBox "! הלקוח נוסף לרשימת הלקוחות", , "הוספת לקוח חדש"
        End If
        rs_dyn_ti_pa.Fields("txt_put_age") = txt_put_age.Text
        rs_dyn_ti_pa.Fields("txt_put_name_first") = txt_put_name_first.Text
        rs_dyn_ti_pa.Fields("txt_put_name_last") = txt_put_name_last.Text
        rs_dyn_ti_pa.Fields("txt_put_phone_2") = txt_put_phone_2.Text
        rs_dyn_ti_pa.Fields("txt_cbo_put_phone_tt") = cbo_put_phone_tt.Text
        rs_dyn_ti_pa.Fields("txt_put_cell") = txt_put_cell.Text
        rs_dyn_ti_pa.Fields("txt_cbo_put_num_web") = cbo_put_num_web.Text
        rs_dyn_ti_pa.Fields("txt_put_st") = txt_put_st.Text
        rs_dyn_ti_pa.Fields("txt_put_num_home") = txt_put_num_home.Text
        rs_dyn_ti_pa.Fields("txt_put_zip") = txt_put_zip.Text
        rs_dyn_ti_pa.Fields("txt_cbo_put_area") = cbo_put_area.Text
        rs_dyn_ti_pa.Fields("txt_cbo_put_city") = cbo_put_city.Text
        rs_dyn_ti_pa.Update
        new_ti_pa = False
        
      yes_num = False
        
        Unload frm_ticket_paitiont
        frm_welcome.Show
End If 'of ok_ticket_paitiont=True

End Sub
Private Sub cbo_put_area_Change()
    cbo_put_area.BackColor = &H80000005
End Sub

Private Sub cbo_put_area_click()
    cbo_put_area.BackColor = &H80000005
    'Call a_area_to_00(cbo_put_area, cbo_put_phone_tt)
    
    If cmd_edit_ticket.Enabled = False Then 'לא צריך לבדוק אם במצב צפייה
        '***this change      the function       what get
        cbo_put_phone_tt.Text = a_area_to_00(cbo_put_area.Text)
        Call area_to_city(cbo_put_area.Text, cbo_put_city)
    End If
    
End Sub
Private Sub cbo_put_num_web_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cbo_put_phone_tt_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 'לא מאפשר לכתוב
End Sub

Private Sub cbo_put_phone_tt_Change()
    cbo_put_phone_tt.BackColor = &H80000005
End Sub

Private Sub cbo_put_city_Click()
'??????
    Dim else_city As String
    cbo_put_city.BackColor = &H80000005
    If cbo_put_city.Text = "אחר" Then
        else_city = InputBox("הכנס שם עיר אחרת", "עיר אחרת")
        If Not (else_city = "") Then 'זאת אומרת באמת הכניס ערך
            cbo_put_city.AddItem else_city
            cbo_put_city.Text = else_city
        End If
    End If
    If cbo_put_city.Text = "אחר" Then 'לחץ לאישור
    cbo_put_city.Text = ""
    End If
End Sub

Private Sub cbo_put_city_KeyPress(KeyAscii As Integer)
KeyAscii = 0 'יכול לשנות כותרת רק באמצעות אחר
End Sub

Private Sub cbo_put_num_web_Click()
cbo_put_num_web.BackColor = &H80000005
End Sub

Private Sub cbo_put_phone_tt_Click()
    cbo_put_phone_tt.BackColor = &H80000005
    cbo_put_area.Text = a_00_to_area(cbo_put_phone_tt.Text)
    Call area_to_city(cbo_put_area.Text, cbo_put_city)
End Sub


Private Sub Chk_cell_Click()
' כאשר לא מסומנת לא מאפשרת כתיבת ערכי רשות של מספר פלאפון
If cmd_edit_ticket.Enabled = False Then 'יש צורך בתנאי למטה רק עם לא במצב צפייה
    If Chk_cell.Value = 1 Then
        txt_put_cell.Text = ""
        txt_put_cell.Enabled = True
        lbl_put_cell.Enabled = True
        cbo_put_num_web.Enabled = True
        lbl_put_num_web.Enabled = True
    Else
        txt_put_cell.Text = "לא חובה"
        cbo_put_num_web.Text = ""
        txt_put_cell.Enabled = False
        lbl_put_cell.Enabled = False
        cbo_put_num_web.Enabled = False
        lbl_put_num_web.Enabled = False
    End If
End If 'of cmd_edit_ticket.Enabled = False
End Sub


Private Sub cmd_edit_ticket_Click()
'לחצן שמאפשר עריכת הנתונים במידה שמתחיל בכרטיס לקוח קיים
MsgBox "!כל שינוי שתבצע ישמר", vbCritical, "!שים לב"
frm_ticket_paitiont.fra_personal_data.Enabled = True
fra_personal_data.Caption = "פרטי הלקוח"
cmd_edit_ticket.Enabled = False
lbl_put_adrs.Enabled = True
End Sub

Private Sub txt_age_KeyPress(KeyAscii As Integer)
'בודק את התו אם אינו מתאים
    Call chr_chk_num(KeyAscii, txt_put_num_home)
End Sub

Private Sub txt_put_age_Change()
txt_put_age.BackColor = &H80000005
End Sub

Private Sub txt_put_cell_KeyPress(KeyAscii As Integer)
'בודק את התו אם אינו מתאים
    Call chr_chk_num(KeyAscii, txt_put_cell)
End Sub

Private Sub txt_put_name_first_Change()
    lbl_caption_frm_ti_pa.Caption = ""
   lbl_caption_frm_ti_pa.Caption = txt_put_name_first.Text + " " + txt_put_name_last.Text
End Sub

Private Sub txt_put_name_first_KeyPress(KeyAscii As Integer)
Dim Char As String
'בודק את התו אם אינו מתאים
Char = Chr(KeyAscii)
 If ((Asc(Char) > 250) Or (Asc(Char) < 224) And (Asc(Char) <> 8) And (Asc(Char) <> 32)) Then
    Beep
    KeyAscii = 0
 End If
txt_put_name_first.BackColor = &H80000005
End Sub

Private Sub txt_put_name_last_Change()
lbl_caption_frm_ti_pa.Caption = ""
lbl_caption_frm_ti_pa.Caption = txt_put_name_first.Text + " " + txt_put_name_last.Text
End Sub

Private Sub txt_put_name_last_KeyPress(KeyAscii As Integer)
'בודק את התו אם אינו מתאים
Dim Char As String
Char = Chr(KeyAscii)
 If ((Asc(Char) > 250) Or (Asc(Char) < 224) And (Asc(Char) <> 8) And (Asc(Char) <> 32)) Then
    Beep
    KeyAscii = 0
 End If
txt_put_name_last.BackColor = &H80000005
End Sub

Private Sub txt_put_num_home_KeyPress(KeyAscii As Integer)
'בודק את התו אם אינו מתאים
    Call chr_chk_num(KeyAscii, txt_put_num_home)
End Sub

Private Sub txt_put_phone_2_keypress(KeyAscii As Integer)
'בודק את התו אם אינו מתאים
    Call chr_chk_num(KeyAscii, txt_put_phone_2)
End Sub

Private Sub txt_put_st_KeyPress(KeyAscii As Integer)
'בודק את התו אם ספרה או אות עברית
Dim Char As String
    Char = Chr(KeyAscii)

If (((Asc(Char)) > 57) Or ((Asc(Char)) < 48) And (Asc(Char) <> 8) And (Asc(Char) <> 32)) Then
    If (((Asc(Char)) > 250) Or ((Asc(Char)) < 224) And (Asc(Char) <> 8) And (Asc(Char) <> 32)) Then
        Beep
        KeyAscii = 0
     End If
End If
    txt_put_st.BackColor = &H80000005

End Sub

Private Sub txt_put_zip_KeyPress(KeyAscii As Integer)
'בודק את התו אם אינו מתאים
    Call chr_chk_num(KeyAscii, txt_put_zip)
End Sub
