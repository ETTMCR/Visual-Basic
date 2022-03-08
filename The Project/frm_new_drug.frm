VERSION 5.00
Begin VB.Form frm_new_drug 
   BackColor       =   &H80000013&
   Caption         =   "כרטיס תרופה"
   ClientHeight    =   7920
   ClientLeft      =   2625
   ClientTop       =   1725
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10125
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
      Height          =   615
      Left            =   3240
      TabIndex        =   18
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_ok_new_drug 
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
      Height          =   615
      Left            =   4800
      TabIndex        =   12
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame fra_drug_data 
      BackColor       =   &H80000003&
      Caption         =   "פרטי התרופה"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   177
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   8775
      Begin VB.Frame fra_mirsham 
         Caption         =   "מירשם רופא ?"
         Height          =   615
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1920
         Width           =   1455
         Begin VB.OptionButton opt_mirsham 
            Alignment       =   1  'Right Justify
            Caption         =   "עם"
            Height          =   255
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton opt_no_mirsham 
            Alignment       =   1  'Right Justify
            Caption         =   "בלי"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fra_warn 
         Caption         =   "אזהרות"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   3000
         Width           =   2415
         Begin VB.CheckBox chk_no 
            Alignment       =   1  'Right Justify
            Caption         =   "ללא אזהרות"
            Height          =   255
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CheckBox chk_1 
            Alignment       =   1  'Right Justify
            Caption         =   "לא לתינוקות"
            Height          =   255
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chk_2 
            Alignment       =   1  'Right Justify
            Caption         =   "לא לזקנים "
            Height          =   255
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chk_3 
            Alignment       =   1  'Right Justify
            Caption         =   "לא לנשים בהריון"
            Height          =   255
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox Chk_warn_else 
            Alignment       =   1  'Right Justify
            Caption         =   "אחר"
            Height          =   255
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CheckBox chk_4 
            Alignment       =   1  'Right Justify
            Caption         =   "לא לבעלי לחץ דם גבוה"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   1320
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmd_put_comp 
         Caption         =   "בחר חברה ממאגר"
         Height          =   495
         Left            =   6480
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame fra_type 
         Caption         =   "סוג תכשיר "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   2175
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            Caption         =   "זריקה"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            Caption         =   "משחה"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            Caption         =   "סירופ"
            Height          =   255
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            Caption         =   "כדור"
            Height          =   255
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "גלולה"
            Height          =   255
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton opt_type_else 
            Alignment       =   1  'Right Justify
            Caption         =   "אחר"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1080
            Width           =   855
         End
      End
      Begin VB.Frame fra_put_use 
         Caption         =   "התויה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3000
         Width           =   4695
         Begin VB.CommandButton cmd_clr_lst 
            Caption         =   "מחק &כל הרשימה"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1920
            TabIndex        =   36
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmd_put_in 
            Caption         =   "הו&סף השימוש לרשימה"
            Height          =   375
            Left            =   1920
            TabIndex        =   17
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CommandButton cmd_del_lst_use 
            Caption         =   "מ&חק שימוש מהרשימה"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3240
            TabIndex        =   15
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox cbo_put_which_use 
            Height          =   315
            Left            =   2040
            RightToLeft     =   -1  'True
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   360
            Width           =   1455
         End
         Begin VB.ListBox lstbx_put_use 
            Height          =   2010
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   10
            ToolTipText     =   "לחץ על שימוש לבחירה"
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox cbo_put_use 
            Height          =   315
            ItemData        =   "frm_new_drug.frx":0000
            Left            =   2040
            List            =   "frm_new_drug.frx":0002
            RightToLeft     =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label lbl_put_use 
            Alignment       =   1  'Right Justify
            Caption         =   "שימוש"
            Height          =   255
            Left            =   3960
            TabIndex        =   14
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lbl_put_which_use 
            Alignment       =   1  'Right Justify
            Caption         =   "שימוש כללי"
            Height          =   255
            Left            =   3600
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txt_put_cost 
         Height          =   285
         Left            =   6480
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txt_put_comp 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "בחר /הוסף ממאגר"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txt_put_name_drug 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   600
         Picture         =   "frm_new_drug.frx":0004
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbl_put_cost 
         Alignment       =   1  'Right Justify
         Caption         =   "מחיר ב-₪"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lbl_put_comp 
         Alignment       =   1  'Right Justify
         Caption         =   "חברה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   6
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbl_put_name_drug 
         Alignment       =   2  'Center
         Caption         =   "שם תרופה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Label lbl_caption_new_drug 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "תרופה חדשה"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   6525
   End
End
Attribute VB_Name = "frm_new_drug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mone_chk As Integer 'שומר בערך בינארי איזה אזהרות נלחצו(כך שאין לעולם טעויות)


Private Sub Form_Load()

sDBloc = App.Path & "\db_prj.mdb"
Set db_prj = OpenDatabase(sDBloc)
Set rs_tbl_n_d = db_prj.OpenRecordset("tbl_n_d", dbOpenTable)
Set rs_dyn_n_d = db_prj.OpenRecordset("tbl_n_d", dbOpenDynaset)
Set rs_dyn_comps = db_prj.OpenRecordset("tbl_comps", dbOpenDynaset)
Set rs_dyn_uses_of_drug = db_prj.OpenRecordset("tbl_uses_of_drug", dbOpenDynaset)

Call add_item_cbo_put_which_use(cbo_put_which_use)

         mone_chk = 0
        If new_drug Then 'אם נלחץ במסך הפתיחה הכפתור הוספת תרופה חדשה
            rs_dyn_comps.AddNew
            rs_dyn_n_d.AddNew
        End If

        If taking_comp Then 'נלחץ הוספת תרופה חדשה לחברה נוכחית
            txt_put_comp.Text = go_comp_name
        Else
             txt_put_comp.Text = "" 'למנוע שאם מיד לחצת חדשה אז ישאר שם חברה קודמת
        End If
        
If new_drug = False Then ' הזנת הנתונים בהתאם לתרופה הקיימת שנבחרה מתוך המאגר-מצב צפייה
          cmd_edit_ticket.Enabled = True
          fra_drug_data.Enabled = False 'מבטל אפשרות עריכה
          fra_drug_data.Caption = "פרטי התרופה - מצב צפייה"
          
         Do While Not rs_dyn_n_d.Fields("txt_put_name_drug") = go_drug_name
                rs_dyn_n_d.MoveNext
            Loop
            
'צריך לעדכן את הרשימה של שימושי התרופה באם נכסנו למצב צפייה
rs_dyn_uses_of_drug.MoveFirst
Do While Not rs_dyn_uses_of_drug.EOF
If rs_dyn_uses_of_drug.Fields("name") = go_drug_name Then
    For i = 0 To rs_dyn_uses_of_drug.Fields("how_much") 'מכניס את השימושים של התרופה לרשימה
        If rs_dyn_uses_of_drug.Fields(i) <> "" Then  'למנוע שגיאה -מוסיף כלום, במקרה שהוספנו לרשימה לקוח חדש וטרם יש לו תרופות
            lstbx_put_use.AddItem rs_dyn_uses_of_drug.Fields(i)
        End If
    Next
End If
    rs_dyn_uses_of_drug.MoveNext
Loop
cmd_del_lst_use.Enabled = True
            
          txt_put_name_drug.Text = rs_dyn_n_d.Fields("txt_put_name_drug")
          txt_put_comp.Text = rs_dyn_n_d.Fields("txt_put_comp")
          txt_put_cost.Text = rs_dyn_n_d.Fields("txt_put_cost")
          cbo_put_which_use.Text = rs_dyn_n_d.Fields("txt_cbo_put_which_use")
         
        If rs_dyn_n_d.Fields("mirsham") = True Then 'לא צויין עם או בלי מירשם רופא
            opt_mirsham.Value = True
        Else
            opt_no_mirsham.Value = True
        End If
        
          Select Case rs_dyn_n_d.Fields("cpt_opt_put_type") 'בדיקה איזה סוג תכשיר נבחר
                Case "גלולה": Option1.Value = True
                Case "כדור": Option2.Value = True
                Case "סירופ": Option3.Value = True
                Case "משחה": Option4.Value = True
                Case "זריקה": Option5.Value = True
                Case Else:
                    opt_type_else.Value = True
                    opt_type_else.Caption = rs_dyn_n_d.Fields("cpt_opt_put_type")
            End Select
            
            ' בדיקות איזה אזהרות סומנו
            'mone_chk משתנה מתי שיש אירוע שינוי בהם
            If rs_dyn_n_d.Fields("Chk_no") = "" Then
                If rs_dyn_n_d.Fields("Chk_warn_else") <> "אחר" Then
                    Chk_warn_else.Value = 1
                    Chk_warn_else.Caption = rs_dyn_n_d.Fields("Chk_warn_else")
                End If
                If rs_dyn_n_d.Fields("Chk_1") = "True" Then
                    chk_1.Value = 1 '
                 End If
                If rs_dyn_n_d.Fields("Chk_2") = "True" Then
                    chk_2.Value = 1
                 End If
                If rs_dyn_n_d.Fields("Chk_3") = "True" Then
                    chk_3.Value = 1
                 End If
                If rs_dyn_n_d.Fields("Chk_4") = "True" Then
                    chk_4.Value = 1
                 End If
            Else 'rs_dyn_n_d.Fields("Chk_no") = "true"
                    chk_no.Value = 1
                  '  mone_chk = 16 + mone_chk
            End If 'rs_dyn_n_d.Fields("Chk_no") = ""
                 
          lbl_caption_new_drug.Caption = txt_put_name_drug.Text 'כותרת הטופס כשם התרופה
'          MsgBox "       לחץ על כפתור העריכה" + vbCrLf + _
'           "! אם רצונך לערוך את פרטי התרופה", , "!שים לב"
           
End If 'of new_drug = False



End Sub

Private Sub cmd_ok_new_drug_Click()
'בדיקות תקינות לקראת סיום
'עדכון נתונים אם הכל תקין

Dim ok_new_drug As Boolean

If cmd_edit_ticket.Enabled = True Then '  זאת אומרת אין צורך בבדיקת תקינות כי זה במצב צפייה לא חל שום שינוי
    Unload frm_new_drug
    frm_welcome.Show
    Exit Sub
End If

    'קורא לפונקציה שתבדוק תקינות לקראת סיום
    Call ok_check_new_drug(txt_put_name_drug, txt_put_comp, txt_put_cost, cbo_put_which_use, lstbx_put_use, ok_new_drug, opt_type_else, Option1, Option2, Option3, Option4, Option5, fra_type)
    
    If mone_chk = 0 Then 'זאת אומרת לא צויינה שום אזהרה
        MsgBox "לא ציינת אזהרה", , "!שים לב "
        fra_warn.BackColor = &HFFC0FF
        Exit Sub 'אין צורך בבדיקת תקינות יותר
    End If
    
'בדיקה האם השתנה שם הלקוח-אינו זהה לשם שאיתו נכנסנו למצב צפייה
'אם כן אז מעדכן שם זה בטבלה של תרופות שקנה הלקוח
If go_drug_name <> txt_put_name_drug.Text Then
Do While Not rs_dyn_uses_of_drug.EOF
    If rs_dyn_uses_of_drug.Fields("name") = go_drug_name Then 'משמע שתא זה צריך לעבור שינוי
        rs_dyn_uses_of_drug.Edit
        rs_dyn_uses_of_drug.Fields("name") = txt_put_name_drug.Text
        'rs_dyn_uses_of_drug.Fields(i) = "null to fix"
        rs_dyn_uses_of_drug.Update
    End If
    rs_dyn_uses_of_drug.MoveNext
Loop
'rs_dyn_uses_of_drug.Update
End If 'השם ישתנה
                        
                        
If ok_new_drug Then  'אם כל הנתונים תקינים מעדכן את טבלת הנתונים
       
        If new_drug = False Then ' cmd_edit_ticket.Enabled'זאת אומרת שזו עריכה
        '#####
            rs_dyn_n_d.Edit '?????
            rs_dyn_comps.Edit
             rs_dyn_comps.AddNew
        End If

    If opt_mirsham.Value = True Then 'לא צויין עם או בלי מירשם רופא
            rs_dyn_n_d.Fields("mirsham") = True
    Else
        If opt_no_mirsham.Value = True Then
            rs_dyn_n_d.Fields("mirsham") = False
        Else 'שתי האופציות לא מסומנות
            fra_mirsham.BackColor = &HFFC0FF
            MsgBox "לא ציינת עם או בלי מירשם רופא", , "! שים לב"
            Exit Sub 'אין צורך בבדיקת תקינות יותר
        End If
    End If
        
'@@@@@@@
'עדכון הטבלה בבסיס הנתונים
        rs_dyn_uses_of_drug.MoveFirst
        rs_dyn_uses_of_drug.Edit
        Do While Not rs_dyn_uses_of_drug.EOF 'מוחק את כל השדות של לקוח זה ובונה אות םחדש לא יעיל אבל הרבה יותר פשוט
        If rs_dyn_uses_of_drug.Fields("name") = txt_put_name_drug.Text Then
            rs_dyn_uses_of_drug.Delete
        End If
        rs_dyn_uses_of_drug.MoveNext
        Loop
Countd = 0
If Not lstbx_put_use.ListCount = 0 Then 'בניית השדות מחדש במקרה שלא נמחקה כל הרשימה
            rs_dyn_uses_of_drug.AddNew
            For i = 0 To lstbx_put_use.ListCount - 1
                If i Mod 5 = 0 Then 'מעבר למס העמודות האפשרי לשדה בודד
                    rs_dyn_uses_of_drug.Update
                    Countd = 0
                    rs_dyn_uses_of_drug.AddNew
                End If
                    rs_dyn_uses_of_drug.Fields("name") = txt_put_name_drug.Text
                    rs_dyn_uses_of_drug.Fields(Countd) = lstbx_put_use.List(i)
                    rs_dyn_uses_of_drug.Fields("how_much") = Countd
                    Countd = Countd + 1
            Next
        rs_dyn_uses_of_drug.Update
End If
'@@@@@@@

        If chk_no.Value = 0 Then
        rs_dyn_n_d.Fields("Chk_no") = ""
         If chk_1.Value = 1 Then
            rs_dyn_n_d.Fields("Chk_1") = "True"
        End If
         If chk_2.Value = 1 Then
            rs_dyn_n_d.Fields("Chk_2") = "True"
        End If
         If chk_3.Value = 1 Then
            rs_dyn_n_d.Fields("Chk_3") = "True"
        End If
         If chk_4.Value = 1 Then
            rs_dyn_n_d.Fields("Chk_4") = "True"
        End If
         If Chk_warn_else.Value = 1 Then
            rs_dyn_n_d.Fields("Chk_warn_else") = Chk_warn_else.Caption
        End If
        rs_dyn_n_d.Fields("Chk_no") = ""
        Else ' chk_no.Value = 1
            rs_dyn_n_d.Fields("Chk_no") = "True" '????
        End If 'of chk_no.Value = 0
        
        rs_dyn_n_d.Fields("txt_cbo_put_which_use") = cbo_put_which_use.Text
        rs_dyn_n_d.Fields("Chk_warn_else") = Chk_warn_else.Caption
        rs_dyn_n_d.Fields("txt_put_name_drug") = txt_put_name_drug.Text
        rs_dyn_n_d.Fields("txt_put_comp") = txt_put_comp.Text
        rs_dyn_comps.Edit
        rs_dyn_comps.Fields("txt_put_name_comp") = txt_put_comp.Text
        rs_dyn_n_d.Fields("txt_put_cost") = txt_put_cost.Text
        
        'בדיקות איזה סוג שימוש נבחר
        If Option1.Value = True Then
            rs_dyn_n_d.Fields("cpt_opt_put_type") = Option1.Caption
        End If
        If Option2.Value = True Then
            rs_dyn_n_d.Fields("cpt_opt_put_type") = Option2.Caption
        End If
        If Option3.Value = True Then
            rs_dyn_n_d.Fields("cpt_opt_put_type") = Option3.Caption
        End If
        If Option4.Value = True Then
            rs_dyn_n_d.Fields("cpt_opt_put_type") = Option4.Caption
        End If
        If Option5.Value = True Then
            rs_dyn_n_d.Fields("cpt_opt_put_type") = Option5.Caption
        End If
        If opt_type_else.Value = True Then
            rs_dyn_n_d.Fields("cpt_opt_put_type") = opt_type_else.Caption
        End If
                            
        If new_drug = False Then 'איזו הודעה להראות
            MsgBox "! פרטי התרופה נערכו בהצלחה", , "עריכת תרופה "
        Else
            MsgBox "! התרופה נוספה לרשימת התרופות", , "הוספת תרופה חדשה"
        End If
        
        rs_dyn_n_d.Update
        rs_dyn_comps.Update
        
        new_drug = False 'איפוס משתנים גלובליים
        taking_comp = False
        taking_name_comp_2_drug = False
                 
        Unload frm_new_drug
        frm_welcome.Show
End If 'of ok_new_drug = True

End Sub
Private Sub cbo_put_which_use_KeyPress(KeyAscii As Integer)
KeyAscii = 0 'לא מאפשר לכתוב
End Sub

Private Sub chk_1_Click()
If chk_1.Value = 1 Then
    mone_chk = mone_chk + 1
Else
    mone_chk = mone_chk - 1
End If
End Sub

Private Sub chk_2_Click()
If chk_2.Value = 1 Then
    mone_chk = mone_chk + 2
Else
    mone_chk = mone_chk - 2
End If
End Sub

Private Sub chk_4_Click()
If chk_4.Value = 1 Then
    mone_chk = mone_chk + 8
Else
    mone_chk = mone_chk - 8
End If
End Sub

Private Sub chk_3_Click()
If chk_3.Value = 1 Then
    mone_chk = mone_chk + 4
Else
    mone_chk = mone_chk - 4
End If
End Sub

Private Sub chk_no_Click()
If chk_no.Value = 1 Then 'And cmd_edit_ticket.Enabled = False Then 'מעכשיו לא צריך את שאר ההזהרות
    mone_chk = mone_chk + 16
    chk_1.Enabled = False
    chk_2.Enabled = False
    chk_3.Enabled = False
    chk_4.Enabled = False
    Chk_warn_else.Caption = "אחר"
    Chk_warn_else.Enabled = False
    chk_1.Value = 0
    chk_2.Value = 0
    chk_3.Value = 0
    chk_4.Value = 0
    Chk_warn_else.Value = 0
Else
    chk_1.Enabled = True
    chk_2.Enabled = True
    chk_3.Enabled = True
    chk_4.Enabled = True
    Chk_warn_else.Enabled = True
    mone_chk = mone_chk - 16
End If
End Sub

Private Sub cbo_put_use_Click()
'?????
'lstbx_put_use.AddItem cbo_put_use.Text' , [cbo_put_use.ListIndex]) = cbo_put_use.Text'?????רק שנבחר
Dim else_use As String
cmd_put_in.Enabled = True
If cbo_put_use.Text = "אחר" Then
    else_use = InputBox("הכנס סוג שימוש אחר", "שימוש אחר")
    If Not (else_use = "") Then
        lstbx_put_use.AddItem else_use
        lstbx_put_use.BackColor = &H80000005
    End If
End If

End Sub


Private Sub cbo_put_which_use_Click()
Call which_to_use(cbo_put_which_use.Text, cbo_put_use)
cbo_put_which_use.BackColor = &H80000005
End Sub

Private Sub Chk_warn_else_Click()
'לחצן הוספת אזהרה אחרת
fra_type.BackColor = &H80000005
Dim else_warn As String

    If Chk_warn_else.Value = 1 And cmd_edit_ticket.Enabled = False Then
        else_warn = InputBox(" כתוב אזהרה אחרת ", " אזהרה אחרת")
        If Not (else_warn = "") Then
            Chk_warn_else.Caption = else_warn
        Else
            Chk_warn_else.Value = 0 'זאת אומרת לחץ אך לא כתב כלום
        End If
    End If
    
    If Chk_warn_else.Value = 1 And cmd_edit_ticket.Enabled = True Then
        mone_chk = mone_chk + 32
    End If
        
    If Chk_warn_else.Value = 0 And cmd_edit_ticket.Enabled = False Then
        Chk_warn_else.Caption = "אחר"
        mone_chk = mone_chk - 32
    End If
    
    If Chk_warn_else.Value = 1 And cmd_edit_ticket.Enabled = False Then
        mone_chk = mone_chk + 32
    End If
    
End Sub

Private Sub cmd_edit_ticket_Click()
'לחצן שמאפשר עריכת הנתונים במידה שהתחיל עם נתוני תרופה קיימת
MsgBox "!כל שינוי שתבצע ישמר", vbCritical, "!שים לב"
fra_drug_data.Caption = "פרטי תרופה"
frm_new_drug.fra_drug_data.Enabled = True
cmd_edit_ticket.Enabled = False
End Sub


Private Sub cmd_del_lst_use_Click()
'מוחק מהרשימה שימוש נבחר
Dim yn As Integer
yn = MsgBox("?האם אתה בטוח שברצונך למחוק שימוש", vbYesNo, "!שים לב") = vbYes
If yn Then
    If lstbx_put_use.ListIndex > -1 Then
        lstbx_put_use.RemoveItem lstbx_put_use.ListIndex
    End If
End If

If lstbx_put_use.ListCount = 0 Then
    cmd_del_lst_use.Enabled = False
End If

End Sub
Private Sub cmd_clr_lst_Click()
'מוחק את כל הרשימה
yn = MsgBox("?האם אתה בטוח שברצונך למחוק שימוש", vbYesNo, "!שים לב") = vbYes
If yn Then
lstbx_put_use.Clear
End If
End Sub

Private Sub cmd_put_comp_Click()
'מכניס את שם החברה מתוך המאגר לתיבת הטקסט
taking_name_comp_2_drug = True
frm_comps.Show

End Sub

Private Sub cmd_put_in_Click()
'מעביר את שם השימוש הנבחר לרשימה מתוך הקומבו
cmd_put_in.Enabled = False 'בשביל שלא יכניס בטעות אותו שימוש פעמיים מיד
If Not cbo_put_use.Text = "אחר" And cbo_put_use.Text <> "" Then
    lstbx_put_use.AddItem cbo_put_use.Text
    cmd_del_lst_use.Enabled = True
    cmd_clr_lst.Enabled = True
    lstbx_put_use.BackColor = &H80000005
End If

End Sub

Private Sub fra_warn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'fra_warn.BackColor = &H8000000F
End Sub
Private Sub Image1_DblClick()
MsgBox "? נו זו לא מספיק הפתעה ", , " ! ביצת הפתעה"
End Sub
Private Sub opt_mirsham_Click()
fra_mirsham.BackColor = &H8000000F
End Sub
Private Sub opt_no_mirsham_Click()
fra_mirsham.BackColor = &H8000000F
End Sub

Private Sub opt_type_else_Click()
'לחצן הוספת אופן שימוש אחר
  Dim else_type As String
fra_type.BackColor = &H8000000F
If cmd_edit_ticket.Enabled = False Then
    else_type = InputBox("כתוב סוג תרופה אחר ", " סוג שימוש אחר")
End If
    If Not (else_type = "") Then
         opt_type_else.Caption = else_type
        'טעויות בתוים?
    End If
End Sub

Private Sub Option1_Click()
fra_type.BackColor = &H8000000F
opt_type_else.Caption = "אחר"
End Sub

Private Sub Option2_Click()
fra_type.BackColor = &H8000000F
opt_type_else.Caption = "אחר"
End Sub

Private Sub Option3_Click()
fra_type.BackColor = &H8000000F
opt_type_else.Caption = "אחר"
End Sub

Private Sub Option4_Click()
fra_type.BackColor = &H8000000F
opt_type_else.Caption = "אחר"
End Sub

Private Sub Option5_Click()
fra_type.BackColor = &H8000000F
opt_type_else.Caption = "אחר"
End Sub

Private Sub txt_put_comp_Change()
txt_put_comp.BackColor = &H80000005
End Sub

Private Sub txt_put_cost_KeyPress(KeyAscii As Integer)
'אחראי על כתיבת מחיר נכונה
Dim Char  As String
Static trml As Boolean 'נקודה נכתבה
Static ml As Integer 'כמה ספרות אחרי הנקודה מקסימום =2

txt_put_cost.BackColor = &H80000005
Char = Chr(KeyAscii)

If cmd_edit_ticket.Enabled = False Then 'יבדוק רק במקרה שזה לא מצב צפייה

 If (((Asc(Char)) > 57) Or ((Asc(Char)) < 48) And (Asc(Char) <> 8) And (Char <> ".")) Then 'לא יאפשר כתיבה של תוים שונים מספרה ומנקודה
    Beep
    KeyAscii = 0
    Exit Sub 'לא צריך לבדוק את כל ההמשך אם זה תו לא נכון
 End If 'I

If ml = 2 And Asc(Char) <> 8 Then 'אם יש בml שתי ספרות אז לא צריך יותר כלום אבל מאפשר למחוק
    KeyAscii = 0
    Exit Sub 'לא צריך יותר לבדוק
End If 'II

If ml = 2 And Asc(Char) = 8 Then 'אם אתה מוחק אחרי שמל=2זה אומר שאפשר לכתוב עוד ספרה
    ml = ml - 1
    Exit Sub 'לא צריך יותר לבדוק
End If 'III

If ml = 1 And Asc(Char) = 8 And trml Then 'אם אתה מוחק אחרי שמל=1זה אומר שאפשר לכתוב עוד שתי ספרות
    ml = ml - 1
    Exit Sub 'לא צריך יותר לבדוק
End If 'IV

If Asc(Char) = 8 And trml And ml = 0 Then 'מחקת את הנקודה
    trml = False
    Exit Sub 'לא צריך יותר לבדוק
End If 'V

If trml And Char = "." Then  'לא מאפשר לכתוב יותר מנקודה אחת
    KeyAscii = 0
End If 'VI

If trml And Char <> "." Then 'And Not nomore Then ' מחבר את הספרות שאחרי הנקודה
    ml = ml + 1
End If 'VII

If Char = "." Then 'ברגע שמופיעה נקודה רק מתו אח"כ צריך לספור לכן מופיע כתנאי אחרון
    trml = True
End If 'VIII

End If 'of cmd_edit_ticket.Enabled = False Then

End Sub

Private Sub txt_put_name_drug_Change()
'משנה את הכןתרת בהתאם לשם התרופה החדשה
lbl_caption_new_drug.Caption = txt_put_name_drug.Text
txt_put_name_drug.BackColor = &H80000005
End Sub
