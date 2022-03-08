VERSION 5.00
Begin VB.Form frm_comps 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "מאגר החברות"
   ClientHeight    =   5025
   ClientLeft      =   5925
   ClientTop       =   2385
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_add_new_drug 
      Caption         =   "הו&ספת חברה חדשה"
      Height          =   495
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ListBox lst_comps 
      Height          =   3180
      Left            =   480
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "לצפייה/עריכה לחץ פעמיים"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lbl_caption_frm_drugs 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "מאגר החברות"
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
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frm_comps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_add_new_drug_Click()
new_comp = True
frm_new_comp.Show
Unload frm_comps
End Sub

Private Sub Form_Load()
Dim i As Integer
sDBloc = App.Path & "\db_prj.mdb"
Set db_prj = OpenDatabase(sDBloc)
Set rs_tbl_comps = db_prj.OpenRecordset("tbl_comps", dbOpenTable)
Set rs_dyn_comps = db_prj.OpenRecordset("tbl_comps", dbOpenDynaset)

rs_dyn_comps.MoveFirst
Do While Not rs_tbl_comps.EOF
    lst_comps.AddItem rs_tbl_comps.Fields("txt_put_name_comp")
    rs_tbl_comps.MoveNext
Loop

For i = 0 To lst_comps.ListCount - 2 'מוחק שמות חברות שנמצאות פעמיים ברשימה
   If lst_comps.List(i) = lst_comps.List(i + 1) Then
        rs_dyn_comps.MoveFirst
        Do While Not rs_dyn_comps.EOF
            'צריך להגיע לרשומה שבה יש שם זה
            If rs_dyn_comps.Fields("txt_put_name_comp") = lst_comps.List(i) Then
                 rs_dyn_comps.Delete
                Exit Do 'אבל מה קורה אם יש יותר מתי שעשיתי המון תרופות חדשות
            End If
            rs_dyn_comps.MoveNext
        Loop
        lst_comps.RemoveItem (i) 'lst_comps.List
   End If 'of lst_comps.List(i) = lst_comps.List(i + 1)
Next
'                rs_dyn_comps.Update '?????

If taking_name_comp_2_drug Or taking_name_comp_2_find_drug Then 'נלחץ בחירת שם חברה ממאגר
        lst_comps.ToolTipText = "להשמה לחץ פעמיים"
End If
End Sub

Private Sub lst_comps_DblClick()
'פותח את הכרטיס של החברה שעליה לחצן
If taking_name_comp_2_drug Then 'נלחץ בחירת שם חברה ממאגר
        frm_new_drug.txt_put_comp.Text = lst_comps.Text
        taking_name_comp_2_drug = False
Else
    If taking_name_comp_2_find_drug Then
         frm_drugs.lbl_find_comp.Caption = lst_comps.Text
         taking_name_comp_2_find_drug = False
    Else 'נבחרה חברה לצפייה בנתוניה
        new_comp = False
        go_comp_name = lst_comps.Text
        frm_new_comp.Show
    End If
End If
    Unload frm_comps
End Sub

