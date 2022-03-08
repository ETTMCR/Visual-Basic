VERSION 5.00
Begin VB.Form frm_new_comp 
   BackColor       =   &H80000013&
   Caption         =   "כרטיס חברה "
   ClientHeight    =   3960
   ClientLeft      =   3900
   ClientTop       =   2235
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   6135
   Begin VB.CommandButton cmd_new_drug 
      Caption         =   "הוספת תרופה חדשה לחברה נוכחית"
      Height          =   615
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox lst_drug_of_comp 
      Height          =   2400
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "לצפייה לחץ פעמיים"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&אישור חברה"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txt_put_name_comp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lbl_frm_n_comp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "חברה"
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
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   4680
      Picture         =   "frm_new_comp.frx":0000
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label lbl_list_drug_of_comp 
      Alignment       =   1  'Right Justify
      Caption         =   "רשימת תרופות החברה"
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
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbl_put_name_comp 
      Alignment       =   1  'Right Justify
      Caption         =   "שם חברה"
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
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frm_new_comp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public last_name As String 'שומר שם חברה באם לפני שינוי
Private Sub cmd_new_drug_Click()
'unload frm_welcome.Hide
go_comp_name = txt_put_name_comp.Text
taking_comp = True
new_drug = True
Unload frm_new_comp
frm_new_drug.Show
End Sub


Private Sub cmd_ok_Click()
    rs_dyn_comps.MoveFirst
    Do While Not rs_dyn_comps.EOF 'Not rs_dyn_comps.Fields("txt_put_name_comp") = last_name
        If rs_dyn_comps.Fields("txt_put_name_comp") = last_name Then
        rs_dyn_comps.Delete
        Exit Do
        End If
        rs_dyn_comps.MoveNext
    Loop
    rs_dyn_comps.AddNew
    rs_dyn_comps.Fields("txt_put_name_comp") = txt_put_name_comp.Text
    rs_dyn_comps.Update
    
    rs_dyn_n_d.MoveFirst
    Do While Not rs_dyn_n_d.EOF
        If rs_dyn_n_d.Fields("txt_put_comp") = last_name Then
            rs_dyn_n_d.Edit
            rs_dyn_n_d.Fields("txt_put_comp") = txt_put_name_comp.Text
            rs_dyn_n_d.Update
        End If
        rs_dyn_n_d.MoveNext
    Loop
    Unload frm_new_comp
    frm_comps.Show
End Sub

Private Sub Form_Load()
sDBloc = App.Path & "\db_prj.mdb"
Set db_prj = OpenDatabase(sDBloc)
Set rs_tbl_comps = db_prj.OpenRecordset("tbl_comps", dbOpenTable)
Set rs_dyn_comps = db_prj.OpenRecordset("tbl_comps", dbOpenDynaset)
Set rs_tbl_n_d = db_prj.OpenRecordset("tbl_n_d", dbOpenTable)
Set rs_dyn_n_d = db_prj.OpenRecordset("tbl_n_d", dbOpenDynaset)

If new_comp = False Then
    rs_dyn_comps.Edit
    rs_dyn_n_d.MoveFirst
    Do While Not rs_dyn_n_d.EOF 'התרופות של אותה חברה
        If rs_dyn_n_d.Fields("txt_put_comp") = go_comp_name Then
            lst_drug_of_comp.AddItem rs_dyn_n_d.Fields("txt_put_name_drug")
            txt_put_name_comp.Text = go_comp_name 'rs_dyn_comps.Fields("txt_put_name_comp")
        End If
        rs_dyn_n_d.MoveNext
    Loop
    txt_put_name_comp.Text = go_comp_name 'rs_dyn_comps.Fields("txt_put_name_comp")
Else 'new_comp = true
    rs_dyn_comps.AddNew
End If 'of new_comp
last_name = txt_put_name_comp.Text
End Sub

Private Sub txt_put_name_comp_change()
lbl_frm_n_comp.Caption = txt_put_name_comp.Text
If txt_put_name_comp.Text = "" Then ' אם לא נכתב שם
    cmd_ok.Enabled = False
Else
    cmd_ok.Enabled = True
End If
End Sub

Private Sub lst_drug_of_comp_DblClick()
    new_drug = False
    go_drug_name = lst_drug_of_comp.Text
    frm_new_drug.Show
    Unload frm_comps
End Sub
