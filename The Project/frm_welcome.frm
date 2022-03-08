VERSION 5.00
Begin VB.Form frm_welcome 
   BackColor       =   &H8000000E&
   Caption         =   "תרופות ולקוחות"
   ClientHeight    =   5775
   ClientLeft      =   3195
   ClientTop       =   2670
   ClientWidth     =   8670
   Icon            =   "frm_welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8670
   Begin VB.CommandButton cmd_new_comp 
      BackColor       =   &H000080FF&
      Caption         =   "הוספת  חברה חדשה"
      Height          =   495
      Left            =   1560
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd_mgr_comps 
      BackColor       =   &H000080FF&
      Caption         =   "מאגר החברות"
      Height          =   495
      Left            =   1560
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6600
      Top             =   5160
   End
   Begin VB.CommandButton cmd_new_drug 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "הוספת תרופה חדשה"
      DisabledPicture =   "frm_welcome.frx":08CA
      Height          =   495
      Left            =   3240
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd_mgr_ti_pa 
      BackColor       =   &H000080FF&
      Caption         =   "מאגר הלקוחות"
      Height          =   495
      Left            =   4920
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd_mgr_drugs 
      BackColor       =   &H000080FF&
      Caption         =   "מאגר התרופות"
      Height          =   495
      Left            =   3240
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd_new_ti_pa 
      BackColor       =   &H000080FF&
      Caption         =   "הוספת  לקוח חדש"
      Height          =   495
      Left            =   4920
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lbl_time 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lbl_date 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_caption_frm_walcome 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "תרופות ולקוחות"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   39
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   960
      TabIndex        =   3
      Top             =   -120
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   0
      Picture         =   "frm_welcome.frx":9F84
      Stretch         =   -1  'True
      Top             =   600
      Width           =   8655
   End
   Begin VB.Menu mnu_mgr 
      Caption         =   "מאגרים"
      Begin VB.Menu mnu_mgr_drugs 
         Caption         =   "מאגר התרופות"
      End
      Begin VB.Menu mnu_mgr_pts 
         Caption         =   "מאגר הלקוחות"
      End
      Begin VB.Menu mnu_mgr_comps 
         Caption         =   "מאגר החברות"
      End
   End
   Begin VB.Menu mnu_news 
      Caption         =   "הוספות"
      Begin VB.Menu mnu_news_drug 
         Caption         =   "הוספת תרופה חדשה"
      End
      Begin VB.Menu mnu_news_ti_pa 
         Caption         =   "הוספת לקוח חדש"
      End
      Begin VB.Menu mnu_news_comp 
         Caption         =   "הוספת חברה חדשה"
      End
   End
   Begin VB.Menu mnu_about 
      Caption         =   "&אודות"
      NegotiatePosition=   3  'Right
   End
   Begin VB.Menu mnu_exit 
      Caption         =   "י&ציאה"
      NegotiatePosition=   2  'Middle
   End
End
Attribute VB_Name = "frm_welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_mgr_drugs_Click()
'frm_welcome.Hide
taking_name_drug_2_lst_ti_pa = False 'מבטל אפשרות לחיצה פעמיים ושיתווסף השם לרשימה של התרופות של הלקוח בטופס כרטיס לקוח
frm_drugs.Show
End Sub

Private Sub cmd_new_comp_Click()
new_comp = True
frm_new_comp.Show
End Sub

Private Sub cmd_new_drug_Click()
'unload frm_welcome.Hide
new_drug = True
frm_new_drug.Show
End Sub

Private Sub cmd_new_ti_pa_Click()
'frm_welcome.Hide
new_ti_pa = True
frm_ticket_paitiont.Show
End Sub

Private Sub cmd_mgr_ti_pa_Click()
'frm_welcome.Hide
frm_paitionts.Show
End Sub

Private Sub cmd_mgr_comps_Click()
taking_name_comp_2_drug = False
frm_comps.Show
End Sub

Private Sub Form_Load()
'לזכור להזכיר את העניין של מקס אורך בתיבות הטקסט למניעת טעות
End Sub

Private Sub mnu_about_Click()
'מציג חלון אודות
MsgBox "The father above look at u :-)", , "Beware !"
End Sub

Private Sub mnu_exit_Click()
'יוצא מן התוכנית
End
End Sub

Private Sub mnu_mgr_comps_Click()
frm_comps.Show
End Sub

Private Sub mnu_mgr_drugs_Click()
taking_name_drug_2_lst_ti_pa = False 'מבטל אפשרות לחיצה פעמיים ושיתווסף השם לרשימה של התרופות של הלקוח בטופס כרטיס לקוח
frm_drugs.Show
End Sub

Private Sub mnu_mgr_pts_Click()
frm_paitionts.Show
End Sub

Private Sub mnu_news_comp_Click()
new_comp = True
frm_new_comp.Show
End Sub

Private Sub mnu_news_drug_Click()
new_drug = True
frm_new_drug.Show
End Sub

Private Sub mnu_news_ti_pa_Click()
new_ti_pa = True
frm_ticket_paitiont.Show
End Sub

Private Sub Timer1_Timer()
'מציג שעה ותאריך
lbl_date.Caption = Date
lbl_time.Caption = Time
If Time = "00:00:00" Then
    MsgBox ":-)"
End If
End Sub
