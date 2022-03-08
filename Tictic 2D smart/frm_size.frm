VERSION 5.00
Begin VB.Form frm_size 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   7185
   ClientTop       =   4695
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox cmb_player 
      Height          =   315
      ItemData        =   "frm_size.frx":0000
      Left            =   1200
      List            =   "frm_size.frx":0007
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmb_columns 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cmb_rows 
      Height          =   315
      ItemData        =   "frm_size.frx":0011
      Left            =   2760
      List            =   "frm_size.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Who's start ?"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "How many rows do you want ?"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "How many columne do you want ?"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "frm_size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unlock_btn_OK As Integer

Private Sub cmb_columns_Click()

num_columns = cmb_columns.ListIndex + 3

unlock_btn_OK = unlock_btn_OK + 1
If unlock_btn_OK = 2 Then
cmd_OK.Enabled = True
End If
End Sub

Private Sub cmb_rows_Click()
num_rows = cmb_rows.ListIndex + 2

unlock_btn_OK = unlock_btn_OK + 1
If unlock_btn_OK = 2 Then
cmd_OK.Enabled = True
End If

End Sub

Private Sub cmd_OK_Click()
Unload Me
frm_tictic.Visible = True
End Sub

Private Sub Form_Load()
unlock_btn_OK = 0
For i = 2 To 10
cmb_rows.AddItem i
Next
For i = 3 To 10
cmb_columns.AddItem i
Next
End Sub

