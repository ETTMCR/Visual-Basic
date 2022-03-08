VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "איך נזכור את כולם"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   1320
   End
   Begin VB.Label line1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Long

Private Sub Timer1_Timer()
num = (num + 1)

Load line1(line1.Count)

  line1(line1.Count - 1).Left = Form1.Width - (Rnd * (Form1.Width))
  line1(line1.Count - 1).Top = Form1.Height - (Rnd * (Form1.Height))
  line1(line1.Count - 1).Visible = True
 ' line1(line1.Count - 1).ForeColor
  line1(line1.Count - 1).Caption = num '* 100
  Form1.Caption = " איך נזכור את כולם " & num '* 100
'  line1(line1.Count - 1).
  If num = 6000000 Then
  InputBox ("")
  End If
End Sub
