VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "צחוק חזק, קשה לי לנשום "
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   Icon            =   "צחוק.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   2280
   End
   Begin VB.Label line1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ח"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()

Load line1(line1.Count)

  line1(line1.Count - 1).Left = Form1.Width - (Rnd * (Form1.Width))
  line1(line1.Count - 1).Top = Form1.Height - (Rnd * (Form1.Height))
  line1(line1.Count - 1).Visible = True
  line1(line1.Count - 1).FontSize = Rnd * 36
  line1(line1.Count - 1).Caption = "ח"
  line1(line1.Count - 1).ToolTipText = "איזה צחוקים איתך"
  
  If (line1.Count Mod 3 = 0) Then
  Form1.Caption = " צחוק חזק, קשה לי לנשום " & " ;-()"
  Else
    Form1.Caption = " צחוק חזק, קשה לי לנשום " & " :-( )"
  End If

End Sub
