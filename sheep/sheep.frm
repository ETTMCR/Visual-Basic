VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   825
   ClientLeft      =   6150
   ClientTop       =   3555
   ClientWidth     =   825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   825
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   240
   End
   Begin VB.Image mg 
      Height          =   825
      Left            =   0
      Picture         =   "sheep.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form1_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii

Select Case KeyAscii
'Case 56: mg = LoadPicture(App.Path & "\pr ().JPG") 'up 8
Case 50: mg = LoadPicture(App.Path & "\pr (4).JPG") 'dn 2
        Form1.Top = Form1.Top + 50
Case 52: mg = LoadPicture(App.Path & "\pr (1).JPG") 'left 4
        Form1.Left = Form1.Left - 50
Case 54: mg = LoadPicture(App.Path & "\pr.JPG") 'rit 6
        Form1.Left = Form1.Left + 50
End Select

End Sub

Private Sub tmr_Timer() 'crazing around !
Randomize
rndy = Int(Rnd * 4)
Select Case rndy
'Case 0: mg = LoadPicture(App.Path & "\pr ().JPG") 'up 8
Case 1: mg = LoadPicture(App.Path & "\pr (4).JPG") 'dn 2
        Form1.Top = Form1.Top + 50
Case 2: mg = LoadPicture(App.Path & "\pr (1).JPG") 'left 4
        Form1.Left = Form1.Left - 50
Case 3: mg = LoadPicture(App.Path & "\pr.JPG") 'rit 6
        Form1.Left = Form1.Left + 50
End Select
End Sub
