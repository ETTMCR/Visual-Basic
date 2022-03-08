VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr_xx 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   2760
   End
   Begin VB.Timer tmr_yy 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Tag             =   "first get on"
      Top             =   2880
   End
   Begin VB.PictureBox lb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   2640
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   780
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   1920
      Width           =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xx, yy, move_x, move_y As Integer

Public Sub uping()
    lb.Top = lb.Top - 10
    If (lb.Top < yy) Then
    tmr_yy.Enabled = False
    tmr_xx.Enabled = True
End If
End Sub
Public Sub downing()
    lb.Top = lb.Top + 10
    If (lb.Top > yy) Then
    tmr_yy.Enabled = False
    tmr_xx.Enabled = True
    End If
End Sub
Public Sub righting()
    lb.Left = lb.Left + 10
    If (lb.Left > xx) Then
    tmr_xx.Enabled = False
    'Form1.MousePointer = 0
    End If
End Sub
Public Sub lefting()
    lb.Left = lb.Left - 10
    If (lb.Left < xx) Then
    tmr_xx.Enabled = False
    'Form1.MousePointer = 0
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xx = X
yy = Y
tmr_yy.Enabled = True
End Sub

Private Sub tmr_yy_Timer()
'Form1.MousePointer = 2
If lb.Top < yy Then
Call downing
Else
Call uping
End If
End Sub

Private Sub tmr_xx_Timer()
    If lb.Left < xx Then
    Call righting
    Else
    Call lefting
    End If
End Sub
