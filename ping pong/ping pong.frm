VERSION 5.00
Begin VB.Form frm_ping_pong 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "שששמממייייאחחחח"
   ClientHeight    =   6165
   ClientLeft      =   3990
   ClientTop       =   2070
   ClientWidth     =   7725
   Icon            =   "ping pong.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   10  'Up Arrow
   RightToLeft     =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ball 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3600
      Picture         =   "ping pong.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   480
   End
   Begin VB.Timer tmr_ball 
      Interval        =   100
      Left            =   3600
      Top             =   840
   End
   Begin VB.Timer tmr_txt2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   5520
   End
   Begin VB.Timer tmr_txt1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   5520
   End
   Begin VB.TextBox txt_player2 
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_player1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      X1              =   7320
      X2              =   7320
      Y1              =   2400
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      X1              =   360
      X2              =   360
      Y1              =   2400
      Y2              =   3360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "player1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "player2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Menu mnu_speed_ball 
      Caption         =   "מהירות כדור"
      Begin VB.Menu mnu_speed_ball_low 
         Caption         =   "נמוכה"
      End
      Begin VB.Menu mnu_speed_ball_med 
         Caption         =   "בינונית"
      End
      Begin VB.Menu mnu_speed_ball_hige 
         Caption         =   "מטורפת"
      End
   End
   Begin VB.Menu mnu_end 
      Caption         =   "יציאה"
   End
End
Attribute VB_Name = "frm_ping_pong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public goal_plr2, goal_plr1 As Integer 'מוני הגולים
Public name1, name2 As String
Public d_l, d_r, u_l, u_r As Boolean 'לאן הולך הכדור
'משום מה שהוא פוגע למטה מהקו הוא חוזר
Private Sub Form_Load()

If goal_plr2 = 10 Then
    MsgBox name2, , "ניצחון"
    End
End If
If goal_plr1 = 10 Then
    MsgBox name1, , "ניצחון"
    End
End If

If goal_plr1 = 0 And goal_plr2 = 0 Then
    name1 = InputBox("שחקן1 שמאלי הכנס שמך", , "שם")
    Label1.Caption = name1
    name2 = InputBox("שחקן2 ימני הכנס שמך", , "שם")
    Label2.Caption = name2
End If

MsgBox "! היכונו", , "ping pong"

Randomize 'activate random movamante
sos = Int(4 * Rnd) + 1
Select Case sos
Case 1: d_l = True
        d_r = False
        u_l = False
        u_r = False
Case 2: d_r = True
        d_l = False
        u_l = False
        u_r = False
Case 3: u_l = True
        d_l = False
        d_r = False
        u_r = False
Case 4: u_r = True
        d_l = False
        d_r = False
        u_l = False
End Select

ball.Left = 3600 'איפוס מקומות
ball.Top = 1320 'שחקן אחד יכול להגיע למטה מספיק מהר
Line1.Y1 = 2400
Line1.Y2 = 3360
Line2.Y1 = 2400
Line2.Y2 = 3360
'אם המסג פה אז כבר מתחיל לזוז
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y > Line2.Y1 Then
Line2.Y1 = Line2.Y1 + 30
Line2.Y2 = Line2.Y2 + 30

Else
Line2.Y1 = Line2.Y1 - 30
Line2.Y2 = Line2.Y2 - 30

End If
End Sub

Private Sub mnu_end_Click()
If goal_plr1 = goal_plr2 Then
    MsgBox "! תיקו", , "? למה באמצע"
Else
    If goal_plr1 < goal_plr2 Then
        MsgBox "! לשחקן 2", , "ניצחון"
    Else
        MsgBox "! לשחקן 1", , "ניצחון"
    End If
End If
End
End Sub

Private Sub mnu_speed_ball_hige_Click()
tmr_ball.Interval = 20
End Sub

Private Sub mnu_speed_ball_med_Click()
tmr_ball.Interval = 50
End Sub
Private Sub mnu_speed_ball_low_Click()
tmr_ball.Interval = 75
End Sub

Private Sub tmr_txt1_Timer()
tmr_txt2.Enabled = True
txt_player2.SetFocus
tmr_txt1.Enabled = False
End Sub

Private Sub tmr_txt2_Timer()
tmr_txt1.Enabled = True
txt_player1.SetFocus
tmr_txt2.Enabled = False
End Sub

Private Sub tmr_ball_Timer()
'תנועות הכדור
'כאשר הם נעים למטה לא צריך לוודא בתנאי נוסף מיקומם למטה או למעלה?

If u_l = True Then 'get up
         ball.Left = ball.Left - 50
       ball.Top = ball.Top - 50
       
       'line1.Y1 =.top
   If ((Line1.Y1 < ball.Top) And _
        ((Line1.Y1 + Line1.Y2) > (ball.Top + ball.Height))) Then
        If (((ball.Left) < (Line1.X1)) And _
            ((ball.Left + ball.Width) > (Line1.X1))) Or _
            (((ball.Left) < (Line1.X1 + Line1.BorderWidth)) And ((ball.Left) > (Line1.X1))) Then
            If (ball.Top < Line1.Y2) Then  'בשביל למטה יכול לעבור מתחתיו
            'bang line1
            u_r = True
            u_l = False
            Exit Sub
            End If
        End If
    End If
    
     
        If (ball.Left < 0) Then
            'משקוף שמאלי
            goal_plr2 = goal_plr2 + 1
            Label2.Caption = name2 & " " & goal_plr2
            Call Form_Load
        End If
    
       If ball.Top < 0 Then
       'תקרה
            d_l = True
             u_l = False
       End If
              
End If 'u_l = True

 If u_r = True Then
        
        ball.Left = ball.Left + 50
       ball.Top = ball.Top - 50
       
    If ((Line2.Y1 < ball.Top) And _
        ((Line2.Y1 + Line2.Y2) > (ball.Top + ball.Height))) Then
        If (((ball.Left) < (Line2.X1)) And _
            ((ball.Left + ball.Width) > (Line2.X1))) Or _
            (((ball.Left) < (Line2.X1 + Line2.BorderWidth)) And ((ball.Left) > (Line2.X1))) Then
            If ball.Top < Line2.Y2 Then  'בשביל למטה יכול לעבור מתחתיו
            'bang line2
            u_l = True
            u_r = False
            Exit Sub
            End If
        End If
    End If
    
        If (ball.Left > frm_ping_pong.Width - (ball.Width)) Then
        'משקוף ימני
            goal_plr1 = goal_plr1 + 1
            Label1.Caption = name1 & " " & goal_plr1
            Call Form_Load
        End If
        
       If (ball.Top) < 0 Then
       
       'תקרה
            d_r = True
             u_r = False
        End If
        
End If 'u_r = True

If d_l = True Then 'get down
 
         ball.Left = ball.Left - 50
       ball.Top = ball.Top + 50

   If ((Line1.Y1 < ball.Top) And _
        ((Line1.Y1 + Line1.Y2) > (ball.Top + ball.Height))) Then
        If (((ball.Left) < (Line1.X1)) And _
            ((ball.Left + ball.Width) > (Line1.X1))) Or _
            (((ball.Left) < (Line1.X1 + Line1.BorderWidth)) And ((ball.Left) > (Line1.X1))) Then
                    If ball.Top < Line1.Y2 Then  'בשביל למטה יכול לעבור מתחתיו

                    'bang line1
                    d_r = True
                    d_l = False
                    Exit Sub
                     End If
        End If
    End If

       If (ball.Left < 0) Then
            goal_plr2 = goal_plr2 + 1
            Label2.Caption = name2 & " " & goal_plr2
            Call Form_Load
            'משקוף שמאלי
       End If
       
        If (ball.Top) > ((frm_ping_pong.Height) - (ball.Height)) Then
        'פוגע למטה
            u_l = True
            d_l = False
        End If
        
End If 'd_l = True

If d_r Then

        ball.Left = ball.Left + 50
       ball.Top = ball.Top + 50

   If ((Line2.Y1 < ball.Top) And _
        ((Line2.Y1 + Line2.Y2) > (ball.Top + ball.Height))) Then
        If (((ball.Left) < (Line2.X1)) And _
            ((ball.Left + ball.Width) > (Line2.X1))) Or _
            (((ball.Left) < (Line2.X1 + Line2.BorderWidth)) And ((ball.Left) > (Line2.X1))) Then
                    If ball.Top < Line2.Y2 Then  'בשביל למטה יכול לעבור מתחתיו

                    'bang line2
                    d_l = True
                    d_r = False
                    Exit Sub
                    End If
        End If
    End If

         If (ball.Left > frm_ping_pong.Width - (ball.Width)) Then
            goal_plr1 = goal_plr1 + 1
            Label1.Caption = name1 & " " & goal_plr1
            Call Form_Load
            'משקוף ימני
         End If
         
        If (ball.Top) > ((frm_ping_pong.Height) - (ball.Height)) Then
        'פוגע למטה
            u_r = True
             d_r = False
        End If
        
End If 'd_r=true


' במקום מקרי קיצון שמתוארים ע"י שני תנאים 8 תנאים שונים
End Sub

Private Sub txt_player1_KeyPress(KeyAscii As Integer)

 If Line1.Y1 < 0 And Not Chr(KeyAscii) = "z" Then  'מגיע למעלה
Exit Sub
End If

If Line1.Y2 > frm_ping_pong.Height _
And Not Chr(KeyAscii) = "q" Then 'מגיע למטה
Exit Sub
End If

Select Case Chr(KeyAscii) 'נורא איטי לעומת המעוכבר
Case "q": Line1.Y1 = Line1.Y1 - 75 'up plr1
          Line1.Y2 = Line1.Y2 - 75
Case "z": Line1.Y1 = Line1.Y1 + 75 'down plr1
          Line1.Y2 = Line1.Y2 + 75
End Select

End Sub

Private Sub txt_player2_KeyPress(KeyAscii As Integer)
 'no use i dont know to do this interface of two text boxes

 If Line2.Y1 < 0 And Not Chr(KeyAscii) = "2" Then  'מגיע למעלה
Exit Sub
End If

If Line2.Y2 > frm_ping_pong.Height _
And Not Chr(KeyAscii) = "8" Then 'מגיע למטה
Exit Sub
End If

Select Case Chr(KeyAscii)
Case "8": Line2.Y1 = Line2.Y1 - 20 'up plr1
          Line2.Y2 = Line2.Y2 - 20
Case "2": Line2.Y1 = Line2.Y1 + 20 'down plr1
          Line2.Y2 = Line2.Y2 + 20
Case "=": End
End Select

End Sub

Private Sub ball_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MsgBox ("" & X & " " & Y & "")
 'from where get the punch
Exit Sub
 'בעצם אפשר למחוק לקראת יצאית המשחק '@@@@@@@

 If X < ball.Height / 2 Then 'get up     1600 1600 downest righti corner of the pic
    If Y < ball.Width / 2 Then
        
        d_r = True 'V
        u_r = False
        u_l = False
        d_l = False
    Else
        u_r = True 'V
        d_l = False
        d_r = False
        u_l = False
    End If
Else
   If Y < ball.Width / 2 Then 'get down    x>800
        d_l = True 'V
        u_r = False
        u_l = False
        d_r = False
    Else
        u_l = True
        d_r = False
        u_r = False
        d_l = False
    End If
End If

End Sub
