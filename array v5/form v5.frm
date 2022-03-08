VERSION 5.00
Begin VB.Form frm_game 
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Bang The World V4"
   ClientHeight    =   5775
   ClientLeft      =   2730
   ClientTop       =   2955
   ClientWidth     =   8085
   Icon            =   "form v5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "form v5.frx":0ECA
   ScaleHeight     =   5775
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic_bull_mon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      CausesValidation=   0   'False
      Height          =   360
      Index           =   3
      Left            =   4440
      Picture         =   "form v5.frx":169D
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   420
   End
   Begin VB.PictureBox pic_bull_mon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      CausesValidation=   0   'False
      Height          =   360
      Index           =   4
      Left            =   5520
      Picture         =   "form v5.frx":1C7F
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   420
   End
   Begin VB.PictureBox pic_bull_mon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   5520
      Picture         =   "form v5.frx":2261
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   12
      Top             =   4680
      Width           =   360
   End
   Begin VB.PictureBox pic_bull_mon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   5520
      Picture         =   "form v5.frx":2843
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   11
      Top             =   4200
      Width           =   360
   End
   Begin VB.PictureBox pic_al 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   6600
      Picture         =   "form v5.frx":2E25
      ScaleHeight     =   600
      ScaleWidth      =   645
      TabIndex        =   10
      Top             =   4680
      Width           =   645
   End
   Begin VB.PictureBox pic_al 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   6600
      Picture         =   "form v5.frx":4307
      ScaleHeight     =   600
      ScaleWidth      =   645
      TabIndex        =   9
      Top             =   3600
      Width           =   645
   End
   Begin VB.PictureBox pic_al 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   6600
      Picture         =   "form v5.frx":57E9
      ScaleHeight     =   600
      ScaleWidth      =   645
      TabIndex        =   8
      Top             =   2400
      Width           =   645
   End
   Begin VB.PictureBox pic_al 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   6600
      Picture         =   "form v5.frx":6CCB
      ScaleHeight     =   600
      ScaleWidth      =   645
      TabIndex        =   7
      Top             =   1200
      Width           =   645
   End
   Begin VB.PictureBox pic_al 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   6600
      Picture         =   "form v5.frx":81AD
      ScaleHeight     =   600
      ScaleWidth      =   645
      TabIndex        =   6
      Top             =   120
      Width           =   645
   End
   Begin VB.PictureBox pic_bull_mon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   5520
      Picture         =   "form v5.frx":968F
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   3600
      Width           =   360
   End
   Begin VB.PictureBox pic_bull_man 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      Picture         =   "form v5.frx":9C71
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer tmr_bull_mon 
      Interval        =   15
      Left            =   5880
      Top             =   3600
   End
   Begin VB.Timer tmr_mon 
      Interval        =   75
      Left            =   7320
      Top             =   2280
   End
   Begin VB.Timer tmr_bull_man 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1800
      Top             =   1440
   End
   Begin VB.PictureBox pic_man 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   360
      Picture         =   "form v5.frx":9FF3
      ScaleHeight     =   1920
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin VB.PictureBox pic_ammo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      Picture         =   "form v5.frx":AB98
      ScaleHeight     =   570
      ScaleWidth      =   525
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pic_ammo_use 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   120
      Picture         =   "form v5.frx":BBE2
      ScaleHeight     =   570
      ScaleWidth      =   525
      TabIndex        =   5
      Top             =   4800
      Width           =   525
   End
   Begin VB.Menu mnu_abt 
      Caption         =   "about"
   End
End
Attribute VB_Name = "frm_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mone_clicks, stuck, gotop, dead_al_i, which_al_i, shot_bl_mon_i, pic_bull_mon_i, how_much_mon As Integer
Dim shot_bl_man, shot_bl_mon(5), dead_al(5) As Boolean
Option Explicit

Private Sub frm_game_Load()
'?=print
End Sub

Private Sub Form_Load()
For dead_al_i = 0 To 4
dead_al(dead_al_i) = False
Next

shot_bl_man = True 'can fire or not
'need to do that more then one mon fire can by doing two pics of bull or array
shot_bl_mon(0) = True 'can fire or not
shot_bl_mon(1) = True 'can fire or not
how_much_mon = 2
'how_much_mon -1 = InputBox("how much monsters u want?")
stuck = 6
dead_al_i = 0
End Sub

Private Sub mnu_abt_Click()
    MsgBox " GF=GF or not ?", , "<($_$)>"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
'MsgBox (KeyAscii)
Call moving_man(KeyAscii, pic_man) ' where to move

mone_clicks = mone_clicks + 1 'how many clicks u pressed
tmr_bull_man.Enabled = False
pic_bull_man.Enabled = True


If shot_bl_man And (KeyAscii = 53) And (stuck > 0) Then 'bull shown
    pic_bull_man.Visible = True
    pic_bull_man.Top = (pic_man.Top + pic_man.Height / 2)
    pic_bull_man.Left = pic_man.Left + pic_man.Width
    tmr_bull_man.Enabled = True
    stuck = stuck - 1
End If

If stuck = 0 Then 'no ammo
    pic_ammo.Visible = True
     pic_ammo.Left = 4010
    Beep
End If

If Not (shot_bl_man) Then
    tmr_bull_man.Enabled = True
End If


If pic_man.Left = pic_ammo.Left Then 'new ammo
    pic_ammo.Visible = False
    pic_ammo.Left = frm_game.Left
    stuck = 6
End If

Select Case stuck 'the stuck view
Case 0: pic_ammo_use.Picture = LoadPicture("C:\data\eliran\My VB\sooting games\pics\ammo0.bmp")
Case 1: pic_ammo_use.Picture = LoadPicture("C:\data\eliran\My VB\sooting games\pics\ammo1.bmp")
Case 2: pic_ammo_use.Picture = LoadPicture("C:\data\eliran\My VB\sooting games\pics\ammo2.bmp")
Case 3: pic_ammo_use.Picture = LoadPicture("C:\data\eliran\My VB\sooting games\pics\ammo3.bmp")
Case 4: pic_ammo_use.Picture = LoadPicture("C:\data\eliran\My VB\sooting games\pics\ammo4.bmp")
Case 5: pic_ammo_use.Picture = LoadPicture("C:\data\eliran\My VB\sooting games\pics\ammo5.bmp")
Case 6: pic_ammo_use.Picture = LoadPicture("C:\data\eliran\My VB\sooting games\pics\ammo.bmp")
End Select

End Sub

Private Sub tmr_bull_man_Timer() 'man bullet
Dim i As Integer
    shot_bl_man = False
   'Text1.Enabled = False '*
   
    pic_bull_man.Left = pic_bull_man.Left + 50 'moving the bull
    For i = 0 To 4
    If (((pic_bull_man.Left + pic_bull_man.Width) > pic_al(i).Left) And ((pic_bull_man.Left) < (pic_al(i).Left + pic_al(i).Width))) And _
        ((pic_bull_man.Top > pic_al(i).Top) And ((pic_bull_man.Top) < (pic_al(i).Top + pic_al(i).Height))) Then
        'MsgBox "u clicked " & mone_clicks & " times. win !", , Text1.Text
        pic_al(i).Enabled = False
        pic_al(i).Visible = False
        dead_al_i = i
        dead_al(dead_al_i) = True
        'End ' mon banged only in the middle
    End If
    Next
    
    If (pic_bull_man.Left > frm_game.Width) Then 'the bull gets to the end
        tmr_bull_man.Enabled = False
        pic_bull_man.Enabled = False
        shot_bl_man = True
        Text1.Enabled = True
        Text1.SetFocus
       
    End If
    

End Sub

Private Sub tmr_mon_Timer()
Text1.SetFocus
Dim i_top, j As Integer
' another time

    If gotop > 50 Then
    For i_top = 0 To 4
        pic_al(i_top).Top = pic_al(i_top).Top + 10

    Next
    End If
    
    If gotop < 50 Then
     For i_top = 0 To 4
        pic_al(i_top).Top = pic_al(i_top).Top - 10

    Next
    End If

   gotop = gotop + 1 'switching the up down
    If gotop = 101 Then
    gotop = 0
    End If
   
For j = 0 To how_much_mon
pic_bull_mon_i = j
Call which_al_fire(which_al_i, pic_al, shot_bl_mon, dead_al, dead_al_i, tmr_bull_mon, pic_bull_mon, shot_bl_mon_i, pic_bull_mon_i)
Next
 
End Sub

Private Sub tmr_bull_mon_Timer() 'monster bullet
Dim j As Integer
 For j = 0 To how_much_mon
pic_bull_mon_i = j
Call bull_mon_hit(pic_man, pic_bull_mon, tmr_bull_mon, shot_bl_mon, mone_clicks, Text1, which_al_i, shot_bl_mon_i, pic_bull_mon_i) ', pic_bull_man, shot_bl_man)
Next
End Sub

