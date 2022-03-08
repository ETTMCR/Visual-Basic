VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Pentago"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "pentago.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_rtt_CW_DR 
      BackColor       =   &H00FF8080&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7395
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Rotate the quareter Clock Wise"
      Top             =   6195
      Width           =   315
   End
   Begin VB.CommandButton cmd_rtt_CCW_DR 
      BackColor       =   &H00FF8080&
      Caption         =   "/\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Rotate the quareter Counter Clock Wise"
      Top             =   5805
      Width           =   315
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   0
      Left            =   4980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3675
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   1
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3675
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   2
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3675
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   3
      Left            =   4965
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4545
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   4
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4515
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   5
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4515
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   6
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5430
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   7
      Left            =   5955
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   8
      Left            =   6915
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_rtt_CW_DL 
      BackColor       =   &H00FF8080&
      Caption         =   "/\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1335
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Rotate the quareter Clock Wise"
      Top             =   5730
      Width           =   315
   End
   Begin VB.CommandButton cmd_rtt_CCW_DL 
      BackColor       =   &H00FF8080&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1755
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Rotate the quareter Counter Clock Wise"
      Top             =   6165
      Width           =   315
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   8
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   7
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5385
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   6
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5385
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   5
      Left            =   3645
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4500
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   4
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4500
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   3
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4515
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   2
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3660
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   1
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3660
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   0
      Left            =   1755
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3660
      Width           =   795
   End
   Begin VB.CommandButton cmd_rtt_CW_UL 
      BackColor       =   &H00FF8080&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Rotate the quareter Clock Wise"
      Top             =   450
      Width           =   315
   End
   Begin VB.CommandButton cmd_rtt_CCW_UL 
      BackColor       =   &H00FF8080&
      Caption         =   "\/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Rotate the quareter Counter Clock Wise"
      Top             =   810
      Width           =   315
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   0
      Left            =   1710
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   825
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   825
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   2
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   825
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   3
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1695
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   4
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   5
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   6
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2550
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   7
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2550
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   8
      Left            =   3645
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2565
      Width           =   795
   End
   Begin VB.CommandButton cmd_rtt_CW_UR 
      BackColor       =   &H00FF8080&
      Caption         =   "\/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Rotate the quareter Clock Wise"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton cmd_rtt_CCW_UR 
      BackColor       =   &H00FF8080&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Rotate the quareter Counter Clock Wise"
      Top             =   465
      Width           =   315
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   8
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   7
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2565
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   6
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2595
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   5
      Left            =   6885
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1695
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   4
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1695
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   3
      Left            =   4965
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1695
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   2
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   1
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   0
      Left            =   4965
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   795
   End
   Begin VB.Label lbl_turn 
      Caption         =   "White Turn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4050
      TabIndex        =   44
      Top             =   180
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim which_player As Boolean
Dim mone_turns As Integer
Private Sub Form_Load()
which_player = True ' white starts
' which_player = True white  player
' which_player = false black  player

mone_turns = 0

For i = 0 To 8
    cmd_UR(i).BackColor = QBColor(7)
    cmd_UL(i).BackColor = QBColor(7)
    cmd_DL(i).BackColor = QBColor(7)
    cmd_DR(i).BackColor = QBColor(7)
Next

End Sub
Function func_mone_turns()
If mone_turns <> 2 Then
MsgBox ("put one or rotate !")
End If
End Function
Function check_victory()
' two functions checks one for rotate one for puting
' and in the mone victory and massaging
'!!!!!' only one when just marking ig from rotating and changing the clr and remeber what was

If which_player = True Then ' white player
clrchk = &HFFFFFF
Else
clrchk = &H80000012
End If
 
 
If ((cmd_DL(1).BackColor = clrchk) And (cmd_DL(2).BackColor = clrchk) And (cmd_DR(0).BackColor = clrchk) And (cmd_DR(1).BackColor = clrchk) And ((cmd_DL(0).BackColor = clrchk) Or (cmd_DR(2).BackColor = clrchk))) Or _
((cmd_DL(4).BackColor = clrchk) And (cmd_DL(5).BackColor = clrchk) And (cmd_DR(3).BackColor = clrchk) And (cmd_DR(4).BackColor = clrchk) And ((cmd_DL(3).BackColor = clrchk) Or (cmd_DR(5).BackColor = clrchk))) Or _
((cmd_DL(7).BackColor = clrchk) And (cmd_DL(8).BackColor = clrchk) And (cmd_DR(6).BackColor = clrchk) And (cmd_DR(7).BackColor = clrchk) And ((cmd_DL(6).BackColor = clrchk) Or (cmd_DR(8).BackColor = clrchk))) Or _
((cmd_UL(1).BackColor = clrchk) And (cmd_UL(2).BackColor = clrchk) And (cmd_UR(0).BackColor = clrchk) And (cmd_UR(1).BackColor = clrchk) And ((cmd_UL(0).BackColor = clrchk) Or (cmd_UR(2).BackColor = clrchk))) Or _
((cmd_UL(4).BackColor = clrchk) And (cmd_UL(5).BackColor = clrchk) And (cmd_UR(3).BackColor = clrchk) And (cmd_UR(4).BackColor = clrchk) And ((cmd_UL(3).BackColor = clrchk) Or (cmd_UR(5).BackColor = clrchk))) Or _
((cmd_UL(7).BackColor = clrchk) And (cmd_UL(8).BackColor = clrchk) And (cmd_UR(6).BackColor = clrchk) And (cmd_UR(7).BackColor = clrchk) And ((cmd_UL(6).BackColor = clrchk) Or (cmd_UR(8).BackColor = clrchk))) Or _
((cmd_UL(3).BackColor = clrchk) And (cmd_UL(6).BackColor = clrchk) And (cmd_DL(0).BackColor = clrchk) And (cmd_DL(3).BackColor = clrchk) And ((cmd_UL(0).BackColor = clrchk) Or (cmd_DL(6).BackColor = clrchk))) Or _
((cmd_UL(4).BackColor = clrchk) And (cmd_UL(7).BackColor = clrchk) And (cmd_DL(1).BackColor = clrchk) And (cmd_DL(4).BackColor = clrchk) And ((cmd_UL(1).BackColor = clrchk) Or (cmd_DL(7).BackColor = clrchk))) Or _
((cmd_UL(5).BackColor = clrchk) And (cmd_UL(8).BackColor = clrchk) And (cmd_DL(2).BackColor = clrchk) And (cmd_DL(5).BackColor = clrchk) And ((cmd_UL(2).BackColor = clrchk) Or (cmd_DL(8).BackColor = clrchk))) Or _
((cmd_UR(3).BackColor = clrchk) And (cmd_UR(6).BackColor = clrchk) And (cmd_DR(0).BackColor = clrchk) And (cmd_DR(3).BackColor = clrchk) And ((cmd_UR(0).BackColor = clrchk) Or (cmd_DR(6).BackColor = clrchk))) Or _
((cmd_UR(4).BackColor = clrchk) And (cmd_UR(7).BackColor = clrchk) And (cmd_DR(1).BackColor = clrchk) And (cmd_DR(4).BackColor = clrchk) And ((cmd_UR(1).BackColor = clrchk) Or (cmd_DR(7).BackColor = clrchk))) Or _
((cmd_UR(5).BackColor = clrchk) And (cmd_UR(8).BackColor = clrchk) And (cmd_DR(2).BackColor = clrchk) And (cmd_DR(5).BackColor = clrchk) And ((cmd_UR(2).BackColor = clrchk) Or (cmd_DR(8).BackColor = clrchk))) Or _
((cmd_DR(5).BackColor = clrchk) And (cmd_DR(1).BackColor = clrchk) And (cmd_UR(6).BackColor = clrchk) And (cmd_UL(5).BackColor = clrchk) And (cmd_UL(1).BackColor = clrchk)) Or _
((cmd_DL(3).BackColor = clrchk) And (cmd_DL(1).BackColor = clrchk) And (cmd_UL(8).BackColor = clrchk) And (cmd_UR(3).BackColor = clrchk) And (cmd_UR(1).BackColor = clrchk)) Or _
((cmd_DL(7).BackColor = clrchk) And (cmd_DL(5).BackColor = clrchk) And (cmd_DR(0).BackColor = clrchk) And (cmd_UR(7).BackColor = clrchk) And (cmd_UR(5).BackColor = clrchk)) Or _
((cmd_DR(7).BackColor = clrchk) And (cmd_DR(3).BackColor = clrchk) And (cmd_DL(2).BackColor = clrchk) And (cmd_UL(7).BackColor = clrchk) And (cmd_UL(3).BackColor = clrchk)) Or _
((cmd_DL(2).BackColor = clrchk) And (cmd_UR(4).BackColor = clrchk) And (cmd_UR(6).BackColor = clrchk) And (cmd_DL(4).BackColor = clrchk) And ((cmd_UR(2).BackColor = clrchk) Or (cmd_DL(6).BackColor = clrchk))) Or _
((cmd_UL(4).BackColor = clrchk) And (cmd_UL(8).BackColor = clrchk) And (cmd_DR(0).BackColor = clrchk) And (cmd_DR(4).BackColor = clrchk) And ((cmd_UL(0).BackColor = clrchk) Or (cmd_DR(8).BackColor = clrchk))) Then


    If clrchk = &HFFFFFF Then
        MsgBox ("The White Player is Victorious !")
    Else ' then check for the other color and boolean for only when checking on roteting
        MsgBox ("The Black Player is Victorious !")
    End If
End If

End Function

Private Sub cmd_UR_Click(Index As Integer)
'mone_turns = mone_turns + 1

'if mone_turns =

If which_player = True Then
    cmd_UR(Index).BackColor = &HFFFFFF 'QBColor(15) '&HFFFFFF    'white
    Call check_victory
    cmd_UR(Index).Enabled = False

'If mone_turns = 2 Then ' can be  func knows who by which_player
    'if which_player = true then
    '    which_player = False
    '    lbl_turn.Caption = ("Black Turn")
    'Else
    '    which_player = true
    '    lbl_turn.Caption = ("White Turn")
    'endif
'    mone_turns = 0
'Else
    'mone_turns = mone_turns + 1
'End If

Else
    cmd_UR(Index).BackColor = &H80000012 'QBColor(0) ' &H80000012  ' black
    Call check_victory
    which_player = True
    cmd_UR(Index).Enabled = False
     lbl_turn.Caption = ("White Turn")
End If

'Call func_mone_turns

End Sub
Private Sub cmd_UL_Click(Index As Integer)
If which_player = True Then
cmd_UL(Index).BackColor = &HFFFFFF 'QBColor(15) '&HFFFFFF    'white
Call check_victory
which_player = False
cmd_UL(Index).Enabled = False
lbl_turn.Caption = ("Black Turn")
Else
cmd_UL(Index).BackColor = &H80000012 'QBColor(0) ' &H80000012  ' black
Call check_victory
which_player = True
cmd_UL(Index).Enabled = False
lbl_turn.Caption = ("White Turn")
End If
End Sub
Private Sub cmd_DL_Click(Index As Integer)
If which_player = True Then
cmd_DL(Index).BackColor = &HFFFFFF 'QBColor(15) '&HFFFFFF    'white
Call check_victory
which_player = False
cmd_DL(Index).Enabled = False
lbl_turn.Caption = ("Black Turn")
Else
cmd_DL(Index).BackColor = &H80000012 'QBColor(0) ' &H80000012  ' black
Call check_victory
which_player = True
cmd_DL(Index).Enabled = False
lbl_turn.Caption = ("White Turn")
End If
End Sub
Private Sub cmd_DR_Click(Index As Integer)
If which_player = True Then
cmd_DR(Index).BackColor = &HFFFFFF 'QBColor(15) '&HFFFFFF    'white
Call check_victory
which_player = False
cmd_DR(Index).Enabled = False
lbl_turn.Caption = ("Black Turn")
Else
cmd_DR(Index).BackColor = &H80000012 'QBColor(0) ' &H80000012  ' black
Call check_victory
which_player = True
cmd_DR(Index).Enabled = False
lbl_turn.Caption = ("White Turn")
End If
End Sub

Private Sub cmd_rtt_CW_UR_Click()
Dim arr_rtt_CW_UR(8) As Long 'Integer 'Double

For i = 0 To 8
   arr_rtt_CW_UR(i) = (cmd_UR(i).BackColor) 'CInt(cmd_UR(i).BackColor)
Next

cmd_UR(0).BackColor = arr_rtt_CW_UR(6)
cmd_UR(1).BackColor = arr_rtt_CW_UR(3)
cmd_UR(2).BackColor = arr_rtt_CW_UR(0)
cmd_UR(3).BackColor = arr_rtt_CW_UR(7)
cmd_UR(4).BackColor = arr_rtt_CW_UR(4)
cmd_UR(5).BackColor = arr_rtt_CW_UR(1)
cmd_UR(6).BackColor = arr_rtt_CW_UR(8)
cmd_UR(7).BackColor = arr_rtt_CW_UR(5)
cmd_UR(8).BackColor = arr_rtt_CW_UR(2)

'cmd_UR(0).BackColor = arr_rtt_CW_UR(6) 'QBColor(arr_rtt_CW_UR(6))
'cmd_UR(1).BackColor = QBColor(arr_rtt_CW_UR(3))
'cmd_UR(2).BackColor = QBColor(arr_rtt_CW_UR(0))
'cmd_UR(3).BackColor = QBColor(arr_rtt_CW_UR(7))
'cmd_UR(4).BackColor = QBColor(arr_rtt_CW_UR(4))
'cmd_UR(5).BackColor = QBColor(arr_rtt_CW_UR(1))
'cmd_UR(6).BackColor = QBColor(arr_rtt_CW_UR(8))
'cmd_UR(7).BackColor = QBColor(arr_rtt_CW_UR(5))
'cmd_UR(8).BackColor = QBColor(arr_rtt_CW_UR(2))
''cmd_UR(0).BackColor = QBColor(arr_rtt_CW_UR(6))

For i = 0 To 8
    If cmd_UR(i).BackColor <> QBColor(7) Then
        cmd_UR(i).Enabled = False
    Else
        cmd_UR(i).Enabled = True
    End If
Next

Call check_victory


End Sub

Private Sub cmd_rtt_CCW_UR_Click()
Dim arr_rtt_CCW_UR(8) As Long 'Integer 'Double

'mone_turns = mone_turns + 1

For i = 0 To 8
   arr_rtt_CCW_UR(i) = (cmd_UR(i).BackColor) 'CInt(cmd_UR(i).BackColor)
Next

cmd_UR(6).BackColor = arr_rtt_CCW_UR(0)
cmd_UR(3).BackColor = arr_rtt_CCW_UR(1)
cmd_UR(0).BackColor = arr_rtt_CCW_UR(2)
cmd_UR(7).BackColor = arr_rtt_CCW_UR(3)
cmd_UR(4).BackColor = arr_rtt_CCW_UR(4)
cmd_UR(1).BackColor = arr_rtt_CCW_UR(5)
cmd_UR(8).BackColor = arr_rtt_CCW_UR(6)
cmd_UR(5).BackColor = arr_rtt_CCW_UR(7)
cmd_UR(2).BackColor = arr_rtt_CCW_UR(8)

'cmd_UR(0).BackColor = arr_rtt_CCW_UR(6) 'QBColor(arr_rtt_CCW_UR(6))
'cmd_UR(1).BackColor = QBColor(arr_rtt_CCW_UR(3))
'cmd_UR(2).BackColor = QBColor(arr_rtt_CCW_UR(0))
'cmd_UR(3).BackColor = QBColor(arr_rtt_CCW_UR(7))
'cmd_UR(4).BackColor = QBColor(arr_rtt_CCW_UR(4))
'cmd_UR(5).BackColor = QBColor(arr_rtt_CCW_UR(1))
'cmd_UR(6).BackColor = QBColor(arr_rtt_CCW_UR(8))
'cmd_UR(7).BackColor = QBColor(arr_rtt_CCW_UR(5))
'cmd_UR(8).BackColor = QBColor(arr_rtt_CCW_UR(2))
''cmd_UR(0).BackColor = QBColor(arr_rtt_CCW_UR(6))

For i = 0 To 8
    If cmd_UR(i).BackColor <> QBColor(7) Then
        cmd_UR(i).Enabled = False
    Else
        cmd_UR(i).Enabled = True
    End If
Next

'If mone_turns = 2 Then
'    which_player = False
'    lbl_turn.Caption = ("Black Turn")
'    mone_turns = 0
''Else
'End If

Call check_victory

End Sub

Private Sub cmd_rtt_CCW_UL_Click()
Dim arr_rtt_CCW_UL(8) As Long 'Integer 'Double

For i = 0 To 8
   arr_rtt_CCW_UL(i) = (cmd_UL(i).BackColor) 'CInt(cmd_UL(i).BackColor)
Next

cmd_UL(6).BackColor = arr_rtt_CCW_UL(0)
cmd_UL(3).BackColor = arr_rtt_CCW_UL(1)
cmd_UL(0).BackColor = arr_rtt_CCW_UL(2)
cmd_UL(7).BackColor = arr_rtt_CCW_UL(3)
cmd_UL(4).BackColor = arr_rtt_CCW_UL(4)
cmd_UL(1).BackColor = arr_rtt_CCW_UL(5)
cmd_UL(8).BackColor = arr_rtt_CCW_UL(6)
cmd_UL(5).BackColor = arr_rtt_CCW_UL(7)
cmd_UL(2).BackColor = arr_rtt_CCW_UL(8)

'cmd_UL(0).BackColor = arr_rtt_CCW_UL(6) 'QBColor(arr_rtt_CCW_UL(6))
'cmd_UL(1).BackColor = QBColor(arr_rtt_CCW_UL(3))
'cmd_UL(2).BackColor = QBColor(arr_rtt_CCW_UL(0))
'cmd_UL(3).BackColor = QBColor(arr_rtt_CCW_UL(7))
'cmd_UL(4).BackColor = QBColor(arr_rtt_CCW_UL(4))
'cmd_UL(5).BackColor = QBColor(arr_rtt_CCW_UL(1))
'cmd_UL(6).BackColor = QBColor(arr_rtt_CCW_UL(8))
'cmd_UL(7).BackColor = QBColor(arr_rtt_CCW_UL(5))
'cmd_UL(8).BackColor = QBColor(arr_rtt_CCW_UL(2))
''cmd_UL(0).BackColor = QBColor(arr_rtt_CCW_UL(6))


For i = 0 To 8
    If cmd_UL(i).BackColor <> QBColor(7) Then
        cmd_UL(i).Enabled = False
    Else
        cmd_UL(i).Enabled = True
    End If
Next

Call check_victory


End Sub

Private Sub cmd_rtt_CW_UL_Click()
Dim arr_rtt_CW_UL(8) As Long 'Integer 'Double

For i = 0 To 8
   arr_rtt_CW_UL(i) = (cmd_UL(i).BackColor) 'CInt(cmd_UL(i).BackColor)
Next

cmd_UL(0).BackColor = arr_rtt_CW_UL(6)
cmd_UL(1).BackColor = arr_rtt_CW_UL(3)
cmd_UL(2).BackColor = arr_rtt_CW_UL(0)
cmd_UL(3).BackColor = arr_rtt_CW_UL(7)
cmd_UL(4).BackColor = arr_rtt_CW_UL(4)
cmd_UL(5).BackColor = arr_rtt_CW_UL(1)
cmd_UL(6).BackColor = arr_rtt_CW_UL(8)
cmd_UL(7).BackColor = arr_rtt_CW_UL(5)
cmd_UL(8).BackColor = arr_rtt_CW_UL(2)

'cmd_UL(0).BackColor = arr_rtt_CW_UL(6) 'QBColor(arr_rtt_CW_UL(6))
'cmd_UL(1).BackColor = QBColor(arr_rtt_CW_UL(3))
'cmd_UL(2).BackColor = QBColor(arr_rtt_CW_UL(0))
'cmd_UL(3).BackColor = QBColor(arr_rtt_CW_UL(7))
'cmd_UL(4).BackColor = QBColor(arr_rtt_CW_UL(4))
'cmd_UL(5).BackColor = QBColor(arr_rtt_CW_UL(1))
'cmd_UL(6).BackColor = QBColor(arr_rtt_CW_UL(8))
'cmd_UL(7).BackColor = QBColor(arr_rtt_CW_UL(5))
'cmd_UL(8).BackColor = QBColor(arr_rtt_CW_UL(2))
''cmd_UL(0).BackColor = QBColor(arr_rtt_CW_UL(6))

For i = 0 To 8
    If cmd_UL(i).BackColor <> QBColor(7) Then
        cmd_UL(i).Enabled = False
    Else
        cmd_UL(i).Enabled = True
    End If
Next

Call check_victory


End Sub

Private Sub cmd_rtt_CW_DL_Click()
Dim arr_rtt_CW_DL(8) As Long 'Integer 'Double

For i = 0 To 8
   arr_rtt_CW_DL(i) = (cmd_DL(i).BackColor) 'CInt(cmd_DL(i).BackColor)
Next

cmd_DL(0).BackColor = arr_rtt_CW_DL(6)
cmd_DL(1).BackColor = arr_rtt_CW_DL(3)
cmd_DL(2).BackColor = arr_rtt_CW_DL(0)
cmd_DL(3).BackColor = arr_rtt_CW_DL(7)
cmd_DL(4).BackColor = arr_rtt_CW_DL(4)
cmd_DL(5).BackColor = arr_rtt_CW_DL(1)
cmd_DL(6).BackColor = arr_rtt_CW_DL(8)
cmd_DL(7).BackColor = arr_rtt_CW_DL(5)
cmd_DL(8).BackColor = arr_rtt_CW_DL(2)

'cmd_DL(0).BackColor = arr_rtt_CW_DL(6) 'QBColor(arr_rtt_CW_DL(6))
'cmd_DL(1).BackColor = QBColor(arr_rtt_CW_DL(3))
'cmd_DL(2).BackColor = QBColor(arr_rtt_CW_DL(0))
'cmd_DL(3).BackColor = QBColor(arr_rtt_CW_DL(7))
'cmd_DL(4).BackColor = QBColor(arr_rtt_CW_DL(4))
'cmd_DL(5).BackColor = QBColor(arr_rtt_CW_DL(1))
'cmd_DL(6).BackColor = QBColor(arr_rtt_CW_DL(8))
'cmd_DL(7).BackColor = QBColor(arr_rtt_CW_DL(5))
'cmd_DL(8).BackColor = QBColor(arr_rtt_CW_DL(2))
''cmd_DL(0).BackColor = QBColor(arr_rtt_CW_DL(6))

For i = 0 To 8
    If cmd_DL(i).BackColor <> QBColor(7) Then
        cmd_DL(i).Enabled = False
    Else
        cmd_DL(i).Enabled = True
    End If
Next

Call check_victory

End Sub


Private Sub cmd_rtt_CCW_DL_Click()
Dim arr_rtt_CCW_DL(8) As Long 'Integer 'Double

For i = 0 To 8
   arr_rtt_CCW_DL(i) = (cmd_DL(i).BackColor) 'CInt(cmd_DL(i).BackColor)
Next

cmd_DL(6).BackColor = arr_rtt_CCW_DL(0)
cmd_DL(3).BackColor = arr_rtt_CCW_DL(1)
cmd_DL(0).BackColor = arr_rtt_CCW_DL(2)
cmd_DL(7).BackColor = arr_rtt_CCW_DL(3)
cmd_DL(4).BackColor = arr_rtt_CCW_DL(4)
cmd_DL(1).BackColor = arr_rtt_CCW_DL(5)
cmd_DL(8).BackColor = arr_rtt_CCW_DL(6)
cmd_DL(5).BackColor = arr_rtt_CCW_DL(7)
cmd_DL(2).BackColor = arr_rtt_CCW_DL(8)

'cmd_DL(0).BackColor = arr_rtt_CCW_DL(6) 'QBColor(arr_rtt_CCW_DL(6))
'cmd_DL(1).BackColor = QBColor(arr_rtt_CCW_DL(3))
'cmd_DL(2).BackColor = QBColor(arr_rtt_CCW_DL(0))
'cmd_DL(3).BackColor = QBColor(arr_rtt_CCW_DL(7))
'cmd_DL(4).BackColor = QBColor(arr_rtt_CCW_DL(4))
'cmd_DL(5).BackColor = QBColor(arr_rtt_CCW_DL(1))
'cmd_DL(6).BackColor = QBColor(arr_rtt_CCW_DL(8))
'cmd_DL(7).BackColor = QBColor(arr_rtt_CCW_DL(5))
'cmd_DL(8).BackColor = QBColor(arr_rtt_CCW_DL(2))
''cmd_DL(0).BackColor = QBColor(arr_rtt_CCW_DL(6))

For i = 0 To 8
    If cmd_DL(i).BackColor <> QBColor(7) Then
        cmd_DL(i).Enabled = False
    Else
        cmd_DL(i).Enabled = True
    End If
Next

Call check_victory

End Sub

Private Sub cmd_rtt_CW_DR_Click()
Dim arr_rtt_CW_DR(8) As Long 'Integer 'Double

For i = 0 To 8
   arr_rtt_CW_DR(i) = (cmd_DR(i).BackColor) 'CInt(cmd_DR(i).BackColor)
Next

cmd_DR(0).BackColor = arr_rtt_CW_DR(6)
cmd_DR(1).BackColor = arr_rtt_CW_DR(3)
cmd_DR(2).BackColor = arr_rtt_CW_DR(0)
cmd_DR(3).BackColor = arr_rtt_CW_DR(7)
cmd_DR(4).BackColor = arr_rtt_CW_DR(4)
cmd_DR(5).BackColor = arr_rtt_CW_DR(1)
cmd_DR(6).BackColor = arr_rtt_CW_DR(8)
cmd_DR(7).BackColor = arr_rtt_CW_DR(5)
cmd_DR(8).BackColor = arr_rtt_CW_DR(2)

'cmd_DR(0).BackColor = arr_rtt_CW_DR(6) 'QBColor(arr_rtt_CW_DR(6))
'cmd_DR(1).BackColor = QBColor(arr_rtt_CW_DR(3))
'cmd_DR(2).BackColor = QBColor(arr_rtt_CW_DR(0))
'cmd_DR(3).BackColor = QBColor(arr_rtt_CW_DR(7))
'cmd_DR(4).BackColor = QBColor(arr_rtt_CW_DR(4))
'cmd_DR(5).BackColor = QBColor(arr_rtt_CW_DR(1))
'cmd_DR(6).BackColor = QBColor(arr_rtt_CW_DR(8))
'cmd_DR(7).BackColor = QBColor(arr_rtt_CW_DR(5))
'cmd_DR(8).BackColor = QBColor(arr_rtt_CW_DR(2))
''cmd_DR(0).BackColor = QBColor(arr_rtt_CW_DR(6))

For i = 0 To 8
    If cmd_DR(i).BackColor <> QBColor(7) Then
        cmd_DR(i).Enabled = False
    Else
        cmd_DR(i).Enabled = True
    End If
Next

Call check_victory

End Sub


Private Sub cmd_rtt_CCW_DR_Click()
Dim arr_rtt_CCW_DR(8) As Long 'Integer 'Double

For i = 0 To 8
   arr_rtt_CCW_DR(i) = (cmd_DR(i).BackColor) 'CInt(cmd_DR(i).BackColor)
Next

cmd_DR(6).BackColor = arr_rtt_CCW_DR(0)
cmd_DR(3).BackColor = arr_rtt_CCW_DR(1)
cmd_DR(0).BackColor = arr_rtt_CCW_DR(2)
cmd_DR(7).BackColor = arr_rtt_CCW_DR(3)
cmd_DR(4).BackColor = arr_rtt_CCW_DR(4)
cmd_DR(1).BackColor = arr_rtt_CCW_DR(5)
cmd_DR(8).BackColor = arr_rtt_CCW_DR(6)
cmd_DR(5).BackColor = arr_rtt_CCW_DR(7)
cmd_DR(2).BackColor = arr_rtt_CCW_DR(8)

'cmd_DR(0).BackColor = arr_rtt_CCW_DR(6) 'QBColor(arr_rtt_CCW_DR(6))
'cmd_DR(1).BackColor = QBColor(arr_rtt_CCW_DR(3))
'cmd_DR(2).BackColor = QBColor(arr_rtt_CCW_DR(0))
'cmd_DR(3).BackColor = QBColor(arr_rtt_CCW_DR(7))
'cmd_DR(4).BackColor = QBColor(arr_rtt_CCW_DR(4))
'cmd_DR(5).BackColor = QBColor(arr_rtt_CCW_DR(1))
'cmd_DR(6).BackColor = QBColor(arr_rtt_CCW_DR(8))
'cmd_DR(7).BackColor = QBColor(arr_rtt_CCW_DR(5))
'cmd_DR(8).BackColor = QBColor(arr_rtt_CCW_DR(2))
''cmd_DR(0).BackColor = QBColor(arr_rtt_CCW_DR(6))

For i = 0 To 8
    If cmd_DR(i).BackColor <> QBColor(7) Then
        cmd_DR(i).Enabled = False
    Else
        cmd_DR(i).Enabled = True
    End If
Next

Call check_victory

End Sub

