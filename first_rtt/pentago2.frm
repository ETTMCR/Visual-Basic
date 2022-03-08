VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Pentago"
   ClientHeight    =   6915
   ClientLeft      =   3420
   ClientTop       =   2190
   ClientWidth     =   8205
   Icon            =   "pentago2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   8205
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4335
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   195
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDlg 
      Left            =   105
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   8
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2670
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   7
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2655
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   6
      Left            =   4425
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2685
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   5
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1785
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   4
      Left            =   5355
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1785
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   3
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1785
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   2
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   930
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   1
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   930
      Width           =   795
   End
   Begin VB.CommandButton cmd_UR 
      Height          =   690
      Index           =   0
      Left            =   4410
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   930
      Width           =   795
   End
   Begin VB.CommandButton cmd_rtt_CW_DR 
      BackColor       =   &H00FF0000&
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
      Left            =   6810
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Rotate the quareter Clock Wise"
      Top             =   6315
      Width           =   315
   End
   Begin VB.CommandButton cmd_rtt_CCW_DR 
      BackColor       =   &H00FF0000&
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
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Rotate the quareter Counter Clock Wise"
      Top             =   5925
      Width           =   315
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   0
      Left            =   4395
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3795
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   1
      Left            =   5325
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3795
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   2
      Left            =   6315
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3795
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   3
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4665
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   4
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4635
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   5
      Left            =   6315
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4635
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   6
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5550
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   7
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5535
      Width           =   795
   End
   Begin VB.CommandButton cmd_DR 
      Height          =   690
      Index           =   8
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_rtt_CW_DL 
      BackColor       =   &H00FF0000&
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
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Rotate the quareter Clock Wise"
      Top             =   5850
      Width           =   315
   End
   Begin VB.CommandButton cmd_rtt_CCW_DL 
      BackColor       =   &H00FF0000&
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
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Rotate the quareter Counter Clock Wise"
      Top             =   6285
      Width           =   315
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   8
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   7
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5505
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   6
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5505
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   5
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4620
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   4
      Left            =   2115
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4620
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   3
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4635
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   2
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3750
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   1
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3780
      Width           =   795
   End
   Begin VB.CommandButton cmd_DL 
      Height          =   690
      Index           =   0
      Left            =   1170
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3780
      Width           =   795
   End
   Begin VB.CommandButton cmd_rtt_CW_UL 
      BackColor       =   &H00FF0000&
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
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Rotate the quareter Clock Wise"
      Top             =   570
      Width           =   315
   End
   Begin VB.CommandButton cmd_rtt_CCW_UL 
      BackColor       =   &H00FF0000&
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
      Left            =   645
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Rotate the quareter Counter Clock Wise"
      Top             =   930
      Width           =   315
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   0
      Left            =   1125
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   945
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   1
      Left            =   2055
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   945
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   2
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   945
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   3
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1815
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   4
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   5
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   6
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2670
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   7
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2670
      Width           =   795
   End
   Begin VB.CommandButton cmd_UL 
      Height          =   690
      Index           =   8
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2685
      Width           =   795
   End
   Begin VB.CommandButton cmd_rtt_CW_UR 
      BackColor       =   &H00FF0000&
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
      Left            =   7215
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Rotate the quareter Clock Wise"
      Top             =   960
      Width           =   315
   End
   Begin VB.CommandButton cmd_rtt_CCW_UR 
      BackColor       =   &H00FF0000&
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
      Left            =   6765
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Rotate the quareter Counter Clock Wise"
      Top             =   585
      Width           =   315
   End
   Begin VB.Label lbl_turn 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "White Turn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2655
      TabIndex        =   44
      Top             =   285
      Width           =   1395
   End
   Begin VB.Menu cmd_file 
      Caption         =   "&FILE"
      Begin VB.Menu cmd_Save 
         Caption         =   "&SAVE"
         Shortcut        =   ^S
      End
      Begin VB.Menu cmd_load 
         Caption         =   "&LOAD"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu cmd_about 
      Caption         =   "&about"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim which_player, mone_recorsion, white_win_bool, black_win_bool As Boolean
Dim mone_which_player  As Integer
Dim clrchk As Double
Dim arr_btns(3, 8), Reply_again_game As Byte


Private Sub Form_Load()
which_player = True ' white starts

mone_which_player = 0

For i = 0 To 8
    cmd_UR(i).BackColor = QBColor(7)
    cmd_UL(i).BackColor = QBColor(7)
    cmd_DL(i).BackColor = QBColor(7)
    cmd_DR(i).BackColor = QBColor(7)
Next

'      For j = 0 To 3
'          For h = 0 To 8
'               arr_btns(j, h) = 2
'          Next
'     Next
     
cmd_rtt_CCW_UR.Enabled = True
Call dis_rtt ' will lock the rtt buttons
clrchk = &HFFFFFF      ' white
'lbl_turn.Caption = ("White Turn, Put button")
 'lbl_turn.Left = (Me.Width / 2) - (lbl_turn.Width)     ' אמצע חלון
 
' MsgBox (cmd_UL(2).BackColor)
'Form1.BackColor = &H808080
End Sub


Function switching_player()

If mone_which_player Mod 2 = 1 Then 'אי  זוגי ז"א שחור
     which_player = False
     clrchk = &H80000012
     lbl_turn.Caption = ("Black Turn, Put button")
Else
    which_player = True
    clrchk = 16777215
     lbl_turn.Caption = ("White Turn, Put button")
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
 
If white_win_bool = True Or black_win_bool = True Then ' התנאי  יכול יהיה בתוקף רק אחרי  יש נצחון של אחרי ריטוט , בתוך הרקורסיה
Else  ' ז"א אין נצחון
     white_win_bool = False
     black_win_bool = False
End If ' nor

 
 
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

'then ' יש נצחון
' !!!!!!!!!!!!!!!!!!!!!! win
    If clrchk = &HFFFFFF Then ' נכנס לפה רק אם יש נצחון כלשהו
         white_win_bool = True
        MsgBox ("The White Player is Victorious !")
'        GoTo again
    Else ' then check for the other color and boolean for only when checking on roteting
        MsgBox ("The Black Player is Victorious !")
       black_win_bool = True
'       GoTo again
    End If
    
   If cmd_rtt_CCW_UR.Enabled = True Then ' רק במקרה של ריטוט שיכול ל היות תיקו
               If white_win_bool = True And black_win_bool = True Then ' להעביר את זה בבדיקת נצחון לריפליי
                    MsgBox ("Both Players are Win !")
                    'GoTo play_again
               End If
     Else ' היה נצחון  רגיל אבל לא כזה  מהסוג של אחרי ריטוט שיכול להיות תיקו
     'GoTo play_again
     
'play_again:
'          Reply_again_game = MsgBox("? want to play again", vbMsgBoxRtlReading + vbMsgBoxRight + vbYesNo, "? another game")
'        If Reply_again_game = 6 Then
'
'          Call Form_Load ' לא בקריאה כי חוזר לפה אחרי הקריאה
'        '  GoTo end_form_loaded למה לא פותר ראה ההערה למטה
'         ' Exit Function ' הולך לדי_רוטט
'      ' Load Form1
'        End If
    
    End If 'תיקו
    
'play_again:
'          Reply_again_game = MsgBox("? want to play again", vbMsgBoxRtlReading + vbMsgBoxRight + vbYesNo, "? another game")
'        If Reply_again_game = 6 Then
'
'          Call Form_Load ' לא בקריאה כי חוזר לפה אחרי הקריאה
'        '  GoTo end_form_loaded למה לא פותר ראה ההערה למטה
'         ' Exit Function ' הולך לדי_רוטט
''      ' Load Form1
 '       End If
    
' !!!!!!!!!!!!!!!!!!!!!! win
Else ' של המצבים בדיקת הנצחון
End If ' של המצבים בדיקת הנצחון


If cmd_rtt_CCW_UR.Enabled = True And (mone_recorsion = True) Then       ' ז"א אפשר עדיין לרטט ולכן צריך עוד בדיקה לשחקן שני קורה רק אחרי ריטוט
'If (cmd_rtt_CCW_UR.Enabled = True And (mone_recorsion = True)) Or Reply_again_game = 6 Then        ' ז"א אפשר עדיין לרטט ולכן צריך עוד בדיקה לשחקן שני קורה רק אחרי ריטוט
     which_player = Not which_player   'הופך את המצב כדי לבדוק לשחקן השני
     mone_recorsion = False ' cuase to out from recorsion
     Call check_victory
Else
     If cmd_rtt_CCW_UR.Enabled = True Then ' קורה רק מתי שיצא מרקורסיה
          which_player = Not which_player
     End If
     mone_recorsion = True
End If

     '#############
       If (white_win_bool = True And black_win_bool = True) Or white_win_bool = True Or black_win_bool = True Then
                    Reply_again_game = MsgBox("? want to play again", vbMsgBoxRtlReading + vbMsgBoxRight + vbYesNo, "? another game")
                    If Reply_again_game = 6 Then
                              white_win_bool = False
                              black_win_bool = False
                              Call Form_Load ' לא בקריאה כי חוזר לפה אחרי הקריאה
                    Else
                    MsgBox "bye byte"
                    End
                    End If
          End If
    '#############

'end_form_loaded: '  Call dis_rtt ' לא עוזר יש  עדיין

End Function

 Function dis_rtt()

' that's NOT depend on which_Player is

If cmd_rtt_CCW_UR.Enabled = True Then
'disabling
cmd_rtt_CCW_UR.Enabled = False
cmd_rtt_CCW_UR.BackColor = &HFFC0C0
cmd_rtt_CCW_UL.Enabled = False
cmd_rtt_CCW_UL.BackColor = &HFFC0C0
cmd_rtt_CCW_DR.Enabled = False
cmd_rtt_CCW_DR.BackColor = &HFFC0C0
cmd_rtt_CCW_DL.Enabled = False
cmd_rtt_CCW_DL.BackColor = &HFFC0C0
cmd_rtt_CW_UR.Enabled = False
cmd_rtt_CW_UR.BackColor = &HFFC0C0
cmd_rtt_CW_UL.Enabled = False
cmd_rtt_CW_UL.BackColor = &HFFC0C0
cmd_rtt_CW_DR.Enabled = False
cmd_rtt_CW_DR.BackColor = &HFFC0C0
cmd_rtt_CW_DL.Enabled = False
cmd_rtt_CW_DL.BackColor = &HFFC0C0

' אחרי ריטוט

For i = 0 To 8
    If cmd_UR(i).BackColor <> QBColor(7) Then
        cmd_UR(i).Enabled = False
    Else
        cmd_UR(i).Enabled = True
    End If
Next

For i = 0 To 8
    If cmd_UL(i).BackColor <> QBColor(7) Then
        cmd_UL(i).Enabled = False
    Else
        cmd_UL(i).Enabled = True
    End If
Next

For i = 0 To 8
    If cmd_DR(i).BackColor <> QBColor(7) Then
        cmd_DR(i).Enabled = False
    Else
        cmd_DR(i).Enabled = True
    End If
Next

For i = 0 To 8
    If cmd_DL(i).BackColor <> QBColor(7) Then
        cmd_DL(i).Enabled = False
    Else
        cmd_DL(i).Enabled = True
    End If
Next


If mone_which_player Mod 2 = 0 Then
lbl_turn.Caption = ("White Turn, Put button")
Else
lbl_turn.Caption = ("Black Turn, Put button")
End If

Else 'enabaling rotating

cmd_rtt_CCW_UR.Enabled = True
cmd_rtt_CCW_UR.BackColor = &HFF0000
cmd_rtt_CCW_UL.Enabled = True
cmd_rtt_CCW_UL.BackColor = &HFF0000
cmd_rtt_CCW_DR.Enabled = True
cmd_rtt_CCW_DR.BackColor = &HFF0000
cmd_rtt_CCW_DL.Enabled = True
cmd_rtt_CCW_DL.BackColor = &HFF0000
cmd_rtt_CW_UR.Enabled = True
cmd_rtt_CW_UR.BackColor = &HFF0000
cmd_rtt_CW_UL.Enabled = True
cmd_rtt_CW_UL.BackColor = &HFF0000
cmd_rtt_CW_DR.Enabled = True
cmd_rtt_CW_DR.BackColor = &HFF0000
cmd_rtt_CW_DL.Enabled = True
cmd_rtt_CW_DL.BackColor = &HFF0000

For i = 0 To 8
cmd_UR(i).Enabled = False
cmd_UL(i).Enabled = False
cmd_DR(i).Enabled = False
cmd_DL(i).Enabled = False
Next

If mone_which_player Mod 2 = 0 Then
lbl_turn.Caption = ("White Turn, Rotate")
Else
lbl_turn.Caption = ("Black Turn, Rotate")
End If

End If


End Function
Function check_victory_COPY_RIGHT()
' two functions checks one for rotate one for puting
' and in the mone victory and massaging
'!!!!!' only one when just marking ig from rotating and changing the clr and remeber what was

If which_player = True Then ' white player
clrchk = &HFFFFFF 'צבע שחור
Else
clrchk = &H80000012 ' צבע לבן
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

Private Sub cmd_UR_Click_COPY_RIGHT(Index As Integer)
'mone_turns = mone_turns + 1

'if mone_turns =


' !!!!!!!! right
If which_player = True Then
cmd_UR(Index).BackColor = &HFFFFFF 'QBColor(15) '&HFFFFFF    'white
Call check_victory
which_player = False
cmd_UR(Index).Enabled = False
lbl_turn.Caption = ("Black Turn")
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
' !!!!!!!! right
'Call func_mone_turns

End Sub

Private Sub cmd_UR_Click_uirca(Index As Integer)
' !!!!!!!! right

If which_player = True Then ' white
     cmd_UR(Index).BackColor = &HFFFFFF 'QBColor(15) '&HFFFFFF    'white
     Call check_victory
' צריכה להיות הפרדה זה לא קורה  השינוי שחקן מיד כי נשאר לרטט
' אולי בהרטטה עצמה שלוחץ יידע לשנות שחקן
' ולעשות פרוצדורה שמשנה שחקן
     which_player = False
     cmd_UR(Index).Enabled = False
     lbl_turn.Caption = ("Black Turn")
Else 'black
    cmd_UR(Index).BackColor = &H80000012 'QBColor(0) ' &H80000012  ' black
    Call check_victory
    which_player = True
    cmd_UR(Index).Enabled = False
     lbl_turn.Caption = ("White Turn")
End If
' !!!!!!!! right

End Sub

Private Sub cmd_UR_Click(Index As Integer) 'UR =in array 1

If clrchk = &HFFFFFF Then ' white = 1
     arr_btns(1, Index) = 1
Else
     arr_btns(1, Index) = 0
End If

     cmd_UR(Index).BackColor = clrchk
     cmd_UR(Index).Enabled = False
     
    Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
    End If
    
End Sub

Private Sub cmd_UL_Click(Index As Integer) 'UL =in array 0

If clrchk = &HFFFFFF Then ' white = 1
     arr_btns(0, Index) = 1
Else
     arr_btns(0, Index) = 0
End If
     
     cmd_UL(Index).BackColor = clrchk
     cmd_UL(Index).Enabled = False
    Call check_victory
    
        If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
    End If
'Call dis_rtt
End Sub
Private Sub cmd_DL_Click(Index As Integer)
'DL =in array 3

If clrchk = &HFFFFFF Then ' white = 1
     arr_btns(3, Index) = 1
Else
     arr_btns(3, Index) = 0
End If

     cmd_DL(Index).BackColor = clrchk
     cmd_DL(Index).Enabled = False
     
    Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
    End If
    'Call dis_rtt
End Sub
Private Sub cmd_DR_Click(Index As Integer) 'DR =in array 2

If clrchk = &HFFFFFF Then ' white = 1
     arr_btns(2, Index) = 1
Else
     arr_btns(2, Index) = 0
End If

     cmd_DR(Index).BackColor = clrchk
     cmd_DR(Index).Enabled = False
     Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
    End If
     'Call dis_rtt
    
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

mone_which_player = mone_which_player + 1

Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
          Call switching_player
          
    End If
'Call dis_rtt
'Call switching_player

End Sub

Private Sub cmd_rtt_CCW_UR_Click()
Dim arr_rtt_CCW_UR(8) As Long 'Integer 'Double

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

mone_which_player = mone_which_player + 1

Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
          Call switching_player
          
    End If
'Call dis_rtt
'Call switching_player
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

mone_which_player = mone_which_player + 1

Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
          Call switching_player
          
    End If
'Call dis_rtt
'Call switching_player
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

mone_which_player = mone_which_player + 1

Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
          Call switching_player
          
    End If
'Call dis_rtt
'Call switching_player
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

mone_which_player = mone_which_player + 1

Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
          Call switching_player
          
    End If
'Call dis_rtt
'Call switching_player
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

mone_which_player = mone_which_player + 1

Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
          Call switching_player
          
    End If
'Call dis_rtt
'Call switching_player
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


mone_which_player = mone_which_player + 1
Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
          Call switching_player
          
    End If
'Call dis_rtt
'Call switching_player
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

mone_which_player = mone_which_player + 1
Call check_victory
    If Reply_again_game = 6 Then ' מונע את הבעיה שאחרי נצחון ואחרי העלאת הפורם שוב, הולך לדיס_רטט
          Reply_again_game = 0
    Else ' בכל מקרה אחר חוצ מזה שבא אחרי העלאת הפורם מחדש אחרי נצחון , נכנס לכאן
          Call dis_rtt
          Call switching_player
          
    End If
'Call dis_rtt
'Call switching_player
End Sub

Private Sub cmd_Save_Click()
' המשתמש בחר לשמור את תוכנה של תיבת הטקסט לקובץ
Dim Reply As Integer
Dim fileName As String

    CommonDlg.fileName = "" 'L אתחול שם הקובץ
   ' CommonDlg.Filter = "Text files (*.txt) | *.txt" 'L פילטר
    CommonDlg.Filter = "Pentago files (*.Pnt) | *.Pnt" 'L פילטר
    CommonDlg.Flags = cdlOFNHideReadOnly 'L הסתר את שמור לקריאה בלבד
    CommonDlg.ShowSave 'L הצג את תיבת הדיאלוג שמירה
    fileName = CommonDlg.fileName
    If fileName = "" Then Exit Sub 'L אם המשתמש בחר בביטול
    ' בדיקה האם הקובץ כבר קיים
    On Error Resume Next
    Name CommonDlg.fileName As CommonDlg.fileName
    If Err = 0 Then
        ' אם הקובץ קיים הצג הודעה האם להחליף אותו
        Reply = MsgBox("? Already exist a file with this name - to replace", vbMsgBoxRtlReading + vbMsgBoxRight + vbYesNo, "? Replacing file")
        If Reply = 7 Then Call cmd_Save_Click  'L המשתמש בחר לא להחליף את הקובץ
    End If
    ' קריאה לפרוצדורה לשמירת תוכנה של תיבת הטקסט לקובץ
    Call SaveTextFile(cmd_UL, cmd_UR, cmd_DL, cmd_DR, fileName, which_player) ', mone_which_player)
End Sub  ' ##########################################


Private Sub cmd_load_Click()

' המשתמש בחר לפתוח קובץ
Dim fileName As String
Dim last, next_line As Integer

    CommonDlg.fileName = "" 'L אתחול שם הקובץ
    CommonDlg.Filter = "Pentago files (*.Pnt) | *.Pnt" 'L פילטר
    CommonDlg.Flags = cdlOFNHideReadOnly 'L הסתר את פתח לקריאה בלבד
    CommonDlg.ShowOpen 'L הצג את תיבת הדיאלוג פתיחה
    fileName = CommonDlg.fileName
    If fileName = "" Then Exit Sub 'L אם המשתמש בחר בביטול
    
    Text1.Text = "" ' איפוס תיבת הטקסט

    ' קרא לפרוצדורה לפתיחת קובץ לתיבת הטקסט
    
Call OpenTextFile(fileName, Text1)
    
'מאפס את התאים
' באם העלו שמירה באמצע משחק
'לקרוא לפורם1 לואד
                                            
                      
next_line = 0
last = 1
i = 0
here1:
For i = i To InStr(last, Text1.Text, 1)
     If InStr(last, Text1.Text, 1) = i + 1 Then
          If i = 9 Then
                GoTo next_line2
          End If
          cmd_UL(i).BackColor = &HFFFFFF
          last = 1 + InStr(last, Text1.Text, 1) '!+i
          i = i + 1
          GoTo here1
     Else
            If i = 9 Then
                GoTo next_line2
            End If
            
            
        If InStr(last, Text1.Text, 0) = i + 1 Then
          If i = 9 Then
                GoTo next_line2
          End If
          cmd_UL(i).BackColor = 0
          last = 1 + InStr(last, Text1.Text, 0) '!+i
          i = i + 1
          GoTo here1
        Else
            If i = 9 Then
                GoTo next_line2
            End If
                    cmd_UL(i).BackColor = 12632256
          End If
    End If
        
        
Next 'UL


next_line2:
next_line = next_line + 9
last = 1 + 9 * 1
i = 0 + 9 * 1
here2:
For i = i To InStr(last, Text1.Text, 1)
     If InStr(last, Text1.Text, 1) = i + 1 Then
          If i = 18 Then
                GoTo next_line3
          End If
          cmd_UR(i - 9 * 1).BackColor = &HFFFFFF
          last = 1 + InStr(last, Text1.Text, 1) '!+i
          i = i + 1
          GoTo here2
     Else
            If i = 18 Then
                GoTo next_line3
            End If
            
            
        If InStr(last, Text1.Text, 0) = i + 1 Then
          If i = 18 Then
                GoTo next_line3
          End If
          cmd_UR(i - 9 * 1).BackColor = 0
          last = 1 + InStr(last, Text1.Text, 0) '!+i
          i = i + 1
          GoTo here2
        Else
            If i = 18 Then
                GoTo next_line3
            End If
                    cmd_UR(i - 9 * 1).BackColor = 12632256
          End If
    End If
        
        
Next ' UR


'!!!!!!!!!!!!!
'next_line2:
'next_line = next_line + 9
'last = 1 + 9 * 1
'i = 0 + 9 * 1
'here2:
'For i = i To InStr(last, Text1.Text, 1)
'        If InStr(last, Text1.Text, 1) = i + 1 Then ' white
'                If i = next_line + 9 Then
'                GoTo next_line3
'            End If
'            cmd_UR(i - 9 * 1).BackColor = 0
'            last = 1 + InStr(last, Text1.Text, 1) '!+i
'            i = i + 1
'            GoTo here2
'        Else
'            If i = 9 * 2 Then
'                GoTo next_line3
'            End If
'            cmd_UR(i - 9 * 1).BackColor = 12632256
'        End If
'   Next
   
   
next_line3:
next_line = next_line + 18
last = 1 + 9 * 2
i = 0 + 9 * 2
here3:
For i = i To InStr(last, Text1.Text, 1)
     If InStr(last, Text1.Text, 1) = i + 1 Then
          If i = 27 Then
                GoTo next_line4
          End If
          cmd_DL(i - 9 * 2).BackColor = &HFFFFFF
          last = 1 + InStr(last, Text1.Text, 1) '!+i
          i = i + 1
          GoTo here3
     Else
            If i = 27 Then
                GoTo next_line4
            End If
            
            
        If InStr(last, Text1.Text, 0) = i + 1 Then
          If i = 27 Then
                GoTo next_line4
          End If
          cmd_DL(i - 9 * 2).BackColor = 0
          last = 1 + InStr(last, Text1.Text, 0) '!+i
          i = i + 1
          GoTo here3
        Else
            If i = 27 Then
                GoTo next_line4
            End If
                    cmd_DL(i - 9 * 2).BackColor = 12632256
          End If
    End If
        
        
Next ' DL
     ' till  here  good!
     
     
next_line4:
next_line = next_line + 27
last = 1 + 9 * 3
i = 0 + 9 * 3
here4:
For i = i To InStr(last, Text1.Text, 1)
     If InStr(last, Text1.Text, 1) = i + 1 Then
          If i = 36 Then
                GoTo next_line5
          End If
          cmd_DR(i - 9 * 3).BackColor = &HFFFFFF
          last = 1 + InStr(last, Text1.Text, 1) '!+i
          i = i + 1
          GoTo here4
     Else
            If i = 36 Then
                GoTo next_line5
            End If
            
            
        If InStr(last, Text1.Text, 0) = i + 1 Then
          If i = 36 Then
                GoTo next_line5
          End If
          cmd_DR(i - 9 * 3).BackColor = 0
          last = 1 + InStr(last, Text1.Text, 0) '!+i
          i = i + 1
          GoTo here4
        Else
            If i = 36 Then
                GoTo next_line5
            End If
                    cmd_DR(i - 9 * 3).BackColor = 12632256
          End If
    End If
        
        
Next ' DR

next_line5:
End Sub
