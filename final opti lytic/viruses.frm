VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "viruses"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   ForeColor       =   &H00000000&
   Icon            =   "viruses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr_mitozis 
      Interval        =   1000
      Left            =   4395
      Top             =   1545
   End
   Begin VB.Timer tmr_vi 
      Interval        =   10
      Left            =   1815
      Top             =   1455
   End
   Begin VB.CommandButton cmd_stop 
      Caption         =   "Stop !"
      Height          =   495
      Left            =   1770
      TabIndex        =   1
      Top             =   75
      Width           =   1215
   End
   Begin VB.Image img_cell_sick 
      Height          =   480
      Left            =   3105
      Picture         =   "viruses.frx":08CA
      Stretch         =   -1  'True
      Top             =   2580
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lbl_vi 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Viruses :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   30
      TabIndex        =   5
      Top             =   585
      Width           =   915
   End
   Begin VB.Label lbl_viruses_mone 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1000
      TabIndex        =   4
      Top             =   585
      Width           =   135
   End
   Begin VB.Label lbl_cells_mone 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   700
      TabIndex        =   3
      Top             =   285
      Width           =   135
   End
   Begin VB.Label lbl_cells 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cells :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   30
      TabIndex        =   2
      Top             =   285
      Width           =   630
   End
   Begin VB.Label lbl_timing 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " timing"
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   60
      Width           =   450
   End
   Begin VB.Image img_cell 
      Height          =   480
      Index           =   0
      Left            =   3810
      Picture         =   "viruses.frx":1194
      Stretch         =   -1  'True
      Top             =   1545
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Dim timing As Integer
' האמת שזה  מחזור חיים ליטי של וירוס - ולא כמו שרציתי ליסוגני
'בו הוירוס משתכפל עם התא(חיידק)
' זמן אחוז ממסך צבוע


'vb Syntax (Toggle Plain Text)
Private Sub Form_Load()
Randomize
'Me.AutoRedraw = True 'need to be false
Me.ScaleMode = vbPixels
Me.Width = Screen.Width ' \ 15
 Me.Height = Screen.Height '\ 15
Me.Left = 0
Me.Top = 0
lbl_timing.Left = 0
lbl_timing.Top = 0



img_cell_wdt = img_cell(0).Width
img_cell_hgt = img_cell(0).Height
 
' img_cell(0).Left = -100
' img_cell(0).Top = 200
' MsgBox Form1.Point(last_x, last_y)
'
 img_cell(0).Visible = False ' saving problems at the end ,can't be unloaded
 img_cell(0).Tag = "dead"
  img_cell(0).Enabled = False
  
'centering
'last_x = Screen.Width \ 30
'last_y = Screen.Height \ 30
'SetPixel Me.hdc, last_x, last_y, 0

 'creating new randomly cells
max_cells = (Rnd * 20) + 1
'max_cells = 3
scr_wdt = Screen.Width \ 15
scr_hgt = Screen.Height \ 15
ReDim ary_cells(max_cells) As Integer
For i = 1 To (UBound(ary_cells))
          Load img_cell(i)
          img_cell(i).Visible = True
          img_cell(i).Left = Int(Rnd * ((scr_wdt - 200)) - img_cell_wdt) + 100
          img_cell(i).Top = Int(Rnd * ((scr_hgt - 200)) - img_cell_hgt) + 100
          img_cell(i).Tag = "healthy"
          ary_cells(i) = Int(Rnd * 10) + 1 ' timing sec for mitozis
          'img_cell(i).ToolTipText = i
          
Next

  'creating new randomly pixel viruses start position
max_vi = (Rnd * 150) + 50 '
'max_vi = 150
ReDim ary_vi_x_y(2, max_vi) As Integer

 For i = 1 To max_vi  ' 0 ->1 !!!!!!
ary_vi_x_y(1, i) = (Rnd * (Screen.Width \ 15)) ' x
ary_vi_x_y(2, i) = (Rnd * (Screen.Height \ 15)) 'y
Next

lbl_cells_mone.Caption = max_cells '+1
lbl_viruses_mone.Caption = max_vi

End Sub

Private Sub tmr_mitozis_Timer()

Randomize
'temp_max_cells = max_cells ' let the timer done it's work
'For i = 1 To temp_max_cells '
For i = 1 To (UBound(ary_cells))
          On Error GoTo err_pass ' does the for mone will automaticly run till the up range you told him even in the middle it change ?
          ary_cells(i) = ary_cells(i) - 1 ' one sec down
          If ary_cells(i) = 0 Then 'timer acount down gets to 0
                     If Not img_cell(i).Tag = "dead" Then ' time to split
                    '(ary_cells(i) = 0 Or ary_cells(i) = -1)
                              If img_cell(i).Tag = "sick" Then 'checking if sick or not
                                        'lytic
                                        
                                        'vi newbies
                                        new_vi = Int(Rnd * 10)
                                        max_vi = max_vi + new_vi
                                        ReDim Preserve ary_vi_x_y(2, max_vi) As Integer
                             
                                        For j = (max_vi - new_vi) To max_vi ' for vi newbies
                                        ary_vi_x_y(1, j) = img_cell(i).Left + Rnd * img_cell_wdt
                                        ary_vi_x_y(2, j) = img_cell(i).Top + Rnd * img_cell_hgt
                                        Form1.PSet (ary_vi_x_y(1, j), ary_vi_x_y(2, j)), 0
                                        Next
                                        cell_IDi = i
                                         Call killing(cell_IDi)
                              Else
                                        If img_cell(i).Tag = "healthy" Then 'won't be dead one ,because his img still exsit at memory - unless I fix the unloaded problem
                                        '(And img_cell(i).Visible = True )
                                        'healthy
                                        Call mitozis(i, img_cell(i).Left, img_cell(i).Top)
                                        'ary_cells(i) = Int(Rnd * 10) + 1 'timing sec for mitozis
                                        'ary_cells(UBound(ary_cells)) = Int(Rnd * 10) + 1 'timing sec for mitozis
                                        'lbl_cells_mone.Caption = max_cells
                                        End If
                              End If ' 'checking if sick or not
                    End If
          End If
'dead_pass2:
Next
err_pass:
lbl_cells_mone.Caption = UBound(ary_cells)
lbl_viruses_mone.Caption = max_vi

End Sub

Private Sub tmr_vi_Timer()

For i = 0 To max_vi


rnd_num = Int((Rnd * 4) + 1)

Select Case rnd_num

Case 1 'up
'last_clr = Form1.Point(ary_vi_x_y(1, i), ary_vi_x_y(2, i) - 1)
 Call does_pix_clr_0(ary_vi_x_y(1, i), ary_vi_x_y(2, i) - 1)
Form1.PSet (ary_vi_x_y(1, i), ary_vi_x_y(2, i) - 1), 0
Form1.PSet (ary_vi_x_y(1, i), ary_vi_x_y(2, i)), vbWhite 'last_clr '&HFFFFFF
ary_vi_x_y(2, i) = ary_vi_x_y(2, i) - 1

Case 2 'down

 Call does_pix_clr_0(ary_vi_x_y(1, i), ary_vi_x_y(2, i) + 1)
Form1.PSet (ary_vi_x_y(1, i), ary_vi_x_y(2, i) + 1), 0
Form1.PSet (ary_vi_x_y(1, i), ary_vi_x_y(2, i)), &HFFFFFF
ary_vi_x_y(2, i) = ary_vi_x_y(2, i) + 1

Case 3 'rigjt

 Call does_pix_clr_0(ary_vi_x_y(1, i) + 1, ary_vi_x_y(2, i))
Form1.PSet (ary_vi_x_y(1, i) + 1, ary_vi_x_y(2, i)), 0
Form1.PSet (ary_vi_x_y(1, i), ary_vi_x_y(2, i)), &HFFFFFF
ary_vi_x_y(1, i) = ary_vi_x_y(1, i) + 1

Case 4 'left

 Call does_pix_clr_0(ary_vi_x_y(1, i) - 1, ary_vi_x_y(2, i))
Form1.PSet (ary_vi_x_y(1, i) - 1, ary_vi_x_y(2, i)), 0
Form1.PSet (ary_vi_x_y(1, i), ary_vi_x_y(2, i)), &HFFFFFF
ary_vi_x_y(1, i) = ary_vi_x_y(1, i) - 1

End Select


Next 'For i = 1 To max_vi

lbl_timing.Caption = timing
timing = timing + 1



End Sub

Private Sub cmd_stop_Click()
If tmr_vi.Enabled = False Then
cmd_stop.Caption = "Stop"
tmr_vi.Enabled = True
tmr_mitozis.Enabled = True
Else
tmr_vi.Enabled = False
tmr_mitozis.Enabled = False
cmd_stop.Caption = "Run !"
End If

End Sub
