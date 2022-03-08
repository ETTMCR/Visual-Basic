VERSION 5.00
Begin VB.Form frm_tictic 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TicTic"
   ClientHeight    =   7365
   ClientLeft      =   6675
   ClientTop       =   3855
   ClientWidth     =   7665
   Icon            =   "frm_tictic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Button 
      BackColor       =   &H00C0C0C0&
      Height          =   450
      Index           =   0
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   450
   End
   Begin VB.CommandButton cmd_new_board 
      BackColor       =   &H0000FF00&
      Caption         =   "New Board"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6400
      Width           =   975
   End
   Begin VB.CommandButton cmd_dead 
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      Height          =   450
      Left            =   720
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   450
   End
   Begin VB.CommandButton cms_rst 
      BackColor       =   &H0000FFFF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frm_tictic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd_bck_clr As Double
Dim how_much As Integer
Public Sub func_mark(cmd_Indx, cmd_clr)

For i = (cmd_Indx \ num_rows) + 1 To num_columns
    For j = 1 To (cmd_Indx Mod num_rows) + 1
        If arrControls(i, j).BackColor = &HC0C0C0 Then
            arrControls(i, j).BackColor = cmd_clr
            arrControls(i, j).Enabled = False
            how_much = how_much - 1
            
        End If
    Next
Next

'Label1.Caption = (cmd_Indx \ num_rows) & vbCrLf & (cmd_Indx Mod num_rows)

End Sub

Private Sub cmd_new_board_Click()
Unload Me
frm_size.Visible = True
End Sub

Private Sub cms_rst_Click()

'num_of_btn = num_columns * num_rows '0 ' the first object index is 0
For i = 1 To num_columns
    For j = 1 To num_rows
        arrControls(i, j).BackColor = 12632256
        arrControls(i, j).Enabled = True
    Next
Next
arrControls(1, num_rows).Enabled = False
cmd_bck_clr = &HFFFFFF

how_much = num_columns * num_rows - 1

End Sub


Private Sub Form_Load()

'Label1.Caption = num_columns & " columns" & vbCrLf & num_rows & " rows"
cmd_bck_clr = &HFFFFFF
how_much = num_columns * num_rows - 1

'Redim "solving constent expression requierd" if was just Dim
ReDim arrControls(1 To num_columns, 1 To num_rows) As CommandButton  ' as Object'To avoid late binding you could also use "As <Control-Class>"
Dim num_of_btn As Integer

'Set arrControls(1, 1) = Button(0)

'Set arrControls(1, 2) = Button(1)
'Button.Item(1).Top = Button.Item(0).Top + 500
'arrControls(1, 2).Visible = False
num_of_btn = 0 ' the first object index is 0
For i = 1 To num_columns
    For j = 1 To num_rows
        Set arrControls(i, j) = Button(num_of_btn)
        num_of_btn = num_of_btn + 1
         '  max_cells = (UBound(ary_cells)) + 1 ' changing # of cells in array
         ' ReDim Preserve ary_cells(max_cells) As Integer 'must  redim Preserve , else zeroes all the ary cells
         ' Load Form1.img_cell(max_cells)
        'ReDim Preserve Button(num_of_btn) As Integer 'must  redim Preserve , else zeroes all the ary cells
        Load Button(num_of_btn)
        arrControls(i, j).Visible = True
        arrControls(i, j).Top = j * 450 + 1470
        arrControls(i, j).Left = i * 450 + 1470
        'arrControls(i, j).Tag = i & vbCrLf & " " & j
        'arrControls(i, j).ToolTipText = arrControls(i, j).Tag
        'arrControls(i, j).Caption = num_of_btn - 1
    Next
Next

cmd_dead.Top = arrControls(1, num_rows).Top
cmd_dead.Left = arrControls(1, 1).Left
arrControls(1, num_rows).Enabled = False

End Sub

Private Sub Button_Click(Index As Integer)
'Label1.Caption = Index & vbCrLf & ""

If cmd_bck_clr = &HFFFFFF Then
    Button.Item(Index).BackColor = cmd_bck_clr
    Button.Item(Index).Enabled = False
    cmd_bck_clr = &H0&
Else
    Button.Item(Index).BackColor = cmd_bck_clr
    Button.Item(Index).Enabled = False
    cmd_bck_clr = &HFFFFFF
End If

cmd_Indx = Index
cmd_clr = Button.Item(Index).BackColor
how_much = how_much - 1
Call func_mark(cmd_Indx, cmd_clr)

'****winning ?****
If how_much = 0 Then
    If cmd_clr = &HFFFFFF Then
        MsgBox "The black player lose !"
    Else
        MsgBox "The white player lose !"
    End If
End If
'MsgBox how_much
End Sub

