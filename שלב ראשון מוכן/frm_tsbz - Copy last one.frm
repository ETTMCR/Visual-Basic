VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_tsbz 
   Caption         =   "תשבצי"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1
      AutoSize        =   -1  'True
      Height          =   2508
      Index           =   0
      Left            =   720
      Picture         =   "frm_tsbz.frx":0000
      ScaleHeight     =   2460
      ScaleWidth      =   3072
      TabIndex        =   16
      Top             =   480
      Width           =   3120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_try_word 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   15
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   3600
      TabIndex        =   14
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   3120
      TabIndex        =   13
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   2640
      TabIndex        =   12
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   2160
      TabIndex        =   11
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1680
      TabIndex        =   10
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1200
      TabIndex        =   9
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmd_frsh 
      Caption         =   "Refrash"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   720
      TabIndex        =   7
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3600
      TabIndex        =   6
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   3120
      TabIndex        =   5
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmd_tfz 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   4800
      Width           =   495
   End
End
Attribute VB_Name = "frm_tsbz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As Long, ByVal nMaxFile As Long, ByVal lpstrInitialDir As Long, ByVal lpstrDefExt As Long, ByVal lpstrFilter As Long, ByVal lpstrTitle As Long) As Long
Private Declare Function GetFileNameFromBrowseA Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As String, ByVal nMaxFile As Long, ByVal lpstrInitialDir As String, ByVal lpstrDefExt As String, ByVal lpstrFilter As String, ByVal lpstrTitle As String) As Long

Private Sub Form_Load()
hm_words = 0
'simple sol
'CommonDialog1.Filter = "Apps (*.txt)|*.txt|All files (*.*)|*.*"
'CommonDialog1.DefaultExt = "txt"
'CommonDialog1.DialogTitle = "Select File"
'CommonDialog1.ShowOpen
'
''The FileName property gives you the variable you need to use
'MsgBox CommonDialog1.FileName
'simple sol

loading : 
    Dim sSave As String
    sSave = Space(255)
    'If we're on WinNT, call the unicode version of the function
    If IsWinNT Then
        'GetFileNameFromBrowseW Me.hWnd, StrPtr(sSave), 255, StrPtr("c:\"), StrPtr("jpg"), StrPtr("picture files " + Chr$(0) + "*.jpg" + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)), StrPtr("Choose a picture ")
        GetFileNameFromBrowseW Me.hWnd, StrPtr(sSave), 255, StrPtr("c:\"), StrPtr("jpg"), StrPtr("picture files " + Chr$(0) +  Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)), StrPtr("Choose a picture ")
    'If we're not on WinNT, call the ANSI version of the function
    Else
        GetFileNameFromBrowseA Me.hWnd, sSave, 255, "c:\", "jpg", "picture files " + Chr$(0) + "*.jpg" + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0), "The Title"
    End If
            
   
    Dim my_fileName As String
     my_fileName = sSave

      
     '@  my_fileName = Mid(my_fileName, InStrRev(my_fileName, "\") + 1, Len(my_fileName))
      '@  MsgBox (my_fileName)
       '@ MsgBox (Len(my_fileName))
        '@ my_fileName = Mid(my_fileName, InStrRev(my_fileName, ".") + 1, 3)  ' gives the perfix file type
        
        '@@@@@ checks file type correct ' can be delete and works
            chk_file_type = Mid(my_fileName, InStrRev(my_fileName, ".") + 1, 3)  ' gives the perfix file type
            Do While Not ((chk_file_type = "jpg") Or (chk_file_type = "JPG") Or (chk_file_type = "bmp") Or (chk_file_type = "BMP"))

                MsgBox ("Wrong file type !")
                sSave = Space(255)
                GetFileNameFromBrowseW Me.hWnd, StrPtr(sSave), 255, StrPtr("c:\"), StrPtr("jpg"), StrPtr("picture files " + Chr$(0)  + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)), StrPtr("Choose a picture ")
                chk_file_type = Mid(sSave, InStrRev(sSave, ".") + 1, 3)
                
            Loop
            
           ' still problem if you first choose null
         '@@@@@ checks file type correct
        
      ' MsgBox (Len(my_fileName))
        
        my_fileName_for_loading_pic = sSave
        my_fileName = sSave
         
         my_fileName = Mid(my_fileName, InStrRev(my_fileName, "\") + 1, Len(my_fileName))
         org_my_fileName = my_fileName
         hwmc1 = Len(my_fileName)
         
        my_fileName = Mid(my_fileName, InStrRev(my_fileName, ".") + 1, Len(my_fileName))
        hwmc2 = (Len(my_fileName))
        hwmc3 = hwmc1 - hwmc2
        
        
       ' str_the_wwwwword = Mid(org_my_fileName, 1, hwmc1 + 4)
        
     '*   MsgBox (hwmc3 - 1) 'how many letters in the word
     '*   MsgBox (org_my_fileName)
        str_the_word = Mid(org_my_fileName, InStrRev(org_my_fileName, "\") + 1, hwmc3 - 1)
      '*  MsgBox (str_the_word) ' the word !!!!!!!!!!!!!!!!!
      '*  MsgBox Len(str_the_word) 'how many letters in the word for sure

   
        '@@@@@ checks file number of letters
        Do While Len(str_the_word) > 14
        MsgBox ("change file name to less than 15 letters")
        
        sSave = Space(255)
        GetFileNameFromBrowseW Me.hWnd, StrPtr(sSave), 255, StrPtr("c:\"), StrPtr("jpg"), StrPtr("picture files " + Chr$(0) + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)), StrPtr("Choose a picture ")
        'chk_file_type = Mid(sSave, InStrRev(sSave, ".") + 1, 3)
        my_fileName_for_loading_pic = sSave
        my_fileName = sSave
         
         my_fileName = Mid(my_fileName, InStrRev(my_fileName, "\") + 1, Len(my_fileName))
         org_my_fileName = my_fileName
         hwmc1 = Len(my_fileName)
         
        my_fileName = Mid(my_fileName, InStrRev(my_fileName, ".") + 1, Len(my_fileName))
        hwmc2 = (Len(my_fileName))
        hwmc3 = hwmc1 - hwmc2
 
        str_the_word = Mid(org_my_fileName, InStrRev(org_my_fileName, "\") + 1, hwmc3 - 1)
        
        Loop
        '@@@@@ checks file number of letters

'@@@@@@@ inserting str_the_word into array
    ReDim Preserve arr_words(hm_words)
    arr_words(hm_words) = str_the_word
    hm_words = hm_words + 1
	Msg = "Do you want to choose more words ?"   ' Define message.
	Style = vbYesNo + vbCritical + vbDefaultButton2 ' Define buttons.
	'Title = "MsgBox Demonstration"  ' Define title.
	'Ctxt = 1000 ' Define topic
        ' context.
        ' Display message.
	Response = MsgBox(Msg, Style, Ctxt)
	If Response = vbYes Then    ' User chose Yes.
	     GoTo loading
	'Else    ' User chose No.
	'    MyString = "No" ' Perform some action.
	End If
    
'@@@@@@@ inserting str_the_word into array

'@@@@@ loading image into imagebox
              
	my_fileName_for_loading_pic = Mid(sSave, 1, 255- (hwmc2 - 3))
	Picture1(0).Picture = LoadPicture(my_fileName_for_loading_pic)

'@@@@@ loading image into imagebox



how_much_cmd_try_word = 1

'@@@@@ making new cmd btn as the number os letters of str_the _word

'Redim "solving constent expression requierd" if was just Dim
ReDim arrControls(1 To Len(str_the_word)) As CommandButton  ' as Object'To avoid late binding you could also use "As <Control-Class>"
Dim num_of_btn As Integer
'ReDim arrControls(1 To num_columns) As Object 'To avoid late binding you could also use "As <Control-Class>"

'Set arrControls(1, 1) = Button(0)
'Set arrControls(1, 2) = Button(1)
'Button.Item(1).Top = Button.Item(0).Top + 500
'arrControls(1, 2).Visible = False
num_of_btn = 0 ' the first object index is 0
For i = 1 To Len(str_the_word)
    
        Set arrControls(i) = cmd_try_word(num_of_btn)
         num_of_btn = num_of_btn + 1
         '  max_cells = (UBound(ary_cells)) + 1 ' changing # of cells in array
         ' ReDim Preserve ary_cells(max_cells) As Integer 'must  redim Preserve , else zeroes all the ary cells
         ' Load Form1.img_cell(max_cells)
        'ReDim Preserve Button(num_of_btn) As Integer 'must  redim Preserve , else zeroes all the ary cells
        Load cmd_try_word(i)
        arrControls(i).Visible = True
        'arrControls(i).Top = j * 450 + 1470
        arrControls(i).Left = (i - 1) * 495 + cmd_tfz(0).Left  '+ ((Picture1(0).Left + Picture1(0).Width) / 2)' need to be relative
        'arrControls(i, j).Tag = i & vbCrLf & " " & j
        'arrControls(i, j).ToolTipText = arrControls(i, j).Tag
        'arrControls(i, j).Caption = num_of_btn - 1
   
Next

'@@@@@ making new cmd btn as the number os letters of str_the _word

'For i = 0 To 6
'        cmd_tfz.Item(i).Left = cmd_tfz.Item(i).Left + ((Picture1(0).Left + Picture1(0).Width) / 2)
'Next
'For i = 7 To 13
'       cmd_tfz.Item(i).Left = cmd_tfz.Item(i).Left + ((Picture1(0).Left + Picture1(0).Width) / 2)
'Next

'@@@@@randoming the word leters into cmd_tfz.

    For i = 1 To Len(str_the_word)
        'MyValue = Int(((Len(str_the_word)) * Rnd) + 1)   ' Generate random value between 0 and Len(str_the_word)- 1.
        Randomize
        MyValue = Int(cmd_tfz.Count - 1) * Rnd
        Do While True
            If cmd_tfz(MyValue).Caption = "" Then
                cmd_tfz(MyValue).Caption = Mid(str_the_word, i, 1)
                Exit Do
            Else
                    MyValue = Int(cmd_tfz.Count - 1) * Rnd
                    Randomize
            End If
        Loop

    Next
    
    For i = 0 To (cmd_tfz.Count - 1) ' generate random leteres for next blank tfz commands
        If cmd_tfz(i).Caption = "" Then
            'xyz = Int(25 * Rnd) + 97 ' 122 - 97 = 25
	xyz = Int(25 * Rnd) + 224' 250 - 224 = 26
            cmd_tfz(i).Caption = Chr(xyz)
        End If
        Randomize
        cmd_tfz(MyValue).Top = Picture1(0).Top + Picture1(0).Height + 750
    Next
    
'@@@@@randoming the word leters into cmd_tfz

'@@@@
    For i = 0 To (cmd_try_word.Count - 1)
       cmd_try_word(i).Top = Picture1(0).Top + Picture1(0).Height + 500
    Next

     For i = 0 To 6 '(cmd_try_word.Count / 2)
        cmd_tfz(i).Top = Picture1(0).Top + Picture1(0).Height + 1250
    Next
    For i = 7 To 13 '(cmd_try_word.Count / 2) To cmd_try_word.Count - 1 ' why error ?
        cmd_tfz(i).Top = cmd_tfz(3).Top + cmd_tfz(3).Height
    Next

     cmd_frsh.Top = cmd_tfz(13).Top + cmd_tfz(13).Height + 250	
'@@@@


'@@@@

End Sub
Private Sub cmd_frsh_Click()
    how_much_cmd_try_word = 1
    For i = 1 To cmd_tfz.Count
        cmd_tfz.Item(i - 1).Enabled = True
        
    Next
    For i = 1 To cmd_try_word.Count
        cmd_try_word.Item(i - 1).Caption = ""
        
    Next

End Sub

Private Sub cmd_tfz_Click(Index As Integer)
      cmd_try_word(Len(str_the_word) - how_much_cmd_try_word).Caption = cmd_tfz(Index).Caption
     If how_much_cmd_try_word < cmd_try_word.Count - 1  Then ' !!!!!!!!!!!!!!!!! an difference beetwin HEB to ENG
        how_much_cmd_try_word = how_much_cmd_try_word + 1
     End If
     
    cmd_tfz(Index).Enabled = False
 
     
    '@@@@ check win
    For i = cmd_try_word.Count - 1 To 0 Step -1
        str_does_win = str_does_win + cmd_try_word.Item(i).Caption
    Next
    If str_the_word = str_does_win Then
        MsgBox "you win !"
        Call cmd_frsh_Click	
        'Beep
    End If
     '@@@@ check
     
End Sub

'Private Sub cmd_tfz_Click(Index As Integer) 'ENGLISH VERSION
'     cmd_try_word(how_much_cmd_try_word - 1).Caption = cmd_tfz(Index).Caption
'     If how_much_cmd_try_word < cmd_try_word.Count Then
'        how_much_cmd_try_word = how_much_cmd_try_word + 1
'     End If
'     cmd_tfz(Index).Enabled = False
'     
'     '@@@@ check win
'    For i = 0 To cmd_try_word.Count - 1
'        str_does_win = str_does_win + cmd_try_word.Item(i).Caption
'    Next
'    If str_the_word = str_does_win Then
'        MsgBox "you win !"
'        Call cmd_frsh_Click
'        'Beep
'    End If
'    '@@@@ check win
'End Sub




Public Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function
