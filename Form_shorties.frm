VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   2505
   Icon            =   "Form_shorties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   2505
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_top 
      Caption         =   "TOP"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txt_add2list 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Tag             =   "write the word to be added"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmd_add_2_list 
      Caption         =   "add 2 list"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4545
      ItemData        =   "Form_shorties.frx":048A
      Left            =   240
      List            =   "Form_shorties.frx":048C
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "Double Click to copy"
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "word copied :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1770
   End
   Begin VB.Menu mnu_about 
      Caption         =   "About"
   End
   Begin VB.Menu mnu_help 
      Caption         =   "Help"
   End
   Begin VB.Menu mnu_my_list 
      Caption         =   "my_list.txt"
      Begin VB.Menu mnu_location_list 
         Caption         =   "location"
         Index           =   1
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu_open_list 
         Caption         =   "open"
         Index           =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_open_list 
         Caption         =   "refresh"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---iprovments---
' *opacity ?
' *the load from the file is too way long time ' in pentago there is solution
' *AppActivate "WINWORD"      ' - why does'nt work - I want by name of app not file - a good and quick solution is to write the file name you work with in my_list.txt fie, and the AppActivate will be based on that name
' * nubering the list, while loading, on the list  / label (array)/ other list. in a way that there is auto focus for the list, and you just need to (after focusing to the form1) press the number, instead of double click.
'* bold only the first letter for A B C 'https://forums.codeguru.com/showthread.php?373211-How-to-set-a-line-bold-in-a-listbox
'* vb6 detecting (current OR active) language english - for hebrew
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
'Public Const FLAGS = SWP_NOMOVE ' Or SWP_NOSIZE ' https://stackoverflow.com/questions/17651725/vb6-how-to-make-a-floating-window-top-most

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Dim FileNum As String
Dim last_added As String
Dim file_name As String
Dim Topmost As Boolean
Dim Ext As String
Dim path, SourceFile As String
Dim i As Integer
Private Declare Function GetThreadLocale Lib "kernel32" () As Long

Private Sub Form_Load()
 
 Ext = "txt"
last_added = ""
Dim i As Integer


'CommonDialog1.Filter = Ext '"Files (*." + Ext + ")|*."
CommonDialog1.Filter = "my_list files (*.txt)|*.txt|All Files (*.*)|*.*"
CommonDialog1.DefaultExt = Ext
CommonDialog1.DialogTitle = "Choose where is your my_list_yourdoc.txt file - should be within the yourdoc-file's folder"


  CommonDialog1.CancelError = True ' when click on cancel
  'https://stackoverflow.com/questions/31696676/detecting-cancel-button-in-commondialog-control
   On Error Resume Next
   CommonDialog1.ShowOpen
    If Err.Number = &H7FF3 Then
        ' Cancel clicked
        End
    End If
  
 path = CommonDialog1.fileName ' the app can be run from every where to any my_ist file.

 
Topmost = True
cmd_top.Caption = "NOT ON TOP"

'Private Sub Command1_Click()
' always On Top
Call SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'End Sub
'load from text file "C:\my_list.txt" into list1

Dim TextLines() As String

Dim fn As Integer


Dim lngA As Integer

'path = App.path & "\Text Files\"
'path = "C:\"

'path = App.path & "\" ' this way the app need to be at the same dir of the my_list file.
       ' SourceFile = path & "my_list.txt"
        SourceFile = path
        fn = FreeFile
       '  On Error
         

        
Open SourceFile For Binary Access Read As #fn ' slow way ?
    TextLines = Split(Input(LOF(fn), fn), vbNewLine)
    'For lngA = 1 To UBound(TextLines) '*** no need because of GetFilenameWithoutExtension, and the txt file have the same name of the doc you work with
    'the upper was when there is the name of the doc, isdie the list, the first row
    For lngA = 0 To UBound(TextLines) - 1 'no last blank row on the list
       Form1.List1.AddItem TextLines(lngA) ' old fashion - no upperCase issues
        'Form1.List1.AddItem lngA & " " & TextLines(lngA) ' what about two digits row ? ' - ....just need to (after focusing to the form1) press the number, instead of double click.
    Next lngA
Close #fn
'file_name = TextLines(0) '*** no need because of GetFilenameWithoutExtension ' can be taken by path, if you open in the dialuge the doc file - https://stackoverflow.com/questions/28632331/vb6-get-filename-from-pat

 GetFilenameWithoutExtension (CommonDialog1.FileTitle) ' than no need to write the name of the file.doc inside the my_list.txt, but instead, the my_list gets the name of the doc you want to use.
 
Call uprcase
' detect active window languge
'https://stackoverflow.com/questions/4677752/user-defined-type-not-defined-error-in-vb-6-under-windows-7
'  Dim myCurrentLanguage As InputLanguage
'  myCurrentLanguage = InputLanguage.CurrentInputLanguage
'    Me.Text = "Current input language is: " & _
'        myCurrentLanguage.Culture.EnglishName


'Beep
'MsgBox List1.List(1)
End Sub

Function uprcase() ' after the list is sorted ABC (by default), then switching to upper cases only the first word of its first letter
Dim old_ltr, new_ltr As String

  old_ltr = ""
    For i = 0 To List1.ListCount - 1 'no last blank row on the list
        new_ltr = (Left(List1.List(i), Len(List1.List(i)) - (Len(List1.List(i)) - 1))) ' - first char
       If old_ltr = new_ltr Then
             List1.List(i) = List1.List(i)  'just add to new line
             'old_ltr = (Left(TextLines(lngA), Len(TextLines(lngA)) - (Len(TextLines(lngA)) - 1)))
       Else
            List1.List(i) = StrConv((List1.List(i)), vbProperCase) 'new fasion - upper case for easier navigation
       End If
       old_ltr = new_ltr
       Next
End Function

'PrivateSub Timer1_Tick(sender As Object, e As EventArgs) HandlesTimer1.Tick

'End Sub

Public Function GetFilenameWithoutExtension(fileName As String)
'function to get file name without an extension.

' this commented part below is for when the name of the txt file is as the name of the doc file, but in that way you can't have the same two files with samne name to work properly
'If InStr(1, StrReverse(FileName), ".") = 0 Then
'    GetFilenameWithoutExtension = FileName 'no extension
'Else
'    GetFilenameWithoutExtension = Left(FileName, Len(FileName) - InStr(1, StrReverse(FileName), ".", vbTextCompare))
'    'same as -  GetFilenameWithoutExtension = Left(FileName, Len(FileName) - 4)
'End If
'file_name = GetFilenameWithoutExtension

'use this form my_list_docname
GetFilenameWithoutExtension = fileName
GetFilenameWithoutExtension = Right(fileName, Len(fileName) - 8) '@ that way you don't confuse the app with two files with the same name = my_list_Copy
GetFilenameWithoutExtension = Left(GetFilenameWithoutExtension, Len(GetFilenameWithoutExtension) - 4) '@ that way you don't confuse the app with two files with the same name = Copy
file_name = GetFilenameWithoutExtension

End Function
Private Sub cmd_top_Click()

     'If Topmost = True Then 'Make the window topmost
     If cmd_top.Caption = "NOT ON TOP" Then
      Call SetWindowPos(Form1.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
         '  Topmost = False
         cmd_top.Caption = "ON TOP"
     Else
        Call SetWindowPos(Form1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
       ' Topmost = True
       cmd_top.Caption = "NOT ON TOP"
     End If
End Sub

Private Sub cmd_add_2_list_Click()
'adding new word to the file my_list

'My.Computer.FileSystem.WriteAllText("C:\my_list.txt", txt_add2list.Text, True) = 1
'Dim FileNum As String

'If Not (txt_add2list.Text = "word 2 e added") Then ' when txt_add2list lost focus, then immidatly txt_add2list.Text = "word 2 e added" comes back again
    If Not (txt_add2list.Text = "") Then
        If Not (last_added = txt_add2list.Text) Then
        'Start append text to file
            FileNum = FreeFile
            'Open "C:\my_list.txt" For Append As FileNum
            Open SourceFile For Append As FileNum
            Print #FileNum, txt_add2list.Text
            Close FileNum
            List1.AddItem txt_add2list.Text
            last_added = txt_add2list.Text
        End If
    End If
'End If
End Sub



Private Sub List1_Click()
'Label1.Caption = List1.ListIndex
'Label1.Caption = List1.List(List1.ListIndex)
End Sub

Private Sub List1_DblClick()
Label1.Caption = StrConv((List1.List(List1.ListIndex)), vbLowerCase)
 
Dim i As Integer
Dim strClip As String


    strClip = ""

    With List1
        For i = 0 To .ListCount - 1
           ' If .Selected(i) Then strClip = strClip & .List(i) & vbNewLine
            If .Selected(i) Then strClip = strClip & StrConv((.List(i)), vbLowerCase) & vbNewLine
        Next i
    End With
    
    If strClip <> "" Then
        Clipboard.Clear
        Clipboard.SetText strClip
    End If
        'https://www.oreilly.com/library/view/vb-vba/1565923588/1565923588_ch07-37-fm2xml.html
        'https://stackoverflow.com/questions/25362563/open-word-document-and-bring-to-front
        'AppActivate ("Notepad")             'works
        'AppActivate ("Microsoft Word (32 bit)")
      '  AppActivate "Microsoft Word (Product Activation Failed)" ' works somtimes
    '    AppActivate ("Doc1")             ' WORKS !
    
    'AppActivate "Microsoft Word (Product Activation Failed)", True
    'AppActivate "Microsoft Word" ' - why does'nt work - I want by name of app, not the file
    'AppActivate "WINWORD"      ' - why does'nt work - I want by name of app not file
    
       On Error GoTo ErrorHandler
        ' Insert code that might generate an error here
       ' this way you can use the app without the need for specific file
       'but then, the auto ctrl+V doesn't work. and you need to do it manually
  
      'AppActivate "proposal", True


  AppActivate file_name, True
  ' problem only with notepad doc.txt files, when is change, it get * to the title of the form.
  ' another problem (solved) - what if you have two files with the same name that open at the same time. in that case we can still
  ' problem above solved - my_list_doc-name.txt
  
ErrorHandler:
   ' Insert code to handle the error here
    Dim WshShell As Object
Set WshShell = CreateObject("WScript.Shell")
'https://www.developerfusion.com/article/57/sendkeys-command/
WshShell.SendKeys ("^v") ' keyboard must be on ENG
'https://bytes.com/topic/visual-basic/answers/888139-any-idea-how-change-system-keyboard-language-vba
 'https://www.codeguru.com/visual-basic/switching-input-languages-dynamically-with-visual-basic/
 'https://microsoft.public.vb.database.dao.narkive.com/sET5Rgr0/how-can-i-switch-keyboard-language-from-vb-code
WshShell.SendKeys (" ")
  ' Resume Next frequency


    

End Sub



Private Sub mnu_about_Click()
MsgBox ("Too many long and complicated words in my thesis..." & vbCrLf & "eliran_t@walla.com")
End Sub

Private Sub mnu_help_Click()
MsgBox "Your doc should be open, and not minimized, in order to make auto paste, where the caret is." & vbCrLf & "Otherwise, you have to paste it yourself. i.e. you can use it where ever you like to." & vbCrLf & "No UpperCases are copied, just to help navigate the list."
End Sub

Private Sub mnu_location_list_Click(Index As Integer)
'path

    Dim floder_Path As String
   
    '
    'Function pathOfFile(fileName As String) As String
        Dim posn As Integer
        posn = InStrRev(path, "\")
        If posn > 0 Then
            floder_Path = Left$(path, posn)
        Else
            floder_Path = ""
        End If
    'End Function

'Shell "explorer " & myPath
ShellExecute 0, vbNullString, floder_Path, vbNullString, vbNullString, 1
End Sub

'Private Sub txt_add2list_GotFocus()
'txt_add2list.Text = ""
'End Sub
'
'Private Sub txt_add2list_LostFocus()
'txt_add2list.Text = "word 2 b added"
'End Sub
Private Sub mnu_my_list_Click()
'SHIPUR refresh - need to do function of loading the list from the file
'im intFile As Integer
'intFile = FreeFile
'open App.path For Input As intFile
'strData = Input(LOF(intFile), intFile)
'Close intFile


End Sub

Private Sub mnu_open_list_Click(Index As Integer)
Dim SW_SHOWNORMAL As Integer
SW_SHOWNORMAL = 1
ShellExecute hWnd, "open", path, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub
